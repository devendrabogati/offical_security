# Copyright 2026 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0

"""
Presentation orchestrator for the Capgemini Talk2Docs brand-aligned
presentation agent.

CHANGELOG (v2):
  - Removed text truncation. Bullets and titles are written in full and
    PowerPoint auto-shrinks the font to fit. This fixes the "..." ellipsis
    appearing on every bullet.
  - Stopped removing empty placeholders. The master's prompt text on the
    Capgemini template is non-visible in the rendered slide, so removing
    the placeholder was unnecessary and risky.
  - Never modify DATE / FOOTER / SLIDE_NUMBER placeholders — they self-
    populate from the master, and removing them strips footer branding
    ("Presentation Title | Author | Date" and "Company Confidential
    © Capgemini 2026. All rights reserved.").
  - Bullet writing now preserves the layout's inherited bullet character.
    Instead of clearing the text frame (which resets paragraph properties),
    we replace text content while keeping the first paragraph's formatting.
"""

import asyncio
import json
import os
import tempfile
import uuid

from google.adk.tools.tool_context import ToolContext
from google.genai import types
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_PARAGRAPH_ALIGNMENT
from pptx.util import Inches, Pt

from ..shared_libraries.config import (
    DEFAULT_TEMPLATE_URI,
    GCS_BUCKET_NAME,
    get_logger,
)
from ..shared_libraries.models import CoverSpec, DeckSpec, SlideSpec
from ..shared_libraries.utils import _insert_image
from .artifact_utils import get_gcs_file_as_local_path, save_presentation
from .layout_router import (
    LAYOUT_MAP,
    PREFERRED_LAYOUTS_FOR_SYNTHESIZER,
    SlotMap,
    get_placeholder_for_role,
    get_slot_map,
    is_layout_known,
)
from .pptx_editor import _insert_visual_into_slide
from .visual_generator import generate_visual


# Placeholders we must never touch — they self-populate from the master.
UNTOUCHABLE_PLACEHOLDER_TYPES = {
    PP_PLACEHOLDER_TYPE.DATE,
    PP_PLACEHOLDER_TYPE.FOOTER,
    PP_PLACEHOLDER_TYPE.SLIDE_NUMBER,
    PP_PLACEHOLDER_TYPE.HEADER,
}


def _is_untouchable(placeholder) -> bool:
    if placeholder is None:
        return False
    try:
        return placeholder.placeholder_format.type in UNTOUCHABLE_PLACEHOLDER_TYPES
    except Exception:
        return False


# =============================================================================
# Layout selection
# =============================================================================

def get_smart_layout(prs: Presentation, requested_name: str):
    """Map a requested layout name to the best matching Capgemini layout."""
    log = get_logger("layout_mapper")
    if not requested_name:
        requested_name = "Content 1 Chapterbox"
    original_request = requested_name
    requested_name_lower = requested_name.lower()

    layouts = prs.slide_layouts

    # 1. Exact match
    for layout in layouts:
        if layout.name == requested_name:
            return layout
    for layout in layouts:
        if layout.name.lower() == requested_name_lower:
            return layout

    # 2. Keyword-based mapping to Capgemini layouts
    keyword_map = [
        ("cover", "Title Slide 1"),
        ("title slide", "Title Slide 1"),
        ("opening", "Title Slide 1"),
        ("section header", "Title Slide Message left"),
        ("divider", "Title Slide Message left"),
        ("transition", "Title Slide Message left"),
        ("comparison", "Comparison Chapterbox"),
        ("compare", "Comparison Chapterbox"),
        ("side by side", "Comparison Chapterbox"),
        ("two content", "Comparison Chapterbox"),
        ("three", "Content 3 Boxes Chapterbox"),
        ("3 boxes", "Content 3 Boxes Chapterbox"),
        ("agenda", "Agenda 3"),
        ("toc", "Agenda 3"),
        ("roadmap", "Agenda 3"),
        ("closing", "Conclusion 1"),
        ("thank you", "End Slide 1"),
        ("end slide", "End Slide 1"),
        ("contact", "Conclusion 1"),
        ("title only", "Title Only Chapterbox"),
        ("image", "Content 2 Chapterbox"),
        ("picture", "Content 2 Chapterbox"),
        ("photo", "Content 2 Chapterbox"),
        ("title and content", "Content 1 Chapterbox"),
        ("title subtitle", "Title Subtitle Chapterbox"),
        ("content", "Content 1 Chapterbox"),
        ("bullet", "Content 1 Chapterbox"),
    ]

    for keyword, target_layout_name in keyword_map:
        if keyword in requested_name_lower:
            for layout in layouts:
                if layout.name == target_layout_name:
                    log.info(f"Mapped '{original_request}' -> '{target_layout_name}'")
                    return layout

    # 3. Last resort
    for layout in layouts:
        if layout.name == "Content 1 Chapterbox":
            log.info(f"No match for '{original_request}', falling back to Content 1 Chapterbox")
            return layout

    log.warning(f"No matching layout for '{original_request}', using first layout: {layouts[0].name}")
    return layouts[0]


# =============================================================================
# Text rendering helpers
# =============================================================================

def _rm_md(t: str) -> str:
    if not t:
        return ""
    return t.replace("**", "")


def _enable_autofit(tf):
    """Enable PowerPoint's built-in auto-fit so the font shrinks if needed."""
    try:
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass


def _write_simple_text(placeholder, text: str):
    """
    Write a single string into a placeholder, preserving inherited formatting.
    No truncation — auto-fit handles overflow by shrinking the font.
    """
    if placeholder is None or not text:
        return
    if _is_untouchable(placeholder):
        return
    try:
        placeholder.text = _rm_md(text)
        _enable_autofit(placeholder.text_frame)
    except Exception:
        pass


def _write_bullets_preserving_format(placeholder, bullets):
    """
    Write bullets WITHOUT clearing the text frame, so the layout's inherited
    bullet character (•) and paragraph formatting survive.

    Approach:
      1. Reuse the first existing paragraph (which carries the layout's
         bullet formatting via inheritance).
      2. Replace its text content (remove existing runs, add new ones).
      3. For subsequent bullets, add_paragraph() — new paragraphs inherit
         the same level-0 formatting including bullet character.
    """
    if placeholder is None or not bullets:
        return
    if _is_untouchable(placeholder):
        return

    tf = placeholder.text_frame
    bullets = [b for b in bullets if b and b.strip()]
    if not bullets:
        return

    def _set_paragraph_text(p, bullet_text, level=0):
        is_sub = (
            bullet_text.startswith("  ")
            or bullet_text.startswith("\t")
            or bullet_text.startswith("- ")
        )
        clean = bullet_text.strip(" \t-•*")

        # Remove any existing runs in this paragraph (keeps <a:pPr> bullet props).
        for r in list(p.runs):
            r._r.getparent().remove(r._r)

        p.level = 1 if is_sub else level

        parts = clean.split("**")
        for j, part in enumerate(parts):
            if not part:
                continue
            run = p.add_run()
            run.text = part
            if j % 2 != 0:
                run.font.bold = True

    # First bullet → reuse first paragraph (preserves bullet character).
    first_p = tf.paragraphs[0]
    _set_paragraph_text(first_p, bullets[0], level=0)

    # Subsequent bullets → add_paragraph() inherits formatting.
    for bullet in bullets[1:]:
        p = tf.add_paragraph()
        _set_paragraph_text(p, bullet, level=0)

    _enable_autofit(tf)


def _insert_picture_into_placeholder(placeholder, image_source):
    if placeholder is None or not image_source:
        return False
    if _is_untouchable(placeholder):
        return False
    try:
        placeholder.insert_picture(image_source)
        return True
    except Exception:
        return False


# =============================================================================
# Main render function
# =============================================================================

def render_slide_content(slide, spec_obj, layout_name: str, prs=None,
                         log=None, is_cover: bool = False):
    """
    Populate a slide based on its layout and the spec's content.

    Critical rules:
      * Never write into or remove DATE/FOOTER/SLIDE_NUMBER placeholders.
      * Never clear() a bullet text frame (strips the bullet character).
      * Never truncate; PowerPoint auto-shrinks the font instead.
    """
    slot_map: SlotMap = get_slot_map(layout_name)
    treat_as_cover = is_cover or slot_map.is_cover

    if slot_map.is_end_slide or slot_map.is_blank:
        if log:
            log.info(f"Layout '{layout_name}' has no text slots — skipping.")
        return

    # 1. CHAPTER LABEL
    chapter_ph = get_placeholder_for_role(slide, layout_name, "chapter_label")
    chapter_text = getattr(spec_obj, "chapter_label", None)
    if chapter_ph is not None and chapter_text:
        _write_simple_text(chapter_ph, chapter_text)

    # 2. TITLE
    title_ph = get_placeholder_for_role(slide, layout_name, "title")
    title_text = getattr(spec_obj, "title", None)
    if title_ph is not None and title_text:
        _write_simple_text(title_ph, title_text)
        if treat_as_cover and title_ph.text_frame.paragraphs:
            try:
                title_ph.text_frame.paragraphs[0].alignment = (
                    PP_PARAGRAPH_ALIGNMENT.CENTER
                )
            except Exception:
                pass

    # 3. SUBHEAD
    subhead_ph = get_placeholder_for_role(slide, layout_name, "subhead")
    subhead_text = getattr(spec_obj, "subhead", None)
    if subhead_ph is not None and subhead_text:
        _write_simple_text(subhead_ph, subhead_text)

    # 4. COVER-ONLY SLOTS
    if treat_as_cover:
        cover_sub_ph = get_placeholder_for_role(slide, layout_name, "cover_subtitle")
        if cover_sub_ph is not None:
            sub_text = (
                getattr(spec_obj, "subhead", None)
                or getattr(spec_obj, "cover_subtitle", None)
            )
            if sub_text:
                _write_simple_text(cover_sub_ph, sub_text)

        cover_author_ph = get_placeholder_for_role(slide, layout_name, "cover_author")
        if cover_author_ph is not None:
            author_text = getattr(spec_obj, "author_line", None)
            if author_text:
                _write_simple_text(cover_author_ph, author_text)

        image_source = (
            getattr(spec_obj, "image_data", None)
            or getattr(spec_obj, "image_file_path", None)
        )
        if image_source:
            picture_ph = get_placeholder_for_role(slide, layout_name, "picture")
            _insert_picture_into_placeholder(picture_ph, image_source)
        return

    # 5. BODY BULLETS
    has_bullets = (
        hasattr(spec_obj, "bullets") and bool(getattr(spec_obj, "bullets", None))
    )

    if has_bullets:
        body_ph = get_placeholder_for_role(slide, layout_name, "body")
        if body_ph is not None:
            _write_bullets_preserving_format(body_ph, spec_obj.bullets)

    bullets_right = getattr(spec_obj, "bullets_right", None)
    if bullets_right:
        body_right_ph = get_placeholder_for_role(slide, layout_name, "body_right")
        if body_right_ph is not None:
            _write_bullets_preserving_format(body_right_ph, bullets_right)

    bullets_third = getattr(spec_obj, "bullets_third", None)
    if bullets_third:
        body_third_ph = get_placeholder_for_role(slide, layout_name, "body_third")
        if body_third_ph is not None:
            _write_bullets_preserving_format(body_third_ph, bullets_third)

    # 6. IMAGE
    image_source = (
        getattr(spec_obj, "image_data", None)
        or getattr(spec_obj, "image_file_path", None)
    )
    if image_source:
        picture_ph = get_placeholder_for_role(slide, layout_name, "picture")
        if picture_ph is not None:
            inserted = _insert_picture_into_placeholder(picture_ph, image_source)
            if not inserted and log:
                log.warning(f"Failed to insert picture into '{layout_name}'")
        elif prs is not None:
            try:
                box_hint = (
                    int(prs.slide_width * 0.55),
                    int(prs.slide_height * 0.30),
                    int(prs.slide_width * 0.40),
                    int(prs.slide_height * 0.55),
                )
                _insert_image(prs, slide, image_source, box_hint=box_hint)
            except Exception as e:
                if log:
                    log.warning(f"Failed to float image: {e}")


# =============================================================================
# Visual layout coercion
# =============================================================================

LAYOUTS_WITH_PICTURE = [
    "Content 2 Chapterbox",
    "Content 2",
    "Divider Photo 1",
    "Divider Photo 2",
    "Title Slide 2",
    "Title Slide 3",
    "Conclusion 1 Photo",
    "Conclusion 2 Photo",
]

DEFAULT_IMAGE_LAYOUT = "Content 2 Chapterbox"


# =============================================================================
# render_deck_from_spec
# =============================================================================

async def render_deck_from_spec(
    spec_dict: dict,
    out_pptx: str,
    tool_context: ToolContext,
    template_pptx: str | None = None,
) -> str:
    """Render a presentation from a spec, using the provided template."""
    log = get_logger("render_deck_from_spec")
    try:
        if template_pptx and os.path.exists(template_pptx):
            log.info(f"Using user template '{template_pptx}' as the foundation.")
            working_template = template_pptx
        else:
            log.error("No valid user template provided. Aborting.")
            return "Error: No valid template provided."

        prs = Presentation(working_template)

        # COVER
        cover_data = spec_dict.get("cover", {"title": "Strategic Research & Analysis"})
        cover_spec = CoverSpec(**cover_data)

        if len(prs.slides) > 0:
            for i in range(len(prs.slides) - 1, 0, -1):
                rId = prs.slides._sldIdLst[i].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[i]

        cover_layout_name = getattr(cover_spec, "layout_name", None) or "Title Slide 1"

        if len(prs.slides) > 0:
            log.info("Template has an existing slide. Using it as the Cover Page.")
            cover_slide = prs.slides[0]
            try:
                existing_layout_name = cover_slide.slide_layout.name
                if is_layout_known(existing_layout_name):
                    cover_layout_name = existing_layout_name
            except Exception:
                pass
            try:
                render_slide_content(
                    cover_slide, cover_spec, cover_layout_name,
                    prs=prs, log=log, is_cover=True,
                )
            except Exception as e:
                log.warning(f"Could not render cover on existing slide: {e}")
        else:
            log.info("Template is empty. Generating a new Cover Page.")
            try:
                cover_slide = prs.slides.add_slide(get_smart_layout(prs, cover_layout_name))
                render_slide_content(
                    cover_slide, cover_spec, cover_layout_name,
                    prs=prs, log=log, is_cover=True,
                )
            except Exception as e:
                log.warning(f"Could not generate/render cover slide: {e}")

        # BODY SLIDES
        for s_data in spec_dict.get("slides", []):
            if "title" not in s_data or not s_data["title"]:
                s_data["title"] = "Slide Content"

            try:
                s_spec = SlideSpec(**s_data)
                layout = get_smart_layout(prs, s_spec.layout_name)
                slide = prs.slides.add_slide(layout)
                actual_layout_name = layout.name
                render_slide_content(
                    slide, s_spec, actual_layout_name,
                    prs=prs, log=log,
                )
            except Exception as e:
                log.error(f"Failed to render slide '{s_data.get('title')}': {e}")
                continue

            if getattr(s_spec, "speaker_notes", None) or getattr(s_spec, "citations", None):
                try:
                    notes = slide.notes_slide.notes_text_frame
                    text = s_spec.speaker_notes or ""
                    if s_spec.citations:
                        if text:
                            text += "\n\n---\nCitations:\n"
                        else:
                            text += "Citations:\n"
                        for citation in s_spec.citations:
                            text += f"- {citation}\n"
                    notes.text = text
                except Exception:
                    pass

        # CLOSING SLIDE
        try:
            closing_layout_name = spec_dict.get("closing_layout_name", "End Slide 1")
            closing_layout = get_smart_layout(prs, closing_layout_name)
            closing = prs.slides.add_slide(closing_layout)
            actual_closing_name = closing_layout.name

            closing_slot_map = get_slot_map(actual_closing_name)
            if not closing_slot_map.is_end_slide and not closing_slot_map.is_blank:
                closing_title = spec_dict.get("closing_title", "Thank You")

                class _ClosingSpec:
                    pass

                closing_spec = _ClosingSpec()
                closing_spec.title = closing_title
                closing_spec.subhead = spec_dict.get("closing_subhead")
                closing_spec.chapter_label = None
                closing_spec.bullets = None

                render_slide_content(
                    closing, closing_spec, actual_closing_name,
                    prs=prs, log=log,
                )
        except Exception as e:
            log.warning(f"Could not generate closing slide: {e}")

        # SAVE
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            prs.save(tmp.name)
            return tmp.name
    except Exception as e:
        log.error(f"Render failed: {e}", exc_info=True)
        return f"Error: Render failed. {e}"


# =============================================================================
# generate_and_render_deck
# =============================================================================

async def generate_and_render_deck(
    tool_context: ToolContext,
    deck_spec: dict | None = None,
    spec_artifact_name: str | None = None,
    template_path: str | None = None,
) -> dict:
    """Orchestrate the entire deck generation process."""
    log = get_logger("generate_and_render_deck_tool")
    try:
        spec_dict = deck_spec

        if not spec_dict and not spec_artifact_name:
            spec_dict = tool_context.state.get("current_deck_spec")
            if spec_dict:
                log.info("Loaded DeckSpec from session state.")

        if not spec_dict and spec_artifact_name:
            log.info(f"Loading DeckSpec from artifact: '{spec_artifact_name}'")
            try:
                artifact = await tool_context.load_artifact(spec_artifact_name)
                if not artifact:
                    log.warning(f"Artifact '{spec_artifact_name}' not found. Waiting 2s...")
                    await asyncio.sleep(2.0)
                    artifact = await tool_context.load_artifact(spec_artifact_name)

                if artifact:
                    spec_json = (
                        artifact.inline_data.data
                        if isinstance(artifact, types.Part)
                        else artifact
                    )
                    if isinstance(spec_json, (bytes, bytearray)):
                        spec_dict = json.loads(spec_json.decode("utf-8"))
                    elif isinstance(spec_json, str):
                        spec_dict = json.loads(spec_json)
                    else:
                        spec_dict = spec_json
            except Exception as e:
                log.error(f"Failed to load named spec artifact: {e}")

        if not spec_dict:
            return {
                "status": "Failed",
                "message": (
                    "No active presentation plan found in session state. "
                    "Please provide deck_spec or ensure an outline was generated."
                ),
            }

        working_template = template_path
        if not working_template or not os.path.exists(working_template):
            log.info("Template path invalid or lost. Re-downloading from GCS...")
            working_template = await get_gcs_file_as_local_path(DEFAULT_TEMPLATE_URI)

        if isinstance(spec_dict.get("slides"), dict):
            spec_dict["slides"] = list(spec_dict["slides"].values())
        if "closing_title" not in spec_dict:
            spec_dict["closing_title"] = "Thank You"

        validated_spec = DeckSpec(**spec_dict)
        all_content = [validated_spec.cover] + validated_spec.slides

        hard_limit = 5
        visuals_kept = 0
        for slide in validated_spec.slides:
            if slide.visual_prompt:
                if visuals_kept < hard_limit:
                    visuals_kept += 1
                    if slide.layout_name not in LAYOUTS_WITH_PICTURE:
                        log.info(
                            f"Coercing slide '{slide.title}' from "
                            f"'{slide.layout_name}' to '{DEFAULT_IMAGE_LAYOUT}'"
                        )
                        slide.layout_name = DEFAULT_IMAGE_LAYOUT
                else:
                    slide.visual_prompt = None

        tasks = []
        slides_with_visuals = []
        for item in all_content:
            if hasattr(item, "visual_prompt") and item.visual_prompt:
                tasks.append(
                    asyncio.create_task(
                        asyncio.wait_for(generate_visual(item.visual_prompt), timeout=60.0)
                    )
                )
                slides_with_visuals.append(item)

        images = await asyncio.gather(*tasks, return_exceptions=True)
        for s, img in zip(slides_with_visuals, images):
            if not isinstance(img, Exception):
                s.image_data = img

        out_name = f"{validated_spec.cover.title}_{uuid.uuid4().hex[:6]}.pptx"

        local_path = await render_deck_from_spec(
            validated_spec.model_dump(),
            out_name,
            tool_context,
            working_template,
        )
        if local_path.startswith("Error:"):
            return {"status": "Failed", "message": local_path}

        msg = await save_presentation(
            tool_context, out_name, local_path, GCS_BUCKET_NAME
        )
        os.remove(local_path)
        return {"status": "Success", "message": msg}
    except Exception as e:
        log.error(f"Generation failed: {e}", exc_info=True)
        return {"status": "Failed", "message": str(e)}
