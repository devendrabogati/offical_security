"""
Patched render_slide_content for presentation_orchestrator.py

This replaces the existing render_slide_content inside the
generate_and_render_deck (and render_deck_from_spec) tools.

What changed vs the original:
  1. Uses the layout_router to find the *correct* placeholder for each
     semantic role (title, subhead, body, chapter_label, etc.), instead of
     grabbing the first TITLE-typed placeholder it sees.
  2. Adds MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE to every text frame so PowerPoint
     auto-shrinks the font if text exceeds the placeholder box.
  3. Truncates text at the prompt-side limit AND at the render-side limit
     (belt-and-suspenders).
  4. Removes empty placeholders so master prompt text ("Click to edit Master
     subtitle style") never leaks through.
  5. Handles is_end_slide and is_blank layouts cleanly — just adds the slide
     without trying to write text into placeholders that don't exist.
"""

from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.util import Pt

# Import the router (adjust import path to match your project layout)
from .layout_router import (
    SlotMap,
    get_slot_map,
    get_placeholder_for_role,
)


def _rm_md(t: str) -> str:
    """Strip basic markdown markers from a string."""
    if not t:
        return ""
    return t.replace("**", "")


def _truncate(text: str, max_chars: int, suffix: str = "…") -> str:
    """Hard-truncate text at max_chars, preferring to break on word boundary."""
    if not text or len(text) <= max_chars:
        return text
    cut = text[: max_chars - len(suffix)]
    # try to back off to the last space so we don't slice mid-word
    last_space = cut.rfind(" ")
    if last_space > max_chars * 0.6:
        cut = cut[:last_space]
    return cut + suffix


def _configure_text_frame(tf, autosize: bool = True):
    """Apply safe defaults to a text frame so text fits its placeholder."""
    tf.word_wrap = True
    if autosize:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.vertical_anchor = MSO_ANCHOR.TOP


def _write_plain_text(placeholder, text: str, max_chars: int = None):
    """Write a single string into a placeholder safely."""
    if placeholder is None or not text:
        return
    if max_chars:
        text = _truncate(text, max_chars)
    placeholder.text = _rm_md(text)
    _configure_text_frame(placeholder.text_frame)


def _apply_bullets(text_frame, bullets, max_bullets: int = 6,
                   bullet_max_chars: int = 120):
    """Write a list of bullet strings into a text frame, with **bold** parsing."""
    text_frame.clear()
    _configure_text_frame(text_frame)

    # Cap the count
    bullets = list(bullets)[:max_bullets]

    for i, bullet_text in enumerate(bullets):
        # Detect sub-bullet indent
        is_sub = (bullet_text.startswith("  ")
                  or bullet_text.startswith("\t")
                  or bullet_text.startswith("- "))

        clean_text = bullet_text.strip(" \t-•*")
        clean_text = _truncate(clean_text, bullet_max_chars)

        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.level = 1 if is_sub else 0

        # Parse **bold** segments
        parts = clean_text.split("**")
        for j, part in enumerate(parts):
            if not part:
                continue
            r = p.add_run()
            r.text = part
            if j % 2 != 0:
                r.font.bold = True


def _remove_placeholder(placeholder):
    """Remove a placeholder shape from its slide entirely.

    Used when a placeholder would otherwise show the master's prompt text
    ('Click to edit ...') because we have no content for it.
    """
    if placeholder is None:
        return
    try:
        sp = placeholder._element
        sp.getparent().remove(sp)
    except Exception:
        pass


# =============================================================================
# Main entry point — drop this into presentation_orchestrator.py in place of
# the existing render_slide_content nested function.
# =============================================================================

def render_slide_content(slide, spec_obj, layout_name: str, log=None):
    """
    Populate a slide based on its layout and the spec's content.

    Args:
        slide: a python-pptx Slide created from the chosen layout
        spec_obj: a SlideSpec or DeckCover/DeckClosing pydantic model with
                  fields: title, subhead, bullets, chapter_label (optional),
                  speaker_notes (optional), image_data/image_file_path
        layout_name: the layout name string used to create this slide.
                     MUST match a key in the layout_router.LAYOUT_MAP.
    """
    slot_map: SlotMap = get_slot_map(layout_name)

    # ---------------------------------------------------------------
    # Short-circuit: layouts with no text content
    # ---------------------------------------------------------------
    if slot_map.is_end_slide or slot_map.is_blank:
        if log:
            log.info(f"Layout '{layout_name}' has no text slots — skipping.")
        return

    # ---------------------------------------------------------------
    # 1. CHAPTER LABEL (Chapterbox layouts only)
    # ---------------------------------------------------------------
    chapter_ph = get_placeholder_for_role(slide, layout_name, "chapter_label")
    chapter_text = getattr(spec_obj, "chapter_label", None)
    if chapter_ph is not None:
        if chapter_text:
            _write_plain_text(chapter_ph, chapter_text, max_chars=50)
        else:
            # No chapter text generated — remove placeholder to hide prompt text
            _remove_placeholder(chapter_ph)

    # ---------------------------------------------------------------
    # 2. TITLE
    # ---------------------------------------------------------------
    title_ph = get_placeholder_for_role(slide, layout_name, "title")
    title_text = getattr(spec_obj, "title", None)
    if title_ph is not None and title_text:
        _write_plain_text(title_ph, title_text,
                          max_chars=slot_map.title_max_chars)
        # Center the title on cover slides
        if slot_map.is_cover and title_ph.text_frame.paragraphs:
            title_ph.text_frame.paragraphs[0].alignment = (
                PP_PARAGRAPH_ALIGNMENT.CENTER
            )

    # ---------------------------------------------------------------
    # 3. SUBHEAD (the blue 20pt accent line on workhorse layouts)
    # ---------------------------------------------------------------
    subhead_ph = get_placeholder_for_role(slide, layout_name, "subhead")
    subhead_text = getattr(spec_obj, "subhead", None)
    if subhead_ph is not None:
        if subhead_text:
            _write_plain_text(subhead_ph, subhead_text,
                              max_chars=slot_map.subhead_max_chars)
        else:
            _remove_placeholder(subhead_ph)

    # ---------------------------------------------------------------
    # 4. COVER-ONLY SLOTS (subtitle + author/date)
    # ---------------------------------------------------------------
    if slot_map.is_cover:
        cover_sub_ph = get_placeholder_for_role(
            slide, layout_name, "cover_subtitle"
        )
        if cover_sub_ph is not None:
            sub_text = getattr(spec_obj, "subhead", None) \
                       or getattr(spec_obj, "cover_subtitle", None)
            if sub_text:
                _write_plain_text(cover_sub_ph, sub_text, max_chars=60)
            else:
                _remove_placeholder(cover_sub_ph)

        cover_author_ph = get_placeholder_for_role(
            slide, layout_name, "cover_author"
        )
        if cover_author_ph is not None:
            author_text = getattr(spec_obj, "author_line", None)
            if author_text:
                _write_plain_text(cover_author_ph, author_text, max_chars=60)
            else:
                _remove_placeholder(cover_author_ph)

        # Covers have no further body content — we're done
        return

    # ---------------------------------------------------------------
    # 5. BODY BULLETS
    # ---------------------------------------------------------------
    has_bullets = hasattr(spec_obj, "bullets") and bool(spec_obj.bullets)

    if has_bullets:
        body_ph = get_placeholder_for_role(slide, layout_name, "body")
        if body_ph is not None:
            _apply_bullets(
                body_ph.text_frame,
                spec_obj.bullets,
                max_bullets=slot_map.body_max_bullets,
                bullet_max_chars=slot_map.body_bullet_max_chars,
            )

        # 2-column / 3-column layouts: distribute extra bullets if provided
        bullets_right = getattr(spec_obj, "bullets_right", None)
        if bullets_right:
            body_right_ph = get_placeholder_for_role(
                slide, layout_name, "body_right"
            )
            if body_right_ph is not None:
                _apply_bullets(
                    body_right_ph.text_frame,
                    bullets_right,
                    max_bullets=slot_map.body_max_bullets,
                    bullet_max_chars=slot_map.body_bullet_max_chars,
                )

        bullets_third = getattr(spec_obj, "bullets_third", None)
        if bullets_third:
            body_third_ph = get_placeholder_for_role(
                slide, layout_name, "body_third"
            )
            if body_third_ph is not None:
                _apply_bullets(
                    body_third_ph.text_frame,
                    bullets_third,
                    max_bullets=slot_map.body_max_bullets,
                    bullet_max_chars=slot_map.body_bullet_max_chars,
                )

    # ---------------------------------------------------------------
    # 6. IMAGE (if the layout supports one and we have one)
    # ---------------------------------------------------------------
    image_source = (getattr(spec_obj, "image_data", None)
                    or getattr(spec_obj, "image_file_path", None))
    if image_source:
        picture_ph = get_placeholder_for_role(slide, layout_name, "picture")
        if picture_ph is not None:
            try:
                picture_ph.insert_picture(image_source)
            except Exception as e:
                if log:
                    log.warning(f"Failed to insert picture: {e}")

    # ---------------------------------------------------------------
    # 7. SPEAKER NOTES (unchanged from original)
    # ---------------------------------------------------------------
    notes_text = getattr(spec_obj, "speaker_notes", None)
    if notes_text:
        try:
            slide.notes_slide.notes_text_frame.text = notes_text
        except Exception as e:
            if log:
                log.warning(f"Failed to set speaker notes: {e}")
