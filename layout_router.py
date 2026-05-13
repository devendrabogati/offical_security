# Copyright 2026 Capgemini
#
# Licensed under the Apache License, Version 2.0 (the "License");

"""
Layout router for the Capgemini Talk2Docs presentation template.

This module is the single source of truth for mapping each layout in the
official Capgemini PowerPoint template to its semantic content roles.

Background
----------
The template ships with 59 layouts. Several of them use the TITLE-typed
placeholder for decoration (e.g. a giant "01" number on a section divider)
rather than for the actual slide title. Naive renderers that grab the first
TITLE-typed placeholder and stuff slide-title text into it produce broken
slides where text overflows or is cropped.

This router fixes that by defining, per layout name, exactly which placeholder
`idx` value receives each semantic content role:

  - chapter_label : tiny breadcrumb at top of slide (~0.18" tall)
  - title         : main heading (1.02" tall at 32pt Ubuntu Medium)
  - subhead       : blue accent line below title (0.34" tall at 20pt)
  - body          : main content area (bullets at 14pt)
  - body_right    : right-side content for 2-column layouts
  - body_third    : third column for 3-column layouts
  - picture       : main image placeholder
  - picture_2     : secondary image placeholder
  - cover_subtitle: cover-only subtitle (idx=10 on Title Slide 1)
  - cover_author  : cover-only author/date line (idx=11 on Title Slide 1)

If a layout has no slot for a given role, the value is None. The renderer
should skip writing that content (and the synthesizer should not generate it).

The router also exposes character/word limits per slot, derived from the
actual placeholder dimensions. These should be enforced in the synthesizer
prompt to prevent overflow at generation time, not just at render time.
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass(frozen=True)
class SlotMap:
    """Maps semantic content roles to placeholder idx values for one layout."""
    # Core text slots — idx values into slide.placeholders
    chapter_label: Optional[int] = None
    title: Optional[int] = None
    subhead: Optional[int] = None
    body: Optional[int] = None
    body_right: Optional[int] = None
    body_third: Optional[int] = None

    # Cover-only slots
    cover_subtitle: Optional[int] = None
    cover_author: Optional[int] = None

    # Visual slots
    picture: Optional[int] = None
    picture_2: Optional[int] = None

    # Decorative slots — placeholders to leave alone (do not write into)
    decorative: tuple = field(default_factory=tuple)

    # Hard limits enforced in the prompt — derived from placeholder size
    title_max_chars: int = 60
    subhead_max_chars: int = 90
    body_max_bullets: int = 6
    body_bullet_max_chars: int = 120

    # Behaviour hints
    is_cover: bool = False
    is_section_divider: bool = False
    is_end_slide: bool = False
    is_blank: bool = False


# =============================================================================
# Per-layout slot maps. Keys are the EXACT layout name strings from the .pptx.
# =============================================================================

LAYOUT_MAP: dict[str, SlotMap] = {

    # -------------------------------------------------------------------------
    # COVER SLIDES
    # -------------------------------------------------------------------------
    # Layout 0: 'Title Slide 1' — main cover.
    #   idx=0  TITLE (8.00" x 4.49") — title area, fits ~3 lines at 72pt
    #   idx=10 BODY  (7.09" x 0.24") — subtitle, fits 1 line of small text
    #   idx=11 BODY  (7.09" x 0.24") — author/date, fits 1 line
    "Title Slide 1": SlotMap(
        title=0,
        cover_subtitle=10,
        cover_author=11,
        title_max_chars=45,  # 3 lines of ~15 chars at 72pt
        subhead_max_chars=60,
        is_cover=True,
    ),

    # Layout 1: 'Title Slide Logo Cut' — cover variant, title is narrower
    "Title Slide Logo Cut": SlotMap(
        title=0,
        cover_subtitle=10,
        cover_author=11,
        title_max_chars=40,
        subhead_max_chars=60,
        is_cover=True,
    ),

    # Layout 3, 4: 'Title Slide 2' / 'Title Slide 3' — cover with bg image
    #   idx=13 PICTURE (full slide background)
    #   idx=0  TITLE (overlay)
    #   idx=14, 15 BODY (subtitle / author)
    "Title Slide 2": SlotMap(
        title=0,
        picture=13,
        cover_subtitle=14,
        cover_author=15,
        title_max_chars=45,
        is_cover=True,
    ),
    "Title Slide 3": SlotMap(
        title=0,
        picture=13,
        cover_subtitle=14,
        cover_author=15,
        title_max_chars=45,
        is_cover=True,
    ),

    # Layouts 5, 6, 7: 'Title Slide 4/5/6' — covers with 2 image regions
    "Title Slide 4": SlotMap(
        title=0,
        picture=13,
        picture_2=14,
        cover_subtitle=15,
        cover_author=16,
        title_max_chars=45,
        is_cover=True,
    ),
    "Title Slide 5": SlotMap(
        title=0,
        picture=13,
        picture_2=14,
        cover_subtitle=16,
        cover_author=17,
        title_max_chars=45,
        is_cover=True,
    ),
    "Title Slide 6": SlotMap(
        title=0,
        picture=13,
        picture_2=14,
        cover_subtitle=15,
        cover_author=16,
        title_max_chars=45,
        is_cover=True,
    ),

    # -------------------------------------------------------------------------
    # SECTION DIVIDERS — TITLE placeholder is DECORATIVE here. Real content
    # lives in BODY placeholders.
    # -------------------------------------------------------------------------
    # Layout 2: 'Title Slide Message left'
    #   idx=0  TITLE (8.44" x 6.78") — DECORATIVE (giant "01" or letter)
    #   idx=10 BODY  (4.13" x 0.92") — actual message/title on the right
    #   idx=11 BODY  (4.13" x 0.60") — subhead/description
    "Title Slide Message left": SlotMap(
        title=10,            # real title goes into BODY idx=10
        subhead=11,
        decorative=(0,),     # skip the giant TITLE placeholder
        title_max_chars=40,
        subhead_max_chars=80,
        is_section_divider=True,
    ),

    # Layout 12: 'Divider 1'
    #   idx=0  TITLE (7.09" x 1.83") at (5.64, 5.09) — the actual title here
    #   idx=19 BODY  (5.09" x 7.50") at (0,0) — full-height left panel deco
    "Divider 1": SlotMap(
        title=0,
        decorative=(19,),
        title_max_chars=50,
        is_section_divider=True,
    ),

    # Layout 13: 'Divider Photo 1'
    "Divider Photo 1": SlotMap(
        title=0,
        picture=13,
        decorative=(19, 28, 40),
        title_max_chars=50,
        is_section_divider=True,
    ),

    # Layout 14: 'Divider 2'
    "Divider 2": SlotMap(
        title=0,
        decorative=(23,),
        title_max_chars=70,
        is_section_divider=True,
    ),

    # Layout 15: 'Divider Photo 2'
    "Divider Photo 2": SlotMap(
        title=0,
        picture=13,
        decorative=(22, 28, 40),
        title_max_chars=70,
        is_section_divider=True,
    ),

    # -------------------------------------------------------------------------
    # BLANK / TITLE ONLY
    # -------------------------------------------------------------------------
    # Layout 16: 'Blank' — no content placeholders, only footer/date/page
    "Blank": SlotMap(is_blank=True),

    # Layout 17: 'Title Only'
    "Title Only": SlotMap(title=0, title_max_chars=80),

    # Layout 18: 'Title Only Chapterbox'
    "Title Only Chapterbox": SlotMap(
        title=0,
        chapter_label=13,
        title_max_chars=80,
    ),

    # -------------------------------------------------------------------------
    # WORKHORSE CONTENT LAYOUTS
    # All share the standard grammar:
    #   chapter_label (idx=13, optional, "Chapterbox" variants only)
    #   title         (idx=0)   1.02" tall — ~8 words
    #   subhead       (idx=22)  0.34" tall — 1 sentence
    #   body          (idx=31)  4.57" tall — ~6 bullets
    # -------------------------------------------------------------------------
    # Layout 19: 'Content 1' (single column, full width)
    "Content 1": SlotMap(
        title=0, subhead=22, body=31,
        title_max_chars=80,
    ),
    # Layout 20: 'Content 1 Chapterbox'
    "Content 1 Chapterbox": SlotMap(
        title=0, subhead=22, body=31, chapter_label=13,
        title_max_chars=80,
    ),
    # Layout 21: 'Title Subtitle' — title + subhead only, no bullets
    "Title Subtitle": SlotMap(
        title=0, subhead=22,
        title_max_chars=80,
    ),
    # Layout 22: 'Title Subtitle Chapterbox'
    "Title Subtitle Chapterbox": SlotMap(
        title=0, subhead=22, chapter_label=13,
        title_max_chars=80,
    ),

    # 2-column layouts: title on left, second column on right
    # Layout 23: 'Content 2' (1/3 left, 2/3 right)
    #   idx=0  title (3.62" wide)         idx=27 subhead_right (7.64" wide)
    #   idx=22 subhead (3.62" wide)       idx=28 body_right
    #   idx=31 body (3.62" wide)
    "Content 2": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28,
        title_max_chars=40,         # narrower title here
        body_bullet_max_chars=80,
    ),
    "Content 2 Chapterbox": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28, chapter_label=13,
        title_max_chars=40,
        body_bullet_max_chars=80,
    ),

    # Layout 25: 'Content 3' (50/50 split)
    "Content 3": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28,
        title_max_chars=60,
        body_bullet_max_chars=100,
    ),
    "Content 3 Chapterbox": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28, chapter_label=13,
        title_max_chars=60,
        body_bullet_max_chars=100,
    ),

    # Layout 27: 'Content 4' (2/3 left, 1/3 right)
    "Content 4": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28,
        title_max_chars=70,
        body_bullet_max_chars=100,
    ),
    "Content 4 Chapterbox": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28, chapter_label=13,
        title_max_chars=70,
        body_bullet_max_chars=100,
    ),

    # Layout 29: 'Content 5' (75/25 split — sidebar style)
    "Content 5": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28,
        title_max_chars=75,
        body_bullet_max_chars=110,
    ),
    "Content 5 Chapterbox": SlotMap(
        title=0, subhead=22, body=31,
        body_right=28, chapter_label=13,
        title_max_chars=75,
        body_bullet_max_chars=110,
    ),

    # Layout 31: 'Content 6' — narrower title, no body shown in inspection
    "Content 6": SlotMap(
        title=0, subhead=22,
        title_max_chars=90,
    ),

    # -------------------------------------------------------------------------
    # 3-COLUMN / SPECIALIZED
    # -------------------------------------------------------------------------
    # Layout 36: 'Content 3 Boxes Chapterbox'
    #   3 sub-headers (22, 35, 36) + 3 content boxes (37, 43, 44)
    "Content 3 Boxes Chapterbox": SlotMap(
        title=0,
        subhead=22, body=37,         # left box
        body_right=43,               # middle box
        body_third=44,               # right box
        chapter_label=13,
        title_max_chars=80,
        body_bullet_max_chars=70,    # narrower columns
        body_max_bullets=4,
    ),

    # Layout 37: 'Comparison' (2 equal columns)
    "Comparison": SlotMap(
        title=0,
        subhead=22, body=23,         # left column
        body_right=25,               # right column
        title_max_chars=80,
        body_bullet_max_chars=90,
    ),
    "Comparison Chapterbox": SlotMap(
        title=0,
        subhead=22, body=23,
        body_right=25, chapter_label=13,
        title_max_chars=80,
        body_bullet_max_chars=90,
    ),

    # -------------------------------------------------------------------------
    # AGENDAS — special, multi-item TOC layouts
    # -------------------------------------------------------------------------
    "Agenda 1": SlotMap(
        title=0, body=1, picture=10,
        title_max_chars=70,
    ),
    "Agenda 2": SlotMap(
        title=0, picture=18,
        title_max_chars=50,
        # Items go into idx=25,26,30,31,32,34,35 — handled specially
    ),
    "Agenda 3": SlotMap(
        title=0,
        title_max_chars=80,
    ),
    "Agenda 4": SlotMap(
        title=0,
        title_max_chars=80,
    ),

    # -------------------------------------------------------------------------
    # CONCLUSION SLIDES
    # -------------------------------------------------------------------------
    "Conclusion 1": SlotMap(
        title=0, body=17,
        title_max_chars=40,
    ),
    "Conclusion 1 Photo": SlotMap(
        title=0, body=17, picture=13,
        title_max_chars=40,
    ),
    "Conclusion 2": SlotMap(
        title=0, body=17,
        title_max_chars=40,
    ),
    "Conclusion 2 Photo": SlotMap(
        title=0, body=17, picture=13,
        title_max_chars=40,
    ),
    "Conclusion 3": SlotMap(
        title=0,
        title_max_chars=40,
    ),
    "Conclusion 3 Photo": SlotMap(
        title=0, picture=13,
        title_max_chars=40,
    ),

    # -------------------------------------------------------------------------
    # END / THANK YOU SLIDES — no placeholders; pure decoration
    # -------------------------------------------------------------------------
    "End Slide 1": SlotMap(is_end_slide=True),
    "End Slide 2": SlotMap(is_end_slide=True),
    "End Slide 3": SlotMap(is_end_slide=True),
}


# =============================================================================
# Public API
# =============================================================================

# Canonical layouts that the synthesizer agent is allowed to choose from.
# Anything else gets remapped to one of these defaults.
PREFERRED_LAYOUTS_FOR_SYNTHESIZER = [
    "Title Slide 1",              # Cover
    "Title Subtitle Chapterbox",  # Title + subhead, no bullets
    "Content 1 Chapterbox",       # Standard content workhorse (use this often)
    "Content 2 Chapterbox",       # Content + sidebar
    "Comparison Chapterbox",      # Two-column compare
    "Content 3 Boxes Chapterbox", # 3 named boxes
    "Title Slide Message left",   # Section divider
    "Title Only Chapterbox",      # For big custom visuals
    "Conclusion 1",               # Closing
    "End Slide 1",                # Final thank-you
]


def get_slot_map(layout_name: str) -> SlotMap:
    """Return the SlotMap for a layout name. Falls back to Content 1 if unknown."""
    if layout_name in LAYOUT_MAP:
        return LAYOUT_MAP[layout_name]
    # Conservative fallback — assume the layout follows the standard grammar
    return LAYOUT_MAP["Content 1"]


def get_placeholder_for_role(slide, layout_name: str, role: str):
    """
    Resolve the actual placeholder shape on a slide for a given semantic role.

    Args:
        slide: a python-pptx Slide object
        layout_name: the layout's name string
        role: one of 'title', 'subhead', 'body', 'body_right', 'body_third',
              'chapter_label', 'cover_subtitle', 'cover_author', 'picture',
              'picture_2'

    Returns:
        The placeholder shape, or None if this layout has no slot for the role.
    """
    slot_map = get_slot_map(layout_name)
    target_idx = getattr(slot_map, role, None)
    if target_idx is None:
        return None
    try:
        return slide.placeholders[target_idx]
    except KeyError:
        return None


def remove_decorative_placeholders(slide, layout_name: str):
    """
    No-op: decorative placeholders are part of the layout's visual design
    (e.g. the big '01' number on a section divider). We don't remove them —
    we just avoid writing into them. This function exists as a hook in case
    we ever need to override.
    """
    pass


def is_layout_known(layout_name: str) -> bool:
    """Whether this layout has an explicit slot map."""
    return layout_name in LAYOUT_MAP
