"""
Patch for presentation_agent/shared_libraries/models.py

Adds the new content tier fields to SlideSpec that the Capgemini template
requires: chapter_label (breadcrumb), bullets_right and bullets_third
(for multi-column layouts), and author_line (for cover slides).

Apply these as additions to the existing SlideSpec class — do NOT remove
existing fields, just add the new optional ones below.
"""

from typing import Optional, List


# ----------------------------------------------------------------------------
# ADD THESE FIELDS TO YOUR EXISTING SlideSpec PYDANTIC MODEL
# ----------------------------------------------------------------------------

ADDITIONS_TO_SLIDESPEC = """
    chapter_label: Optional[str] = Field(
        default=None,
        description=(
            "Breadcrumb shown above the slide title on 'Chapterbox' layouts. "
            "1-3 words naming the current section/chapter. Use the same value "
            "across all slides in a section. Set to null for cover and "
            "section-divider slides."
        ),
        max_length=50,
    )

    bullets_right: Optional[List[str]] = Field(
        default=None,
        description=(
            "Bullets for the right column on 2-column layouts "
            "(Content 2/3/4, Comparison). Each item 10-25 words."
        ),
    )

    bullets_third: Optional[List[str]] = Field(
        default=None,
        description=(
            "Bullets for the third column on 3-column layouts "
            "(Content 3 Boxes). Each item 10-25 words."
        ),
    )

    author_line: Optional[str] = Field(
        default=None,
        description=(
            "Author name and date for cover slides. "
            "Format: 'Jane Doe | November 2025'. Cover slides only."
        ),
        max_length=60,
    )
"""


# ----------------------------------------------------------------------------
# COMPLETE EXAMPLE — what the updated SlideSpec should look like
# ----------------------------------------------------------------------------

EXAMPLE_FULL_SLIDESPEC = '''
from typing import Optional, List
from pydantic import BaseModel, Field


class SlideSpec(BaseModel):
    """A single slide in a presentation deck, aligned to the Capgemini template."""
    
    # Existing fields (keep as-is)
    title: str = Field(..., description="Main slide heading")
    layout_name: str = Field(..., description="One of the canonical layouts")
    bullets: Optional[List[str]] = Field(default=None)
    visual_prompt: Optional[str] = Field(default=None)
    speaker_notes: Optional[str] = Field(default=None)
    citations: Optional[List[str]] = Field(default=None)
    image_data: Optional[bytes] = Field(default=None)
    image_file_path: Optional[str] = Field(default=None)
    
    # NEW fields for Capgemini template tier hierarchy
    chapter_label: Optional[str] = Field(
        default=None,
        max_length=50,
        description="Section breadcrumb. 1-3 words. Null on covers/dividers."
    )
    
    subhead: Optional[str] = Field(
        default=None,
        max_length=90,
        description="Blue 20pt accent line below title. The slide's takeaway."
    )
    
    bullets_right: Optional[List[str]] = Field(
        default=None,
        description="Right-column bullets for 2-column layouts."
    )
    
    bullets_third: Optional[List[str]] = Field(
        default=None,
        description="Third-column bullets for 3-box layouts."
    )
    
    author_line: Optional[str] = Field(
        default=None,
        max_length=60,
        description="Cover slides only — 'Name | Date'."
    )
'''
