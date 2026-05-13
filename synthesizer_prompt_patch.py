"""
Patch for presentation_agent/sub_agents/synthesizer/prompt.py

REPLACES the existing SYNTHESIZER_OUTLINE_INSTRUCTION and
SYNTHESIZER_SLIDE_INSTRUCTION rules around title length, layout selection,
and the new chapter_label / subhead fields.

The key changes:
  1. New content tiers: chapter_label, title, subhead, bullets are now
     distinct fields with distinct purposes and hard length limits.
  2. Layout names are restricted to the canonical Capgemini set defined in
     layout_router.PREFERRED_LAYOUTS_FOR_SYNTHESIZER.
  3. Hard character limits per layout — these mirror the placeholder
     dimensions so generated content fits without overflow.
"""


# ----------------------------------------------------------------------------
# Drop-in addition to SYNTHESIZER_OUTLINE_INSTRUCTION — append after rule #2,
# or replace rule #2 with this expanded version.
# ----------------------------------------------------------------------------

CAPGEMINI_TITLE_RULES = """
    ## CAPGEMINI TEMPLATE — Content Tier Hierarchy (CRITICAL) ##
    
    Every slide in this template has up to FOUR content tiers. You MUST
    distribute content across them correctly, not stuff everything into one.
    
    1. **chapter_label** (1-3 words, max 50 chars):
       A breadcrumb showing what section/chapter this slide belongs to.
       Examples: "Market Overview", "Strategy", "Case Studies", "Recommendations".
       Use the SAME chapter_label across all slides in the same section so
       readers know which chapter they're in.
       OMIT this field (set to null) on cover slides and section dividers.
    
    2. **title** (max 8 words, max chars depends on layout — see below):
       The main heading. ONE crisp statement of what this slide is about.
       Do NOT write a thesis here. Save the argument for the subhead.
       GOOD: "The Cloud Market Landscape"
       BAD:  "The 2024 Cloud Landscape: A Strategic Guide to AWS, Azure, and GCP"
    
    3. **subhead** (one sentence, max 90 chars):
       The blue accent line below the title. Summarizes the slide's TAKEAWAY
       — what should the reader walk away knowing? This is where the
       strategic insight goes, NOT in the title.
       GOOD: "Three vendors capture 66% of a $79.8B market in Q1 2024"
       BAD:  "Overview" (too vague — the title already says this)
    
    4. **bullets** (3-6 items, 10-25 words each):
       The supporting evidence and detail. Each bullet should ground a claim
       in a specific metric, citation, or example. Use **bold** for the
       leading concept.
    
    ## TITLE LENGTH LIMITS BY LAYOUT ##
    
    Different layouts have differently sized title boxes. You MUST respect
    these limits or text will overflow:
    
    | Layout name                    | title_max_chars | use when                        |
    |--------------------------------|-----------------|----------------------------------|
    | Title Slide 1                  | 45              | Cover slide                     |
    | Title Subtitle Chapterbox      | 80              | Title + subhead only            |
    | Content 1 Chapterbox           | 80              | Standard content (DEFAULT)      |
    | Content 2 Chapterbox           | 40              | Sidebar + main content          |
    | Content 3 Chapterbox           | 60              | 2 equal columns                 |
    | Content 4 Chapterbox           | 70              | Main + narrow sidebar           |
    | Comparison Chapterbox          | 80              | Two-option compare              |
    | Content 3 Boxes Chapterbox     | 80              | Three labeled boxes             |
    | Title Slide Message left       | 40              | Section divider                 |
    | Conclusion 1                   | 40              | Closing summary                 |
    
    ## LAYOUT SELECTION — APPROVED LIST ##
    
    You may ONLY choose layout_name from this list (anything else is an error):
    
    - "Title Slide 1"              → for the cover slide ONLY (slide 1)
    - "Title Slide Message left"   → for major section transitions
    - "Title Subtitle Chapterbox"  → title + subhead, no bullets (intros, transitions)
    - "Content 1 Chapterbox"       → DEFAULT body slide (use most often)
    - "Content 2 Chapterbox"       → when one side has supporting detail
    - "Comparison Chapterbox"      → comparing TWO options side-by-side
    - "Content 3 Boxes Chapterbox" → comparing THREE options
    - "Title Only Chapterbox"      → for slides with a big custom diagram (no bullets)
    - "Conclusion 1"               → for the closing summary slide
    - "End Slide 1"                → for the final thank-you slide (no text)
    
    Prefer "Chapterbox" variants over non-Chapterbox — they include the chapter
    breadcrumb which gives the reader navigation context.
"""


# ----------------------------------------------------------------------------
# Drop-in replacement for SYNTHESIZER_SLIDE_INSTRUCTION rule #1.
# ----------------------------------------------------------------------------

CAPGEMINI_SLIDE_CONTENT_RULES = """
    1. **Content Generation (Capgemini Template Constraints):**
       
       a. **chapter_label** (REQUIRED for Chapterbox layouts, omit for covers):
          1-3 words, max 50 characters. Use the same label across all slides
          in the same section. Example: "Market Overview".
       
       b. **title** — HARD LIMIT depends on the layout you selected:
          - Cover slides (Title Slide 1): max 45 chars (about 6 words)
          - Section dividers: max 40 chars  
          - Comparison/3-Boxes: max 80 chars
          - Default (Content 1 Chapterbox): max 80 chars
          If your strategic message needs more words, MOVE THEM to the subhead.
       
       c. **subhead** — one sentence, max 90 characters:
          Summarize the slide's takeaway. Include the key number or insight.
          This becomes the blue 20pt accent line. ALWAYS include this on
          content slides — leaving it empty causes the master prompt text
          ("Click to edit...") to leak through in the rendered output.
       
       d. **bullets** — 3 to 6 items, each 10-25 words:
          For single-column layouts (Content 1 Chapterbox): aim for 5-6 bullets.
          For 2-column layouts (Content 2/3/4): split content between `bullets`
          and `bullets_right`, ~3-4 each.
          For 3-box layouts: use `bullets`, `bullets_right`, `bullets_third`
          with ~2-3 short items each.
          Each bullet should include a specific metric, year, or attribution.
          Use **bold** at the start of each bullet for the leading concept.
"""


# ----------------------------------------------------------------------------
# Example before/after — include this in the prompt so the model has a
# concrete reference for the right shape of content.
# ----------------------------------------------------------------------------

CAPGEMINI_EXAMPLE = """
    ## EXAMPLE: BAD vs GOOD content distribution ##
    
    User asks for a slide about the 2024 cloud market.
    
    BAD (what the agent currently produces — text overflows):
    {
        "layout_name": "Title Slide 1",
        "title": "The 2024 Cloud Landscape: A Strategic Guide to AWS, Azure, and GCP",
        "subhead": null,
        "bullets": [...]
    }
    
    GOOD (correctly distributes content across tiers):
    {
        "layout_name": "Content 1 Chapterbox",
        "chapter_label": "Market Overview",
        "title": "The 2024 Cloud Market Landscape",
        "subhead": "Three vendors capture 66% of a $79.8B market growing 21% YoY",
        "bullets": [
            "**AWS** holds 31% share but shows a slight downward trend after years of dominance",
            "**Microsoft Azure** has doubled to 25% share over seven years, now firmly in second place",
            "**Google Cloud** is the fastest-growing at 11%, having doubled its stake in six years",
            "**Combined Big Three** account for 66% of $79.8 billion Q1 2024 spending",
            "**Profitability inflection**: Google Cloud recently achieved profitability, joining AWS and Azure"
        ]
    }
    
    Notice: title is 31 chars (fits easily), subhead carries the headline
    insight, chapter_label provides navigation, bullets ground each claim.
"""
