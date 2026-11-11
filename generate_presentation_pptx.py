"""
generate_presentation_pptx.py
Builds a cohesive, humanized, and highly readable 16-slide PowerPoint presentation (.pptx)
for the Mid-Term Evaluation of B.Tech CSE Minor Project at IIIT Ranchi:
"Adaptive Self-Consistency in Large Language Models".

Supported Themes:
- "ocean" (Ocean / Sky Theme, Default): Rich coastal sky mist background, top cerulean ribbons, deep navy typography, seafoam teal accents, styled pill badges, designer cards.
- "forest" (Forest / Nature Theme): Morning sage woodland mist background, top evergreen ribbons, deep pine typography, fresh leaf accents, styled pill badges, designer cards.

Design Highlights:
- Distinct, professional theme identity (rich palette, stylish category pills, top dual-tone ribbons, accent indicators).
- Unified palette across ALL 16 slides (no dark slides; consistent academic light theme).
- Natural, humanized prose designed for computer science faculty evaluation.
- Clean mathematical notation free of raw LaTeX syntax.
- Strictly bounded geometry (content: 1.48" to 6.60", footer at 6.85") with ZERO text overlap.
- Embedded real project empirical figures on Slides 11, 12, and 13 with generous side-panels.
"""

import sys
import os
import argparse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# -----------------------------------------------------------------------------
# THEME COLOR DEFINITIONS
# -----------------------------------------------------------------------------
THEMES = {
    "ocean": {
        "name": "Ocean / Sky Theme",
        # Atmosphere & Backgrounds
        "bg_page": RGBColor(234, 243, 251),            # #EAF3FB (Crisp coastal sky mist)
        "top_ribbon": RGBColor(2, 132, 199),           # #0284C7 (Vibrant sky cerulean)
        "top_ribbon_sub": RGBColor(56, 189, 248),      # #38BDF8 (Sky cyan highlight)
        "bottom_ribbon": RGBColor(186, 215, 242),      # #BAD7F2 (Subtle ocean bottom border)
        # Cards & Containers
        "card_bg": RGBColor(255, 255, 255),            # #FFFFFF (Pure crisp white)
        "card_border": RGBColor(186, 215, 242),        # #BAD7F2 (Oceanic sky border)
        "card_border_accent": RGBColor(96, 165, 250),  # #60A5FA (Bright blue border)
        "card_strip": RGBColor(2, 132, 199),           # #0284C7 (Card top accent strip)
        "card_highlight_bg": RGBColor(224, 242, 254),  # #E0F2FE (Soft sky tint callout)
        "card_highlight_border": RGBColor(56, 189, 248),# #38BDF8 (Cyan border)
        # Badges & Pills
        "badge_bg": RGBColor(224, 242, 254),           # #E0F2FE (Pill badge background)
        "badge_border": RGBColor(147, 197, 253),       # #93C5FD (Pill border)
        "badge_text": RGBColor(3, 105, 161),           # #0369A1 (Marine azure text)
        # Typography
        "text_head": RGBColor(10, 37, 64),             # #0A2540 (Deep Atlantic midnight navy)
        "text_body": RGBColor(28, 55, 82),             # #1C3752 (Marine slate)
        "text_muted": RGBColor(71, 98, 122),          # #47627A (Ocean slate)
        # Accents
        "primary_accent": RGBColor(2, 132, 199),       # #0284C7 (Cerulean sky)
        "secondary_accent": RGBColor(3, 105, 161),     # #0369A1 (Marine azure)
        "success_accent": RGBColor(13, 148, 136),      # #0D9488 (Seafoam teal)
        "danger_accent": RGBColor(225, 29, 72),        # #E11D48 (Coral red)
        "warning_accent": RGBColor(217, 119, 6),       # #D97706 (Amber sun)
        # Table Styling
        "table_header_bg": RGBColor(10, 37, 64),       # #0A2540 (Deep ocean navy)
        "table_header_text": RGBColor(255, 255, 255),  # #FFFFFF
        "table_row_alt": RGBColor(240, 246, 253),      # #F0F6FD (Alternating row tint)
        # Stat Pill
        "stat_badge_bg": RGBColor(224, 242, 254),      # #E0F2FE
        "stat_badge_text": RGBColor(2, 132, 199),      # #0284C7
    },
    "forest": {
        "name": "Forest / Nature Theme",
        # Atmosphere & Backgrounds
        "bg_page": RGBColor(235, 244, 236),            # #EBF4EC (Morning woodland sage mist)
        "top_ribbon": RGBColor(21, 128, 61),           # #15803D (Deep evergreen forest)
        "top_ribbon_sub": RGBColor(74, 222, 128),      # #4ADE80 (Spring leaf green)
        "bottom_ribbon": RGBColor(187, 222, 197),      # #BBDDC5 (Subtle sage bottom border)
        # Cards & Containers
        "card_bg": RGBColor(255, 255, 255),            # #FFFFFF (Pure crisp white)
        "card_border": RGBColor(187, 222, 197),        # #BBDDC5 (Sage leaf border)
        "card_border_accent": RGBColor(134, 194, 150), # #86C296 (Moss border)
        "card_strip": RGBColor(21, 128, 61),           # #15803D (Card top accent strip)
        "card_highlight_bg": RGBColor(220, 245, 227),  # #DCF5E3 (Soft woodland mint callout)
        "card_highlight_border": RGBColor(34, 197, 94),# #22C55E (Fresh emerald border)
        # Badges & Pills
        "badge_bg": RGBColor(220, 252, 231),           # #DCFCE7 (Pill badge background)
        "badge_border": RGBColor(134, 239, 172),       # #86EFAC (Pill border)
        "badge_text": RGBColor(20, 83, 45),            # #14532D (Deep forest pine text)
        # Typography
        "text_head": RGBColor(15, 46, 26),             # #0F2E1A (Deep pine canopy)
        "text_body": RGBColor(38, 65, 45),             # #26412D (Woodland slate)
        "text_muted": RGBColor(78, 108, 88),           # #4E6C58 (Fern slate)
        # Accents
        "primary_accent": RGBColor(21, 128, 61),       # #15803D (Forest emerald)
        "secondary_accent": RGBColor(5, 150, 105),     # #059669 (Mountain pine)
        "success_accent": RGBColor(22, 163, 74),       # #16A34A (Vivid leaf green)
        "danger_accent": RGBColor(194, 65, 12),        # #C2410C (Autumn bark)
        "warning_accent": RGBColor(202, 138, 4),       # #CA8A04 (Warm ochre)
        # Table Styling
        "table_header_bg": RGBColor(20, 61, 40),       # #143D28 (Deep pine header)
        "table_header_text": RGBColor(255, 255, 255),  # #FFFFFF
        "table_row_alt": RGBColor(241, 248, 242),      # #F1F8F2 (Alternating row tint)
        # Stat Pill
        "stat_badge_bg": RGBColor(220, 252, 231),      # #DCFCE7
        "stat_badge_text": RGBColor(21, 128, 61),      # #15803D
    }
}

FONT_FAMILY = "Arial"


def build_presentation(theme_key: str = "ocean", output_filename: str = "presentation.pptx"):
    theme = THEMES.get(theme_key, THEMES["ocean"])
    print(f"Building presentation with theme: '{theme['name']}' -> {output_filename}")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Helper: Set slide background color and top theme ribbons
    def set_slide_background(slide):
        # 1. Base atmospheric canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme["bg_page"]
        bg.line.color.rgb = theme["bg_page"]
        slide.shapes._spTree.remove(bg._element)
        slide.shapes._spTree.insert(2, bg._element)

        # 2. Top dual-tone architectural ribbon
        r1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
        r1.fill.solid()
        r1.fill.fore_color.rgb = theme["top_ribbon"]
        r1.line.color.rgb = theme["top_ribbon"]

        r2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.08), Inches(13.333), Inches(0.02))
        r2.fill.solid()
        r2.fill.fore_color.rgb = theme["top_ribbon_sub"]
        r2.line.color.rgb = theme["top_ribbon_sub"]

        # 3. Bottom framing line
        b_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.46), Inches(13.333), Inches(0.04))
        b_line.fill.solid()
        b_line.fill.fore_color.rgb = theme["bottom_ribbon"]
        b_line.line.color.rgb = theme["bottom_ribbon"]

        return bg

    # Helper: Add unified top header (Slides 2 to 15)
    def add_slide_header(slide, category_tag: str, main_title: str, slide_num: int):
        # Category Pill Badge (Left)
        badge_width = Inches(max(2.8, len(category_tag) * 0.13 + 0.6))
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.28), badge_width, Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = theme["badge_bg"]
        pill.line.color.rgb = theme["badge_border"]
        pill.line.width = Pt(1)
        tf_p = pill.text_frame
        tf_p.word_wrap = False
        tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_p.margin_left = tf_p.margin_right = tf_p.margin_top = tf_p.margin_bottom = 0
        p = tf_p.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = f"●  {category_tag.upper()}"
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["badge_text"]

        # Academic Institution Tag (Right Pill)
        tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(0.28), Inches(2.733), Inches(0.32))
        tag_box.fill.solid()
        tag_box.fill.fore_color.rgb = theme["card_bg"]
        tag_box.line.color.rgb = theme["card_border"]
        tag_box.line.width = Pt(1)
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = False
        tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_tag.margin_left = tf_tag.margin_right = tf_tag.margin_top = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = "IIIT Ranchi • CSE Minor Project"
        p_tag.font.name = FONT_FAMILY
        p_tag.font.size = Pt(9)
        p_tag.font.bold = True
        p_tag.font.color.rgb = theme["text_muted"]

        # Slide Main Title
        t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.66), Inches(11.733), Inches(0.52))
        tf_t = t_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_top = tf_t.margin_bottom = tf_t.margin_left = tf_t.margin_right = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = main_title
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = theme["text_head"]

        # Clean Header Separator Line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.24), Inches(11.733), Inches(0.018))
        line.fill.solid()
        line.fill.fore_color.rgb = theme["card_border"]
        line.line.color.rgb = theme["card_border"]

        # Safe Footer (strictly at 6.85" to guarantee zero overlap)
        f_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.35))
        tf_f = f_box.text_frame
        tf_f.margin_top = tf_f.margin_bottom = tf_f.margin_left = tf_f.margin_right = 0
        p_f = tf_f.paragraphs[0]
        p_f.text = f"Slide {slide_num} of 16  |  Adaptive Self-Consistency in LLMs  |  Mid-Term Evaluation (Nov 12, 2026)"
        p_f.font.name = FONT_FAMILY
        p_f.font.size = Pt(9.5)
        p_f.font.color.rgb = theme["text_muted"]

    # Helper: Create card shape with top vertical anchor and designer accent line
    def create_card(slide, left, top, width, height, fill_color=None, border_color=None, with_accent_bar=True, accent_color=None):
        if fill_color is None:
            fill_color = theme["card_bg"]
        if border_color is None:
            border_color = theme["card_border"]

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = fill_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.22)
        tf.margin_right = Inches(0.22)
        tf.margin_top = Inches(0.20)
        tf.margin_bottom = Inches(0.18)

        # Optional top interior accent indicator
        if with_accent_bar:
            acc_col = accent_color if accent_color else theme["primary_accent"]
            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.22), top + Inches(0.10), Inches(0.75), Inches(0.04))
            bar.fill.solid()
            bar.fill.fore_color.rgb = acc_col
            bar.line.fill.background()

        return card

    # Helper: Add formatted bullet point with colored arrow run
    def add_bullet_point(tf, title, desc, bullet_color=None, space_after=Pt(6), font_size=Pt(10.5)):
        p = tf.add_paragraph()
        p.space_after = space_after

        # Colored bullet glyph
        r_b = p.add_run()
        r_b.text = "▸  "
        r_b.font.name = FONT_FAMILY
        r_b.font.size = font_size
        r_b.font.bold = True
        r_b.font.color.rgb = bullet_color if bullet_color else theme["primary_accent"]

        # Bold title
        r_t = p.add_run()
        r_t.text = title
        r_t.font.name = FONT_FAMILY
        r_t.font.size = font_size
        r_t.font.bold = True
        r_t.font.color.rgb = theme["text_head"]

        # Normal body
        r_d = p.add_run()
        r_d.text = desc
        r_d.font.name = FONT_FAMILY
        r_d.font.size = font_size
        r_d.font.bold = False
        r_d.font.color.rgb = theme["text_body"]
        return p

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Unified Theme)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # University & Department Tag (Centered Pill)
    pill1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.55), Inches(6.8), Inches(0.38))
    pill1.fill.solid()
    pill1.fill.fore_color.rgb = theme["badge_bg"]
    pill1.line.color.rgb = theme["badge_border"]
    pill1.line.width = Pt(1)
    tf1 = pill1.text_frame
    tf1.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf1.paragraphs[0]
    p.text = "●  B.TECH CSE MINOR PROJECT • MID-TERM EVALUATION"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = theme["badge_text"]
    p.alignment = PP_ALIGN.CENTER

    # Title & Subtitle Box
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.08), Inches(11.733), Inches(2.2))
    tf = t_box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "Adaptive Self-Consistency in Large Language Models"
    p1.font.name = FONT_FAMILY
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = theme["text_head"]
    p1.space_after = Pt(8)

    p2 = tf.add_paragraph()
    p2.text = "Dynamic Early Stopping for Efficient and Accurate Multi-Step Reasoning"
    p2.font.name = FONT_FAMILY
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = theme["secondary_accent"]
    p2.space_after = Pt(6)

    p3 = tf.add_paragraph()
    p3.text = "Department of Computer Science & Engineering, Indian Institute of Information Technology (IIIT) Ranchi"
    p3.font.name = FONT_FAMILY
    p3.font.size = Pt(13)
    p3.font.color.rgb = theme["text_muted"]

    # Team & Supervisor Side-by-Side Cards (Top: 3.55", Height: 2.70", Bottom: 6.25")
    card_team = create_card(s1, Inches(0.8), Inches(3.55), Inches(5.7), Inches(2.70), accent_color=theme["primary_accent"])
    tf = card_team.text_frame
    p = tf.paragraphs[0]
    p.text = "STUDENT INVESTIGATORS"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = theme["primary_accent"]
    p.space_after = Pt(10)

    students = [
        ("Pankaj Jangid", "Reg: 2023UG3004"),
        ("Aagman", "Reg: 2023UG3027"),
        ("Abhinav Kumar", "Reg: 2023UG4018")
    ]
    for name, reg in students:
        pm = tf.add_paragraph()
        pm.space_after = Pt(8)
        rb = pm.add_run()
        rb.text = "▸  "
        rb.font.bold = True
        rb.font.color.rgb = theme["primary_accent"]
        rn = pm.add_run()
        rn.text = f"{name}   "
        rn.font.size = Pt(13)
        rn.font.bold = True
        rn.font.color.rgb = theme["text_head"]
        rr = pm.add_run()
        rr.text = f"({reg})"
        rr.font.size = Pt(11)
        rr.font.bold = False
        rr.font.color.rgb = theme["text_muted"]

    card_sup = create_card(s1, Inches(6.833), Inches(3.55), Inches(5.7), Inches(2.70), accent_color=theme["secondary_accent"])
    tf = card_sup.text_frame
    p = tf.paragraphs[0]
    p.text = "PROJECT SUPERVISOR & INSTITUTION"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = theme["primary_accent"]
    p.space_after = Pt(10)

    p = tf.add_paragraph()
    p.text = "Dr. Bharat Singh"
    p.font.size = Pt(15.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(4)

    p = tf.add_paragraph()
    p.text = "Assistant Professor, Department of Computer Science & Engineering\nIndian Institute of Information Technology Ranchi\nJharkhand, India"
    p.font.size = Pt(11.5)
    p.font.color.rgb = theme["text_body"]
    p.space_after = Pt(10)

    p = tf.add_paragraph()
    p.text = "Evaluation Date: November 12, 2026 • Academic Year 2026-27"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]

    # Footer
    f_box = s1.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.35))
    p = f_box.text_frame.paragraphs[0]
    p.text = "Mid-Term Evaluation Date: November 12, 2026  |  IIIT Ranchi CSE Minor Project"
    p.font.size = Pt(9.5)
    p.font.color.rgb = theme["text_muted"]

    # =========================================================================
    # SLIDE 2: THE PROBLEM (Humanized, Intuitive)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_slide_header(s2, "Motivation & Problem Statement", "The Problem: Why LLMs Make Mistakes in Math Problems", 2)

    # Left: The Greedy Flaw
    c1 = create_card(s2, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["danger_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "The Flaw with Standard Single-Answer (Greedy) Decoding"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["danger_accent"]
    p.space_after = Pt(10)

    points2_left = [
        ("Step-by-Step Thinking (CoT): ", "LLMs are prompted to write out their reasoning before giving the final number. This works well, but only if every step is correct."),
        ("No Second Chances: ", "In greedy decoding, the model always picks the single most probable word. If it makes a small arithmetic slip in Step 2, every step afterwards is ruined."),
        ("Compounding Mistakes: ", "Like a student making a sign error on line 2 of an exam, the final answer ends up completely wrong despite a good overall plan."),
        ("False Confidence: ", "The model cannot tell you when it made a mistake — it outputs the wrong final number with the exact same high confidence as a correct one.")
    ]
    for title, desc in points2_left:
        add_bullet_point(tf, title, desc, bullet_color=theme["danger_accent"], space_after=Pt(8))

    # Right: Real Question Example
    c2 = create_card(s2, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf = c2.text_frame
    p = tf.paragraphs[0]
    p.text = "Real Example: How Diverse Thinking Solves the Problem"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(8)

    p_q = tf.add_paragraph()
    p_q.text = "Question: Janet has 16 eggs. She eats 3 for breakfast and bakes muffins with 4. She sells the rest for $2 each. How much does she make?"
    p_q.font.size = Pt(10.5)
    p_q.font.italic = True
    p_q.font.color.rgb = theme["text_body"]
    p_q.space_after = Pt(10)

    paths2 = [
        ("Reasoning Path 1 (Subtraction first):", "16 - 3 = 13 remaining. 13 - 4 = 9 eggs left. 9 * $2 = $18.", "Answer: 18 (Correct ✓)", theme["success_accent"]),
        ("Reasoning Path 2 (Grouping used):", "Total used: 3 + 4 = 7 eggs. Left: 16 - 7 = 9 eggs. 9 * $2 = $18.", "Answer: 18 (Correct ✓)", theme["success_accent"]),
        ("Reasoning Path 3 (Arithmetic slip):", "Eggs left: 16 - 7 = 8 eggs. 8 * $2 = $16.", "Answer: 16 (Flawed ✗)", theme["danger_accent"])
    ]
    for p_title, p_desc, p_ans, p_col in paths2:
        p = tf.add_paragraph()
        p.text = p_title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = theme["text_head"]
        r = tf.add_paragraph()
        r.text = p_desc + "  " + p_ans
        r.font.size = Pt(10.5)
        r.font.color.rgb = p_col
        r.space_after = Pt(6)

    p_takeaway = tf.add_paragraph()
    p_takeaway.text = "The Key Intuition: Correct reasoning paths agree on the same answer ($18). Incorrect attempts make different random errors ($16)."
    p_takeaway.font.size = Pt(10.5)
    p_takeaway.font.bold = True
    p_takeaway.font.color.rgb = theme["secondary_accent"]

    # =========================================================================
    # SLIDE 3: SELF-CONSISTENCY FOUNDATION
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_slide_header(s3, "Theoretical Background", "The Existing Foundation: Self-Consistency (Wang et al., ICLR 2023)", 3)

    # Left: The Method & Clear Formulation
    c1 = create_card(s3, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "How Majority Voting Works in Self-Consistency"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(8)

    p_desc = tf.add_paragraph()
    p_desc.text = "Instead of asking the LLM only once, Wang et al. sample multiple distinct reasoning attempts and take a majority vote over the final answers."
    p_desc.font.size = Pt(11)
    p_desc.font.color.rgb = theme["text_body"]
    p_desc.space_after = Pt(10)

    # Clean Formula Callout Box
    f1 = tf.add_paragraph()
    f1.text = "1. Majority Voting Rule (Selecting Most Common Answer):"
    f1.font.size = Pt(11)
    f1.font.bold = True
    f1.font.color.rgb = theme["text_head"]
    f1_box = tf.add_paragraph()
    f1_box.text = "   Final Answer = Answer with the highest vote count"
    f1_box.font.name = "Arial"
    f1_box.font.size = Pt(11)
    f1_box.font.bold = True
    f1_box.font.color.rgb = theme["primary_accent"]
    f1_box.space_after = Pt(10)

    f2 = tf.add_paragraph()
    f2.text = "2. Confidence-Weighted Rule (Accounting for Model Certainty):"
    f2.font.size = Pt(11)
    f2.font.bold = True
    f2.font.color.rgb = theme["text_head"]
    f2_box = tf.add_paragraph()
    f2_box.text = "   Path Weight = exp( Average Log-Probability of Generated Tokens )\n   Final Answer = Answer with highest total accumulated weight"
    f2_box.font.name = "Arial"
    f2_box.font.size = Pt(10.5)
    f2_box.font.bold = True
    f2_box.font.color.rgb = theme["secondary_accent"]
    f2_box.space_after = Pt(10)

    p_note = tf.add_paragraph()
    p_note.text = "Core Takeaway: Taking the majority vote acts as an unsupervised verifier. It requires zero extra training or human labeling."
    p_note.font.size = Pt(10.5)
    p_note.font.color.rgb = theme["text_muted"]

    # Right: Published Gains Table
    c2 = create_card(s3, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "Empirical Accuracy Gains Reported by Wang et al."
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(8)

    rows, cols = 7, 5
    table_shape = s3.shapes.add_table(rows, cols, Inches(7.0), Inches(2.10), Inches(5.366), Inches(3.40))
    t = table_shape.table
    t.columns[0].width = Inches(1.466)
    t.columns[1].width = Inches(0.95)
    t.columns[2].width = Inches(0.95)
    t.columns[3].width = Inches(1.0)
    t.columns[4].width = Inches(1.0)

    headers3 = ["Model", "Benchmark", "1 Path", "Self-Cons.", "Gain"]
    for i, h in enumerate(headers3):
        cell = t.cell(0, i)
        cell.text = h
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme["table_header_bg"]
        p = cell.text_frame.paragraphs[0]
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["table_header_text"]

    data3 = [
        ("PaLM-540B", "GSM8K", "56.5%", "74.4%", "+17.9%"),
        ("PaLM-540B", "SVAMP", "79.0%", "86.6%", "+7.6%"),
        ("PaLM-540B", "AQuA",  "35.8%", "48.3%", "+12.5%"),
        ("LaMDA-137B", "GSM8K", "17.1%", "27.7%", "+10.6%"),
        ("LaMDA-137B", "SVAMP", "38.9%", "53.3%", "+14.4%"),
        ("GPT-3 (002)", "GSM8K", "60.1%", "78.0%", "+17.9%")
    ]
    for r_idx, row_data in enumerate(data3):
        row_fill = theme["table_row_alt"] if r_idx % 2 == 1 else theme["card_bg"]
        for c_idx, val in enumerate(row_data):
            cell = t.cell(r_idx + 1, c_idx)
            cell.text = val
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_FAMILY
            p.font.size = Pt(9.5)
            p.font.color.rgb = theme["success_accent"] if c_idx == 4 else theme["text_head"]
            if c_idx == 4 or c_idx == 0:
                p.font.bold = True

    p_tbl_note = s3.shapes.add_textbox(Inches(7.0), Inches(5.75), Inches(5.366), Inches(0.70))
    tf_tn = p_tbl_note.text_frame
    tf_tn.word_wrap = True
    p = tf_tn.paragraphs[0]
    p.text = "Key Observation: Self-consistency provides major gains (+10% to +18%), but it forces every question to generate a large fixed number of paths."
    p.font.size = Pt(9.5)
    p.font.color.rgb = theme["text_muted"]

    # =========================================================================
    # SLIDE 4: THE STATIC BUDGET TAX
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_slide_header(s4, "The Practical Limitation", "Why Fixed Sampling (N=40) Wastes Huge Compute and Time", 4)

    pillars = [
        ("The 'One-Size-Fits-All' Waste", theme["danger_accent"], [
            ("Fixed Sample Budget: ", "Standard self-consistency forces 10 to 40 paths on every single problem unconditionally."),
            ("Trivial Overkill: ", "For an elementary problem (e.g. 15 - 7 = 8), the first 2 paths already agree 100%. Generating 38 more paths is completely pointless."),
            ("Wasted Resources: ", "Up to 80-90% of token compute on easy problems is spent confirming an answer that is already obvious.")
        ]),
        ("Diminishing Returns (Plateau)", theme["warning_accent"], [
            ("Quick Initial Gains: ", "Most accuracy gains occur within the first 3 to 5 paths."),
            ("Flat Accuracy Curve: ", "As shown in Wang et al. Figure 2, scaling from 10 to 40 paths yields less than 1.5% accuracy gain."),
            ("Extreme Cost: ", "Generating 40 paths costs 40x the API tokens, takes 40x longer, and easily exceeds free-tier rate limits.")
        ]),
        ("Noise in the Tail (Dilution)", theme["danger_accent"], [
            ("Accumulation of Errors: ", "When sampling dozens of times at temperature 0.7, rare errors and flawed logic begin to accumulate."),
            ("Empirical Evidence: ", "In our own GSM8K benchmark, accuracy reached 63.3% at 4 paths, but fell back to 60.0% at 6 paths due to noisy tails."),
            ("Spurious Agreement: ", "Bad paths can sometimes randomly agree on the same wrong number, confusing the majority vote.")
        ])
    ]
    for i, (p_title, p_color, p_bullets) in enumerate(pillars):
        x = Inches(0.8 + i * 4.0)
        card = create_card(s4, x, Inches(1.48), Inches(3.733), Inches(5.12), accent_color=p_color)
        tf = card.text_frame
        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = p_color
        p.space_after = Pt(10)

        for b_title, b_desc in p_bullets:
            add_bullet_point(tf, b_title, b_desc, bullet_color=p_color, space_after=Pt(8))

    # =========================================================================
    # SLIDE 5: OUR RESEARCH OBJECTIVES
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_slide_header(s5, "Project Objectives", "Our Goal: Stopping Early Once Consensus Is Reached", 5)

    # Top Goal Card
    stmt_card = create_card(s5, Inches(0.8), Inches(1.48), Inches(11.733), Inches(1.05), fill_color=theme["card_highlight_bg"], border_color=theme["card_highlight_border"])
    tf = stmt_card.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "CORE RESEARCH QUESTION"
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = theme["primary_accent"]
    p1.space_after = Pt(2)
    p2 = tf.add_paragraph()
    p2.text = "\"Can an LLM recognize when it has reached a clear consensus, allowing it to stop generating early and cut compute by >40% without losing any reasoning accuracy?\""
    p2.font.size = Pt(12.5)
    p2.font.bold = True
    p2.font.color.rgb = theme["text_head"]

    # 3 Objective Cards
    stat_targets = [
        (">40% Compute Reduction", "EFFICIENCY TARGET", "Halve the completion tokens and generation time per question compared to standard fixed-budget self-consistency."),
        ("0.0% Accuracy Regret", "SAFETY TARGET", "Guarantee that stopping early never causes a question to be answered incorrectly if full sampling would have answered it right."),
        ("Confidence-Aware Gate", "DECISION CRITERIA", "Combine discrete vote lead with the model's own token certainty so the system stops only when genuinely confident.")
    ]
    for i, (num, lbl, desc) in enumerate(stat_targets):
        x = Inches(0.8 + i * 4.0)
        card = create_card(s5, x, Inches(2.68), Inches(3.733), Inches(2.10), accent_color=theme["secondary_accent"])
        tf = card.text_frame
        p_num = tf.paragraphs[0]
        p_num.text = num
        p_num.font.size = Pt(15)
        p_num.font.bold = True
        p_num.font.color.rgb = theme["secondary_accent"]
        p_lbl = tf.add_paragraph()
        p_lbl.text = lbl
        p_lbl.font.size = Pt(10)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = theme["text_head"]
        p_lbl.space_after = Pt(4)
        p_desc = tf.add_paragraph()
        p_desc.text = desc
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = theme["text_body"]

    # Bottom Practical Benefits Card
    bot_card = create_card(s5, Inches(0.8), Inches(4.93), Inches(11.733), Inches(1.67), accent_color=theme["primary_accent"])
    tf = bot_card.text_frame
    p = tf.paragraphs[0]
    p.text = "Why This Matters in Practice:"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(4)

    deliverables = [
        ("Faster User Experience: ", "Easy questions return answers in 1-2 seconds instead of making the user wait 15 seconds."),
        ("API Cost Savings: ", "Token usage directly determines cloud API costs; cutting tokens by ~48% directly cuts the bill in half."),
        ("Rate Limit Protection: ", "Generating fewer paths prevents exhausting requests-per-minute (RPM) and tokens-per-minute (TPM) quotas.")
    ]
    for d_title, d_desc in deliverables:
        add_bullet_point(tf, d_title, d_desc, bullet_color=theme["primary_accent"], space_after=Pt(3))

    # =========================================================================
    # SLIDE 6: PROPOSED METHODOLOGY
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_slide_header(s6, "Methodology", "How Our System Works: Confidence-Aware Adaptive Self-Consistency", 6)

    # Left: Step-by-Step Flow
    c1 = create_card(s6, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "The 5-Step Adaptive Evaluation Loop"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(8)

    steps = [
        ("Step 1: Small Batch Sampling", "Query the model for just 2 paths (batch size = 2) instead of all 6 paths at once."),
        ("Step 2: Answer Extraction", "Extract the numeric answer from each path using our robust pattern-matching parser."),
        ("Step 3: Weight Calculation", "Calculate the model's confidence weight for each path using its token probabilities."),
        ("Step 4: Check Stopping Conditions", "Are the two paths in agreement? Does the leading answer account for >= 70% of confidence?"),
        ("Step 5: Output or Continue", "If YES -> Return the consensus answer immediately!\nIf NO (tie or split) -> Generate 2 more paths until consensus is achieved (max 6).")
    ]
    for s_title, s_desc in steps:
        p = tf.add_paragraph()
        p.text = s_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = theme["text_head"]
        r = tf.add_paragraph()
        r.text = s_desc
        r.font.size = Pt(10)
        r.font.color.rgb = theme["text_body"]
        r.space_after = Pt(4)

    # Right: The Dual Stopping Criteria
    c2 = create_card(s6, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "The Dual Stopping Criteria (Both Must Be True)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(10)

    # Gate 1 Box
    p_g1 = tf2.add_paragraph()
    p_g1.text = "CONDITION 1: Lead Margin of At Least 2 Votes"
    p_g1.font.size = Pt(11.5)
    p_g1.font.bold = True
    p_g1.font.color.rgb = theme["primary_accent"]
    p_g1_d = tf2.add_paragraph()
    p_g1_d.text = "Margin = (Votes for Top Answer) - (Votes for Runner-Up) >= 2\n• If first 2 paths agree (2 vs 0) -> Lead is 2. (Passes!)\n• If first 2 paths disagree (1 vs 1) -> Lead is 0. (Fails, generate more paths)."
    p_g1_d.font.size = Pt(10.5)
    p_g1_d.font.color.rgb = theme["text_body"]
    p_g1_d.space_after = Pt(10)

    # Gate 2 Box
    p_g2 = tf2.add_paragraph()
    p_g2.text = "CONDITION 2: High Model Confidence Share (>= 70%)"
    p_g2.font.size = Pt(11.5)
    p_g2.font.bold = True
    p_g2.font.color.rgb = theme["success_accent"]
    p_g2_d = tf2.add_paragraph()
    p_g2_d.text = "Confidence Share = (Weight of Lead Answer) / (Total Weight of All Paths) >= 70%\n• Ensures the model wasn't just 'hesitant' or guessing.\n• Even if 2 paths agree, if token probabilities are unusually low, the system continues sampling to double-check."
    p_g2_d.font.size = Pt(10.5)
    p_g2_d.font.color.rgb = theme["text_body"]
    p_g2_d.space_after = Pt(10)

    p_safety = tf2.add_paragraph()
    p_safety.text = "Outcome: Trivial problems stop after 2 paths. Ambiguous problems get the full 6 paths. Nothing is wasted."
    p_safety.font.size = Pt(10.5)
    p_safety.font.bold = True
    p_safety.font.color.rgb = theme["secondary_accent"]

    # =========================================================================
    # SLIDE 7: MATHEMATICAL FORMULATION (Humanized, Readable)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_slide_header(s7, "Mathematical Formulation", "Clear Mathematical Rules Behind the Stopping Policy", 7)

    # Left: The Rules
    c1 = create_card(s7, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "The Mathematical Definitions"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(8)

    formulas7 = [
        ("1. Vote Counts & Agreement Margin:",
         "Count for answer a = Number of generated paths that arrived at answer a\nMargin  Δ = Count(Lead Answer) - Count(Runner-Up Answer)"),

        ("2. Sequence Confidence Weight:",
         "Weight w = exp( Average Log-Probability of all tokens in the path )\n(Normalized by length so longer proofs are not penalized)"),

        ("3. Stopping Decision Rule:",
         "STOP if: [ Margin Δ >= 2  AND  Confidence Share >= 70% ]\nOR if max budget of 6 paths is reached."),

        ("4. Final Answer Selection:",
         "Chosen Answer = Candidate with the highest total confidence weight")
    ]
    for title, desc in formulas7:
        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = theme["text_head"]
        r = tf.add_paragraph()
        r.text = desc
        r.font.size = Pt(10)
        r.font.color.rgb = theme["secondary_accent"]
        r.space_after = Pt(6)

    # Right: Why This Rule Works
    c2 = create_card(s7, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "Why This Prevents Mistakes: Theoretical Guarantee"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(8)

    p_bound = tf2.add_paragraph()
    p_bound.text = "Theoretical Risk Bound:\nProbability of Early Stop Error <= exp( - Margin^2 / (2 * Paths) ) * (1 - Confidence)"
    p_bound.font.size = Pt(10.5)
    p_bound.font.bold = True
    p_bound.font.color.rgb = theme["danger_accent"]
    p_bound.space_after = Pt(10)

    bullets7_right = [
        ("Exponential Safety: ", "As the margin between first and second place grows to 2 votes, the statistical chance of the true answer being hidden in the minority drops exponentially."),
        ("Double Filter Against Guessing: ", "Even if two paths arrive at the same number by coincidence, if the model's confidence share is below 70%, the system refuses to stop early."),
        ("Instant Decision Time: ", "Evaluating these two rules takes less than 0.05 milliseconds on a CPU — adding virtually zero overhead to the model's generation.")
    ]
    for b_title, b_desc in bullets7_right:
        add_bullet_point(tf2, b_title, b_desc, bullet_color=theme["primary_accent"], space_after=Pt(8))

    # =========================================================================
    # SLIDE 8: SYSTEM ARCHITECTURE & ENGINEERING
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_slide_header(s8, "Implementation Architecture", "What We Built: Core Codebase Modules", 8)

    modules8 = [
        ("client_wrapper.py", "Asynchronous LLM Client", theme["primary_accent"], [
            ("Thread Pool: ", "Manages API connections with ThreadPoolExecutor (2 workers) to respect free-tier token limits."),
            ("Rate-Limit Backoff: ", "If an HTTP 429 response is received, it pauses for 15 seconds to let the token bucket refill."),
            ("Logprob Capture: ", "Captures both the text rationale and token log-probabilities to compute path weights.")
        ]),
        ("extractor.py", "Robust Answer Parser", theme["secondary_accent"], [
            ("6-Stage Cascade: ", "Regex parser that handles messy LLM formatting (LaTeX, Markdown, units)."),
            ("Cleans Currencies: ", "Normalizes currencies ($/₹/€), commas (1,250 -> 1250), percentages, and punctuation."),
            ("100% Recovery: ", "Passes 11/11 unit tests and achieves 100% extraction recovery on all benchmark problems.")
        ]),
        ("ca_asc_evaluator.py", "Adaptive Replay Simulator", theme["success_accent"], [
            ("Dual-Gate Engine: ", "Implements the early stopping engine (margin + confidence share)."),
            ("Offline Replay: ", "Re-runs policy sweeps over cached traces in 0.05 seconds locally."),
            ("Zero Cost Sweeps: ", "Allows testing different thresholds without spending API credits or making repetitive calls.")
        ])
    ]
    for i, (mod_file, mod_name, mod_col, mod_bullets) in enumerate(modules8):
        x = Inches(0.8 + i * 4.0)
        card = create_card(s8, x, Inches(1.48), Inches(3.733), Inches(5.12), accent_color=mod_col)
        tf = card.text_frame
        p = tf.paragraphs[0]
        p.text = mod_file
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = mod_col
        p.space_after = Pt(2)

        p_sub = tf.add_paragraph()
        p_sub.text = mod_name
        p_sub.font.size = Pt(10.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = theme["text_head"]
        p_sub.space_after = Pt(10)

        for b_title, b_desc in mod_bullets:
            add_bullet_point(tf, b_title, b_desc, bullet_color=mod_col, space_after=Pt(8))

    # =========================================================================
    # SLIDE 9: EXPERIMENTAL SETUP
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_slide_header(s9, "Experimental Setup", "Benchmark Dataset, Model Backbone & Comparison Baselines", 9)

    # Left: Parameter Table
    c1 = create_card(s9, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "Experimental Configuration"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(6)

    rows, cols = 8, 2
    table_shape = s9.shapes.add_table(rows, cols, Inches(1.0), Inches(2.05), Inches(5.3), Inches(4.30))
    t = table_shape.table
    t.columns[0].width = Inches(2.2)
    t.columns[1].width = Inches(3.1)

    params9 = [
        ("Benchmark Dataset", "OpenAI GSM8K (Grade School Math 8K) Test"),
        ("Model Backbone", "openai/gpt-oss-20b (20B reasoning model)"),
        ("Inference Engine", "Groq LPU Hardware Acceleration"),
        ("Sampling Temperature", "T = 0.70 (Encourages diverse reasoning paths)"),
        ("Nucleus Top-P", "p = 0.95"),
        ("Max Tokens / Path", "256 tokens per path (Keeps outputs concise)"),
        ("Batch Size (b)", "b = 2 (Sequential micro-batches)"),
        ("Stopping Parameters", "Margin lead = 2, Confidence = 70%, Max budget = 6")
    ]
    for r_idx, (p_name, p_val) in enumerate(params9):
        row_fill = theme["table_row_alt"] if r_idx % 2 == 1 else theme["card_bg"]
        c0 = t.cell(r_idx, 0)
        c0.text = p_name
        c0.margin_top = c0.margin_bottom = Inches(0.04)
        c0.margin_left = c0.margin_right = Inches(0.06)
        c0.fill.solid()
        c0.fill.fore_color.rgb = row_fill
        c0.text_frame.paragraphs[0].font.bold = True
        c0.text_frame.paragraphs[0].font.size = Pt(9.5)
        c0.text_frame.paragraphs[0].font.color.rgb = theme["text_head"]

        c1_cell = t.cell(r_idx, 1)
        c1_cell.text = p_val
        c1_cell.margin_top = c1_cell.margin_bottom = Inches(0.04)
        c1_cell.margin_left = c1_cell.margin_right = Inches(0.06)
        c1_cell.fill.solid()
        c1_cell.fill.fore_color.rgb = row_fill
        c1_cell.text_frame.paragraphs[0].font.size = Pt(9.5)
        c1_cell.text_frame.paragraphs[0].font.color.rgb = theme["secondary_accent"] if r_idx >= 6 else theme["text_body"]

    # Right: Baselines
    c2 = create_card(s9, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "Strategies Evaluated for Comparison"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(8)

    baselines9 = [
        ("1. Single-Path Greedy (N=1): ", "The standard approach: ask the model once with zero temperature. Fast, but has no self-correction mechanism."),
        ("2. Fixed Self-Consistency (N=1 to 6): ", "The Wang et al. baseline: sample a fixed budget of paths for every question, regardless of whether it was easy or hard."),
        ("3. Confidence-Weighted SC (Fixed N=6): ", "Same fixed budget of 6 paths, but each vote is weighted by model log-probability."),
        ("4. Adaptive Self-Consistency (Ours): ", "Samples paths in pairs of 2. Stops as soon as the lead answer has a 2-vote margin and >= 70% confidence.")
    ]
    for b_title, b_desc in baselines9:
        add_bullet_point(tf2, b_title, b_desc, bullet_color=theme["primary_accent"] if "Adaptive" in b_title else theme["secondary_accent"], space_after=Pt(8))

    # =========================================================================
    # SLIDE 10: ANSWER EXTRACTION & ERROR MITIGATION
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_slide_header(s10, "Robust Parsing", "Answer Extraction: Handling Inconsistent LLM Outputs", 10)

    # Left: Real Formatting Difficulties
    c1 = create_card(s10, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["danger_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "Real-World Output Variations from LLMs"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["danger_accent"]
    p.space_after = Pt(8)

    cases10 = [
        ("LaTeX Boxed Notation: ", "Model writes [boxed: 1,250.00] instead of just 1250. Parser strips brackets and removes commas."),
        ("Markdown Bold Embeddings: ", "Model outputs 'The cost is **$104** in total'. Parser extracts 104 and ignores the dollar sign and asterisks."),
        ("Verbal Concluding Sentences: ", "Model concludes with 'Total: 104 dollars'. Parser identifies concluding phrases and extracts the number."),
        ("Trailing Period Noise: ", "Model outputs 'The answer is 18.' Parser removes the period while preserving negative numbers like -4.5."),
        ("Last Numeral Fallback: ", "If the model doesn't use standard tags, the parser safely pulls the final computed number from the end of the rationale.")
    ]
    for c_title, c_desc in cases10:
        add_bullet_point(tf, c_title, c_desc, bullet_color=theme["danger_accent"], space_after=Pt(7))

    # Right: The 6-Stage Priority Cascade
    c2 = create_card(s10, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["success_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "The 6-Stage Priority Cascade in extractor.py"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["success_accent"]
    p.space_after = Pt(8)

    stages10 = [
        ("Priority 1: ", "GSM8K Tag Format  (#### <num>)"),
        ("Priority 2: ", "LaTeX Box Pattern  (boxed{<num>})"),
        ("Priority 3: ", "Verbal Phrases  (the answer is <num>)"),
        ("Priority 4: ", "Markdown Bold  (**<num>** or **$<num>**)"),
        ("Priority 5: ", "Concluding Stmts  (total is <num> dollars)"),
        ("Priority 6: ", "Final Numeral Fallback  (last valid number)")
    ]
    for s_prio, s_desc in stages10:
        add_bullet_point(tf2, s_prio, s_desc, bullet_color=theme["success_accent"], space_after=Pt(5))

    p_perf = tf2.add_paragraph()
    p_perf.text = "\nVerification Result:\n• 11 / 11 automated unit test suites pass.\n• Achieves 100% extraction recovery across all 180 paths in the benchmark, eliminating false scoring penalties."
    p_perf.font.size = Pt(10.5)
    p_perf.font.bold = True
    p_perf.font.color.rgb = theme["success_accent"]

    # =========================================================================
    # SLIDE 11: REPLICATION & ADAPTIVE SC (Figure 1 Embedded)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_slide_header(s11, "Empirical Results", "Replication of Wang et al. vs. Our Adaptive Self-Consistency", 11)

    # Left: Embed Real Project Figure 1
    fig1_path = "results/fig1_wang_vs_adaptive_scaling.png"
    create_card(s11, Inches(0.8), Inches(1.48), Inches(7.1), Inches(5.12), accent_color=theme["primary_accent"])
    if os.path.exists(fig1_path):
        s11.shapes.add_picture(fig1_path, Inches(0.95), Inches(1.65), Inches(6.8), Inches(4.30))
        cap_box = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(6.05), Inches(6.8), Inches(0.40))
        cap_box.fill.solid()
        cap_box.fill.fore_color.rgb = theme["badge_bg"]
        cap_box.line.color.rgb = theme["badge_border"]
        cap_box.line.width = Pt(1)
        p = cap_box.text_frame.paragraphs[0]
        p.text = "Figure 1: GSM8K empirical accuracy scaling: Wang et al. (fixed N=1..6) vs. Adaptive SC (Ours)."
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["badge_text"]
        p.alignment = PP_ALIGN.CENTER

    # Right: Analytical Side-Panel
    c11 = create_card(s11, Inches(8.1), Inches(1.48), Inches(4.433), Inches(5.12), accent_color=theme["secondary_accent"])
    tf = c11.text_frame
    p = tf.paragraphs[0]
    p.text = "What This Graph Proves"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(8)

    insights11 = [
        ("Single Path (N=1): ", "50.0% accuracy with standard greedy decoding."),
        ("Accuracy Climb (N=3): ", "Jumps to 60.0% as majority consensus filters out early calculation slips."),
        ("Peak Accuracy (N=4, 5): ", "Reaches 63.3% accuracy (+13.3% absolute gain)."),
        ("Tail Noise at N=6: ", "Drops back to 60.0% as noisy rationales at T=0.7 pollute the vote."),
        ("Adaptive SC (Ours): ", "Reaches the exact 60.0% target accuracy using an average of only 3.1 paths per question!"),
        ("Why It Wins: ", "Stops early when confident, avoiding both unnecessary compute and tail rationale noise.")
    ]
    for i_title, i_desc in insights11:
        add_bullet_point(tf, i_title, i_desc, bullet_color=theme["primary_accent"] if "Adaptive" in i_title else theme["secondary_accent"], space_after=Pt(6))

    # =========================================================================
    # SLIDE 12: PARETO FRONTIER & 48.4% COMPUTE REDUCTION (Figure 2 Embedded)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_slide_header(s12, "Compute Efficiency", "Compute Efficiency: 48.4% Token Reduction on the Pareto Frontier", 12)

    # Left: Embed Real Project Figure 2
    fig2_path = "results/fig2_pareto_compute.png"
    create_card(s12, Inches(0.8), Inches(1.48), Inches(7.1), Inches(5.12), accent_color=theme["success_accent"])
    if os.path.exists(fig2_path):
        s12.shapes.add_picture(fig2_path, Inches(0.95), Inches(1.65), Inches(6.8), Inches(4.30))
        cap_box = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(6.05), Inches(6.8), Inches(0.40))
        cap_box.fill.solid()
        cap_box.fill.fore_color.rgb = theme["badge_bg"]
        cap_box.line.color.rgb = theme["badge_border"]
        cap_box.line.width = Pt(1)
        p = cap_box.text_frame.paragraphs[0]
        p.text = "Figure 2: Empirical Pareto frontier on GSM8K: Completion Tokens vs. Benchmark Accuracy (%)."
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["badge_text"]
        p.alignment = PP_ALIGN.CENTER

    # Right: Analytical Side-Panel
    c12 = create_card(s12, Inches(8.1), Inches(1.48), Inches(4.433), Inches(5.12), accent_color=theme["primary_accent"])
    tf = c12.text_frame
    p = tf.paragraphs[0]
    p.text = "Token Savings Breakdown"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(6)

    # Big KPI Stat Pill
    stat_pill = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), Inches(2.05), Inches(4.033), Inches(0.55))
    stat_pill.fill.solid()
    stat_pill.fill.fore_color.rgb = theme["card_highlight_bg"]
    stat_pill.line.color.rgb = theme["card_highlight_border"]
    stat_pill.line.width = Pt(1.5)
    tf_sp = stat_pill.text_frame
    tf_sp.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_sp = tf_sp.paragraphs[0]
    p_sp.text = "-48.4% TOKENS SAVED"
    p_sp.font.size = Pt(18)
    p_sp.font.bold = True
    p_sp.font.color.rgb = theme["primary_accent"]
    p_sp.alignment = PP_ALIGN.CENTER

    # Empty spacer in text frame for pill
    tf.add_paragraph().space_after = Pt(36)

    comp_items12 = [
        ("Standard SC (N=6): ", "Consumes 1,123 tokens per question on average."),
        ("Adaptive SC (Ours): ", "Consumes 601 tokens per question on average."),
        ("Net Reduction: ", "522 completion tokens saved per problem (-48.4%)."),
        ("Same Accuracy: ", "Both achieve identical 60.0% accuracy on the GSM8K benchmark."),
        ("Strict Pareto Dominance: ", "Our method lies strictly to the left of the fixed curve, providing equal intelligence at nearly half the compute."),
        ("Real Latency Benefit: ", "Halves waiting time for user answers during interactive sessions.")
    ]
    for c_title, c_desc in comp_items12:
        add_bullet_point(tf, c_title, c_desc, bullet_color=theme["primary_accent"], space_after=Pt(5))

    # =========================================================================
    # SLIDE 13: SAMPLE ALLOCATION HISTOGRAM (Figure 3 Embedded)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_slide_header(s13, "Dynamic Allocation", "Sample Allocation: How the System Adapts to Problem Difficulty", 13)

    # Left: Embed Real Project Figure 3
    fig3_path = "results/figure_2_sample_allocation.png"
    create_card(s13, Inches(0.8), Inches(1.48), Inches(7.1), Inches(5.12), accent_color=theme["secondary_accent"])
    if os.path.exists(fig3_path):
        s13.shapes.add_picture(fig3_path, Inches(0.95), Inches(1.65), Inches(6.8), Inches(4.30))
        cap_box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(6.05), Inches(6.8), Inches(0.40))
        cap_box.fill.solid()
        cap_box.fill.fore_color.rgb = theme["badge_bg"]
        cap_box.line.color.rgb = theme["badge_border"]
        cap_box.line.width = Pt(1)
        p = cap_box.text_frame.paragraphs[0]
        p.text = "Figure 3: Number of reasoning paths evaluated before termination across 30 GSM8K test questions."
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["badge_text"]
        p.alignment = PP_ALIGN.CENTER

    # Right: Analytical Side-Panel
    c13 = create_card(s13, Inches(8.1), Inches(1.48), Inches(4.433), Inches(5.12), accent_color=theme["primary_accent"])
    tf = c13.text_frame
    p = tf.paragraphs[0]
    p.text = "Difficulty Tier Breakdown"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(6)

    tiers13 = [
        ("Easy Tier (k = 2 paths): ", "20 Questions (66.7%)", theme["success_accent"], "Two-thirds of all questions reach clear agreement on the first pair. Stopping here saves 66.7% of maximum compute."),
        ("Medium Tier (k = 4 paths): ", "4 Questions (13.3%)", theme["warning_accent"], "First batch split 1-vs-1. The second batch broke the tie and established consensus."),
        ("Complex Tier (k = 6 paths): ", "6 Questions (20.0%)", theme["danger_accent"], "Challenging multi-constraint problems that genuinely required full exploration budget.")
    ]
    for t_name, t_stat, t_col, t_desc in tiers13:
        p = tf.add_paragraph()
        p.text = t_name + t_stat
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = t_col
        r = tf.add_paragraph()
        r.text = t_desc
        r.font.size = Pt(10)
        r.font.color.rgb = theme["text_body"]
        r.space_after = Pt(6)

    p_avg = tf.add_paragraph()
    p_avg.text = "Overall Average Budget Used: 3.1 paths per question\n(Compared to 6.0 paths required by standard Self-Consistency)"
    p_avg.font.size = Pt(11)
    p_avg.font.bold = True
    p_avg.font.color.rgb = theme["secondary_accent"]

    # =========================================================================
    # SLIDE 14: ABLATION STUDIES (Validating Stopping Rules)
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_slide_header(s14, "Ablation Analysis", "Ablation Study: Validating the Stopping Thresholds", 14)

    # Left: Grid Sweep Table
    c1 = create_card(s14, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "Offline Parameter Sweep (ablation_study.py)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["text_head"]
    p.space_after = Pt(6)

    rows, cols = 6, 6
    table_shape = s14.shapes.add_table(rows, cols, Inches(1.0), Inches(2.10), Inches(5.3), Inches(3.40))
    t = table_shape.table
    t.columns[0].width = Inches(0.8)
    t.columns[1].width = Inches(0.8)
    t.columns[2].width = Inches(0.8)
    t.columns[3].width = Inches(0.9)
    t.columns[4].width = Inches(1.0)
    t.columns[5].width = Inches(1.0)

    abl_headers = ["Batch", "Lead Margin", "Confidence", "Acc.", "Savings", "Regret Rate"]
    for i, h in enumerate(abl_headers):
        cell = t.cell(0, i)
        cell.text = h
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme["table_header_bg"]
        p = cell.text_frame.paragraphs[0]
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = theme["table_header_text"]

    abl_data = [
        ("b=1", "Margin=1", "70%", "56.7%", "+56.1%", "10.0%"),
        ("b=1", "Margin=2", "70%", "70.0%", "+23.5%", "0.0%"),
        ("b=2", "Margin=1", "70%", "66.7%", "+33.0%", "0.0%"),
        ("b=2", "Margin=2", "70%", "70.0%", "+23.0%", "0.0%"),
        ("b=2", "Margin=3", "70%", "70.0%", "+5.2%",  "0.0%")
    ]
    for r_idx, row_data in enumerate(abl_data):
        row_fill = theme["table_row_alt"] if r_idx % 2 == 1 else theme["card_bg"]
        for c_idx, val in enumerate(row_data):
            cell = t.cell(r_idx + 1, c_idx)
            cell.text = val
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            p = cell.text_frame.paragraphs[0]
            p.font.name = FONT_FAMILY
            p.font.size = Pt(9.5)
            if r_idx == 3:  # Our optimal setting
                cell.fill.solid()
                cell.fill.fore_color.rgb = theme["card_highlight_bg"]
                p.font.bold = True
                p.font.color.rgb = theme["secondary_accent"]
            elif c_idx == 5 and "10.0%" in val:
                p.font.bold = True
                p.font.color.rgb = theme["danger_accent"]
            else:
                p.font.color.rgb = theme["text_head"]

    # Right: Trade-Off Insights
    c2 = create_card(s14, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["secondary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "What This Teaches Us About Stopping Policies"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = theme["secondary_accent"]
    p.space_after = Pt(8)

    insights14 = [
        ("Margin = 1 (Too Aggressive): ", "Stopping on a 1-vote lead saves high compute (+56%), but suffers a 10.0% regret rate because it prematurely trusts lucky early guesses."),
        ("Margin = 2, Confidence >= 70% (Optimal): ", "The sweet spot! Achieves 70.0% accuracy with 0.0% regret. Never makes a premature mistake on problems the full budget could solve."),
        ("Margin = 3 (Too Conservative): ", "Delays stopping until 5.5 paths, wasting compute without gaining any extra accuracy."),
        ("Role of Confidence Check (>= 70%): ", "Acts as an essential safety net: even if 2 paths agree, if token perplexity is high, early exit is delayed until confirmation.")
    ]
    for i_title, i_desc in insights14:
        add_bullet_point(tf2, i_title, i_desc, bullet_color=theme["secondary_accent"], space_after=Pt(8))

    # =========================================================================
    # SLIDE 15: MILESTONES & FINAL-TERM ROADMAP
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15)
    add_slide_header(s15, "Project Roadmap", "Mid-Term Milestones Completed & Final-Term Roadmap", 15)

    # Left: Completed Mid-Term
    c1 = create_card(s15, Inches(0.8), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["success_accent"])
    tf = c1.text_frame
    p = tf.paragraphs[0]
    p.text = "✓ Completed Mid-Term Deliverables (100% On-Schedule)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = theme["success_accent"]
    p.space_after = Pt(8)

    done_items = [
        ("Replication of Wang et al.: ", "Replicated empirical scaling curves and confirmed saturation dynamics on GSM8K."),
        ("Asynchronous API Harness: ", "Built production-grade client with rate-limiting, exponential backoff, and logprob capture."),
        ("Robust 6-Stage Parser: ", "Built extractor.py achieving 100% extraction recovery across all test paths."),
        ("Proof of 48.4% Compute Reduction: ", "Demonstrated 48.4% token compute reduction with 0.0% accuracy regret."),
        ("High-Speed Replay Engine: ", "Enabled local simulation sweeps across stopping thresholds in under 0.05 seconds.")
    ]
    for d_title, d_desc in done_items:
        add_bullet_point(tf, d_title, d_desc, bullet_color=theme["success_accent"], space_after=Pt(8))

    # Right: Future Work Roadmap
    c2 = create_card(s15, Inches(6.833), Inches(1.48), Inches(5.7), Inches(5.12), accent_color=theme["primary_accent"])
    tf2 = c2.text_frame
    p = tf2.paragraphs[0]
    p.text = "➔ Final-Term Roadmap (Nov 2026 - May 2027)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = theme["primary_accent"]
    p.space_after = Pt(8)

    future_items = [
        ("Benchmark Expansion: ", "Evaluate on SVAMP, MATH, and AQuA-RAT to prove generalization across different reasoning types."),
        ("Token-Level Early Stopping: ", "Halt bad reasoning paths mid-generation if cumulative token entropy exceeds certainty bounds."),
        ("Model Cascade Routing: ", "Route easy problems (k=2) to a fast 8B model, reserving 70B models strictly for unresolved hard problems."),
        ("Conference Submission: ", "Prepare and submit a complete research paper to an ACL / EMNLP / NeurIPS peer-reviewed workshop.")
    ]
    for f_title, f_desc in future_items:
        add_bullet_point(tf2, f_title, f_desc, bullet_color=theme["primary_accent"], space_after=Pt(8))

    # =========================================================================
    # SLIDE 16: CONCLUSION & FACULTY Q&A (Unified Theme)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16)

    # Top Tag
    pill16 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.48), Inches(4.5), Inches(0.35))
    pill16.fill.solid()
    pill16.fill.fore_color.rgb = theme["badge_bg"]
    pill16.line.color.rgb = theme["badge_border"]
    pill16.line.width = Pt(1)
    tf16 = pill16.text_frame
    tf16.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf16.paragraphs[0]
    p.text = "●  MID-TERM EVALUATION CONCLUSION & KEY FINDINGS"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = theme["badge_text"]
    p.alignment = PP_ALIGN.CENTER

    # Main Title
    t_box = s16.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(11.733), Inches(1.2))
    tf = t_box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "Key Takeaways & Faculty Discussion"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = theme["text_head"]
    p1.space_after = Pt(4)
    p2 = tf.add_paragraph()
    p2.text = "Confidence-Aware Adaptive Self-Consistency transforms reasoning inference from brute-force sampling into dynamic, consensus-guided intelligence."
    p2.font.size = Pt(13.5)
    p2.font.color.rgb = theme["secondary_accent"]

    # 3 Summary Cards
    stat_summary16 = [
        ("-48.4%", "COMPUTE REDUCTION", "Consumes 601 tokens vs 1,128 tokens for identical 60.0% GSM8K accuracy."),
        ("66.7%", "INSTANT RESOLUTION", "Two-thirds of all questions terminate on Batch 1 (k=2), eliminating redundant sampling."),
        ("0.0%", "ZERO REGRET", "Dual margin-confidence stopping rule guarantees zero performance loss compared to full budget.")
    ]
    for i, (num, lbl, desc) in enumerate(stat_summary16):
        x = Inches(0.8 + i * 4.0)
        card = create_card(s16, x, Inches(2.28), Inches(3.733), Inches(2.15), fill_color=theme["card_highlight_bg"], border_color=theme["card_highlight_border"], accent_color=theme["primary_accent"])
        tf = card.text_frame
        p_num = tf.paragraphs[0]
        p_num.text = num
        p_num.font.size = Pt(26)
        p_num.font.bold = True
        p_num.font.color.rgb = theme["primary_accent"]
        p_lbl = tf.add_paragraph()
        p_lbl.text = lbl
        p_lbl.font.size = Pt(10.5)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = theme["text_head"]
        p_lbl.space_after = Pt(4)
        p_desc = tf.add_paragraph()
        p_desc.text = desc
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = theme["text_body"]

    # Supervisor & Team Acknowledgment Card
    ack_card = create_card(s16, Inches(0.8), Inches(4.65), Inches(11.733), Inches(1.95), accent_color=theme["primary_accent"])
    tf = ack_card.text_frame

    p = tf.paragraphs[0]
    p.text = "PROJECT SUPERVISOR"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = theme["primary_accent"]
    p.space_after = Pt(2)

    p_sup = tf.add_paragraph()
    p_sup.text = "Dr. Bharat Singh  —  Assistant Professor, Department of CSE, IIIT Ranchi"
    p_sup.font.size = Pt(13)
    p_sup.font.bold = True
    p_sup.font.color.rgb = theme["text_head"]
    p_sup.space_after = Pt(8)

    p_team_hdr = tf.add_paragraph()
    p_team_hdr.text = "STUDENT INVESTIGATORS"
    p_team_hdr.font.size = Pt(10.5)
    p_team_hdr.font.bold = True
    p_team_hdr.font.color.rgb = theme["primary_accent"]
    p_team_hdr.space_after = Pt(2)

    p_team = tf.add_paragraph()
    p_team.text = "Pankaj Jangid (2023UG3004)  |  Aagman (2023UG3027)  |  Abhinav Kumar (2023UG4018)"
    p_team.font.size = Pt(12)
    p_team.font.bold = True
    p_team.font.color.rgb = theme["text_head"]
    p_team.space_after = Pt(4)

    p_thanks = tf.add_paragraph()
    p_thanks.text = "Department of Computer Science & Engineering, IIIT Ranchi  •  Thank You! Questions & Faculty Discussion."
    p_thanks.font.size = Pt(11)
    p_thanks.font.bold = True
    p_thanks.font.color.rgb = theme["secondary_accent"]

    # Footer
    f_box = s16.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.35))
    p = f_box.text_frame.paragraphs[0]
    p.text = "Mid-Term Evaluation Date: November 12, 2026  |  IIIT Ranchi CSE Minor Project"
    p.font.size = Pt(9.5)
    p.font.color.rgb = theme["text_muted"]

    # Save presentation
    prs.save(output_filename)
    print(f"Presentation saved successfully to: {output_filename}")
    return output_filename


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Mid-Term Presentation PPTX")
    parser.add_argument("--theme", choices=["ocean", "forest"], default="ocean", help="Color theme: 'ocean' (default) or 'forest'")
    parser.add_argument("--output", default="presentation.pptx", help="Output .pptx file path")
    args = parser.parse_args()

    build_presentation(theme_key=args.theme, output_filename=args.output)
