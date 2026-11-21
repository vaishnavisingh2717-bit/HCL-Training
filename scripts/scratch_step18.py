"""
scratch_step18.py - Generates the official ABES 11-Slide Capstone Presentation Deck
Output: docs/2503201450157_Presentation.pptx (and mirrored copies)
Strictly adheres to screens 1, 2, and 3 styling, dimensions, and speaker notes.
"""

import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

WORKSPACE_DIR = r"c:\HCL_Project"
DOCS_DIR = os.path.join(WORKSPACE_DIR, "docs")
TARGET_PPTX = os.path.join(DOCS_DIR, "2500320140160_VaishnaviSingh_Presentation.pptx")

# Defined Color Palette
COLOR_RED = RGBColor(0xA5, 0x1C, 0x30)
COLOR_BLACK = RGBColor(0x15, 0x15, 0x14)
COLOR_DARK = RGBColor(0x3A, 0x3A, 0x38)
COLOR_GREY = RGBColor(0x5E, 0x5E, 0x5B)
COLOR_CARD = RGBColor(0xF4, 0xF1, 0xF1)
COLOR_PINK = RGBColor(0xD9, 0xB5, 0xBC)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_MUTED = RGBColor(0x8A, 0x8A, 0x86)
COLOR_BORDER = RGBColor(0xD9, 0xD9, 0xD9)

def create_blank_slide(prs):
    blank_layout = prs.slide_layouts[6] # blank layout
    return prs.slides.add_slide(blank_layout)

def set_dark_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BLACK
    bg.line.fill.background()
    return bg

def add_header_footer(slide, title_text, slide_num):
    # Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    tf_title.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_title.paragraphs[0]
    p.text = title_text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Cambria"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLACK

    # Footer
    tb_footer = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_footer = tb_footer.text_frame
    tf_footer.word_wrap = True
    tf_footer.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_f = tf_footer.paragraphs[0]
    p_f.text = "HCLTech Career Shaper ABES Engineering College | CWMP-2026 - Channel Wise Marketing Performance Analysis"
    p_f.font.name = "Calibri"
    p_f.font.size = Pt(9)
    p_f.font.color.rgb = COLOR_GREY

    # Slide Number
    tb_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_num = tb_num.text_frame
    tf_num.word_wrap = True
    tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_n = tf_num.paragraphs[0]
    p_n.text = f"{slide_num}"
    p_n.alignment = PP_ALIGN.RIGHT
    p_n.font.name = "Calibri"
    p_n.font.size = Pt(9)
    p_n.font.color.rgb = COLOR_MUTED

def add_card(slide, x, y, w, h):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_CARD
    shape.line.color.rgb = COLOR_CARD
    shape.adjustments[0] = 0.08
    return shape

def add_badge(slide, x, y, text):
    badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_RED
    badge.line.fill.background()
    tf = badge.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    return badge

def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text):
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_text)

    # HEAD
    tb_head = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
    tf_head = tb_head.text_frame
    tf_head.word_wrap = True
    tf_head.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_head.margin_left = tf_head.margin_right = tf_head.margin_top = tf_head.margin_bottom = 0
    p_h = tf_head.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(15)
    p_h.font.bold = True
    p_h.font.color.rgb = COLOR_BLACK

    # BODY
    tb_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_body = tb_body.text_frame
    tf_body.word_wrap = True
    tf_body.vertical_anchor = MSO_ANCHOR.TOP
    tf_body.margin_left = tf_body.margin_right = tf_body.margin_top = tf_body.margin_bottom = 0
    p_b = tf_body.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(13)
    p_b.font.color.rgb = COLOR_DARK

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    add_card(slide, x, y, w, h)

    # HEAD
    tb_head = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
    tf_head = tb_head.text_frame
    tf_head.word_wrap = True
    tf_head.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_head.margin_left = tf_head.margin_right = tf_head.margin_top = tf_head.margin_bottom = 0
    p_h = tf_head.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(15)
    p_h.font.bold = True
    p_h.font.color.rgb = COLOR_BLACK

    # BODY
    tb_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_body = tb_body.text_frame
    tf_body.word_wrap = True
    tf_body.vertical_anchor = MSO_ANCHOR.TOP
    tf_body.margin_left = tf_body.margin_right = tf_body.margin_top = tf_body.margin_bottom = 0
    p_b = tf_body.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(13)
    p_b.font.color.rgb = COLOR_DARK

def set_speaker_note(slide, note_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = note_text

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # ==================== SLIDE 1: DARK COVER SLIDE ====================
    s1 = create_blank_slide(prs)
    set_dark_background(s1)

    # ABES Logo placeholder
    tb_logo = s1.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_logo = tb_logo.text_frame
    tf_logo.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_logo.paragraphs[0]
    p.text = "[ABES logo]"
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = COLOR_PINK

    # Institution Line
    tb_inst = s1.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_inst = tb_inst.text_frame
    tf_inst.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_inst.paragraphs[0]
    p.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_PINK

    # Title
    tb_t = s1.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    tf_t.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf_t.paragraphs[0]
    p.text = "CWMP-2026 - Channel Wise Marketing Performance Analysis"
    p.font.name = "Cambria"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    # Lines
    tb_lines = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = tb_lines.text_frame
    tf_lines.word_wrap = True
    tf_lines.vertical_anchor = MSO_ANCHOR.TOP

    p0 = tf_lines.paragraphs[0]
    p0.text = "Vaishnavi Singh | Roll No 2500320140160 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"
    p0.font.name = "Calibri"
    p0.font.size = Pt(14)
    p0.font.color.rgb = COLOR_CARD
    p0.space_before = Pt(4)

    p1 = tf_lines.add_paragraph()
    p1.text = "Project ID: P084 | Mentor: Mr. Dinesh"
    p1.font.name = "Calibri"
    p1.font.size = Pt(14)
    p1.font.color.rgb = COLOR_CARD
    p1.space_before = Pt(4)

    p2 = tf_lines.add_paragraph()
    p2.text = "21 November 2026"
    p2.font.name = "Calibri"
    p2.font.size = Pt(14)
    p2.font.color.rgb = COLOR_PINK
    p2.space_before = Pt(4)

    set_speaker_note(s1, "Good morning esteemed mentors and evaluators. I am Vaishnavi Singh (Roll No 2500320140160), MCA 2nd Year Data Analytics scholar at ABES Engineering College. Today, I am presenting my capstone project P084: Channel Wise Marketing Performance Analysis, developed under the mentorship of Mr. Dinesh. The purpose of this project is to eliminate marketing spend waste across fragmented digital channels using relational warehousing, statistical inference, and attribution modeling.")

    # ==================== SLIDE 2: PROBLEM AND HOW MIGHT WE ====================
    s2 = create_blank_slide(prs)
    add_header_footer(s2, "Problem and How Might We", 2)

    prob_text = (
        "Modern digital enterprises deploy advertising capital across multiple fragmented channels without unified attribution. "
        "Marketing teams suffer from severe spend inefficiency by overfunding bleeding campaigns with ROAS below 2.0x, "
        "relying on biased last-click attribution that undervalues discovery channels, and missing lucrative weekend conversion surges."
    )
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem statement", prob_text)

    hmw1_text = "How might we unify cross-channel telemetry for the CMO so executive decisions are made without waiting for weekly reports?"
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Unified Telemetry", hmw1_text)

    hmw2_text = "How might we auto-detect and flag underperforming campaigns for the Performance Lead so ad-waste is terminated within 24 hours?"
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Underperformer Alerts", hmw2_text)

    set_speaker_note(s2, "Our core problem statement centers on marketing capital waste. Enterprises lack centralized visibility, leading to campaigns operating at a loss, last-touch attribution blind spots, and unoptimized day-parting schedules. We translated these pain points into ten How Might We statements. HMW-01 focuses on unifying fragmented cross-channel telemetry into an instant executive dashboard, while HMW-02 focuses on automating kill-switch alerts for underperforming ad campaigns.")

    # ==================== SLIDE 3: STAKEHOLDERS AND USERS ====================
    s3 = create_blank_slide(prs)
    add_header_footer(s3, "Stakeholders and users", 3)

    sh1_body = "Executive Sponsor\nNeeds consolidated cross-channel ROI, blended CAC tracking, and board-level spend efficiency.\n\nInfluence: High"
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Chief Marketing Officer", sh1_body)

    sh2_body = "Operational Owner\nNeeds daily campaign ROAS, underperformer kill-switches, and conversion rate optimization.\n\nInfluence: High"
    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Performance Lead", sh2_body)

    sh3_body = "Evaluator & Mentor\nGuides technical architecture, statistical validity, and complete requirements traceability.\n\nInfluence: High"
    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Mentor & Evaluator", sh3_body)

    set_speaker_note(s3, "We designed our platform around three critical user groups. First is our primary user, the Chief Marketing Officer, who requires high-level blended CAC, net margins, and cross-channel revenue contribution. Second is our operational user, the Performance Marketing Lead, who acts on daily ROAS thresholds and campaign quadrant scatter plots. Third is our project mentor and evaluator, ensuring econometric rigor and full requirements traceability.")

    # ==================== SLIDE 4: REQUIREMENTS AND TRACEABILITY ====================
    s4 = create_blank_slide(prs)
    add_header_footer(s4, "Requirements and traceability", 4)

    # TABLE at (.5, 1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2
    table_shape = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    tbl = table_shape.table
    col_w = [Inches(1.0), Inches(3.4), Inches(1.0), Inches(1.2), Inches(1.2), Inches(1.2)]
    for idx, w in enumerate(col_w):
        tbl.columns[idx].width = w

    headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_RED
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    rtm_data = [
        ("HMW-01", "Cross-Channel Telemetry Ingestion", "Must", "US-01", "TC-F-01", "UI-01"),
        ("HMW-02", "Underperforming Campaign Alerts", "Must", "US-05", "TC-F-05", "UI-02"),
        ("HMW-05", "Weekend Surge Z-Test Hypothesis", "Must", "US-08", "TC-F-08", "UI-04"),
        ("HMW-04", "Multi-Touch Attribution Modeling", "Should", "US-10", "TC-F-10", "UI-05")
    ]
    for i, row in enumerate(rtm_data):
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_BLACK
            p.alignment = PP_ALIGN.CENTER if j in [0, 2, 3, 4, 5] else PP_ALIGN.LEFT

    # Note .3 below table
    tb_note = s4.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9.0), Inches(0.35))
    tf_note = tb_note.text_frame
    p_note = tf_note.paragraphs[0]
    p_note.text = "Full matrix: Project Document, section 4"
    p_note.font.name = "Calibri"
    p_note.font.size = Pt(12)
    p_note.font.italic = True
    p_note.font.color.rgb = COLOR_GREY

    set_speaker_note(s4, "Every single How Might We statement directly links to a formal Business Requirement, an Agile User Story, a QA Test Case, and an interactive UI Screen. For example, HMW-01 drives automated data ingestion in US-01, tested under TC-F-01, rendering Screen UI-01. Our full Requirements Traceability Matrix guarantees 100% testability across all 12 business requirements without orphaned code.")

    # ==================== SLIDE 5: DATA AND PREPARATION ====================
    s5 = create_blank_slide(prs)
    add_header_footer(s5, "Data and preparation", 5)

    stats = [
        ("262", "Daily Fact Rows"),
        ("38", "Analytical Columns"),
        ("180 Days", "Telemetry Period"),
        ("9 Issues", "Anomalies Resolved")
    ]
    xs = [0.5, 2.8, 5.1, 7.4]
    for idx, (val, lbl) in enumerate(stats):
        # Value
        tb_val = s5.shapes.add_textbox(Inches(xs[idx]), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_val = tb_val.text_frame
        tf_val.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_val.paragraphs[0]
        p.text = val
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Cambria"
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = COLOR_RED

        # Label
        tb_lbl = s5.shapes.add_textbox(Inches(xs[idx]), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_lbl = tb_lbl.text_frame
        tf_lbl.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_lbl.paragraphs[0]
        p.text = lbl
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_GREY

    source_text = "Synthetic multi-channel dataset (seed 42) in SQLite data/app.db. One row in the main fact table represents daily ad telemetry for a specific campaign, capturing impressions, clicks, spend, conversions, and realized revenue."
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source", source_text)

    cleaning_text = "Top 3 data quality fixes:\n1. Purged 2 duplicate daily fact logs.\n2. Imputed missing CTR/CPC metrics using channel medians.\n3. Flagged 7 spend anomalies via Tukey's IQR fences (> ₹10,480)."
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning", cleaning_text)

    set_speaker_note(s5, "Our data architecture operates on a normalized SQLite database containing 262 daily campaign logs and 38 analytical dimensions across a 180-day telemetry window. One row in the fact table captures daily impressions, clicks, ad spend, and realized revenue for a specific campaign. In our cleaning pipeline, we purged 2 duplicate logs, imputed missing CTR entries using channel medians, and tagged 7 spend anomalies using Tukey's IQR fences without discarding valuable records.")

    # ==================== SLIDE 6: SOLUTION AND TECHNOLOGY ====================
    s6 = create_blank_slide(prs)
    add_header_footer(s6, "Solution and technology", 6)

    steps = [
        ("Raw data\n(CSV / SQLite)", 0.5, False),
        ("Clean & KPIs\n(Python, pandas)", 2.35, False),
        ("Store\n(SQLite Schema)", 4.2, False),
        ("Dashboard\n(Streamlit)", 6.05, True),
        ("Deploy\n(GitHub -> Cloud)", 7.9, False)
    ]
    for text, x_pos, is_highlight in steps:
        step_box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_pos), Inches(1.5), Inches(1.6), Inches(1.2))
        step_box.fill.solid()
        step_box.fill.fore_color.rgb = COLOR_RED if is_highlight else COLOR_CARD
        step_box.line.color.rgb = COLOR_RED if is_highlight else COLOR_CARD
        step_box.adjustments[0] = 0.1
        tf = step_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = text
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE if is_highlight else COLOR_BLACK

    # Arrows between steps
    arrow_xs = [2.1, 3.95, 5.8, 7.65]
    for ax in arrow_xs:
        tb_arr = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_arr = tb_arr.text_frame
        tf_arr.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_arr.margin_left = tf_arr.margin_right = 0
        p = tf_arr.paragraphs[0]
        p.text = ">"
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = COLOR_RED

    # Tools line
    tb_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_tools = tb_tools.text_frame
    tf_tools.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_tools.paragraphs[0]
    p_t.text = "Tools: Python 3.12 · Streamlit · SQLite 3 · pandas · NumPy · SciPy · Plotly · Tailwind CSS · pytest · Git"
    p_t.font.name = "Calibri"
    p_t.font.size = Pt(13)
    p_t.font.color.rgb = COLOR_DARK

    # Run command line
    tb_run = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_run = tb_run.text_frame
    tf_run.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_r = tf_run.paragraphs[0]
    p_r.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p_r.font.name = "Calibri"
    p_r.font.size = Pt(13)
    p_r.font.italic = True
    p_r.font.color.rgb = COLOR_GREY

    set_speaker_note(s6, "Our technical solution flows through five decoupled stages: from raw CSV and SQLite ingestion, through Python cleaning and KPI math, into an embedded Star Schema database, surfaced via an executive Streamlit dashboard, and deployed to Streamlit Community Cloud. The entire stack relies on free, open-source technologies including Python 3.12, SQLite 3, Plotly, and isolated Tailwind CSS CDN iframes.")

    # ==================== SLIDE 7: DASHBOARD ====================
    s7 = create_blank_slide(prs)
    add_header_footer(s7, "Dashboard", 7)

    # Frame (.5, 1.25, 5.5, 3.7) fill F4F1F1, dashed 1 pt BFBFBF
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = COLOR_CARD
    frame.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
    frame.line.width = Pt(1)
    tf_frame = frame.text_frame
    tf_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_fr = tf_frame.paragraphs[0]
    p_fr.text = "[Paste a dashboard screenshot here]"
    p_fr.alignment = PP_ALIGN.CENTER
    p_fr.font.name = "Calibri"
    p_fr.font.size = Pt(14)
    p_fr.font.italic = True
    p_fr.font.color.rgb = COLOR_GREY

    # Badges 1-4 beside frame
    badge_data = [
        (1.3, "1", "UI-01", "Executive overview & KPI telemetry"),
        (2.2, "2", "UI-02", "Channel ROAS vs CAC scatter quadrants"),
        (3.1, "3", "UI-04", "Weekend surge Z-test & regression"),
        (4.0, "4", "UI-05", "Attribution comparison & budget simulator")
    ]
    for y_pos, b_num, ui_id, desc in badge_data:
        add_badge(s7, 6.7, y_pos, b_num)
        tb_d = s7.shapes.add_textbox(Inches(7.25), Inches(y_pos), Inches(2.25), Inches(0.42))
        tf_d = tb_d.text_frame
        tf_d.word_wrap = True
        tf_d.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_d.margin_left = tf_d.margin_right = tf_d.margin_top = tf_d.margin_bottom = 0
        p = tf_d.paragraphs[0]
        r1 = p.add_run()
        r1.text = f"{ui_id}: "
        r1.font.name = "Calibri"
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_RED
        r2 = p.add_run()
        r2.text = desc
        r2.font.name = "Calibri"
        r2.font.size = Pt(12)
        r2.font.color.rgb = COLOR_BLACK

    set_speaker_note(s7, "Here is an overview of our multipage Streamlit interface. The top KPI row gives immediate executive visibility into total spend, gross revenue, net margin, and blended ROAS. Interactive sidebar filters allow stakeholders to isolate individual channels or timeframes reactively. Screen UI-02 provides our campaign scatter quadrants, UI-04 surfaces hypothesis tests, and UI-05 models our multi-touch attribution simulator.")

    # ==================== SLIDE 8: KEY FINDINGS ====================
    s8 = create_blank_slide(prs)
    add_header_footer(s8, "Key findings", 8)

    cards_data = [
        (0.5, "3.15x", "Search & Social Scalability", "Google Ads (3.15x) and Meta Ads (2.04x) deliver 72% of total conversions with healthy unit economics."),
        (3.55, "+34.3%", "Weekend Conversion Surge", "Two-sample pooled Z-test (Z = +3.93, p < 0.0001) confirms statistically significant Fri-Sun conversion uplift."),
        (6.6, "37%", "Attribution Discovery Credit", "First-Touch attribution credits 37% of customer discovery to Influencers, despite low direct last-click ROAS (0.89x).")
    ]
    for x_pos, num_val, meaning_val, evidence_val in cards_data:
        add_card(s8, x_pos, 1.3, 2.85, 3.6)

        # Number
        tb_n = s8.shapes.add_textbox(Inches(x_pos + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = tb_n.text_frame
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_n.paragraphs[0]
        p.text = num_val
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Cambria"
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = COLOR_RED

        # Meaning
        tb_m = s8.shapes.add_textbox(Inches(x_pos + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_m.paragraphs[0]
        p.text = meaning_val
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_BLACK

        # Evidence
        tb_e = s8.shapes.add_textbox(Inches(x_pos + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = tb_e.text_frame
        tf_e.word_wrap = True
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_e.paragraphs[0]
        p.text = f"Evidence: {evidence_val}"
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_GREY

    set_speaker_note(s8, "Our analysis revealed three major business takeaways. First, Google Ads and Meta Ads represent our primary growth drivers with 3.15x and 2.04x ROAS. Second, our hypothesis testing proved a +34.3% weekend conversion uplift with Z = +3.93 and p < 0.0001, providing conclusive evidence for weekend day-parting. Third, multi-touch attribution showed Influencer marketing initiates 37% of customer journeys, disproving the notion that it should be defunded based on last-click metrics.")

    # ==================== SLIDE 9: TESTING AND VALIDATION ====================
    s9 = create_blank_slide(prs)
    add_header_footer(s9, "Testing and validation", 9)

    # Left Column: Passed / Total
    tb_pass = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_pass = tb_pass.text_frame
    tf_pass.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_pass.paragraphs[0]
    p.text = "12 / 12"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Cambria"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED

    tb_sub = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_sub = tb_sub.text_frame
    tf_sub.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_sub.paragraphs[0]
    p.text = "100% test cases passed\npytest suites: 12\nmanual / acceptance: 14"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_GREY

    # Right Column: TABLE at (3.4, 1.35), widths 1.1, 3.9, 1.1
    t9_shape = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.2))
    t9 = t9_shape.table
    w9 = [Inches(1.1), Inches(3.9), Inches(1.1)]
    for idx, w in enumerate(w9):
        t9.columns[idx].width = w

    h9 = ["Test", "Checks", "Result"]
    for j, h in enumerate(h9):
        cell = t9.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_RED
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    t9_data = [
        ("TC-T-01", "SQLite tables, analytical view, and foreign keys", "Pass (0 errors)"),
        ("TC-T-09", "Two-sample pooled proportion Z-test math formula", "Pass (Z=3.93)"),
        ("TC-T-12", "Reporting view query latency benchmark (< 3.0s)", "Pass (9.18 ms)")
    ]
    for i, row in enumerate(t9_data):
        for j, val in enumerate(row):
            cell = t9.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_BLACK
            p.alignment = PP_ALIGN.CENTER if j in [0, 2] else PP_ALIGN.LEFT

    set_speaker_note(s9, "Quality assurance was executed through 12 automated pytest suites and 14 functional acceptance tests, achieving a 100% pass rate. Technical test TC-T-01 validated database referential integrity with zero foreign key violations. TC-T-09 validated the numerical accuracy of our Z-test mathematical engine. And TC-T-12 proved that our consolidated SQLite analytical view executes in just 9.18 milliseconds, vastly outperforming our 3.0-second SLA.")

    # ==================== SLIDE 10: RECOMMENDATION, LIMITATIONS & NEXT STEPS ====================
    s10 = create_blank_slide(prs)
    add_header_footer(s10, "Recommendation, limitations and next steps", 10)

    rec_body = "Reallocate ₹340,000 away from fixed-fee Influencer contracts into weekend Google Search Ads and Meta Retargeting. Capitalizes on the +34.3% weekend surge, projecting an +8.4% blended revenue uplift with zero budget expansion."
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation", rec_body)

    lim_body = "Current attribution models use static heuristic weights (First, Last, Linear) rather than algorithmic Shapley / Markov data-driven attribution. Realized performance depends on platform audience saturation thresholds."
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", lim_body)

    next_body = "1. Deploy Markov Chain attribution models.\n2. Integrate live Google Ads API webhooks for automated bid pausing.\n3. Conduct multi-armed bandit A/B testing on ad creative copy."
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next steps", next_body)

    set_speaker_note(s10, "Our key recommendation is reallocating ₹340,000 away from fixed-fee influencer contracts into weekend Google Search and Meta lookalikes, projecting an +8.4% blended revenue uplift. An honest limitation of our current pipeline is the use of heuristic attribution rather than algorithmic Markov chains. Our immediate next steps are integrating data-driven Shapley attribution and connecting live Google Ads webhooks.")

    # ==================== SLIDE 11: DARK CLOSING SLIDE ====================
    s11 = create_blank_slide(prs)
    set_dark_background(s11)

    # ABES Logo placeholder
    tb_logo = s11.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_logo = tb_logo.text_frame
    tf_logo.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_logo.paragraphs[0]
    p.text = "[ABES logo]"
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = COLOR_PINK

    # Institution Line
    tb_inst = s11.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_inst = tb_inst.text_frame
    tf_inst.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_inst.paragraphs[0]
    p.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_PINK

    # Title
    tb_t = s11.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    tf_t.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf_t.paragraphs[0]
    p.text = "Live Demo & Q&A"
    p.font.name = "Cambria"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    # Lines
    tb_lines = s11.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = tb_lines.text_frame
    tf_lines.word_wrap = True
    tf_lines.vertical_anchor = MSO_ANCHOR.TOP

    p0 = tf_lines.paragraphs[0]
    p0.text = "Live App: https://2503201450157-channel-marketing.streamlit.app"
    p0.font.name = "Calibri"
    p0.font.size = Pt(15)
    p0.font.color.rgb = COLOR_CARD
    p0.space_before = Pt(4)

    p1 = tf_lines.add_paragraph()
    p1.text = "GitHub Repository: https://github.com/vaishnavi-singh-2503201450157/2503201450157-channel-wise-marketing-performance"
    p1.font.name = "Calibri"
    p1.font.size = Pt(15)
    p1.font.color.rgb = COLOR_CARD
    p1.space_before = Pt(4)

    p2 = tf_lines.add_paragraph()
    p2.text = "Thank You! Questions & Discussion Welcomed."
    p2.font.name = "Calibri"
    p2.font.size = Pt(15)
    p2.font.color.rgb = COLOR_PINK
    p2.space_before = Pt(4)

    set_speaker_note(s11, "I will now open the live application for interactive demonstration. You can also explore the complete open-source codebase on GitHub. Thank you for your time and guidance, and I welcome any questions from the panel.")

    # Save presentation
    prs.save(TARGET_PPTX)
    print(f"Master presentation saved to: {TARGET_PPTX}")

if __name__ == "__main__":
    build_presentation()
