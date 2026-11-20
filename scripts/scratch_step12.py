import os
import openpyxl
from openpyxl.styles import Font, Alignment
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, header_bg="1E3A8A", alt_bg="F1F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            for cell in row.cells:
                set_cell_background(cell, header_bg)
                set_cell_margins(cell, 120, 120, 140, 140)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(9.5)
        else:
            bg = alt_bg if i % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                set_cell_background(cell, bg)
                set_cell_margins(cell, 90, 90, 130, 130)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
                        run.font.color.rgb = RGBColor(30, 41, 59)

def generate_step12_doc():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Title Block
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("STREAMLIT UI/UX FRAMEWORK & WIREFRAME SPECIFICATION\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nStep 12 Milestone | Tailwind CSS Integration & Multipage Screen Architecture\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Identifier", "DOC-UI-V1.0 (Step 12 UI Framework Artifact)"),
        ("Author / UI Architect", "Vaishnavi Singh (Roll No: 2503201450157, MCA 2nd Yr DA, ABES)"),
        ("Project Mentor", "Prof. Kalpana Dwivedi (ABES Engineering College)"),
        ("UI Presentation Stack", "Streamlit 1.40+, Plotly 5.24+, Tailwind CSS (CDN in st.components.v1.html)"),
        ("Screen Inventory (7)", "Overview (UI-01), Funnels (UI-02), Cohorts (UI-03), Inference (UI-04), Attribution (UI-05), Quality (UI-06), About (UI-07)")
    ]
    for idx, (k, v) in enumerate(meta_info):
        ck, cv = meta_table.rows[idx].cells
        ck.text = k
        cv.text = v
        set_cell_background(ck, "F1F5F9")
        set_cell_background(cv, "FFFFFF")
        set_cell_margins(ck, 60, 60, 100, 100)
        set_cell_margins(cv, 60, 60, 100, 100)
        ck.paragraphs[0].runs[0].font.bold = True
        ck.paragraphs[0].runs[0].font.size = Pt(9.5)
        cv.paragraphs[0].runs[0].font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # 1. Screen Inventory
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Application Screen Inventory (UI-01 to UI-07)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "The application provides a multi-page dashboard interface structured around core analytical workflows:"
    )

    scr_table = doc.add_table(rows=8, cols=6)
    scr_headers = ["UI ID", "Screen Name / Route", "Operational Purpose", "Primary UI Components", "Interactive Filters", "Linked User Stories"]
    for j, sh in enumerate(scr_headers):
        scr_table.rows[0].cells[j].text = sh

    scr_data = [
        ("UI-01", "Executive Overview (app.py)", "Portfolio acquisition health, spend pacing, blended ROAS/CAC, and threshold breach warnings.", "Tailwind Header, 4 KPI Metric Cards, Channel ROAS Comparison Bar, Exception Table.", "Date Range, Channel Multiselect", "US-01, US-04, US-05, US-06"),
        ("UI-02", "Channel Deep Dive (pages/1_Channel_Deep_Dive.py)", "Funnel stage drop-offs (CTR/CVR) and day-of-week / monthly temporal seasonality patterns.", "Multi-stage Plotly Funnel, 7-day rolling ROAS line chart, Day-of-Week Seasonality Heatmap.", "Channel Tabs, Campaign Type, Date Range", "US-05, US-07, US-08"),
        ("UI-03", "Customer Cohorts & LTV (pages/2_Customer_Cohorts.py)", "Demographic conversion affinities and customer lifetime value (LTV:CAC) ratio benchmarks.", "Age/Gender Stacked Bar, City Tier Sunburst, LTV:CAC Scatter with 3:1 reference line.", "City Tier, Age Brackets, Channel", "US-09, US-10"),
        ("UI-04", "Statistical Inference (pages/3_Statistical_Inference.py)", "Scientific hypothesis testing validating conversion rate disparities and cross-channel variance.", "Channel Pair Selector, Two-Sample Z-Test Card (Z, p-value, CI, Verdict), One-Way ANOVA Table.", "Channel A, Channel B, Alpha (0.05)", "US-11, US-12"),
        ("UI-05", "Attribution & Budget (pages/4_Attribution_&_Budget.py)", "Comparative multi-touch attribution and forward-looking constrained budget reallocation simulation.", "Model Selector (First/Last/Linear), Side-by-side Attribution Bars, Interactive Budget Sliders.", "Attribution Model, Total Budget Slider", "US-13, US-14"),
        ("UI-06", "Data Quality & Export (pages/5_Data_Quality_&_Export.py)", "Data governance, missing UTM audit, outlier detection, and downloadable data exports.", "Completeness Gauge (98%+), Missing UTM Table, Outlier List, Download CSV/Excel Buttons.", "Anomaly Type Selector", "US-02, US-03"),
        ("UI-07", "Project Documentation (pages/6_About.py)", "Academic metadata, problem statement, ABES student/mentor details, and architecture diagrams.", "Student/Mentor Card, Architecture Viewer, Live System Diagnostics, Git Repo Links.", "None", "US-01, US-04")
    ]
    for r_idx, r_data in enumerate(scr_data, start=1):
        for c_idx, val in enumerate(r_data):
            scr_table.rows[r_idx].cells[c_idx].text = val
    style_table(scr_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. ASCII Wireframes for All Screens
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Comprehensive ASCII Wireframe Architecture").font.color.rgb = RGBColor(30, 58, 138)

    wireframes = [
        ("UI-01: Executive Overview & Financial Pacing (app.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR - st.session_state]     | [TAILWIND HEADER] Channel Wise Marketing Performance               |\n"
         "| Date: [2026-01-01 to 2026-12-31] | System: Online | Database: SQLite (app.db) | Author: Vaishnavi     |\n"
         "| Channels: [*] All Selected       +-------------------------------------------------------------------+\n"
         "| Campaign Type: [All]             | [KPI 1: SPEND]      [KPI 2: REVENUE]    [KPI 3: ROAS]     [KPI 4: CAC] |\n"
         "|                                  | ₹1,348,210.50       ₹4,285,420.00       3.18x (+12%)      ₹980 (-8%)   |\n"
         "| [Reset Filters]                  +-------------------------------------------------------------------+\n"
         "|                                  | [CHART: DUAL-AXIS BAR & LINE] Channel Spend vs. Revenue vs. ROAS  |\n"
         "|                                  | (X: Channel Name | Bars: Spend, Revenue | Line: Blended ROAS)     |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [ALERT TABLE: UNDERPERFORMING CAMPAIGNS (ROAS < 2.0x | CAC > ₹1500]|\n"
         "|                                  | Campaign Name         | Actual ROAS | Target | Actual CAC | Status   |\n"
         "|                                  | Meta_Stories_Flash    | 1.55x       | 2.00x  | ₹1,820     | [PAUSE]  |\n"
         "|                                  | Influencer_Celebrity  | 1.38x       | 1.80x  | ₹2,150     | [PAUSE]  |\n"
         "+------------------------------------------------------------------------------------------------------+\n"),

        ("UI-02: Channel Deep Dive & Funnels (pages/1_Channel_Deep_Dive.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR]                        | [PAGE HEADER] Channel Funnel Decomposition & Temporal Dynamics    |\n"
         "| Channel: [Google Ads      v]     +-------------------------------------------------------------------+\n"
         "| Metric:  [Conversion Rate v]     | [FUNNEL CHART: IMPRESSIONS -> CLICKS -> CONVERSIONS]              |\n"
         "|                                  | Impressions: 450,200  ======================================= 100%|\n"
         "|                                  | Clicks (CTR 3.4%):     15,300   ============== 3.4% drop          |\n"
         "|                                  | Conversions (CVR 3.1%):   475   ==== 3.1% final                   |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [HEATMAP: DAY-OF-WEEK VS. MONTH SEASONALITY (Plotly imshow)]      |\n"
         "|                                  | Monday   [  ][  ][  ][  ]   Friday   [##][##][##][##] (+34% CVR)  |\n"
         "|                                  | Tuesday  [  ][  ][  ][  ]   Saturday [##][##][##][##] (Peak Shop) |\n"
         "+------------------------------------------------------------------------------------------------------+\n"),

        ("UI-03: Customer Cohorts & LTV (pages/2_Customer_Cohorts.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR]                        | [PAGE HEADER] Customer Demographic Cohorts & Lifetime Value (LTV) |\n"
         "| City Tier: [Tier 1, Tier 2 v]    +-------------------------------------------------------------------+\n"
         "| Age:       [25-34          v]    | [STACKED BAR: CONVERSION BY AGE]  | [SUNBURST: TIER -> CITY -> CH]|\n"
         "|                                  | 18-24: [Meta 45%][Inf 30%][SEO]   | Tier 1 > Bengaluru > Google   |\n"
         "|                                  | 25-34: [Google 35%][Meta 35%][Em] | Tier 2 > Jaipur > Meta        |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [SCATTER PLOT: LTV VS. CAC WITH BENCHMARK 3:1 RATIO LINE]         |\n"
         "|                                  | (Y: Customer Lifetime Value | X: Acquisition Cost | Ref: LTV=3*CAC)|\n"
         "+------------------------------------------------------------------------------------------------------+\n"),

        ("UI-04: Statistical Inference & Testing (pages/3_Statistical_Inference.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR]                        | [PAGE HEADER] Inferential Statistical Testing Engine              |\n"
         "| Select Test: [Two-Sample Z-Test] +-------------------------------------------------------------------+\n"
         "| Channel A:   [Google Ads       ] | [TWO-SAMPLE PROPORTION Z-TEST RESULTS]                            |\n"
         "| Channel B:   [Meta Ads         ] | Channel A CVR: 3.15% (n=15,300) | Channel B CVR: 2.04% (n=22,400) |\n"
         "| Alpha Level: [0.05 (95% CI)    ] | Z-Score: 6.84 | p-value: 0.0001 < 0.05 | 95% CI: [+0.78%, +1.44%] |\n"
         "|                                  | [INSIGHT BOX: H0 REJECTED] Statistically significant advantage!   |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [ONE-WAY ANOVA F-TEST: CROSS-CHANNEL VARIANCE]                    |\n"
         "|                                  | F-Statistic: 42.18 | p-value: 1.2e-9 | Significant Differences     |\n"
         "+------------------------------------------------------------------------------------------------------+\n"),

        ("UI-05: Attribution & Budget Simulator (pages/4_Attribution_&_Budget.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR]                        | [PAGE HEADER] Multi-Touch Attribution & Prescriptive Budget Model |\n"
         "| Total Budget: [₹1,500,000 Slider]+-------------------------------------------------------------------+\n"
         "| Model: [First vs. Last vs. Linear| [CHART: MULTI-TOUCH ATTRIBUTION SHARE COMPARISON]                  |\n"
         "|                                  | First-Touch: [SEO 35%][Inf 25%][Google 20%][Meta 15%][Email 5%]   |\n"
         "|                                  | Last-Touch:  [Email 38%][Meta Retarget 35%][Google 15%][SEO 8%]   |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [BUDGET REALLOCATION WHAT-IF SIMULATOR (Diminishing Returns)]     |\n"
         "|                                  | Current Projected Revenue: ₹4.28M | Optimized Revenue: ₹4.82M     |\n"
         "|                                  | Projected Blended ROAS Uplift: +12.6% | Capital Saved: ₹185,000   |\n"
         "+------------------------------------------------------------------------------------------------------+\n"),

        ("UI-06: Data Quality & Export (pages/5_Data_Quality_&_Export.py)",
         "+------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR]                        | [PAGE HEADER] Data Governance, Anomaly Audits & Tabular Export    |\n"
         "| Audit Type: [All Quality Checks] +-------------------------------------------------------------------+\n"
         "| Outliers:   [Show Z > 3.0 Only ] | [QUALITY SCORE: 98.4%] | Missing UTMs: 7 (2.7%) | Outliers: 3     |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [AUDIT TABLE: FLAGGED ANOMALIES & OUTLIERS]                       |\n"
         "|                                  | Fact ID | Date       | Campaign           | Anomaly Reason        |\n"
         "|                                  | 42      | 2026-03-12 | Google_Search_Core | Spend Spike (Z=3.6)   |\n"
         "|                                  | 118     | 2026-06-25 | Meta_Lookalike     | Revenue Spike (Z=3.8) |\n"
         "|                                  +-------------------------------------------------------------------+\n"
         "|                                  | [BUTTON: Download Filtered CSV]  [BUTTON: Download Excel Workbook]|\n"
         "+------------------------------------------------------------------------------------------------------+\n")
    ]

    for title, wireframe in wireframes:
        p_w = doc.add_paragraph()
        r_w = p_w.add_run(f"• {title}\n")
        r_w.bold = True
        r_w.font.color.rgb = RGBColor(30, 58, 138)
        
        p_box = doc.add_paragraph()
        r_box = p_box.add_run(wireframe)
        r_box.font.name = 'Consolas'
        r_box.font.size = Pt(7.5)
        p_box.paragraph_format.left_indent = Inches(0.1)

    # 3. Global Layout & State Architecture
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Global Layout & Session State Architecture").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "• Shared Sidebar Filter State: To prevent filter loss during page navigation, global filter parameters "
        "('selected_date_range', 'selected_channels', 'selected_campaign_types') are initialized in 'st.session_state' "
        "inside app.py. When a user changes filters on any page, the state persists across all other page routes.\n"
        "• Corporate Theme Palette:\n"
        "  - Primary Accent: #1E3A8A (Deep Navy Blue)\n"
        "  - App Canvas Background: #F8FAFC (Slate 50)\n"
        "  - Card Container Background: #FFFFFF (Pure White)\n"
        "  - Core Typography: #0F172A (Slate 900)\n"
        "  - Success: #10B981 (Emerald Green), Danger: #EF4444 (Crimson Red), Warning: #F59E0B (Amber)\n"
        "• Responsive Breakpoint System: Desktop view displays 4-column KPI rows; mobile viewports automatically "
        "reflow into stacked 1-column responsive cards using Tailwind grid-cols-1 sm:grid-cols-2 lg:grid-cols-4."
    )

    # 4. Best Chart for Every Operational Question
    h4 = doc.add_heading(level=1)
    h4.add_run("4. Recommended Visualizations for Each Operational Question").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Every operational inquiry is paired with its optimal visual encoding to maximize interpretability:"
    )

    chart_table = doc.add_table(rows=11, cols=4)
    chart_headers = ["OQ ID", "Operational Inquiry", "Optimal Visual Chart Type", "Analytical Justification"]
    for j, ch in enumerate(chart_headers):
        chart_table.rows[0].cells[j].text = ch

    chart_data = [
        ("OQ-01", "Cross-Channel ROAS & CAC", "Dual-Axis Grouped Bar & Line Chart", "Bars display absolute financial scale (Spend/Revenue), while line highlights efficiency ratio (ROAS)."),
        ("OQ-02", "CTR & CVR Funnel Breakdown", "Multi-Stage Conversion Funnel (Plotly funnel)", "Visually pinpoints exact percentage drop-offs between impressions, clicks, and final conversions."),
        ("OQ-03", "Customer LTV:CAC Benchmarking", "Quadrant Scatter Plot with 3:1 Ratio Line", "Plots LTV against CAC; reference diagonal (LTV = 3*CAC) instantly categorizes profitable channels."),
        ("OQ-04", "Temporal Seasonality & Dips", "Day-of-Week vs. Month Matrix Heatmap", "Color intensity instantly reveals cyclical conversion surges (weekends) and seasonal drops."),
        ("OQ-05", "Demographic Segment Affinities", "Hierarchical Sunburst Chart", "Visualizes multi-tiered demographic distribution (City Tier -> Age Group -> Preferred Channel)."),
        ("OQ-06", "Statistical Significance Gaps", "Whisker Box Plot & Error Bar Plot", "Displays median, IQR, and 95% confidence interval bounds, visually proving non-overlapping significance."),
        ("OQ-07", "Underperforming Ad Exception Alert", "Conditional Highlight Exception Table", "Sorts campaigns breaching ROAS < 2.0x or CAC > ₹1500 with red badges for immediate action."),
        ("OQ-08", "Multi-Touch Attribution Shifts", "Grouped Horizontal Bar Chart", "Compares First-Touch, Last-Touch, and Linear credit side-by-side, visually exposing last-click bias."),
        ("OQ-09", "Constrained Budget Optimization", "Diminishing Returns Power Curve Plot", "Plots spend vs. revenue curve, showing current operating point vs. optimal marginal ROAS point."),
        ("OQ-10", "Data Quality & Anomaly Auditing", "Z-Score Scatter Plot with 3-Sigma Bounds", "Plots observation deviations from mean, clearly isolating outliers beyond red +/- 3-sigma lines.")
    ]
    for r_idx, r_data in enumerate(chart_data, start=1):
        for c_idx, val in enumerate(r_data):
            chart_table.rows[r_idx].cells[c_idx].text = val
    style_table(chart_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5. Production Theme Config
    h5 = doc.add_heading(level=1)
    h5.add_run("5. Production Theme Configuration (.streamlit/config.toml)").font.color.rgb = RGBColor(30, 58, 138)

    cfg_text = (
        "[theme]\n"
        "primaryColor = \"#1E3A8A\"\n"
        "backgroundColor = \"#F8FAFC\"\n"
        "secondaryBackgroundColor = \"#FFFFFF\"\n"
        "textColor = \"#0F172A\"\n"
        "font = \"sans serif\"\n\n"
        "[server]\n"
        "headless = true\n"
        "enableCORS = false\n\n"
        "[browser]\n"
        "gatherUsageStats = false"
    )
    p_cfg = doc.add_paragraph()
    r_cfg = p_cfg.add_run(cfg_text)
    r_cfg.font.name = 'Consolas'
    r_cfg.font.size = Pt(8.5)
    p_cfg.paragraph_format.left_indent = Inches(0.2)

    # Save document
    out_word = os.path.abspath("docs/12_UI_Framework.docx")
    doc.save(out_word)
    print(f"Saved: {out_word}")

def update_rtm_with_ui_screens():
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    if not os.path.exists(rtm_path):
        print("RTM not found!")
        return

    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]

    # Mapping from (OQ, HMW, BR) to UI Screen (Column 11)
    ui_mapping = {
        ("OQ-01", "HMW-01", "BR-01"): "UI-01 (Executive Overview)",
        ("OQ-01", "HMW-01", "BR-11"): "UI-01, UI-02, UI-04 (Global Sidebar)",
        ("OQ-01", "HMW-01", "BR-12"): "UI-06 (Data Quality & Export)",
        ("OQ-07", "HMW-02", "BR-02"): "UI-01 (Overview Exception Table)",
        ("OQ-02", "HMW-03", "BR-03"): "UI-02 (Channel Funnels)",
        ("OQ-02", "HMW-03", "BR-11"): "UI-02 (Channel Filters)",
        ("OQ-03", "HMW-04", "BR-04"): "UI-03 (Cohorts & LTV)",
        ("OQ-04", "HMW-05", "BR-05"): "UI-02 (Seasonality Heatmap)",
        ("OQ-05", "HMW-06", "BR-06"): "UI-03 (Demographics)",
        ("OQ-06", "HMW-07", "BR-07"): "UI-04 (Statistical Inference)",
        ("OQ-08", "HMW-08", "BR-08"): "UI-05 (Attribution Modeling)",
        ("OQ-09", "HMW-09", "BR-09"): "UI-05 (Budget Simulator)",
        ("OQ-10", "HMW-10", "BR-10"): "UI-06 (Data Quality Audit)",
        ("OQ-10", "HMW-10", "BR-12"): "UI-06 (Data Quality & Export)"
    }

    font_entry = Font(name="Segoe UI", size=9.5, bold=True, color="1E3A8A")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Column 11 is "UI Screen"
    for r in range(2, ws.max_row + 1):
        oq = ws.cell(row=r, column=1).value
        hmw = ws.cell(row=r, column=2).value
        br = ws.cell(row=r, column=3).value
        key = (oq, hmw, br)
        if key in ui_mapping:
            ui_screen_val = ui_mapping[key]
            cell_ui = ws.cell(row=r, column=11, value=ui_screen_val)
            cell_ui.font = font_entry
            cell_ui.alignment = align_center

    wb.save(rtm_path)
    print(f"RTM updated with UI Screen IDs: {rtm_path}")

if __name__ == "__main__":
    generate_step12_doc()
    update_rtm_with_ui_screens()
