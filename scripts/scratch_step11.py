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

def generate_step11_word():
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
    title_run = title_p.add_run("AGILE USER STORIES & ACCEPTANCE CRITERIA\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nStep 11 Milestone | Epic Grouping, Given/When/Then & Sprint Plan\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Identifier", "DOC-US-V1.0 (Step 11 User Stories Specification)"),
        ("Author / Agile Product Owner", "Vaishnavi Singh (Roll No: 2503201450157, MCA 2nd Yr DA, ABES)"),
        ("Project Mentor", "Prof. Kalpana Dwivedi (ABES Engineering College)"),
        ("Story Count & Velocity", "14 User Stories (39 Story Points, all buildable in <= 0.5 days each)"),
        ("Epics Covered (5)", "Data Engineering, Executive Dashboard, Funnel Analysis, Statistical Inference, Prescriptive Insights")
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

    # 1. Epics Overview
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Agile Epic Breakdown & Architectural Grouping").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "User stories are structured into five cohesive functional epics, ensuring a modular development lifecycle:"
    )
    doc.add_paragraph(
        "• EPIC-DATA (Data Engineering & Quality): Ingestion from SQLite, in-memory caching, anomaly auditing, and tabular export.\n"
        "• EPIC-DASH (Executive Dashboard & Pacing): Real-time financial KPI cards, dynamic sidebar filters, and threshold alerts.\n"
        "• EPIC-ANALYSIS (Channel Funnel & Seasonality): CTR/CVR funnel decomposition, temporal patterns, and cohort demographics.\n"
        "• EPIC-INFERENCE (Statistical Hypothesis Testing): Scientific validation using Two-sample Z-tests and One-way ANOVA.\n"
        "• EPIC-INSIGHTS (Attribution & Prescriptive Simulation): Multi-touch attribution comparison and what-if budget optimization."
    )

    # 2. Master User Stories Table
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Master User Stories Specification Table (US-01 to US-14)").font.color.rgb = RGBColor(30, 58, 138)

    us_table = doc.add_table(rows=15, cols=7)
    us_headers = ["US ID", "Target Persona", "User Story Statement (As a / I want / So that)", "Feature ID", "BR ID", "Priority", "Points"]
    for j, uh in enumerate(us_headers):
        us_table.rows[0].cells[j].text = uh

    stories_data = [
        ("US-01", "Senior Data Analyst", "As a Data Analyst, I want the system to load cached marketing telemetry from SQLite so that analysis executes instantly without database lag.", "F-01", "BR-01", "Must", "3"),
        ("US-02", "Senior Data Analyst", "As a Data Analyst, I want an automated audit of untracked UTMs and outliers so that downstream reporting is free from corrupted data.", "F-10", "BR-10", "Must", "3"),
        ("US-03", "Academic Mentor / Analyst", "As an Analyst or Reviewer, I want to export filtered data and audit summaries to CSV/Excel so that I can audit calculations offline.", "F-12", "BR-12", "Should", "2"),
        ("US-04", "Chief Marketing Officer (CMO)", "As a CMO, I want a unified overview of channel ROAS, CAC, and ad spend so that I can evaluate acquisition efficiency across all channels.", "F-01", "BR-01", "Must", "3"),
        ("US-05", "All Stakeholders", "As a User, I want global sidebar filters for dates, channels, and campaign types so that every dashboard screen updates dynamically.", "F-11", "BR-11", "Must", "3"),
        ("US-06", "Paid Acquisition Lead", "As an Ad Manager, I want automatic alerts for campaigns breaching ROAS < 2.0x or CAC > ₹1,500 so that I can immediately pause bleeding ad sets.", "F-02", "BR-02", "Must", "2"),
        ("US-07", "Paid Acquisition Lead", "As an Ad Manager, I want to decompose CTR and CVR over time so that I can isolate ad creative fatigue from landing page friction.", "F-03", "BR-03", "Must", "3"),
        ("US-08", "Paid Acquisition Lead", "As a Media Planner, I want day-of-week and monthly performance patterns so that I can schedule ad delivery around peak conversion days.", "F-05", "BR-05", "Should", "3"),
        ("US-09", "Retention Marketing Lead", "As a Content Marketer, I want to view channel conversion rates across age, gender, and city tiers so that I can tailor ad creative copy.", "F-06", "BR-06", "Could", "2"),
        ("US-10", "Retention Marketing Lead", "As a Retention Lead, I want to analyze customer repeat purchases and LTV:CAC ratios so that I can prioritize high-LTV acquisition channels.", "F-04", "BR-04", "Should", "3"),
        ("US-11", "Senior Data Analyst", "As a Data Analyst, I want a Two-sample proportion Z-test between channels so that I can prove whether conversion differences are statistically significant.", "F-07", "BR-07", "Must", "3"),
        ("US-12", "Senior Data Analyst / Mentor", "As an Analyst, I want to run One-way ANOVA across all channels so that I can validate overall performance variance at alpha = 0.05.", "F-07", "BR-07", "Must", "3"),
        ("US-13", "Chief Marketing Officer (CMO)", "As a CMO, I want to compare First-Touch, Last-Touch, and Linear attribution so that awareness channels are not penalized by last-click bias.", "F-08", "BR-08", "Must", "3"),
        ("US-14", "Chief Marketing Officer (CMO)", "As a CMO, I want an interactive budget reallocation simulator under fixed spend so that I can project revenue uplifts from marginal return shifts.", "F-09", "BR-09", "Must", "3")
    ]

    for r_idx, r_data in enumerate(stories_data, start=1):
        for c_idx, val in enumerate(r_data):
            us_table.rows[r_idx].cells[c_idx].text = val
    style_table(us_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Acceptance Criteria (Given / When / Then)
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Acceptance Criteria in Given / When / Then Notation").font.color.rgb = RGBColor(30, 58, 138)

    criteria_specs = [
        ("US-01: SQLite Telemetry Loading & Caching",
         "Given the SQLite database 'data/app.db' exists and contains fact tables,\n"
         "When the application initiates or a user navigates to any dashboard screen,\n"
         "Then telemetry records are loaded into an in-memory DataFrame via @st.cache_data in under 1.5 seconds without disk re-queries."),

        ("US-02: Automated Data Quality Auditing",
         "Given marketing telemetry contains injected untracked UTMs and spend spikes,\n"
         "When the user opens the Data Quality page,\n"
         "Then the audit module reports 7 missing UTM records (2.7%), flags 3 statistical outliers (Z > 3.0), and displays diagnostic reasons."),

        ("US-03: Data Export to CSV & Excel",
         "Given an active filtered view of campaign records or audit tables,\n"
         "When the user clicks the 'Download CSV' or 'Download Excel' button,\n"
         "Then a clean tabular file is generated matching the filtered rows and columns with 100% data fidelity."),

        ("US-04: Cross-Channel Financial KPI Summary Cards",
         "Given the Overview page is loaded with active marketing telemetry,\n"
         "When the executive view renders,\n"
         "Then Tailwind-styled cards display Total Spend, Total Revenue, Blended ROAS, and Blended CAC matching mathematical formulas with two decimal precision."),

        ("US-05: Global Multi-Dimensional Sidebar Filtering",
         "Given the user is on any screen of the multi-page Streamlit application,\n"
         "When the user selects a specific Date Range or deselects channels in the sidebar,\n"
         "Then all KPI cards, statistical tables, and Plotly charts across all pages re-render dynamically to reflect only the selected cohort."),

        ("US-06: Automated Threshold Exception Alerts",
         "Given campaign records with varying ROAS and CAC performance,\n"
         "When a campaign has ROAS < 2.0x OR CAC > ₹1,500 over a rolling 7-day period,\n"
         "Then it is highlighted in red in the Exception Table with an 'Underperforming' badge and exact variance gaps displayed."),

        ("US-07: CTR & CVR Funnel Decomposition",
         "Given daily impression, click, and conversion logs,\n"
         "When the user opens the Channel Deep Dive screen,\n"
         "Then a Plotly funnel chart displays Impression-to-Click (CTR) and Click-to-Conversion (CVR) ratios alongside time-series trend lines."),

        ("US-08: Temporal Seasonality Heatmaps",
         "Given 12 months of daily performance data,\n"
         "When the temporal analysis tab is selected,\n"
         "Then a heatmap and 7-day moving average plot reveal day-of-week patterns, clearly highlighting the weekend conversion surge (+34.3% CVR)."),

        ("US-09: Demographic Segment Affinities",
         "Given customer demographic records linked to acquisition touchpoints,\n"
         "When demographic filters (Age Group, Gender, City Tier) are applied,\n"
         "Then stacked bar and sunburst charts display conversion shares across demographic cohorts."),

        ("US-10: Customer LTV:CAC Cohort Evaluation",
         "Given customer repeat order histories linked to initial acquisition channels,\n"
         "When cohort retention analysis is viewed,\n"
         "Then Average Order Value, order frequency, total lifetime revenue, and LTV:CAC ratio (> 3.0:1 benchmark) are calculated per channel."),

        ("US-11: Two-Sample Proportion Z-Test",
         "Given user clicks and conversion totals for any two selected marketing channels,\n"
         "When the user triggers the Two-Sample Z-Test in the Statistical Inference module,\n"
         "Then the system computes the pooled Z-score, two-tailed p-value, and automatically outputs a plain-English conclusion (Reject / Fail to Reject H0 at alpha = 0.05)."),

        ("US-12: One-Way ANOVA Hypothesis Testing",
         "Given daily revenue and spend distributions across all five marketing channels,\n"
         "When the ANOVA module executes,\n"
         "Then the F-statistic, between-group degrees of freedom, and p-value are displayed, proving overall channel efficiency disparity."),

        ("US-13: Multi-Touch Attribution Comparison",
         "Given multi-touch customer journey logs,\n"
         "When the user selects the Attribution tab,\n"
         "Then conversion and revenue credits are displayed side-by-side for First-Touch, Last-Touch, and Equal Linear models in a comparative Plotly bar chart."),

        ("US-14: What-If Constrained Budget Reallocation",
         "Given historical diminishing return curves for each marketing channel,\n"
         "When the user shifts channel budget sliders while enforcing total budget constraint,\n"
         "Then projected portfolio revenue and ROAS uplift percentage (+10% to 15%) update in real time.")
    ]

    for title, ac in criteria_specs:
        p_c = doc.add_paragraph()
        r_c = p_c.add_run(f"• {title}\n")
        r_c.bold = True
        r_c.font.color.rgb = RGBColor(30, 58, 138)
        p_c.add_run(ac)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4. Sprint Schedule
    h4 = doc.add_heading(level=1)
    h4.add_run("4. 10-Day Sprint Implementation Schedule").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Stories are scheduled across a structured 10-day sprint cycle, balancing foundational infrastructure, analytical computation, and UI integration:"
    )

    sprint_table = doc.add_table(rows=11, cols=4)
    sprint_headers = ["Sprint Day", "Scheduled User Stories", "Primary Focus & Deliverables", "Target Epic"]
    for j, sh in enumerate(sprint_headers):
        sprint_table.rows[0].cells[j].text = sh

    sprint_schedule = [
        ("Day 1", "US-01, US-02", "Data loader, in-memory caching (@st.cache_data), and data quality anomaly detection.", "EPIC-DATA"),
        ("Day 2", "US-04, US-05", "Tailwind KPI summary cards on Overview screen and shared sidebar filter state.", "EPIC-DASH"),
        ("Day 3", "US-06, US-03", "Threshold exception filtering table (ROAS < 2.0x, CAC > ₹1,500) and CSV/Excel export tools.", "EPIC-DASH & QUALITY"),
        ("Day 4", "US-07", "Click-Through Rate (CTR) and Conversion Rate (CVR) multi-stage funnel visualizations.", "EPIC-ANALYSIS"),
        ("Day 5", "US-08", "Temporal day-of-week and monthly seasonality heatmaps with 7-day rolling trends.", "EPIC-ANALYSIS"),
        ("Day 6", "US-09, US-10", "Demographic segment cross-tabulations and customer LTV:CAC cohort retention calculations.", "EPIC-ANALYSIS"),
        ("Day 7", "US-11", "Two-sample proportion Z-test mathematical engine with automated hypothesis rejection.", "EPIC-INFERENCE"),
        ("Day 8", "US-12", "One-way ANOVA cross-channel variance calculation with F-statistic and p-value displays.", "EPIC-INFERENCE"),
        ("Day 9", "US-13", "Multi-touch journey attribution models (First-Touch, Last-Touch, Linear) comparative bars.", "EPIC-INSIGHTS"),
        ("Day 10", "US-14", "Interactive what-if budget reallocation slider simulator under fixed spend constraints.", "EPIC-INSIGHTS")
    ]
    for r_idx, r_data in enumerate(sprint_schedule, start=1):
        for c_idx, val in enumerate(r_data):
            sprint_table.rows[r_idx].cells[c_idx].text = val
    style_table(sprint_table)

    # Save document
    out_word = os.path.abspath("docs/11_User_Stories.docx")
    doc.save(out_word)
    print(f"Saved: {out_word}")

def update_rtm_with_user_stories():
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    if not os.path.exists(rtm_path):
        print("RTM not found!")
        return

    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]

    # Mapping from (OQ, HMW, BR) to (Feature ID, User Story ID)
    mapping = {
        ("OQ-01", "HMW-01", "BR-01"): ("F-01", "US-01, US-04"),
        ("OQ-01", "HMW-01", "BR-11"): ("F-11", "US-05"),
        ("OQ-01", "HMW-01", "BR-12"): ("F-12", "US-03"),
        ("OQ-07", "HMW-02", "BR-02"): ("F-02", "US-06"),
        ("OQ-02", "HMW-03", "BR-03"): ("F-03", "US-07"),
        ("OQ-02", "HMW-03", "BR-11"): ("F-11", "US-05"),
        ("OQ-03", "HMW-04", "BR-04"): ("F-04", "US-10"),
        ("OQ-04", "HMW-05", "BR-05"): ("F-05", "US-08"),
        ("OQ-05", "HMW-06", "BR-06"): ("F-06", "US-09"),
        ("OQ-06", "HMW-07", "BR-07"): ("F-07", "US-11, US-12"),
        ("OQ-08", "HMW-08", "BR-08"): ("F-08", "US-13"),
        ("OQ-09", "HMW-09", "BR-09"): ("F-09", "US-14"),
        ("OQ-10", "HMW-10", "BR-10"): ("F-10", "US-02"),
        ("OQ-10", "HMW-10", "BR-12"): ("F-12", "US-03")
    }

    font_entry = Font(name="Segoe UI", size=9.5, bold=True, color="1E3A8A")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Column 8 is "Feature", Column 10 is "User Story"
    for r in range(2, ws.max_row + 1):
        oq = ws.cell(row=r, column=1).value
        hmw = ws.cell(row=r, column=2).value
        br = ws.cell(row=r, column=3).value
        key = (oq, hmw, br)
        if key in mapping:
            feat_id, us_id = mapping[key]
            
            c_feat = ws.cell(row=r, column=8, value=feat_id)
            c_feat.font = font_entry
            c_feat.alignment = align_center
            
            c_us = ws.cell(row=r, column=10, value=us_id)
            c_us.font = font_entry
            c_us.alignment = align_center

    wb.save(rtm_path)
    print(f"RTM updated with Feature and User Story IDs: {rtm_path}")

if __name__ == "__main__":
    generate_step11_word()
    update_rtm_with_user_stories()
