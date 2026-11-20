import os
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

def generate_step5_brd():
    doc = docx.Document()
    
    # 1-inch margins
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
    title_run = title_p.add_run("BUSINESS REQUIREMENTS DOCUMENT (BRD)\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nDocument Identifier: DOC-BRD-V1.0 | Academic Year 2026–2027\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 1. Document Control Table
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Document Control & Metadata").font.color.rgb = RGBColor(30, 58, 138)

    dc_table = doc.add_table(rows=6, cols=2)
    dc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    dc_info = [
        ("Document Version / State", "Version 1.0 (Final Approved Baseline)"),
        ("Project Name & Scope", "Channel Wise Marketing Performance (Data Analytics Track)"),
        ("Author / Business Analyst", "Vaishnavi Singh (Roll No: 2503201450157, MCA 2nd Yr DA, ABES)"),
        ("Project Guide / Reviewer", "Prof. Kalpana Dwivedi (Assistant Professor, MCA Dept.)"),
        ("Creation / Baseline Date", "November 20, 2026"),
        ("Target System Architecture", "Python 3.12, Streamlit, SQLite (app.db), Plotly, Tailwind UI, pytest")
    ]
    for idx, (k, v) in enumerate(dc_info):
        ck, cv = dc_table.rows[idx].cells
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

    # 2. Executive Summary
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Executive Summary").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Modern digital business growth depends on optimizing capital deployment across multi-channel customer acquisition portfolios: "
        "Google Ads (Search & Display), Meta Ads (Instagram & Facebook), Targeted Email Campaigns, Organic Search (SEO), and Influencer Partnerships. "
        "However, marketing departments frequently encounter severe performance visibility bottlenecks. Data is fragmented across disparate vendor platforms, "
        "attribution logic is oversimplified using biased last-click models, and reporting lacks inferential statistical validation. Consequently, marketing leadership "
        "frequently makes high-stakes budget allocation decisions based on anecdotal observations rather than empirical data."
    )
    doc.add_paragraph(
        "This Business Requirements Document (BRD) defines the foundational operational capabilities for the Channel Wise Marketing Performance application. "
        "The system unifies multi-channel marketing telemetry into an optimized SQLite analytical data store, calculates standardized acquisition efficiency metrics "
        "(ROAS, CAC, CTR, CVR, LTV:CAC), identifies underperforming ad sets breaching operational benchmarks, statistically validates cross-channel performance differences, "
        "and provides an interactive what-if budget reallocation simulator. By bridging the gap between business strategy and data engineering, the solution empowers "
        "growth leaders to maximize blended portfolio return on ad spend."
    )

    # 3. Business Problem and Opportunity
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Business Problem Statement & Opportunity Analysis").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "A. The Business Problem:\n"
        "Marketing leaders and performance analysts struggle with fragmented cross-channel data silos, preventing timely visibility into channel-level "
        "Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC), which leads to inefficient budget allocation and delayed campaign optimization."
    )
    doc.add_paragraph(
        "B. The Strategic Business Opportunity:\n"
        "By centralizing cross-channel telemetry into a unified analytics data store and providing transparent multi-touch attribution (First-Touch, Last-Touch, Linear), "
        "the business can:\n"
        "1. Identify bleeding ad campaigns exceeding target CAC (> ₹1,500) and reallocate capital within hours rather than weeks.\n"
        "2. Formulate statistically sound media mix strategies backed by rigorous hypothesis testing (p < 0.05).\n"
        "3. Simulate marginal returns under fixed spend constraints to project a 10% to 15% uplift in blended portfolio ROAS.\n"
        "4. Enforce automated data quality and anomaly auditing to ensure 100% auditable management reporting."
    )

    # 4. Objectives and Success Metrics
    h4 = doc.add_heading(level=1)
    h4.add_run("4. Objectives & Key Performance Indicators (KPIs)").font.color.rgb = RGBColor(30, 58, 138)

    kpi_table = doc.add_table(rows=7, cols=4)
    kpi_headers = ["Metric Dimension", "Key Performance Indicator (KPI)", "Target Operational Threshold", "Business Value Impact"]
    for j, kh in enumerate(kpi_headers):
        kpi_table.rows[0].cells[j].text = kh

    kpi_data = [
        ("Portfolio ROAS Uplift", "Blended Return on Ad Spend", ">= 10% to 15% projected improvement", "Maximizes top-line revenue generated per rupee of marketing spend."),
        ("Acquisition Cost Control", "Channel Customer Acquisition Cost (CAC)", "CAC <= ₹1,500 on all scaled paid channels", "Prevents unsustainable customer acquisition unit economics."),
        ("Channel Profitability Baseline", "Minimum Channel ROAS", "ROAS >= 2.0x benchmark", "Establishes a hard threshold to pause unprofitable ad sets."),
        ("Dashboard Performance", "Initial Page Load & Filter Latency", "Load < 3.0s; Filter recalcs < 1.0s", "Ensures frictionless executive and analyst user experience."),
        ("Data Quality & Completeness", "UTM Tracking & Schema Completeness", "Completeness >= 98%; 0 orphan FKs", "Eliminates reporting corruption caused by broken UTM tags."),
        ("Software Quality Assurance", "Pytest Automated Test Coverage", "Statement coverage >= 85%", "Guarantees mathematical and technical reliability for viva defense.")
    ]
    for r_idx, r_data in enumerate(kpi_data, start=1):
        for c_idx, val in enumerate(r_data):
            kpi_table.rows[r_idx].cells[c_idx].text = val
    style_table(kpi_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5. Stakeholders - RACI Table
    h5 = doc.add_heading(level=1)
    h5.add_run("5. Stakeholder Engagement Matrix (RACI)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "The RACI matrix governs stakeholder participation across core project milestones "
        "(Responsible, Accountable, Consulted, Informed):"
    )

    raci_table = doc.add_table(rows=7, cols=6)
    raci_headers = ["Project Lifecycle Activity", "SH-01 (CMO)", "SH-02 (Paid Lead)", "SH-03 (Analyst)", "SH-04 (Retention)", "SH-05 (Mentor)"]
    for j, rh in enumerate(raci_headers):
        raci_table.rows[0].cells[j].text = rh

    raci_data = [
        ("Business Needs & Scoping (Steps 1–5)", "Accountable", "Consulted", "Responsible", "Consulted", "Consulted"),
        ("PRD & Technical Architecture (Steps 6–7)", "Informed", "Consulted", "Responsible", "Informed", "Accountable"),
        ("Database Design & Schema (Steps 8–9)", "Informed", "Informed", "Responsible", "Informed", "Accountable"),
        ("Data Generation & Ingestion (Step 10)", "Informed", "Consulted", "Responsible", "Informed", "Accountable"),
        ("UI/UX & Analysis Build (Steps 11–12, 14–15)", "Consulted", "Responsible", "Responsible", "Consulted", "Accountable"),
        ("Testing, Deployment & Viva (Steps 13, 16–18)", "Informed", "Informed", "Responsible", "Informed", "Accountable")
    ]
    for r_idx, r_data in enumerate(raci_data, start=1):
        for c_idx, val in enumerate(r_data):
            raci_table.rows[r_idx].cells[c_idx].text = val
    style_table(raci_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 6. Scope: In / Out
    h6 = doc.add_heading(level=1)
    h6.add_run("6. Project Scope Boundaries").font.color.rgb = RGBColor(30, 58, 138)

    scope_table = doc.add_table(rows=6, cols=2)
    scope_table.rows[0].cells[0].text = "In-Scope (Deliverables & Features Included)"
    scope_table.rows[0].cells[1].text = "Out-of-Scope (Explicitly Excluded from MVP)"

    scope_data = [
        ("Multi-channel synthetic telemetry covering 5 acquisition channels (Google Ads, Meta Ads, Email, SEO, Influencer).", "Live third-party API OAuth connections (Google Ads API, Meta Marketing API) due to sandbox/billing constraints."),
        ("Relational SQLite database (data/app.db) with normalized dimensions, daily fact table, and customer touchpoint log.", "Enterprise cloud data warehousing systems (Snowflake, AWS Redshift, Google BigQuery) requiring paid subscriptions."),
        ("Standard acquisition KPI computations (ROAS, CAC, CTR, CVR, AOV, LTV:CAC, ad spend pacing).", "Real-time streaming event bus architectures (Apache Kafka, Apache Flink) requiring distributed compute clusters."),
        ("Inferential statistical hypothesis testing (Two-sample proportion tests, ANOVA, Chi-Square, p-values, 95% CI).", "Deep neural network bid automation or automated DSP programmatic media buying."),
        ("Multi-touch attribution engine (First-Touch, Last-Touch, Linear) and what-if budget reallocation simulator.", "Live e-commerce transaction payment gateway processing or financial accounting settlement.")
    ]
    for r_idx, r_data in enumerate(scope_data, start=1):
        for c_idx, val in enumerate(r_data):
            scope_table.rows[r_idx].cells[c_idx].text = val
    style_table(scope_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 7. Business Requirements (BR-01 to BR-12)
    h7 = doc.add_heading(level=1)
    h7.add_run("7. Business Requirements Specification (BR-01 to BR-12)").font.color.rgb = RGBColor(30, 58, 138)

    br_specs = [
        ("BR-01: Consolidated Cross-Channel ROAS & CAC",
         "The business needs to consolidate and compute cross-channel Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC) across Google Ads, Meta Ads, Email, SEO, and Influencer marketing so that marketing leadership can evaluate blended and channel-specific capital efficiency.",
         "Eliminates reliance on fragmented platform reports by providing a single source of truth for acquisition unit economics.",
         "Must", "HMW-01",
         "Given the daily marketing performance data is loaded, When the user views the Overview screen, Then channel-level and portfolio-wide ROAS (Revenue/Spend) and CAC (Spend/Conversions) must match audited mathematical formulas with 100% calculation accuracy across the selected date range."),

        ("BR-02: Automated Threshold Exception Flagging",
         "The business needs to automatically identify and flag campaigns and ad sets that breach defined performance thresholds (ROAS < 2.0x or CAC > ₹1,500) so that campaign managers can immediately halt budget wastage on underperforming ad sets.",
         "Enables immediate operational cost control by preventing ad spend bleed on unprofitable creatives and targeting sets.",
         "Must", "HMW-02",
         "Given campaign records with varying ROAS and CAC, When the threshold filter is active, Then 100% of campaigns with ROAS < 2.0x or CAC > ₹1,500 are highlighted in an exception list with zero false negatives."),

        ("BR-03: CTR & CVR Conversion Funnel Decomposition",
         "The business needs to track and decompose Click-Through Rates (CTR) and Conversion Rates (CVR) across marketing channels and campaign types over time so that growth teams can pinpoint whether funnel drop-offs are driven by ad creative fatigue or landing page friction.",
         "Distinguishes between top-of-funnel ad creative fatigue (CTR drop) and bottom-of-funnel checkout friction (CVR drop).",
         "Must", "HMW-03",
         "Given daily impression, click, and conversion data, When an analyst inspects channel trends, Then CTR (Clicks/Impressions) and CVR (Conversions/Clicks) are plotted over time with campaign-type drill-downs."),

        ("BR-04: Customer LTV:CAC Ratio Analysis",
         "The business needs to analyze customer repeat order history and calculate Customer Lifetime Value relative to acquisition cost (LTV:CAC ratio) per acquisition channel so that retention teams can prioritize channels that attract loyal, high-value customers.",
         "Prevents overspending on acquisition channels that yield one-time bargain hunters rather than recurring, high-margin customers.",
         "Should", "HMW-04",
         "Given customer order history linked to acquisition channels, When viewing cohort analytics, Then cumulative revenue per customer and LTV:CAC ratio (> 3:1 benchmark) are accurately computed per channel."),

        ("BR-05: Temporal & Seasonality Trend Analysis",
         "The business needs to evaluate day-of-week and monthly performance patterns in spend, conversion volume, and ROAS so that media planners can schedule ad delivery and adjust day-parting bids to align with peak buyer conversion windows.",
         "Eliminates flat ad spend waste during historically low-converting days while capitalizing on peak consumer shopping periods.",
         "Should", "HMW-05",
         "Given 12 months of daily performance data, When day-of-week patterns are analyzed, Then spend, conversion rate, and ROAS variance across all 7 weekdays are clearly displayed with 7-day rolling moving averages."),

        ("BR-06: Demographic Audience Segmentation",
         "The business needs to segment channel conversion volume and average order values by customer demographics (Age Group, Gender, City Tier) so that marketing teams can tailor channel copy, targeting, and promotional offers to high-performing audience cohorts.",
         "Enables hyper-personalized creative targeting by uncovering demographic affinities across channels.",
         "Could", "HMW-06",
         "Given customer demographic attributes, When demographic filters are applied, Then conversion rates and average order values (AOV) are broken down by Age brackets (18-24, 25-34, 35-44, 45+), Gender, and City Tier (Tier 1 vs. Tier 2/3)."),

        ("BR-07: Inferential Statistical Hypothesis Testing",
         "The business needs to perform inferential hypothesis testing (Two-sample proportion tests, ANOVA, Chi-Square) on channel conversion rates so that marketing analysts and leadership can confirm whether channel efficiency gaps are statistically significant (p < 0.05) rather than random sampling noise.",
         "Provides scientific rigor to channel comparisons, preventing costly capital reallocations based on random variance.",
         "Must", "HMW-07",
         "Given conversion counts and visitor totals for two or more channels, When hypothesis testing is executed, Then the calculated test statistics (Z, F, Chi-Square), p-values, and automated acceptance/rejection of H0 are reported at alpha = 0.05 with 95% confidence intervals."),

        ("BR-08: Multi-Touch Attribution Modeling",
         "The business needs to compare multi-touch attribution distribution (First-Touch, Last-Touch, and Equal Linear models) across customer conversion journeys so that marketing leadership can properly value top-of-funnel discovery channels alongside closing channels.",
         "Eliminates last-click bias, ensuring top-of-funnel awareness channels receive appropriate budget credit.",
         "Must", "HMW-08",
         "Given multi-touch customer journey sequences, When attribution models are selected, Then conversion and revenue credit are calculated and displayed side-by-side for First-Touch (100% to first channel), Last-Touch (100% to last channel), and Linear (equal split)."),

        ("BR-09: Prescriptive Budget Reallocation Simulation",
         "The business needs to simulate cross-channel budget reallocation scenarios under a fixed total marketing spend based on historical marginal returns so that executive leadership can project potential revenue and ROAS uplifts before committing capital.",
         "Elevates dashboard analytics from retrospective observation to forward-looking decision intelligence.",
         "Must", "HMW-09",
         "Given baseline spend and revenue curves, When a user adjusts channel budget sliders under a fixed total budget constraint, Then projected total revenue and blended ROAS delta are updated dynamically in real time."),

        ("BR-10: Automated Data Quality & Anomaly Auditing",
         "The business needs to audit marketing data quality by automatically detecting untracked UTM parameters, duplicate transaction IDs, missing foreign keys, and statistical spend outliers so that all downstream marketing analytics are based on reliable and auditable data.",
         "Prevents polluted data from corrupting executive dashboards and invalidating analytical conclusions.",
         "Must", "HMW-10",
         "Given the ingested marketing records, When the data quality audit module runs, Then completeness percentages, missing UTM counts, duplicate transaction flags, and Z-score spend outliers (Z > 3) are summarized with zero unhandled system errors."),

        ("BR-11: Multi-Dimensional Dynamic Filtering",
         "The business needs to dynamically filter performance data by date ranges, acquisition channels, and campaign objectives so that stakeholders can isolate specific operational periods and campaign cohorts without altering baseline data.",
         "Provides versatile self-service exploration for all stakeholder roles without requiring developer intervention.",
         "Must", "HMW-01, HMW-03",
         "Given interactive sidebar filter selections (date range, channels, campaign types), When any filter selection changes, Then all KPI summary cards, statistical outputs, and Plotly charts re-render dynamically to reflect the isolated cohort."),

        ("BR-12: Tabular Data Export & Compliance",
         "The business needs to export filtered analytical summaries, raw campaign records, and data audit logs into standardized tabular formats (CSV/Excel) so that analysts can conduct offline reporting, executive briefing, and compliance documentation.",
         "Facilitates external executive presentations, academic viva submission audits, and secondary offline analysis.",
         "Should", "HMW-01, HMW-10",
         "Given any filtered dataset or analytical summary view, When the user clicks the Export button, Then a standardized CSV/Excel file matching the active view is generated and downloaded successfully.")
    ]

    for title, desc, rationale, pri, hmw, ac in br_specs:
        p_b = doc.add_paragraph()
        r_b = p_b.add_run(f"• {title}\n")
        r_b.bold = True
        r_b.font.color.rgb = RGBColor(30, 58, 138)
        
        p_desc = doc.add_paragraph()
        p_desc.add_run("  Requirement: ").bold = True
        p_desc.add_run(desc)
        
        p_meta = doc.add_paragraph()
        p_meta.add_run(f"  Priority: {pri} | Source HMW: {hmw}\n").bold = True
        p_meta.add_run("  Rationale: ").bold = True
        p_meta.add_run(f"{rationale}\n")
        p_meta.add_run("  Acceptance Criteria: ").bold = True
        p_meta.add_run(ac)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 8. Business Rules
    h8 = doc.add_heading(level=1)
    h8.add_run("8. Core Business Rules (BRUL-01 to BRUL-08)").font.color.rgb = RGBColor(30, 58, 138)

    bruls = [
        ("BRUL-01 (ROAS Mathematical Definition):", "Return on Ad Spend (ROAS) is strictly defined as Attributed Revenue divided by Ad Spend. If Ad Spend equals zero (e.g. pure organic search), ROAS is defined as 'N/A' or mathematically undefined, not infinite."),
        ("BRUL-02 (CAC Mathematical Definition):", "Customer Acquisition Cost (CAC) is strictly defined as Total Channel Spend divided by New Customer Conversions attributed to that channel within the selected period."),
        ("BRUL-03 (Underperforming Ad Identification):", "A campaign or ad set is automatically classified as 'Underperforming' if Actual ROAS < 2.0x OR Actual CAC > ₹1,500 over a rolling 7-day evaluation window."),
        ("BRUL-04 (Attribution Model Credit Allocation):", "First-Touch allocates 100% credit to the first recorded touchpoint in a multi-touch journey. Last-Touch allocates 100% credit to the final touchpoint prior to purchase. Linear allocates exactly 1/N credit to each of the N touchpoints in the path."),
        ("BRUL-05 (Statistical Confidence Standard):", "All inferential hypothesis tests must maintain a 95% confidence level (alpha = 0.05). The null hypothesis (H0: No difference between channels) is rejected if and only if p < 0.05."),
        ("BRUL-06 (Currency & Precision Standardization):", "All financial metrics must be rendered in Indian Rupees (INR / ₹). Monetary values and percentages must be displayed rounded to exactly two decimal places (e.g., ₹1,420.50, 3.45%)."),
        ("BRUL-07 (Outlier Classification Threshold):", "Any spend or revenue observation deviating by more than 3 standard deviations (Z > 3) from the channel 30-day rolling mean or falling outside 1.5x IQR is tagged as a statistical anomaly."),
        ("BRUL-08 (Database Integrity Invariants):", "All monetary spend, impression counts, click counts, and revenue amounts must be strictly non-negative (>= 0). Any record violating this rule is rejected during ingestion.")
    ]
    for b_title, b_desc in bruls:
        p_rule = doc.add_paragraph()
        p_rule.add_run(f"• {b_title} ").bold = True
        p_rule.add_run(b_desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 9. Data Requirements
    h9 = doc.add_heading(level=1)
    h9.add_run("9. Data Architecture & Telemetry Requirements").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To satisfy all analytical and business requirements, the application ingests and manages five core relational entities within SQLite (data/app.db):"
    )

    data_req_table = doc.add_table(rows=6, cols=4)
    data_req_headers = ["Entity Name", "Entity Classification", "Primary Business Fields Captured", "Refresh Cadence & Source"]
    for j, dh in enumerate(data_req_headers):
        data_req_table.rows[0].cells[j].text = dh

    data_req_info = [
        ("Channel Master", "Dimension Table", "Channel_ID (PK), Channel_Name, Channel_Category, Benchmark_CAC, Benchmark_ROAS", "Static / Master (Step 10 Synthetic)"),
        ("Campaign Master", "Dimension Table", "Campaign_ID (PK), Channel_ID (FK), Campaign_Name, Campaign_Type, Target_ROAS, Target_CAC", "Monthly / Master (Step 10 Synthetic)"),
        ("Daily Campaign Performance", "Fact Table (Core)", "Fact_ID (PK), Date, Campaign_ID (FK), Spend, Impressions, Clicks, Conversions, Revenue", "Daily Batch Aggregates (12 Months)"),
        ("Customer Master", "Dimension Table", "Customer_ID (PK), Full_Name, Age_Group, Gender, City_Tier, First_Touch_Channel", "Synthetic Master (250+ Customers)"),
        ("Customer Touchpoints", "Fact Table (Attribution)", "Touchpoint_ID (PK), Customer_ID (FK), Journey_ID, Channel_ID (FK), Sequence_Order, Converted_Flag", "Event-level touchpoint sequences")
    ]
    for r_idx, r_data in enumerate(data_req_info, start=1):
        for c_idx, val in enumerate(r_data):
            data_req_table.rows[r_idx].cells[c_idx].text = val
    style_table(data_req_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 10. Assumptions, Constraints & Dependencies
    h10 = doc.add_heading(level=1)
    h10.add_run("10. Assumptions, Constraints & Dependencies").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "• Assumptions: Synthetic data generation reliably mirrors Indian e-commerce distribution patterns (CTR 1.2%–4.8%, CVR 1.5%–4.2%, ROAS 1.2x–4.5x). SQLite provides microsecond analytical query times for ~1,000 to 10,000 rows without concurrency degradation.\n"
        "• Constraints: Execution strictly constrained to Python 3.12, Streamlit, SQLite (data/app.db), Plotly, and pytest. Deployment bound to Streamlit Community Cloud (free-tier resource caps of 1 vCPU and 1 GB RAM).\n"
        "• Dependencies: Python libraries: streamlit, pandas, numpy, scipy, statsmodels, plotly, pytest, openpyxl, python-docx."
    )

    # 11. Risks & Mitigation Matrix
    h11 = doc.add_heading(level=1)
    h11.add_run("11. Risk Management & Mitigation Matrix").font.color.rgb = RGBColor(30, 58, 138)

    risk_table = doc.add_table(rows=5, cols=5)
    risk_headers = ["Risk ID", "Risk Description", "Severity", "Likelihood", "Mitigation Strategy"]
    for j, rh in enumerate(risk_headers):
        risk_table.rows[0].cells[j].text = rh

    risk_data = [
        ("RSK-01", "Synthetic data lacks realistic statistical correlations or variance, invalidating hypothesis testing.", "High", "Low", "Use seeded mathematical distributions (Beta/Gamma) in Step 10 to embed intentional, realistic conversion variance."),
        ("RSK-02", "Streamlit UI recalculation latency exceeds 3 seconds during dynamic filter changes.", "Medium", "Medium", "Implement @st.cache_data decorators and pre-aggregate database views to achieve sub-second execution."),
        ("RSK-03", "Data quality issues (null UTMs, duplicate keys) distort analytical aggregations.", "High", "Low", "Enforce strict SQLite schema constraints (PK, FK, NOT NULL) and automated data cleaning in Step 10 and Step 15."),
        ("RSK-04", "Streamlit Community Cloud deployment fails due to unpinned dependencies or path mismatches.", "High", "Low", "Pin exact package versions in requirements.txt and use pathlib.Path(__file__).parent for all filesystem references.")
    ]
    for r_idx, r_data in enumerate(risk_data, start=1):
        for c_idx, val in enumerate(r_data):
            risk_table.rows[r_idx].cells[c_idx].text = val
    style_table(risk_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 12. Glossary
    h12 = doc.add_heading(level=1)
    h12.add_run("12. Project Glossary & Acronyms").font.color.rgb = RGBColor(30, 58, 138)

    gloss_table = doc.add_table(rows=10, cols=2)
    gloss_headers = ["Term / Acronym", "Authoritative Business Definition"]
    for j, gh in enumerate(gloss_headers):
        gloss_table.rows[0].cells[j].text = gh

    gloss_data = [
        ("ROAS (Return on Ad Spend)", "Financial efficiency metric calculating revenue generated per rupee spent on advertising (Revenue / Spend)."),
        ("CAC (Customer Acquisition Cost)", "The total marketing capital expended to acquire a single paying customer (Spend / Conversions)."),
        ("CTR (Click-Through Rate)", "The proportion of total ad impressions that resulted in an ad click (Clicks / Impressions)."),
        ("CVR (Conversion Rate)", "The percentage of visitors or clicks that successfully completed a purchase conversion (Conversions / Clicks)."),
        ("AOV (Average Order Value)", "The average monetary revenue generated per completed customer transaction (Total Revenue / Total Orders)."),
        ("LTV (Customer Lifetime Value)", "The projected cumulative revenue generated by a customer throughout their active purchasing relationship."),
        ("First-Touch Attribution", "An attribution rule allocating 100% of conversion and revenue credit to the initial customer discovery channel."),
        ("Last-Touch Attribution", "An attribution rule allocating 100% of conversion and revenue credit to the final touchpoint channel prior to purchase."),
        ("Linear Attribution", "An attribution rule dividing conversion credit equally across all touchpoints involved in the customer journey.")
    ]
    for r_idx, r_data in enumerate(gloss_data, start=1):
        for c_idx, val in enumerate(r_data):
            gloss_table.rows[r_idx].cells[c_idx].text = val
    style_table(gloss_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # 13. Sign-off Block
    h13 = doc.add_heading(level=1)
    h13.add_run("13. Formal BRD Sign-Off & Approvals").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "By signing below, the project stakeholders confirm that this Business Requirements Document (BRD v1.0) accurately reflects "
        "the operational requirements, business rules, and analytical scope of the Channel Wise Marketing Performance project. "
        "This baseline document authorizes transition to Step 6 (Product Requirements Document - PRD)."
    )

    sign_table = doc.add_table(rows=4, cols=4)
    sign_headers = ["Stakeholder Name", "Organizational / Academic Role", "Approval Signature", "Approval Date"]
    for j, sh in enumerate(sign_headers):
        sign_table.rows[0].cells[j].text = sh

    sign_data = [
        ("Vaishnavi Singh", "Learner / Lead Business Analyst", "[Signed Digitally]", "20-Nov-2026"),
        ("Prof. Kalpana Dwivedi", "Academic Project Mentor (ABES)", "[Pending Review]", "20-Nov-2026"),
        ("Marketing Leadership Sponsor", "Chief Marketing Officer (Proxy / Role)", "[Approved]", "20-Nov-2026")
    ]
    for r_idx, r_data in enumerate(sign_data, start=1):
        for c_idx, val in enumerate(r_data):
            sign_table.rows[r_idx].cells[c_idx].text = val
    style_table(sign_table)

    # Save document
    path1 = os.path.abspath("docs/05_BRD.docx")
    doc.save(path1)
    print(f"Saved: {path1}")

if __name__ == "__main__":
    generate_step5_brd()
