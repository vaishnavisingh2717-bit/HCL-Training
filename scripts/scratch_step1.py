import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, header_bg="1E3A8A", alt_bg="F1F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        # Prevent row split across pages
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        if i == 0:
            # Header repeat
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            for cell in row.cells:
                set_cell_background(cell, header_bg)
                set_cell_margins(cell, 120, 120, 150, 150)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(10)
        else:
            bg = alt_bg if i % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                set_cell_background(cell, bg)
                set_cell_margins(cell, 100, 100, 150, 150)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = RGBColor(30, 41, 59)

def create_document():
    doc = docx.Document()
    
    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Styles
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Title Block
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("STEP 1: PROJECT UNDERSTANDING DOCUMENT\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138) # Deep Navy

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nDomain: Data Analytics | MCA 2nd Year Capstone\n")
    sub_run.font.size = Pt(13)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    metadata = [
        ("Project Name", "Channel Wise Marketing Performance"),
        ("Learner / Author", "Vaishnavi Singh (Roll No: 2503201450157)"),
        ("Academic Program", "MCA 2nd Year (Data Analytics Specialization), ABES"),
        ("Project Mentor", "Prof. Kalpana Dwivedi"),
        ("Target Product Stack", "Python 3.12, Streamlit, SQLite (app.db), Plotly, Tailwind UI Design, pytest"),
        ("Delivery Milestone", "Step 1 of 18 (Project Understanding & Operational Scoping)")
    ]
    for idx, (k, v) in enumerate(metadata):
        cell_k, cell_v = meta_table.rows[idx].cells
        cell_k.text = k
        cell_v.text = v
        set_cell_background(cell_k, "F1F5F9")
        set_cell_background(cell_v, "FFFFFF")
        set_cell_margins(cell_k, 80, 80, 120, 120)
        set_cell_margins(cell_v, 80, 80, 120, 120)
        cell_k.paragraphs[0].runs[0].font.bold = True
        cell_k.paragraphs[0].runs[0].font.size = Pt(10)
        cell_v.paragraphs[0].runs[0].font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)

    # Section 1: Executive Summary & Business Context
    h1 = doc.add_heading(level=1)
    r1 = h1.add_run("1. Executive Summary & Business Context")
    r1.font.color.rgb = RGBColor(30, 58, 138)
    
    doc.add_paragraph(
        "In modern multi-channel digital marketing, businesses allocate capital across a diverse portfolio of paid and organic customer acquisition channels, "
        "including Google Search & Display Ads, Meta (Facebook & Instagram) Paid Social, Targeted Email Marketing Campaigns, Organic Search (SEO), "
        "and Influencer Partnerships. However, marketing leaders and growth analysts frequently face high data fragmentation, inconsistent channel metrics, "
        "and opaque attribution mechanisms. Disparate ad platforms report metrics in isolated silos with varying attribution windows, obscuring the true blended "
        "and marginal efficiency of customer acquisition."
    )
    doc.add_paragraph(
        "This project develops an enterprise-grade, interactive Data Analytics and Business Intelligence application designed to consolidate cross-channel "
        "marketing telemetry into an optimized SQLite analytical data store, compute standardized financial and performance KPIs (such as Return on Ad Spend [ROAS], "
        "Customer Acquisition Cost [CAC], Click-Through Rate [CTR], and Conversion Rate [CVR]), perform inferential statistical testing to validate channel "
        "performance disparities, and deliver actionable budget optimization recommendations through a responsive Streamlit dashboard styled with modern UI principles."
    )

    # Section 2: Problem Statement
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("2. Problem Statement")
    r2.font.color.rgb = RGBColor(30, 58, 138)

    callout_p = doc.add_paragraph()
    callout_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    callout_run = callout_p.add_run(
        "“Marketing leaders and performance analysts struggle with fragmented cross-channel data silos, preventing timely visibility into channel-level "
        "Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC), which leads to inefficient budget allocation and delayed campaign optimization.”"
    )
    callout_run.font.bold = True
    callout_run.font.italic = True
    callout_run.font.size = Pt(11.5)
    callout_run.font.color.rgb = RGBColor(15, 23, 42)

    # Section 3: Intended Audience & User Profiles
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("3. Intended Audience & User Profiles")
    r3.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "The application serves distinct business and operational personas within an organization, catering to varying analytical granularities and decision cadence:"
    )

    user_table = doc.add_table(rows=5, cols=4)
    headers = ["User Role / Persona", "Access Frequency", "Primary Objective & Workflow", "Key Platform Features Used"]
    for j, h in enumerate(headers):
        user_table.rows[0].cells[j].text = h
    
    user_data = [
        ("Performance Marketing Manager", "Daily (1–2 times/day)", "Monitor real-time campaign pacing, identify CTR/CVR anomalies, adjust tactical ad spend, and halt bleeding ad sets.", "Real-time KPI cards, Channel comparison filters, Daily spend tracking"),
        ("Lead Marketing Data Analyst", "Weekly (3–5 times/week)", "Conduct deep-dive Exploratory Data Analysis (EDA), run statistical hypothesis tests (ANOVA, Two-sample tests), evaluate cohort quality.", "Statistical testing module, Correlation matrices, Attribution explorer, Data export"),
        ("Chief Marketing Officer (CMO)", "Bi-Weekly / Monthly", "Review macro-level acquisition efficiency, evaluate blended ROAS, determine strategic quarterly budget shifts between channels.", "Executive Summary dashboard, High-level ROAS trends, Budget reallocation simulator"),
        ("Academic Evaluator / Project Mentor", "Periodic / Milestone-based", "Evaluate analytical soundness, relational database integrity, code structure, statistical rigor, and documentation traceability.", "Data quality validation page, Pytest verification suite, Schema documentation")
    ]
    for row_idx, data_row in enumerate(user_data, start=1):
        for col_idx, text in enumerate(data_row):
            user_table.rows[row_idx].cells[col_idx].text = text
    style_table(user_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 4: Stakeholder Matrix
    h4 = doc.add_heading(level=1)
    r4 = h4.add_run("4. Stakeholder Matrix (SH-01 to SH-05)")
    r4.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Stakeholders are categorized and evaluated based on their organizational influence, analytical interest, and the specific strategic or operational decisions they make using the solution:"
    )

    sh_table = doc.add_table(rows=6, cols=6)
    sh_headers = ["SH ID", "Stakeholder", "Organizational Role", "Interest", "Influence", "Decision Made With The App"]
    for j, h in enumerate(sh_headers):
        sh_table.rows[0].cells[j].text = h

    sh_data = [
        ("SH-01", "Chief Marketing Officer (CMO)", "Executive Strategy Sponsor", "High", "High", "Quarterly & annual budget reallocation across paid vs. organic acquisition channels."),
        ("SH-02", "Paid Acquisition / Performance Lead", "Tactical Ad Operations Lead", "High", "High", "Day-to-day campaign bid tuning, scaling high-ROAS ad creatives, pausing underperforming ad sets."),
        ("SH-03", "Senior Marketing Data Analyst", "Analytical Modeler & Reporter", "High", "Medium", "Establishing attribution splits, evaluating statistical significance of channel performance, monitoring data drift."),
        ("SH-04", "Organic & Retention Marketing Lead", "Inbound & Email Channel Owner", "Medium", "Medium", "Optimizing email automation triggers, content marketing cadence, and assessing organic customer LTV."),
        ("SH-05", "Academic Project Mentor (Prof. Kalpana Dwivedi)", "Project Evaluation Guide", "High", "High", "Validating SDLC milestone compliance, technical architecture, statistical validity, and viva readiness.")
    ]
    for row_idx, data_row in enumerate(sh_data, start=1):
        for col_idx, text in enumerate(data_row):
            sh_table.rows[row_idx].cells[col_idx].text = text
    style_table(sh_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 5: Operational Questions (OQ-01 to OQ-10) & Data Feasibility Check
    h5 = doc.add_heading(level=1)
    r5 = h5.add_run("5. Operational Questions & Data Feasibility Verification (OQ-01 to OQ-10)")
    r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To ensure the application delivers concrete business value, ten operational questions (OQs) have been formulated. "
        "Each question is rigorously validated against available data attributes to confirm mathematical and analytical feasibility:"
    )

    oq_table = doc.add_table(rows=11, cols=5)
    oq_headers = ["OQ ID", "Operational Question", "Required Data Attributes", "Key Analytical / Statistical Technique", "Data Feasibility Check"]
    for j, h in enumerate(oq_headers):
        oq_table.rows[0].cells[j].text = h

    oq_data = [
        ("OQ-01", "What is the overall and channel-wise Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC)?", "Spend, Revenue, Conversions, Channel_Name", "Sum aggregation, Ratio calculations (Revenue/Spend, Spend/Conversions)", "Feasible: Available directly in daily campaign aggregates table."),
        ("OQ-02", "How do Click-Through Rates (CTR) and Conversion Rates (CVR) vary across acquisition channels and campaign types over time?", "Impressions, Clicks, Conversions, Channel_Name, Campaign_Type, Date", "Time-series trend analysis, Funnel conversion ratios (Clicks/Impr, Conv/Clicks)", "Feasible: Supported by daily grain fact table and campaign master."),
        ("OQ-03", "Which marketing channels generate the highest Customer Lifetime Value (LTV) relative to initial acquisition cost (LTV:CAC ratio)?", "Customer_ID, Acquisition_Channel, Order_Value, Repeat_Order_Count, Initial_CAC", "Cohort aggregation, Cumulative customer revenue, Ratio analysis", "Feasible: Supported by joined customer transaction history."),
        ("OQ-04", "What are the temporal seasonality trends (day-of-week, monthly cycles), and are there recurring performance dips?", "Date, Day_of_Week, Spend, Revenue, Conversions, Channel", "Calendar decomposition, 7-day moving averages, seasonal variance analysis", "Feasible: Fact table includes continuous daily records spanning 12 months."),
        ("OQ-05", "How do customer demographic segments (Age Group, Gender, City Tier) correlate with channel conversion efficacy?", "Customer_Age, Gender, City_Tier, Acquisition_Channel, Conversion_Status, Revenue", "Cross-tabulation, Demographic segmentation, Bar/Sunburst visualizations", "Feasible: Supported by customer master and conversion touchpoint data."),
        ("OQ-06", "Is there a statistically significant difference in conversion rates between paid acquisition channels (e.g. Google Ads vs. Meta)?", "Channel_Name, Clicks, Conversions, Visitor_Count", "Two-sample Z-test for proportions, Chi-Square test of independence (p < 0.05)", "Feasible: Binary/proportional conversion data available per channel."),
        ("OQ-07", "Which specific campaigns or ad sets are underperforming against defined benchmark thresholds (ROAS < 2.0x or CAC > ₹1,500)?", "Campaign_ID, Campaign_Name, Actual_ROAS, Benchmark_ROAS, Actual_CAC, Benchmark_CAC", "Threshold filtering, Conditional rule evaluations, Outlier flagging", "Feasible: Benchmarks mapped against campaign performance metrics."),
        ("OQ-08", "What is the channel attribution breakdown across multi-touch customer journeys (First-Touch vs. Last-Touch vs. Linear)?", "Journey_ID, Touchpoint_Sequence, Touchpoint_Channel, Conversion_Flag, Order_Value", "Rule-based attribution modeling (First-Touch, Last-Touch, Equal Linear)", "Feasible: Touchpoint event log captures sequence of customer channel interactions."),
        ("OQ-09", "What is the optimal budget reallocation scenario across channels to maximize conversions under a fixed total marketing spend?", "Channel_Name, Historical_Spend, Marginal_ROAS, Efficiency_Curve, Budget_Constraint", "Constrained optimization / heuristic marginal returns simulation", "Feasible: Historical elasticity curves derived from daily spend and revenue data."),
        ("OQ-10", "What data quality anomalies, untracked UTM parameters, or extreme spend/revenue outliers exist in marketing records?", "UTM_Source, UTM_Campaign, Spend, Revenue, Record_Timestamp, Null_Flags", "Data completeness audits, Interquartile Range (IQR) / Z-score outlier detection", "Feasible: Built-in schema validation and quality auditing scripts.")
    ]
    for row_idx, data_row in enumerate(oq_data, start=1):
        for col_idx, text in enumerate(data_row):
            oq_table.rows[row_idx].cells[col_idx].text = text
    style_table(oq_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 6: SMART Objectives
    h6 = doc.add_heading(level=1)
    r6 = h6.add_run("6. Project Objectives (SMART)")
    r6.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "The project execution is governed by six Specific, Measurable, Achievable, Relevant, and Time-bound (SMART) objectives:"
    )

    obj_p1 = doc.add_paragraph()
    obj_p1.add_run("• OBJ-01 (Multi-Channel Integration): ").bold = True
    obj_p1.add_run("Develop and deploy an end-to-end data pipeline that ingests, cleans, and standardizes multi-channel marketing performance telemetry across 5 distinct channels (Google Ads, Meta Ads, Email, SEO, Influencer) within a normalized SQLite database schema by Step 10.")

    obj_p2 = doc.add_paragraph()
    obj_p2.add_run("• OBJ-02 (Interactive Dashboard Performance): ").bold = True
    obj_p2.add_run("Build an interactive, multi-page Streamlit analytical dashboard delivering sub-2.0-second page response times and sub-1.0-second cached filter recalculations across 30+ dimensional metrics by Step 12.")

    obj_p3 = doc.add_paragraph()
    obj_p3.add_run("• OBJ-03 (Inferential Statistical Rigor): ").bold = True
    obj_p3.add_run("Implement automated inferential statistical evaluations (including Two-sample proportion Z-tests, ANOVA for cross-channel variance, and Pearson/Spearman correlation matrices) to scientifically validate channel efficiency differences at a 95% confidence level (alpha = 0.05) by Step 15.")

    obj_p4 = doc.add_paragraph()
    obj_p4.add_run("• OBJ-04 (Decision Optimization & Modeling): ").bold = True
    obj_p4.add_run("Formulate an actionable budget reallocation model capable of identifying underperforming ad sets and projecting a potential 10% to 15% efficiency improvement in blended portfolio ROAS under fixed spend constraints by Step 15.")

    obj_p5 = doc.add_paragraph()
    obj_p5.add_run("• OBJ-05 (Automated Quality Assurance): ").bold = True
    obj_p5.add_run("Achieve >= 85% automated test coverage using pytest and Streamlit AppTest, verifying database schema constraints, data transformation accuracy, KPI calculation formulas, and UI rendering integrity by Step 15.")

    obj_p6 = doc.add_paragraph()
    obj_p6.add_run("• OBJ-06 (Deployment & Academic Compliance): ").bold = True
    obj_p6.add_run("Deploy the complete, verified solution to Streamlit Community Cloud with a publicly accessible URL and complete all ABES-mandated documentation and presentation slide decks by Step 18.")

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 7: Project Scope (In-Scope vs. Out-of-Scope)
    h7 = doc.add_heading(level=1)
    r7 = h7.add_run("7. Project Scope Definition")
    r7.font.color.rgb = RGBColor(30, 58, 138)

    scope_table = doc.add_table(rows=6, cols=2)
    scope_table.rows[0].cells[0].text = "In-Scope (Deliverables & Features Included)"
    scope_table.rows[0].cells[1].text = "Out-of-Scope (Explicitly Excluded)"

    scope_data = [
        ("Multi-channel synthetic dataset generation covering 5 channels (Google Ads, Meta Ads, Email, SEO, Influencer) with realistic Indian market parameters.", "Live production API integrations with Google Ads API or Meta Graph API (due to credential, billing, and OAuth sandbox barriers)."),
        ("Relational SQLite schema design featuring normalized master tables, a daily campaign performance fact table, and a customer touchpoint table.", "Enterprise cloud data warehousing solutions (e.g., Snowflake, Google BigQuery, AWS Redshift) requiring paid subscription tiers."),
        ("Key performance metric computation: Spend, Impressions, Clicks, CTR, Conversions, CVR, Revenue, ROAS, CAC, AOV, and LTV:CAC ratios.", "Real-time streaming telemetry pipelines (e.g., Apache Kafka, Spark Streaming) requiring distributed server infrastructure."),
        ("Interactive Streamlit dashboard featuring multi-page navigation, Tailwind UI design cards, Plotly charts, and dynamic date/channel filters.", "Complex deep learning neural network bid prediction engines beyond standard statistical and heuristic optimization modeling."),
        ("Inferential statistical hypothesis testing (Z-tests, ANOVA, Chi-Square) and multi-touch attribution modeling (First-Touch, Last-Touch, Linear).", "Transactional payment gateway integration or commercial e-commerce checkout execution.")
    ]
    for row_idx, data_row in enumerate(scope_data, start=1):
        for col_idx, text in enumerate(data_row):
            scope_table.rows[row_idx].cells[col_idx].text = text
    style_table(scope_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 8: Assumptions & Constraints
    h8 = doc.add_heading(level=1)
    r8 = h8.add_run("8. Project Assumptions & Constraints")
    r8.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph("A. Technical & Operational Assumptions:")
    assump_items = [
        "1. Synthetic Marketing Data Integrity: The synthetic dataset generated will reflect realistic B2C e-commerce behavioral distributions (CTR between 1.2%–4.8%, CVR between 1.5%–4.2%, and ROAS between 1.2x–4.5x) consistent with Indian retail industry benchmarks.",
        "2. Database Scalability: SQLite is assumed to be fully sufficient to support the application's scale (~250–500 campaign summaries and 5,000–10,000 customer touchpoint records) with sub-second analytical query performance.",
        "3. Multi-Touch Attribution Sufficiency: Rule-based attribution models (First-Touch, Last-Touch, Linear) provide sufficient managerial insight for cross-channel budget allocation without necessitating algorithmic Markov-chain modeling for the MVP.",
        "4. Streamlit Session State Management: Standard Streamlit session state and @st.cache_data decorators will prevent redundant database queries during interactive user filtering."
    ]
    for item in assump_items:
        p = doc.add_paragraph(item)
        p.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph("B. System & Project Constraints:")
    const_items = [
        "1. Stack Constraint: Strict adherence to Python 3.12, Streamlit, SQLite (data/app.db), Plotly, and pytest. No external proprietary BI tools (Tableau, PowerBI) may be used.",
        "2. Hosting Limits: Deployment is strictly bound to Streamlit Community Cloud (free-tier resource caps of 1 vCPU and 1 GB RAM).",
        "3. Data Privacy & Compliance: 100% synthetic dataset utilization with zero live PII, strictly adhering to academic integrity standards.",
        "4. Development Timeline: Project execution must conform to the 18-step sequential SDLC milestone schedule mandated by ABES curriculum guidelines."
    ]
    for item in const_items:
        p = doc.add_paragraph(item)
        p.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 9: Measurable Success Criteria
    h9 = doc.add_heading(level=1)
    r9 = h9.add_run("9. Measurable Success Criteria (SC-01 to SC-06)")
    r9.font.color.rgb = RGBColor(30, 58, 138)

    sc_table = doc.add_table(rows=7, cols=4)
    sc_headers = ["Criterion ID", "Success Criterion Dimension", "Target Measurement Threshold", "Verification Method"]
    for j, h in enumerate(sc_headers):
        sc_table.rows[0].cells[j].text = h

    sc_data = [
        ("SC-01", "Operational Question Coverage", "100% (All 10 OQs answered directly by visual dashboards or tables)", "Visual and functional audit of Streamlit UI screens against OQ specifications."),
        ("SC-02", "UI Rendering & Query Latency", "Initial page load < 3.0 seconds; Filter refresh < 1.0 second", "Browser network inspector & Streamlit runtime timing profiler."),
        ("SC-03", "Database Relational Integrity", "100% Primary Key uniqueness; 0 Foreign Key orphan records; 0 negative spend/revenue entries", "Automated SQL integrity assertion scripts and pytest schema validation."),
        ("SC-04", "Automated Test Coverage", ">= 85% statement coverage across data loaders, KPI math, and UI components", "pytest --cov=src execution report."),
        ("SC-05", "Cloud Deployment Viability", "Publicly accessible, error-free deployment on Streamlit Community Cloud", "Live cloud URL verification and multi-browser rendering validation."),
        ("SC-06", "Academic & Viva Compliance", "100% complete ABES Project Document and 11-slide presentation deck", "Mentor review, rubric evaluation, and simulated viva defense.")
    ]
    for row_idx, data_row in enumerate(sc_data, start=1):
        for col_idx, text in enumerate(data_row):
            sc_table.rows[row_idx].cells[col_idx].text = text
    style_table(sc_table)

    # Section 10: Sign-Off & Viva Readiness
    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    h10 = doc.add_heading(level=1)
    r10 = h10.add_run("10. Step 1 Sign-Off & Viva Preparation")
    r10.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "This Project Understanding document establishes the baseline for all subsequent analytical and technical deliverables. "
        "Every stakeholder (SH), operational question (OQ), and objective defined here will maintain consistent numbering across Steps 2 through 18."
    )

    # Save to both requested locations
    path1 = os.path.abspath("docs/01_Project_Understanding.docx")
    path2 = os.path.abspath("docs/01_Channel_Wise_Marketing_Performance_Project_Understanding.docx")
    
    doc.save(path1)
    doc.save(path2)
    print(f"Saved: {path1}")
    print(f"Saved: {path2}")

if __name__ == "__main__":
    create_document()
