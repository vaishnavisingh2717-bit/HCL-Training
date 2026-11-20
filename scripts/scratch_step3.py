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

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

def generate_step3_doc():
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
    title_run = title_p.add_run("STEP 3: BUSINESS NEEDS ANALYSIS & REQUIREMENTS (BR)\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nLearner: Vaishnavi Singh (2503201450157) | Mentor: Prof. Kalpana Dwivedi\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Block
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Name & ID", "Step 3 — Needs Analysis (03_Needs_Analysis.docx)"),
        ("Requirement Formulation Standard", "'The business needs to <capability> so that <benefit>.' (WHAT, not HOW)"),
        ("MoSCoW Prioritization", "8 Must (66.7%), 3 Should (25.0%), 1 Could (8.3%), 0 Won't"),
        ("Traceability Baseline", "Every BR derived from HMW-01 to HMW-10; 100% testable criteria")
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

    doc.add_paragraph().paragraph_format.space_after = Pt(16)

    # Section 1: Needs Analysis Methodology
    h1 = doc.add_heading(level=1)
    r1 = h1.add_run("1. Business Needs Analysis Methodology")
    r1.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "A rigorous Needs Analysis translates design-thinking opportunity statements (HMW-01 to HMW-10) into formal, testable Business Requirements (BR-01 to BR-12). "
        "In strict compliance with business analysis best practices, each requirement articulates WHAT capability the business requires and WHY it delivers value, "
        "strictly avoiding premature technical prescriptions (such as SQL scripts, Streamlit widgets, or Python syntax)."
    )

    # Section 2: Business Requirements Table
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("2. Business Requirements Specification Table (BR-01 to BR-12)")
    r2.font.color.rgb = RGBColor(30, 58, 138)

    br_table = doc.add_table(rows=13, cols=6)
    br_headers = ["BR ID", "Business Requirement (Capability & Benefit)", "Primary Stakeholder", "Source HMW", "Priority", "Measurement & Verification Method"]
    for j, h in enumerate(br_headers):
        br_table.rows[0].cells[j].text = h

    br_rows = [
        ("BR-01", "The business needs to consolidate and compute cross-channel Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC) across Google Ads, Meta Ads, Email, SEO, and Influencer marketing so that marketing leadership can evaluate blended and channel-specific capital efficiency.",
         "SH-01 (CMO)", "HMW-01", "Must", "Verification that channel ROAS and CAC match audited aggregate formulas (Revenue/Spend and Spend/Conversions) with 100% calculation accuracy across selected date ranges."),
        
        ("BR-02", "The business needs to automatically identify and flag campaigns and ad sets that breach defined performance thresholds (ROAS < 2.0x or CAC > ₹1,500) so that campaign managers can immediately halt budget wastage on underperforming ad sets.",
         "SH-02 (Paid Ads Lead)", "HMW-02", "Must", "100% of campaigns exceeding target CAC or failing minimum ROAS are surfaced in an actionable exception table."),
        
        ("BR-03", "The business needs to track and decompose Click-Through Rates (CTR) and Conversion Rates (CVR) across marketing channels and campaign types over time so that growth teams can pinpoint whether funnel drop-offs are driven by ad creative fatigue or landing page friction.",
         "SH-02 (Paid Ads Lead)", "HMW-03", "Must", "Daily and aggregated CTR (Clicks/Impressions) and CVR (Conversions/Clicks) displayed across all campaign types with drill-down capability."),
        
        ("BR-04", "The business needs to analyze customer repeat order history and calculate Customer Lifetime Value relative to acquisition cost (LTV:CAC ratio) per acquisition channel so that retention teams can prioritize channels that attract loyal, high-value customers.",
         "SH-04 (Retention Lead)", "HMW-04", "Should", "Average customer order value, order frequency, total cumulative revenue, and LTV:CAC ratio computed and benchmarked per channel (>3:1 target)."),
        
        ("BR-05", "The business needs to evaluate day-of-week and monthly performance patterns in spend, conversion volume, and ROAS so that media planners can schedule ad delivery and adjust day-parting bids to align with peak buyer conversion windows.",
         "SH-02 (Paid Ads Lead)", "HMW-05", "Should", "Performance variance across the 7 days of the week and 12 calendar months visualized with rolling 7-day trend baselines."),
        
        ("BR-06", "The business needs to segment channel conversion volume and average order values by customer demographics (Age Group, Gender, City Tier) so that marketing teams can tailor channel copy, targeting, and promotional offers to high-performing audience cohorts.",
         "SH-04 (Retention Lead)", "HMW-06", "Could", "Conversion count and revenue aggregated and compared across defined demographic segments (Tier 1 vs Tier 2/3, Age brackets)."),
        
        ("BR-07", "The business needs to perform inferential hypothesis testing (Two-sample proportion tests, ANOVA, Chi-Square) on channel conversion rates so that marketing analysts and leadership can confirm whether channel efficiency gaps are statistically significant (p < 0.05) rather than random sampling noise.",
         "SH-03 (Analyst) / SH-05 (Mentor)", "HMW-07", "Must", "Computed test statistics (Z-score, F-statistic, Chi-Square), p-values, and automated plain-language rejection/acceptance of null hypotheses displayed with 95% confidence intervals."),
        
        ("BR-08", "The business needs to compare multi-touch attribution distribution (First-Touch, Last-Touch, and Equal Linear models) across customer conversion journeys so that marketing leadership can properly value top-of-funnel discovery channels alongside closing channels.",
         "SH-01 (CMO) / SH-03 (Analyst)", "HMW-08", "Must", "Channel revenue and conversion share calculated and compared side-by-side under First-Touch, Last-Touch, and Linear models for completed customer journeys."),
        
        ("BR-09", "The business needs to simulate cross-channel budget reallocation scenarios under a fixed total marketing spend based on historical marginal returns so that executive leadership can project potential revenue and ROAS uplifts before committing capital.",
         "SH-01 (CMO)", "HMW-09", "Must", "What-if budget shift simulator showing projected incremental revenue and portfolio ROAS delta based on historical diminishing return curves."),
        
        ("BR-10", "The business needs to audit marketing data quality by automatically detecting untracked UTM parameters, duplicate transaction IDs, missing foreign keys, and statistical spend outliers so that all downstream marketing analytics are based on reliable and auditable data.",
         "SH-03 (Analyst) / SH-05 (Mentor)", "HMW-10", "Must", "Summary audit report displaying completeness percentages (>98%), outlier counts (Z > 3 or 1.5x IQR), and zero unhandled integrity violations."),
        
        ("BR-11", "The business needs to dynamically filter performance data by date ranges, acquisition channels, and campaign objectives so that stakeholders can isolate specific operational periods and campaign cohorts without altering baseline data.",
         "SH-01 (CMO) / SH-02 (Paid Lead) / SH-03 (Analyst)", "HMW-01, HMW-03", "Must", "All analytical metrics and charts update dynamically when filter criteria are applied, reflecting isolated subsets accurately."),
        
        ("BR-12", "The business needs to export filtered analytical summaries, raw campaign records, and data audit logs into standardized tabular formats (CSV/Excel) so that analysts can conduct offline reporting, executive briefing, and compliance documentation.",
         "SH-03 (Analyst) / SH-05 (Mentor)", "HMW-01, HMW-10", "Should", "Downloadable file generation verifying that exported records match filtered UI data row-for-row and column-for-column.")
    ]

    for r_idx, r_data in enumerate(br_rows, start=1):
        for c_idx, val in enumerate(r_data):
            br_table.rows[r_idx].cells[c_idx].text = val
    style_table(br_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 3: Priority Breakdown
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("3. Requirement Prioritization Analysis (MoSCoW)")
    r3.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Requirements are categorized using the industry-standard MoSCoW framework to ensure high operational delivery focus for the MVP:"
    )

    moscow_p = doc.add_paragraph()
    moscow_p.add_run("• MUST Have (8 Requirements / 66.7%): ").bold = True
    moscow_p.add_run("BR-01 (ROAS/CAC), BR-02 (Threshold Alerts), BR-03 (CTR/CVR Funnel), BR-07 (Statistical Hypothesis Tests), BR-08 (Attribution Modeling), BR-09 (Budget Optimization), BR-10 (Data Quality Auditing), BR-11 (Dynamic Filtering). Essential for functional MVP viability.\n")
    
    moscow_p.add_run("• SHOULD Have (3 Requirements / 25.0%): ").bold = True
    moscow_p.add_run("BR-04 (LTV:CAC Analysis), BR-05 (Seasonality/Day-of-Week Trends), BR-12 (Data Export). High business value, enhancing analytical depth.\n")

    moscow_p.add_run("• COULD Have (1 Requirement / 8.3%): ").bold = True
    moscow_p.add_run("BR-06 (Demographic Segmentation). Valuable cohort insight, implemented if data granularity permits.\n")

    moscow_p.add_run("• WON'T Have (Out-of-Scope for MVP): ").bold = True
    moscow_p.add_run("Live external ad network OAuth API connections, live streaming Kafka consumers, or automated programmatic ad bidding execution.")

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 4: Data Required by Each Requirement
    h4 = doc.add_heading(level=1)
    r4 = h4.add_run("4. Data Attributes Required by Business Requirement")
    r4.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Each requirement is explicitly mapped to the necessary business entities and data fields to guarantee that data engineering in Step 8 and Step 10 satisfies all analytical dependencies:"
    )

    data_table = doc.add_table(rows=13, cols=4)
    data_headers = ["BR ID", "Requirement Short Title", "Primary Data Entities", "Mandatory Data Fields Required"]
    for j, h in enumerate(data_headers):
        data_table.rows[0].cells[j].text = h

    data_mappings = [
        ("BR-01", "Cross-Channel ROAS & CAC", "Daily Campaign Fact, Channel Master", "Spend, Revenue, Conversions, Channel_Name, Date"),
        ("BR-02", "Threshold Exception Flagging", "Campaign Master, Daily Performance Fact", "Campaign_ID, Actual_ROAS, Target_ROAS, Actual_CAC, Target_CAC"),
        ("BR-03", "CTR & CVR Funnel Decomposition", "Daily Campaign Fact, Campaign Master", "Impressions, Clicks, Conversions, Campaign_Type, Channel_Name"),
        ("BR-04", "Customer LTV:CAC Ratio Analysis", "Customer Master, Order Transactions Fact", "Customer_ID, Acquisition_Channel, Order_Value, Order_Date, Repeat_Orders"),
        ("BR-05", "Temporal Seasonality Trends", "Daily Campaign Fact", "Date, Day_of_Week, Month, Daily_Spend, Daily_Revenue, Conversions"),
        ("BR-06", "Demographic Cohort Segmentation", "Customer Master, Order Transactions Fact", "Customer_Age, Gender, City_Tier, Acquisition_Channel, Order_Value"),
        ("BR-07", "Inferential Statistical Tests", "Daily Campaign Fact, Channel Aggregates", "Channel_Name, Clicks/Visitors, Conversions, Conversion_Proportions"),
        ("BR-08", "Multi-Touch Attribution Modeling", "Touchpoints Event Fact, Customer Orders", "Journey_ID, Touchpoint_Sequence, Channel_Name, Conversion_Flag, Revenue"),
        ("BR-09", "Prescriptive Budget Simulation", "Channel Historical Fact, Elasticity Model", "Historical_Spend, Historical_Revenue, Marginal_ROAS, Budget_Limit"),
        ("BR-10", "Data Quality & Integrity Auditing", "All System Tables & Ingestion Logs", "UTM_Source, UTM_Campaign, Transaction_ID, Spend, Revenue, Null_Flags"),
        ("BR-11", "Multi-Dimensional Filtering", "All Fact & Dimension Tables", "Date_Range, Channel_Selection, Campaign_Objective"),
        ("BR-12", "Data Export & Compliance", "Filtered Tabular Data Aggregates", "Consolidated Tabular Export Columns, Summary Metrics")
    ]

    for r_idx, r_data in enumerate(data_mappings, start=1):
        for c_idx, val in enumerate(r_data):
            data_table.rows[r_idx].cells[c_idx].text = val
    style_table(data_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 5: Gap Analysis
    h5 = doc.add_heading(level=1)
    r5 = h5.add_run("5. Bidirectional Traceability & Gap Analysis")
    r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "A formal gap audit was conducted to confirm complete alignment between Design Thinking (Step 2) and Business Requirements (Step 3):"
    )

    p_gap1 = doc.add_paragraph()
    p_gap1.add_run("A. HMWs with no corresponding BR: ").bold = True
    p_gap1.add_run("NONE (0 Gaps). Every statement from HMW-01 through HMW-10 is directly addressed by at least one business requirement.")

    p_gap2 = doc.add_paragraph()
    p_gap2.add_run("B. BRs with no originating HMW: ").bold = True
    p_gap2.add_run("NONE (0 Gaps). All requirements BR-01 to BR-12 trace directly back to verified stakeholder HMW statements. (BR-11 and BR-12 provide cross-cutting enablement for HMW-01, HMW-03, and HMW-10).")

    p_gap3 = doc.add_paragraph()
    p_gap3.add_run("C. Testability Verification: ").bold = True
    p_gap3.add_run("100% of defined business requirements feature unambiguous verification criteria that will be directly validated in Step 13 (Test Cases) and Step 15 (Pytest Suite).")

    # Save document
    path1 = os.path.abspath("docs/03_Needs_Analysis.docx")
    doc.save(path1)
    print(f"Saved: {path1}")

if __name__ == "__main__":
    generate_step3_doc()
