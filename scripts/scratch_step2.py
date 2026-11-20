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

def generate_step2_doc():
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
    title_run = title_p.add_run("STEP 2: DESIGN THINKING & HOW MIGHT WE (HMW) FRAMEWORK\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nLearner: Vaishnavi Singh (2503201450157) | Mentor: Prof. Kalpana Dwivedi\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Summary
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document ID & Name", "Step 2 — How Might We Statements (02_HMW_Statements.docx)"),
        ("Methodological Framework", "Design Thinking Problem Framing (Empathize -> Define -> Ideate Baseline)"),
        ("Stakeholder Scope", "SH-01 (CMO), SH-02 (Paid Ads), SH-03 (Analyst), SH-04 (Retention), SH-05 (Mentor)"),
        ("Coverage Requirement", "100% of Operational Questions (OQ-01 to OQ-10) directly mapped to HMWs")
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

    # Section 1: Stakeholder Pain Points
    h1 = doc.add_heading(level=1)
    r1 = h1.add_run("1. Detailed Stakeholder Pain Point Analysis")
    r1.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Using the Stakeholder Matrix established in Step 1, a structured design-thinking empathy assessment was conducted. "
        "Each stakeholder group experiences acute friction points that prevent data-driven marketing decisions:"
    )

    stakeholder_pain_points = [
        ("SH-01: Chief Marketing Officer (CMO)", [
            "Cross-Channel Data Fragmentation: Ad performance data is locked in isolated platform silos (Google Ads, Meta Ad Manager, Klaviyo), making it virtually impossible to view true holistic ROI across channels.",
            "Attribution Window Distortion: Reliance on default last-click attribution causes over-investment in bottom-funnel retargeting ads while starving top-of-funnel brand awareness channels.",
            "Capital Allocation Uncertainty: Leadership lacks quantitative visibility into diminishing marginal returns, resulting in suboptimal quarterly budget distribution."
        ]),
        ("SH-02: Performance Marketing Lead", [
            "Delayed Ad Fatigue Detection: Identifying underperforming ad sets where CAC spikes and ROAS drops takes days of manual cross-referencing, leading to wasted ad spend.",
            "Conversion Funnel Ambiguity: Campaign managers struggle to discern whether poor campaign performance is due to low creative engagement (CTR) or downstream landing page friction (CVR).",
            "Inefficient Ad Scheduling: Without granular day-of-week and hourly trend analysis, budgets are distributed evenly, missing peak conversion windows."
        ]),
        ("SH-03: Senior Marketing Data Analyst", [
            "Lack of Statistical Rigor: Management makes major budget decisions based on surface-level conversion differences that may simply be random sampling noise.",
            "Tedious Data Quality Audits: Analysts spend hours cleaning corrupted tracking parameters, missing UTM tags, and anomalous spend spikes before reporting can begin.",
            "Rigid Attribution Reporting: Inability to dynamically toggle and compare First-Touch, Last-Touch, and Linear attribution models for customer conversion paths."
        ]),
        ("SH-04: Retention & Content Marketing Lead", [
            "Acquisition-Retention Disconnect: High upfront ad spend frequently acquires low-value, one-off purchasers instead of repeat, high-LTV customers.",
            "Demographic Blind Spots: Lack of unified insights on how demographic segments (age, gender, tier) convert across organic vs. paid touchpoints."
        ]),
        ("SH-05: Academic Project Mentor (Prof. Kalpana Dwivedi)", [
            "Superficial Analytics Projects: Student projects frequently build decorative UI dashboards without robust relational data design, data validation, or automated tests.",
            "Assertions Lacking Statistical Backing: Failure to formulate and test formal hypotheses with p-values and confidence intervals to validate analytical conclusions."
        ])
    ]

    for sh_title, points in stakeholder_pain_points:
        p_sh = doc.add_paragraph()
        run_sh = p_sh.add_run(sh_title)
        run_sh.bold = True
        run_sh.font.color.rgb = RGBColor(37, 99, 235)
        for pt in points:
            p_pt = doc.add_paragraph(f"• {pt}")
            p_pt.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 2: How Might We (HMW) Statements Table
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("2. 'How Might We' (HMW) Problem Formulation Table")
    r2.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Design-thinking HMW statements reframe observed pain points into constructive, actionable problem spaces without dictating rigid technical implementations. "
        "Every statement is linked to its originating stakeholder and operational question (OQ):"
    )

    hmw_table = doc.add_table(rows=11, cols=5)
    hmw_headers = ["HMW ID", "How Might We Statement", "Primary Stakeholder", "Addressed Pain Point", "Linked OQ"]
    for j, h in enumerate(hmw_headers):
        hmw_table.rows[0].cells[j].text = h

    hmw_rows = [
        ("HMW-01", "How might we provide marketing leadership with a unified, standardized view of channel-level ROAS and CAC to eliminate guesswork during capital allocation?", "SH-01 (CMO)", "Siloed channel data obscuring holistic ROI and true capital efficiency.", "OQ-01"),
        ("HMW-02", "How might we enable campaign managers to detect underperforming ad sets breaching efficiency thresholds (ROAS < 2.0x, CAC > ₹1,500) in time to halt budget waste?", "SH-02 (Paid Ads Lead)", "Delayed detection of bleeding ad sets causing capital loss.", "OQ-07"),
        ("HMW-03", "How might we isolate whether poor channel conversion stems from upstream ad creative disengagement (CTR) or downstream checkout friction (CVR)?", "SH-02 (Paid Ads Lead)", "Inability to pinpoint exact funnel bottlenecks across channels.", "OQ-02"),
        ("HMW-04", "How might we identify which acquisition channels deliver the highest lifetime customer value (LTV:CAC) rather than just cheap initial acquisitions?", "SH-04 (Retention Lead)", "Acquiring low-value one-time buyers instead of high-LTV repeat customers.", "OQ-03"),
        ("HMW-05", "How might we reveal cyclical day-of-week and monthly performance patterns so media planners can optimize ad delivery schedules?", "SH-02 (Paid Ads Lead)", "Inefficient flat ad delivery ignoring temporal/seasonal conversion cycles.", "OQ-04"),
        ("HMW-06", "How might we uncover demographic segment affinities across marketing channels to tailor creative messaging by age group, gender, and city tier?", "SH-04 (Retention Lead)", "Generic campaign targeting ignoring demographic channel preferences.", "OQ-05"),
        ("HMW-07", "How might we statistically validate whether conversion rate differences between marketing channels reflect true performance advantages rather than random sampling variation?", "SH-03 (Analyst) / SH-05 (Mentor)", "Unverified assertions about channel performance without statistical hypothesis testing.", "OQ-06"),
        ("HMW-08", "How might we expose multi-touch attribution credit across customer journeys so top-of-funnel discovery channels are not penalized by last-click bias?", "SH-01 (CMO) / SH-03 (Analyst)", "Last-click attribution undervaluing brand awareness and discovery channels.", "OQ-08"),
        ("HMW-09", "How might we simulate optimal cross-channel budget reallocations based on historical marginal returns to maximize total revenue under fixed spend constraints?", "SH-01 (CMO)", "Uncertainty in reallocating capital across channels to maximize blended ROAS.", "OQ-09"),
        ("HMW-10", "How might we automatically detect tracking errors, missing UTM tags, and anomalous spend spikes before they corrupt marketing reporting?", "SH-03 (Analyst) / SH-05 (Mentor)", "Corrupted analytical insights due to undetected tracking gaps and data quality anomalies.", "OQ-10")
    ]

    for r_idx, r_data in enumerate(hmw_rows, start=1):
        for c_idx, val in enumerate(r_data):
            hmw_table.rows[r_idx].cells[c_idx].text = val
    style_table(hmw_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 3: Priority HMW Selection
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("3. Top 5 Priority HMW Statements & Justification")
    r3.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To ensure high project impact and rigorous academic differentiation, five HMW statements are designated as top priority:"
    )

    priorities = [
        ("1. HMW-01: Unified Channel-Level ROAS & CAC Analysis (Linked to OQ-01)",
         "PRIORITY REASON: Serves as the core financial bedrock of the entire analytics platform, directly addressing executive decision-making on marketing capital efficiency."),
        ("2. HMW-02: Automated Threshold-Based Underperforming Ad Detection (Linked to OQ-07)",
         "PRIORITY REASON: Directly drives cost savings by providing operational leads with an immediate mechanism to halt bleeding campaigns and reallocate wasted budget."),
        ("3. HMW-07: Inferential Statistical Hypothesis Testing of Channel Conversion Rates (Linked to OQ-06)",
         "PRIORITY REASON: Elevates the project beyond cosmetic reporting into rigorous scientific analytics by statistically proving significance (p < 0.05)."),
        ("4. HMW-08: Multi-Touch Customer Journey Attribution Modeling (Linked to OQ-08)",
         "PRIORITY REASON: Resolves the pervasive industry pitfall of last-click bias, ensuring top-of-funnel discovery channels receive fair performance credit."),
        ("5. HMW-09: Prescriptive Budget Reallocation Simulation Under Fixed Spend (Linked to OQ-09)",
         "PRIORITY REASON: Transforms the platform from passive historical reporting to prescriptive decision support, demonstrating measurable business optimization.")
    ]

    for title, reason in priorities:
        p_p = doc.add_paragraph()
        r_title = p_p.add_run(f"★ {title}\n")
        r_title.bold = True
        r_title.font.color.rgb = RGBColor(180, 83, 9) # Amber-Dark
        r_reason = p_p.add_run(f"   {reason}")
        r_reason.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 4: Operational Question Coverage Matrix
    h4 = doc.add_heading(level=1)
    r4 = h4.add_run("4. OQ-to-HMW Traceability & Coverage Verification")
    r4.font.color.rgb = RGBColor(30, 58, 138)

    cov_table = doc.add_table(rows=11, cols=4)
    cov_headers = ["Operational Question ID", "Description Summary", "Covering HMW Statement(s)", "Coverage Status"]
    for j, h in enumerate(cov_headers):
        cov_table.rows[0].cells[j].text = h

    cov_rows = [
        ("OQ-01", "Overall & channel-wise ROAS and CAC", "HMW-01 (Priority)", "Fully Covered (100%)"),
        ("OQ-02", "CTR and CVR variations across channels & time", "HMW-03", "Fully Covered (100%)"),
        ("OQ-03", "Channel Customer Lifetime Value (LTV:CAC)", "HMW-04", "Fully Covered (100%)"),
        ("OQ-04", "Temporal seasonality trends & performance dips", "HMW-05", "Fully Covered (100%)"),
        ("OQ-05", "Demographic segment correlations with channel conversion", "HMW-06", "Fully Covered (100%)"),
        ("OQ-06", "Statistically significant differences in conversion rates", "HMW-07 (Priority)", "Fully Covered (100%)"),
        ("OQ-07", "Underperforming ad set identification against thresholds", "HMW-02 (Priority)", "Fully Covered (100%)"),
        ("OQ-08", "Multi-touch journey attribution breakdown", "HMW-08 (Priority)", "Fully Covered (100%)"),
        ("OQ-09", "Optimal budget reallocation scenario under fixed spend", "HMW-09 (Priority)", "Fully Covered (100%)"),
        ("OQ-10", "Data quality anomalies, missing UTMs & spend outliers", "HMW-10", "Fully Covered (100%)")
    ]

    for r_idx, r_data in enumerate(cov_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cov_table.rows[r_idx].cells[c_idx].text = val
    style_table(cov_table)

    # Save documents
    path1 = os.path.abspath("docs/02_HMW_Statements.docx")
    path2 = os.path.abspath("docs/02_HMW_Framework.docx")
    doc.save(path1)
    doc.save(path2)
    print(f"Saved: {path1}")
    print(f"Saved: {path2}")

if __name__ == "__main__":
    generate_step2_doc()
