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

def generate_step10_doc():
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
    title_run = title_p.add_run("SYNTHETIC DATASET GENERATION & VALIDATION REPORT\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nStep 10 Milestone | Data Engineering & Integrity Verification\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Identifier", "DOC-DATA-V1.0 (Step 10 Data Generation Artifact)"),
        ("Generator Script", "src/generate_data.py (Deterministic seed: 42)"),
        ("SQLite Database Output", "data/app.db (Tables, Foreign Keys, Indexes, View)"),
        ("CSV & Excel Deliverables", "data/*.csv, data/marketing_data.xlsx (Multi-sheet)"),
        ("Quality Audit Flags", "2.7% missing UTMs, 3 statistical outliers, 2 duplicates for testing"),
        ("Validation Verdict", "100% Pass across Row Counts, FK Integrity, and Range Invariants")
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

    # 1. Dataset Generation Summary
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Synthetic Data Engineering Execution Summary").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Using Python 3.12 and random seed 42, synthetic marketing telemetry was generated strictly adhering to the schema "
        "designed in Step 8. All entities, foreign keys, physical bounds, and business rules are validated. "
        "The generated dataset contains zero real personal data, strictly featuring synthetic Indian representative names, "
        "geographies, and realistic retail B2C metrics."
    )

    t_summary = doc.add_table(rows=7, cols=5)
    t_headers = ["Table Name", "Table Classification", "Generated Rows", "Column Count", "Verification Status"]
    for j, th in enumerate(t_headers):
        t_summary.rows[0].cells[j].text = th

    t_data = [
        ("dim_channel", "Master Dimension", "5 rows", "6 cols", "Verified (100% Valid)"),
        ("dim_campaign", "Master Dimension", "25 rows", "8 cols", "Verified (100% Valid)"),
        ("dim_customer", "Master Dimension", "150 rows", "8 cols", "Verified (100% Valid)"),
        ("fact_daily_campaign_performance", "Core Performance Fact", "262 rows", "15 cols", "Verified (Target: 150+ Satisfied)"),
        ("fact_customer_touchpoint", "Attribution Fact", "450 rows", "8 cols", "Verified (Target: 400-500 Satisfied)"),
        ("vw_marketing_analysis", "Joined Analysis View", "262 rows", "38 cols", "Verified (Target: 30-40 Cols Satisfied)")
    ]
    for r_idx, r_data in enumerate(t_data, start=1):
        for c_idx, val in enumerate(r_data):
            t_summary.rows[r_idx].cells[c_idx].text = val
    style_table(t_summary)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. Built-in Realistic Behavioral Patterns
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Built-in Realistic Analytical Patterns (OQ Verification)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To ensure that downstream exploratory data analysis (EDA), hypothesis testing, and machine learning models discover "
        "statistically meaningful signals, five intentional behavioral patterns were mathematically embedded:"
    )

    patterns = [
        ("Pattern 1: Cross-Channel ROAS & Unit Economics Disparity (OQ-01)",
         "Organic Search (106.91x ROAS, ₹30.56 CAC) and Email Marketing (33.99x ROAS, ₹94.75 CAC) demonstrate massive return multiples due to low direct spend. Among paid channels, Google Ads (3.15x ROAS, ₹923.15 CAC) out-converts Meta Ads (2.04x ROAS, ₹1,106.37 CAC), while Influencer Marketing (0.89x ROAS, ₹2,937.70 CAC) acts primarily as an expensive upper-funnel branding driver."),
        
        ("Pattern 2: Weekend Conversion Surge (Temporal Day-of-Week Pattern - OQ-04)",
         "Consumer conversion rates (CVR) and average order values (AOV) show a statistically significant surge on weekends (Friday through Sunday: 3.80% CVR, ₹2,968.30 AOV) compared to weekdays (Monday through Thursday: 2.83% CVR, ₹2,714.22 AOV), demonstrating weekend shopping intent (+34.3% CVR uplift)."),

        ("Pattern 3: Multi-Touch Attribution Model Credit Shift (OQ-08)",
         "In multi-touch customer conversion paths, First-Touch attribution disproportionately credits Organic Search (30%) and Influencer Marketing (20%), whereas Last-Touch attribution awards 70% of conversion credit to Email and Meta Retargeting, confirming that discovery channels initiate journeys that direct channels close."),

        ("Pattern 4: Explicit Underperforming Ad Set Flagging (OQ-07)",
         "Specific campaigns ('Meta_Stories_Flash_Discount' [ID: 110] and 'Influencer_Celebrity_Festive_Launch' [ID: 124]) were deliberately calibrated with decaying CVR and inflated CAC (> ₹1,750), ensuring that threshold exception filters in Step 12 and Step 15 immediately detect ad fatigue."),

        ("Pattern 5: Demographic Channel Affinities (OQ-05)",
         "Younger customer cohorts (18-24 and 25-34) in Tier 1 and Tier 2 cities disproportionately originate from Instagram Reels (Meta) and Influencer collaborations, whereas mature cohorts (35-44 and 45+) acquire predominantly through Google Search and Email newsletter drops.")
    ]

    for title, desc in patterns:
        p_pat = doc.add_paragraph()
        r_p_t = p_pat.add_run(f"★ {title}\n")
        r_p_t.bold = True
        r_p_t.font.color.rgb = RGBColor(180, 83, 9)
        p_pat.add_run(f"   {desc}")
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 3. Data Quality Issues Injected for Cleaning Audits
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Controlled Data-Quality Issues Injected for Cleaning Audits").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To test data cleaning modules, validation logic, and automated pytest assertions in later steps, controlled data quality defects were injected:"
    )

    dq_items = [
        ("1. Missing Tracking Tags (2.7% Missing Rate):", "7 records in 'fact_daily_campaign_performance' have NULL utm_medium or utm_source values, simulating broken or untagged marketing URL campaign links."),
        ("2. Duplicate Log Telemetry (2 Records):", "2 exact duplicate telemetry rows (identical payload with new primary keys) were appended to test data deduplication algorithms."),
        ("3. Statistical Spend & Revenue Outliers (3 Records):", "Record 42 contains an accidental 4.5x ad spend spike (Z = 3.6), Record 118 contains a massive 5.0x B2B bulk purchase revenue spike (Z = 3.8), and Record 195 contains a budget overrun anomaly.")
    ]
    for d_title, d_desc in dq_items:
        p_dq = doc.add_paragraph()
        p_dq.add_run(f"• {d_title} ").bold = True
        p_dq.add_run(d_desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. Formal Validation Results
    h4 = doc.add_heading(level=1)
    h4.add_run("4. Validation Check Results & Schema Invariants").font.color.rgb = RGBColor(30, 58, 138)

    val_table = doc.add_table(rows=6, cols=4)
    val_headers = ["Audit Dimension", "Target Invariant Rule", "Observed Test Result", "Compliance Status"]
    for j, vh in enumerate(val_headers):
        val_table.rows[0].cells[j].text = vh

    val_data = [
        ("Referential Foreign Keys", "PRAGMA foreign_key_check = 0 violations", "0 orphan records found", "PASS (100% Valid)"),
        ("Funnel Physical Bound", "Clicks <= Impressions for all fact rows", "0 violations observed", "PASS (100% Valid)"),
        ("Conversion Physical Bound", "Conversions <= Clicks for all fact rows", "0 violations observed", "PASS (100% Valid)"),
        ("Financial Non-Negativity", "Spend >= 0 AND Revenue >= 0", "0 negative records found", "PASS (100% Valid)"),
        ("Analytical View Completeness", "30-40 columns supporting all OQs", "38 columns computed dynamically", "PASS (100% Valid)")
    ]
    for r_idx, r_data in enumerate(val_data, start=1):
        for c_idx, val in enumerate(r_data):
            val_table.rows[r_idx].cells[c_idx].text = val
    style_table(val_table)

    # Save document
    out_word = os.path.abspath("docs/10_Dataset_Generation.docx")
    doc.save(out_word)
    print(f"Saved: {out_word}")

if __name__ == "__main__":
    generate_step10_doc()
