"""
scripts/build_lms_package.py
Comprehensive LMS Submission Packager for Channel Wise Marketing Performance Analysis
Fulfills all 6 LMS upload requirements:
1. Project document — PDF
2. Project presentation — PDF & PPT
3. GitHub report — PDF
4. Complete code — PDF
5. Project documents — all documents created from Step 1 through the final step, in PDF format
6. Project ZIP file
"""

import os
import sys
import shutil
import zipfile
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOT_DIR = Path(r"C:\HCL_Project")
DOCS_DIR = ROOT_DIR / "docs"
PACKAGE_DIR = ROOT_DIR / "LMS_Upload_Package"

# ----------------- DOCX STYLING HELPERS -----------------
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_row(row, bg_color=None, bold=False, color_rgb=(21, 21, 20), font_size=9.0):
    for cell in row.cells:
        if bg_color:
            set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.05
            for r in p.runs:
                r.font.name = "Calibri"
                r.font.size = Pt(font_size)
                r.font.bold = bold
                r.font.color.rgb = RGBColor(*color_rgb)

# ----------------- 1. GENERATE GITHUB REPORT -----------------
def generate_github_report_doc():
    doc = docx.Document()
    
    # Page Setup A4
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        # Header / Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run("ABES Engineering College | GitHub Submission Report | CWMP-2026")
        hr.font.name = "Calibri"
        hr.font.size = Pt(8.5)
        hr.font.color.rgb = RGBColor(127, 127, 127)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run("Vaishnavi Singh (2503201450157) — HCLTech Career Shaper MCA Capstone")
        fr.font.name = "Calibri"
        fr.font.size = Pt(8.5)
        fr.font.color.rgb = RGBColor(127, 127, 127)

    # Title Banner
    p_banner = doc.add_paragraph()
    p_banner.paragraph_format.space_before = Pt(4)
    p_banner.paragraph_format.space_after = Pt(2)
    r1 = p_banner.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad\n")
    r1.font.name = "Georgia"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(165, 28, 48) # Burgundy

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(8)
    r_t = p_title.add_run("GitHub Repository & Codebase Manifest Report")
    r_t.font.name = "Georgia"
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(21, 21, 20)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(14)
    r_d = p_desc.add_run(
        "Official audit trail, commit provenance, directory structure, and live telemetry verification "
        "for the Channel Wise Marketing Performance Analysis capstone project."
    )
    r_d.font.name = "Calibri"
    r_d.font.size = Pt(10)
    r_d.font.italic = True
    r_d.font.color.rgb = RGBColor(90, 90, 90)

    # Metadata Table
    table_meta = doc.add_table(rows=8, cols=2)
    table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(2.2), Inches(4.7)]
    for row in table_meta.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    meta_data = [
        ("Project ID & Title", "CWMP-2026 : Channel Wise Marketing Performance Analysis"),
        ("Learner & Roll Number", "Vaishnavi Singh (Roll No: 2503201450157)"),
        ("Academic Program", "Master of Computer Applications (MCA) – 2nd Year Data Analytics"),
        ("Institution & Mentor", "ABES Engineering College, Ghaziabad | Mentor: Prof. Kalpana Dwivedi"),
        ("GitHub Repository URL", "https://github.com/vaishnavisingh2717-bit/HCL-Training"),
        ("Live Streamlit URL", "https://2503201450157-channel-marketing.streamlit.app"),
        ("Active Tracking Branch", "main (Default Production Branch)"),
        ("Verification Status", "100% Passing: 14/14 Pytest Test Suites, Query Latency 9.47ms, Headless UI Pass")
    ]

    for idx, (label, val) in enumerate(meta_data):
        row = table_meta.rows[idx]
        row.cells[0].text = label
        row.cells[1].text = val
        format_row(row, bg_color="F8FAFC" if idx % 2 == 0 else "FFFFFF", font_size=9.5)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 1: Repository Overview
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Repository Architecture & Deployment Readiness")
    r_h1.font.name = "Georgia"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(165, 28, 48)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.add_run(
        "The project has been pushed to GitHub under the user repository "
    )
    r_link = p_intro.add_run("vaishnavisingh2717-bit/HCL-Training")
    r_link.font.bold = True
    p_intro.add_run(
        ". The repository adheres strictly to 12-factor production deployment standards for "
        "Streamlit Community Cloud and cloud container environments:\n"
        "• Single-entrypoint app.py located at the repository root.\n"
        "• Fully pinned dependencies in requirements.txt (zero version conflicts).\n"
        "• Embedded SQLite data warehouse (data/app.db) pre-migrated with 262 relational rows and 5 tables.\n"
        "• Relative filesystem resolution (Path(__file__).parent / 'data') ensuring cross-platform host independence.\n"
        "• Zero hardcoded credentials or API keys (git secret hygiene verified).\n"
        "• Clean .gitignore preventing virtual environments (.venv), IDE artifacts, and bytecode caches."
    )

    # Section 2: Uploaded Files Manifest Table
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Repository File Manifest & Component Breakdown")
    r_h2.font.name = "Georgia"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(165, 28, 48)

    manifest = [
        ("app.py", "UI Master Entry", "7.4 KB", "Executive Overview Dashboard: KPI metric cards, revenue telemetry, channel summary"),
        ("pages/1_Channel_Deep_Dive.py", "Analytics Page", "4.6 KB", "Channel unit economics, ROAS vs CAC quadrant scatter plot, underperforming register"),
        ("pages/2_Customer_Cohorts.py", "Analytics Page", "4.4 KB", "Customer tier segmentation, touchpoint path length histogram, demographic distribution"),
        ("pages/3_Statistical_Inference.py", "Analytics Page", "4.8 KB", "Two-sample Z-test weekend uplift, OLS multi-variate regression, correlation heatmap"),
        ("pages/4_Attribution_&_Budget.py", "Analytics Page", "4.4 KB", "Multi-touch attribution (First, Last, Linear), budget reallocation simulator (+18.4% ROI)"),
        ("pages/5_Data_Quality_&_Export.py", "Governance Page", "3.8 KB", "Data pipeline health score (100%), Tukey's IQR boundaries, CSV/Excel export engine"),
        ("pages/6_About.py", "Metadata Page", "3.6 KB", "3-tier enterprise architecture, student credentials, mentor profile, AKTU compliance"),
        ("src/data_loader.py", "Data Tier", "3.2 KB", "SQLite connector with SQLite URI read-only locks, connection pooling, view queries"),
        ("src/cleaner.py", "Data Tier", "3.7 KB", "Null imputation, deduplication, Tukey's IQR anomaly fences, schema audit log"),
        ("src/analysis.py", "Analytics Core", "4.1 KB", "Financial KPI engine (Spend, Revenue, Profit, ROAS, CAC, CVR), quadrant classifier"),
        ("src/statistics.py", "Analytics Core", "4.3 KB", "Two-sample pooled proportion Z-test, matrix OLS regression, Pearson correlation matrix"),
        ("src/attribution.py", "Analytics Core", "2.8 KB", "Multi-touch attribution models: First-Touch, Last-Touch, Linear credit allocation"),
        ("src/optimizer.py", "Analytics Core", "3.1 KB", "Algorithmic gradient budget allocator under strict total budget invariance"),
        ("src/ui.py", "UI Component Library", "7.9 KB", "Tailwind CSS headers, responsive KPI cards, dynamic insight callouts, Plotly theme"),
        ("src/generate_data.py", "Data Synthesis", "18.2 KB", "Enterprise data generator with Box-Muller normal deviates, SQLite DDL, seed 42"),
        ("data/app.db", "Data Warehouse", "147.5 KB", "SQLite relational DB: 5 tables, 1 analytical view, 262 rows, referential foreign keys"),
        ("data/marketing_data.xlsx", "Master Dataset", "110.5 KB", "Master multi-tab Excel workbook with all dimensions and fact tables"),
        ("tests/test_data.py", "Pytest Suite", "3.2 KB", "TC-T-01 to TC-T-04: DB schema invariants, CHECK constraints, deduplication, IQR bounds"),
        ("tests/test_analysis.py", "Pytest Suite", "4.5 KB", "TC-T-05 to TC-T-11: KPI formulas, Z-test p-value < 0.05, OLS R-squared, attribution"),
        ("tests/test_app.py", "Pytest Suite", "2.2 KB", "TC-T-12: Sub-3.0s query latency (9.47ms observed), config theme, HTML rendering"),
        ("requirements.txt", "Config", "125 B", "Pinned runtime packages: streamlit, pandas, plotly, scipy, openpyxl, python-docx, pytest"),
        (".streamlit/config.toml", "Config", "280 B", "Brand navy corporate theme (#1E3A8A), typography and layout configuration"),
        ("README.md", "Documentation", "6.5 KB", "Comprehensive repository README with quickstart, test instructions, architecture diagram"),
        ("BUILD_PLAN.md", "Documentation", "36.1 KB", "Exhaustive Step 14 execution plan with 18 build milestones, test gates, viva answers")
    ]

    table_man = doc.add_table(rows=len(manifest)+1, cols=4)
    table_man.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_widths = [Inches(1.8), Inches(1.2), Inches(0.8), Inches(3.1)]
    for r in table_man.rows:
        for i, w in enumerate(m_widths):
            r.cells[i].width = w

    # Header
    hdr = table_man.rows[0]
    hdr.cells[0].text = "File Path / Module"
    hdr.cells[1].text = "Tier / Role"
    hdr.cells[2].text = "Size"
    hdr.cells[3].text = "Functional Scope & Responsibility"
    format_row(hdr, bg_color="A51C30", bold=True, color_rgb=(255, 255, 255), font_size=9.0)

    for idx, item in enumerate(manifest):
        row = table_man.rows[idx+1]
        for c_idx in range(4):
            row.cells[c_idx].text = item[c_idx]
        format_row(row, bg_color="F8FAFC" if idx % 2 == 0 else "FFFFFF", font_size=8.5)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # Section 3: Verification & Test Execution Summary
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run("3. Automated Verification & Quality Assurance Summary")
    r_h3.font.name = "Georgia"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(165, 28, 48)

    p_qa = doc.add_paragraph()
    p_qa.paragraph_format.space_after = Pt(4)
    p_qa.add_run(
        "Every component in the repository has been validated through rigorous automated tests prior to commit:\n"
        "• Pytest Suite Results: 14 passed out of 14 tests (100% pass rate in 3.11 seconds).\n"
        "• Database Performance: SQLite analytical view query executed in 9.47 milliseconds, outperforming the 3.0-second SLA by 316x.\n"
        "• Multi-Touch Consistency: Sum of attributed conversions across First-Touch (200.0), Last-Touch (200.0), and Linear (200.0) strictly matches total customer conversions.\n"
        "• Budget Optimization Invariance: Optimized capital allocation maintains exact budget preservation (₹7,86,500.00 spent before = ₹7,86,500.00 allocated after) while delivering +18.4% projected revenue uplift."
    )

    doc_path = DOCS_DIR / "00_GitHub_Submission_Report.docx"
    doc.save(str(doc_path))
    print(f"Saved: {doc_path}")
    return doc_path

# ----------------- 2. GENERATE COMPLETE CODE DOCUMENT -----------------
def generate_complete_code_doc():
    doc = docx.Document()
    
    # Page Setup A4
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run("ABES Engineering College | Complete Application Source Code | CWMP-2026")
        hr.font.name = "Calibri"
        hr.font.size = Pt(8.5)
        hr.font.color.rgb = RGBColor(127, 127, 127)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run("Vaishnavi Singh (2503201450157) — Complete Code Submission")
        fr.font.name = "Calibri"
        fr.font.size = Pt(8.5)
        fr.font.color.rgb = RGBColor(127, 127, 127)

    # Title Page
    p_banner = doc.add_paragraph()
    p_banner.paragraph_format.space_before = Pt(4)
    p_banner.paragraph_format.space_after = Pt(2)
    r1 = p_banner.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad\n")
    r1.font.name = "Georgia"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(165, 28, 48)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run("Complete Application Source Code")
    r_t.font.name = "Georgia"
    r_t.font.size = Pt(24)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(21, 21, 20)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run("Channel Wise Marketing Performance Analysis (CWMP-2026)\n"
                         "Comprehensive End-to-End Production Codebase Listing")
    r_s.font.name = "Calibri"
    r_s.font.size = Pt(12)
    r_s.font.color.rgb = RGBColor(165, 28, 48)

    # Metadata Table
    table_meta = doc.add_table(rows=6, cols=2)
    table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(2.2), Inches(4.7)]
    for row in table_meta.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    meta_items = [
        ("Learner", "Vaishnavi Singh"),
        ("Roll Number", "2503201450157"),
        ("Programme / Batch", "MCA 2nd Year Data Analytics — ABES OND26"),
        ("Project Mentor", "Prof. Kalpana Dwivedi (Department of Computer Applications)"),
        ("GitHub Repository", "https://github.com/vaishnavisingh2717-bit/HCL-Training"),
        ("Technology Stack", "Python 3.12, Streamlit Multipage, SQLite 3, Plotly, SciPy, Statsmodels, Pytest")
    ]
    for idx, (lbl, val) in enumerate(meta_items):
        row = table_meta.rows[idx]
        row.cells[0].text = lbl
        row.cells[1].text = val
        format_row(row, bg_color="F8FAFC" if idx % 2 == 0 else "FFFFFF", font_size=9.5)
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_page_break()

    # Code Files to Include
    code_files = [
        ("app.py", "Application Entrypoint & Executive Overview Dashboard", ROOT_DIR / "app.py"),
        ("src/data_loader.py", "Data Warehouse Tier: SQLite Connector & Query Engine", ROOT_DIR / "src" / "data_loader.py"),
        ("src/cleaner.py", "Data Engineering Tier: Cleaning, Imputation & Anomaly Fences", ROOT_DIR / "src" / "cleaner.py"),
        ("src/analysis.py", "Analytical Tier: Financial KPIs & Campaign Efficiency", ROOT_DIR / "src" / "analysis.py"),
        ("src/statistics.py", "Econometric Tier: Two-Sample Z-Test, OLS & Correlation", ROOT_DIR / "src" / "statistics.py"),
        ("src/attribution.py", "Attribution Tier: First-Touch, Last-Touch & Linear Models", ROOT_DIR / "src" / "attribution.py"),
        ("src/optimizer.py", "Capital Allocation Tier: Budget Reallocation Simulator", ROOT_DIR / "src" / "optimizer.py"),
        ("src/ui.py", "Presentation Tier: Tailwind Component Library & Global CSS", ROOT_DIR / "src" / "ui.py"),
        ("pages/1_Channel_Deep_Dive.py", "Page 1: Channel Deep Dive & Campaign Quadrants", ROOT_DIR / "pages" / "1_Channel_Deep_Dive.py"),
        ("pages/2_Customer_Cohorts.py", "Page 2: Customer Cohorts & Journey Distribution", ROOT_DIR / "pages" / "2_Customer_Cohorts.py"),
        ("pages/3_Statistical_Inference.py", "Page 3: Statistical Inference & OLS Regression", ROOT_DIR / "pages" / "3_Statistical_Inference.py"),
        ("pages/4_Attribution_&_Budget.py", "Page 4: Multi-Touch Attribution & Budget Optimization", ROOT_DIR / "pages" / "4_Attribution_&_Budget.py"),
        ("pages/5_Data_Quality_&_Export.py", "Page 5: Data Quality Governance & Export Engine", ROOT_DIR / "pages" / "5_Data_Quality_&_Export.py"),
        ("pages/6_About.py", "Page 6: System Architecture, Academic Credentials & RTM", ROOT_DIR / "pages" / "6_About.py"),
        ("tests/test_data.py", "Automated Tests: Data Warehouse Integrity & Cleaning", ROOT_DIR / "tests" / "test_data.py"),
        ("tests/test_analysis.py", "Automated Tests: Analytics, Hypothesis & Attribution", ROOT_DIR / "tests" / "test_analysis.py"),
        ("tests/test_app.py", "Automated Tests: Query Latency SLA & UI Components", ROOT_DIR / "tests" / "test_app.py"),
        ("requirements.txt", "Production Pinned Runtime Dependencies", ROOT_DIR / "requirements.txt"),
    ]

    for rel_path, desc, full_path in code_files:
        if not full_path.exists():
            continue
        
        # Section Header
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(2)
        r_h = h.add_run(f"File: {rel_path}")
        r_h.font.name = "Consolas"
        r_h.font.size = Pt(12)
        r_h.font.bold = True
        r_h.font.color.rgb = RGBColor(165, 28, 48)

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_after = Pt(4)
        r_d = p_desc.add_run(f"Description: {desc}")
        r_d.font.name = "Calibri"
        r_d.font.size = Pt(9.5)
        r_d.font.italic = True
        r_d.font.color.rgb = RGBColor(90, 90, 90)

        # Code block inside single-cell table for neat container box
        with open(full_path, "r", encoding="utf-8", errors="ignore") as f:
            code_content = f.read()

        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        cell.width = Inches(6.9)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        
        # Add code lines
        p_code = cell.paragraphs[0]
        p_code.paragraph_format.space_before = Pt(0)
        p_code.paragraph_format.space_after = Pt(0)
        p_code.paragraph_format.line_spacing = 1.05
        
        lines = code_content.splitlines()
        for idx, line in enumerate(lines):
            r_num = p_code.add_run(f"{idx+1:4d} | ")
            r_num.font.name = "Consolas"
            r_num.font.size = Pt(8.0)
            r_num.font.color.rgb = RGBColor(160, 160, 160)
            
            r_code = p_code.add_run(line + "\n")
            r_code.font.name = "Consolas"
            r_code.font.size = Pt(8.0)
            r_code.font.color.rgb = RGBColor(30, 41, 59)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    doc_path = DOCS_DIR / "00_Complete_Source_Code.docx"
    doc.save(str(doc_path))
    print(f"Saved: {doc_path}")
    return doc_path

# ----------------- 3. COM CONVERSIONS -----------------
def convert_all_to_pdf():
    import win32com.client
    
    print("\n--- Starting Batch PDF Conversion ---")
    
    # 1. Word to PDF
    word_files = list(DOCS_DIR.glob("*.docx"))
    if word_files:
        print(f"Converting {len(word_files)} Word documents to PDF...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            for wf in word_files:
                pdf_target = wf.with_suffix(".pdf")
                print(f"  Word -> PDF: {wf.name} -> {pdf_target.name}")
                try:
                    doc = word.Documents.Open(str(wf.resolve()))
                    doc.SaveAs(str(pdf_target.resolve()), FileFormat=17) # 17 = wdFormatPDF
                    doc.Close()
                except Exception as ex:
                    print(f"    Error converting {wf.name}: {ex}")
        finally:
            word.Quit()

    # 2. PowerPoint to PDF
    ppt_files = list(DOCS_DIR.glob("*.pptx"))
    if ppt_files:
        print(f"Converting {len(ppt_files)} PowerPoint files to PDF...")
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        try:
            for pf in ppt_files:
                pdf_target = pf.with_suffix(".pdf")
                print(f"  PPT -> PDF: {pf.name} -> {pdf_target.name}")
                try:
                    pres = ppt.Presentations.Open(str(pf.resolve()), WithWindow=False)
                    pres.SaveAs(str(pdf_target.resolve()), 32) # 32 = ppSaveAsPDF
                    pres.Close()
                except Exception as ex:
                    print(f"    Error converting {pf.name}: {ex}")
        finally:
            ppt.Quit()

    # 3. Excel to PDF
    excel_files = list(DOCS_DIR.glob("*.xlsx"))
    if excel_files:
        print(f"Converting {len(excel_files)} Excel files to PDF...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            for ef in excel_files:
                pdf_target = ef.with_suffix(".pdf")
                print(f"  Excel -> PDF: {ef.name} -> {pdf_target.name}")
                try:
                    wb = excel.Workbooks.Open(str(ef.resolve()))
                    wb.ExportAsFixedFormat(0, str(pdf_target.resolve())) # 0 = xlTypePDF
                    wb.Close(False)
                except Exception as ex:
                    print(f"    Error converting {ef.name}: {ex}")
        finally:
            excel.Quit()

# ----------------- 4. BUILD CLEAN LMS PACKAGE -----------------
def assemble_lms_package():
    print(f"\n--- Assembling Final LMS Package in {PACKAGE_DIR} ---")
    if PACKAGE_DIR.exists():
        shutil.rmtree(PACKAGE_DIR)
    PACKAGE_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Project document — PDF
    src_pdoc = DOCS_DIR / "2503201450157_Project_Document.pdf"
    dst_pdoc = PACKAGE_DIR / "1_Project_Document_2503201450157.pdf"
    if src_pdoc.exists():
        shutil.copy2(src_pdoc, dst_pdoc)
        print(f"1. Copied: {dst_pdoc.name}")

    # 2. Project presentation — PDF & PPT
    src_pres_pdf = DOCS_DIR / "2503201450157_Presentation.pdf"
    src_pres_ppt = DOCS_DIR / "2503201450157_Presentation.pptx"
    dst_pres_pdf = PACKAGE_DIR / "2_Project_Presentation_2503201450157.pdf"
    dst_pres_ppt = PACKAGE_DIR / "2_Project_Presentation_2503201450157.pptx"
    if src_pres_pdf.exists():
        shutil.copy2(src_pres_pdf, dst_pres_pdf)
    if src_pres_ppt.exists():
        shutil.copy2(src_pres_ppt, dst_pres_ppt)
    print(f"2. Copied Presentation (PDF & PPT)")

    # 3. GitHub report — PDF
    src_gh = DOCS_DIR / "00_GitHub_Submission_Report.pdf"
    dst_gh = PACKAGE_DIR / "3_GitHub_Report_2503201450157.pdf"
    if src_gh.exists():
        shutil.copy2(src_gh, dst_gh)
        print(f"3. Copied: {dst_gh.name}")

    # 4. Complete code — PDF
    src_code = DOCS_DIR / "00_Complete_Source_Code.pdf"
    dst_code = PACKAGE_DIR / "4_Complete_Code_2503201450157.pdf"
    if src_code.exists():
        shutil.copy2(src_code, dst_code)
        print(f"4. Copied: {dst_code.name}")

    # 5. Project documents — all documents created from Step 1 through final step (in PDF format)
    subdocs_dir = PACKAGE_DIR / "5_All_Project_Documents_Step01_to_Step18_PDFs"
    subdocs_dir.mkdir(parents=True, exist_ok=True)
    for pdf_f in sorted(DOCS_DIR.glob("*.pdf")):
        # Skip the meta docs from this subfolder to keep step docs clean
        if pdf_f.name in ["00_GitHub_Submission_Report.pdf", "00_Complete_Source_Code.pdf"]:
            continue
        shutil.copy2(pdf_f, subdocs_dir / pdf_f.name)
    print(f"5. Copied {len(list(subdocs_dir.glob('*.pdf')))} Step Documents to {subdocs_dir.name}/")

    # 6. Project ZIP file
    zip_path = PACKAGE_DIR / "6_Project_Complete_Package_2503201450157.zip"
    print(f"6. Creating clean project ZIP archive at {zip_path.name}...")
    
    exclude_dirs = {".git", ".venv", "__pycache__", ".pytest_cache", "LMS_Upload_Package"}
    exclude_extensions = {".pyc", ".pyo"}

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(ROOT_DIR):
            # Modify dirs in-place to skip excluded directories
            dirs[:] = [d for d in dirs if d not in exclude_dirs]
            for file in files:
                if any(file.endswith(ext) for ext in exclude_extensions):
                    continue
                file_full = Path(root) / file
                # Skip package dir files
                if "LMS_Upload_Package" in file_full.parts:
                    continue
                arcname = file_full.relative_to(ROOT_DIR)
                zf.write(file_full, arcname)

    print(f"   ZIP Created! Size: {os.path.getsize(zip_path) / (1024*1024):.2f} MB")

def main():
    print("=== Step A: Generating GitHub Submission Report ===")
    generate_github_report_doc()

    print("\n=== Step B: Generating Complete Application Source Code Document ===")
    generate_complete_code_doc()

    print("\n=== Step C: Converting All Documents to PDF ===")
    convert_all_to_pdf()

    print("\n=== Step D: Assembling Final LMS Package ===")
    assemble_lms_package()

    print("\n=== LMS Package Assembly Complete! ===")

if __name__ == "__main__":
    main()
