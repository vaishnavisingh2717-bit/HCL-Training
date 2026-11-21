"""
scripts/prepare_final_submission.py
Prepares the complete LMS final submission package with exact RollNo_Name_Item naming:
Roll No: 2500320140160
Name: VaishnaviSingh
Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
Project ID: P084
Project Mentor: Mr. Dinesh
"""

import os
import sys
import shutil
import zipfile
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOT_DIR = Path(r"C:\HCL_Project")
DOCS_DIR = ROOT_DIR / "docs"
FINAL_DIR = ROOT_DIR / "final submission"
ROLL_NO = "2500320140160"
NAME = "VaishnaviSingh"
PREFIX = f"{ROLL_NO}_{NAME}"

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
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

# 1. CREATE GITHUB AND STREAMLIT DETAILS DOCX
def create_github_streamlit_details_doc():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Logo if exists
    logo_path = DOCS_DIR / "abes_logo.jpg"
    if logo_path.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        p_logo.add_run().add_picture(str(logo_path), width=Inches(1.0))

    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(1)
    r_inst = p1.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\n")
    r_inst.font.name = "Georgia"
    r_inst.font.size = Pt(14)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(165, 28, 48)
    
    r_sub = p1.add_run("Department of Computer Applications | HCLTech Career Shaper Programme\n")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
    r_sub.font.bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(10)
    r_t = p_title.add_run("GitHub Repository & Streamlit Deployment Verification Details")
    r_t.font.name = "Georgia"
    r_t.font.size = Pt(16)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(21, 21, 20)

    # Meta Table
    tbl = doc.add_table(rows=7, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta = [
        ("Learner Name & Roll No:", f"Vaishnavi Singh  |  Roll No: {ROLL_NO}"),
        ("Batch & Project ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  |  Project ID: P084"),
        ("Project Mentor:", "Mr. Dinesh"),
        ("Project Title:", "Channel Wise Marketing Performance Analysis (CWMP-2026)"),
        ("GitHub Repository URL:", "https://github.com/vaishnavisingh2717-bit/HCL-Training"),
        ("Live Streamlit App URL:", "https://2503201450157-channel-marketing.streamlit.app"),
        ("Deployment Verification:", "Public GitHub Repo (Clean, No Secrets) | All 7 Pages Tested Headlessly & Incognito")
    ]
    for idx, (k, v) in enumerate(meta):
        row = tbl.rows[idx]
        row.cells[0].width = Inches(2.3)
        row.cells[1].width = Inches(4.5)
        row.cells[0].paragraphs[0].add_run(k).font.bold = True
        row.cells[1].paragraphs[0].add_run(v)
        set_cell_margins(row.cells[0], 50, 50, 80, 80)
        set_cell_margins(row.cells[1], 50, 50, 80, 80)
        if idx % 2 == 0:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 1: GitHub Repository File List Screenshot
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. GitHub Repository File List & Verification Screenshot")
    r_h1.font.name = "Georgia"
    r_h1.font.size = Pt(12)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(165, 28, 48)

    p_repo_desc = doc.add_paragraph()
    p_repo_desc.paragraph_format.space_after = Pt(4)
    p_repo_desc.add_run(
        "Public Repository: https://github.com/vaishnavisingh2717-bit/HCL-Training\n"
        "Branch: main | Includes app.py, pages/, src/, tests/, data/, requirements.txt, and README.md."
    )

    repo_shot = DOCS_DIR / "github_repo_screenshot.png"
    if repo_shot.exists():
        p_img1 = doc.add_paragraph()
        p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img1.add_run().add_picture(str(repo_shot), width=Inches(6.2))

    doc.add_page_break()

    # Section 2: Live Streamlit Application Screenshot
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. Live Streamlit Application Deployment Screenshot")
    r_h2.font.name = "Georgia"
    r_h2.font.size = Pt(12)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(165, 28, 48)

    p_app_desc = doc.add_paragraph()
    p_app_desc.paragraph_format.space_after = Pt(4)
    p_app_desc.add_run(
        "Live Cloud Application: https://2503201450157-channel-marketing.streamlit.app\n"
        "Tested in incognito mode. All 7 multipage views render with active telemetry filters, Plotly visualizations, and zero console errors."
    )

    app_shot = DOCS_DIR / "streamlit_app_screenshot.png"
    if app_shot.exists():
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.add_run().add_picture(str(app_shot), width=Inches(6.2))

    out_path = DOCS_DIR / f"{PREFIX}_GitHub_and_Streamlit_Details.docx"
    doc.save(str(out_path))
    print(f"Created: {out_path.name}")
    return out_path

# 2. CREATE VALIDATION REPORT DOCX
def create_validation_report_doc():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    logo_path = DOCS_DIR / "abes_logo.jpg"
    if logo_path.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        p_logo.add_run().add_picture(str(logo_path), width=Inches(0.9))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nStep 10: Dataset Validation & Quality Governance Report")
    r.font.name = "Georgia"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(165, 28, 48)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(6)
    p_meta.add_run(f"Project ID: P084 | Learner: Vaishnavi Singh ({ROLL_NO}) | Mentor: Mr. Dinesh\n")

    val_txt_path = ROOT_DIR / "data" / "validation_report.txt"
    if val_txt_path.exists():
        with open(val_txt_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        cell.width = Inches(6.8)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, 100, 100, 120, 120)
        p_code = cell.paragraphs[0]
        p_code.paragraph_format.line_spacing = 1.05
        for line in lines:
            run = p_code.add_run(line)
            run.font.name = "Consolas"
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor(30, 41, 59)

    out_path = DOCS_DIR / f"{PREFIX}_10_Data_Validation_Report.docx"
    doc.save(str(out_path))
    print(f"Created: {out_path.name}")
    return out_path

# 3. CREATE SOURCE CODE DOCX
def create_source_code_doc():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    logo_path = DOCS_DIR / "abes_logo.jpg"
    if logo_path.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        p_logo.add_run().add_picture(str(logo_path), width=Inches(1.0))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nComplete Application Source Code\n")
    r.font.name = "Georgia"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(165, 28, 48)

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(12)
    p_meta.add_run(f"Project ID: P084 — Channel Wise Marketing Performance Analysis\n"
                   f"Learner: Vaishnavi Singh  |  Roll No: {ROLL_NO}  |  Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01\n"
                   f"Mentor: Mr. Dinesh  |  Session: 2026–2027\n")

    code_order = [
        ("app.py", "Master Entrypoint & Executive Overview Dashboard", ROOT_DIR / "app.py"),
        ("pages/1_Channel_Deep_Dive.py", "Page 1: Channel Deep Dive & Campaign Matrix", ROOT_DIR / "pages" / "1_Channel_Deep_Dive.py"),
        ("pages/2_Customer_Cohorts.py", "Page 2: Customer Cohort & Demographic Distribution", ROOT_DIR / "pages" / "2_Customer_Cohorts.py"),
        ("pages/3_Statistical_Inference.py", "Page 3: Econometric Hypothesis Testing & Regression", ROOT_DIR / "pages" / "3_Statistical_Inference.py"),
        ("pages/4_Attribution_&_Budget.py", "Page 4: Multi-Touch Attribution & Capital Simulator", ROOT_DIR / "pages" / "4_Attribution_&_Budget.py"),
        ("pages/5_Data_Quality_&_Export.py", "Page 5: Data Quality Governance & Data Export", ROOT_DIR / "pages" / "5_Data_Quality_&_Export.py"),
        ("pages/6_About.py", "Page 6: System Architecture & Student Profile", ROOT_DIR / "pages" / "6_About.py"),
        ("src/data_loader.py", "Data Loader Tier: SQLite Connector", ROOT_DIR / "src" / "data_loader.py"),
        ("src/cleaner.py", "Cleaning Tier: Deduplication & IQR Fences", ROOT_DIR / "src" / "cleaner.py"),
        ("src/analysis.py", "Analytics Tier: Financial KPIs Engine", ROOT_DIR / "src" / "analysis.py"),
        ("src/statistics.py", "Statistics Tier: Two-sample Z-test & OLS", ROOT_DIR / "src" / "statistics.py"),
        ("src/attribution.py", "Attribution Tier: First, Last & Linear Models", ROOT_DIR / "src" / "attribution.py"),
        ("src/optimizer.py", "Optimizer Tier: Budget Allocation Simulator", ROOT_DIR / "src" / "optimizer.py"),
        ("src/ui.py", "UI Tier: Tailwind Components & Responsive CSS", ROOT_DIR / "src" / "ui.py"),
        ("tests/test_data.py", "Pytest Suite: Data Integrity & Invariants", ROOT_DIR / "tests" / "test_data.py"),
        ("tests/test_analysis.py", "Pytest Suite: Analytical Calculations & Models", ROOT_DIR / "tests" / "test_analysis.py"),
        ("tests/test_app.py", "Pytest Suite: SLA Query Latency & Config Theme", ROOT_DIR / "tests" / "test_app.py"),
        ("requirements.txt", "Pinned Dependencies", ROOT_DIR / "requirements.txt"),
    ]

    for rel_p, desc, full_p in code_order:
        if not full_p.exists():
            continue
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(2)
        r_h = h.add_run(f"File: {rel_p}")
        r_h.font.name = "Consolas"
        r_h.font.size = Pt(11)
        r_h.font.bold = True
        r_h.font.color.rgb = RGBColor(165, 28, 48)

        p_d = doc.add_paragraph()
        p_d.paragraph_format.space_after = Pt(4)
        p_d.add_run(f"Module Scope: {desc}").font.italic = True

        with open(full_p, "r", encoding="utf-8", errors="ignore") as f:
            code_txt = f.read()

        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        cell.width = Inches(6.8)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, 80, 80, 100, 100)
        p_c = cell.paragraphs[0]
        p_c.paragraph_format.line_spacing = 1.05
        for idx, line in enumerate(code_txt.splitlines()):
            r_num = p_c.add_run(f"{idx+1:4d} | ")
            r_num.font.name = "Consolas"
            r_num.font.size = Pt(7.5)
            r_num.font.color.rgb = RGBColor(160, 160, 160)
            r_code = p_c.add_run(line + "\n")
            r_code.font.name = "Consolas"
            r_code.font.size = Pt(7.5)
            r_code.font.color.rgb = RGBColor(30, 41, 59)

    out_path = DOCS_DIR / f"{PREFIX}_Source_Code.docx"
    doc.save(str(out_path))
    print(f"Created: {out_path.name}")
    return out_path

# 4. CONVERT DOCX & PPTX & XLSX TO PDF VIA WIN32COM
def convert_doc_to_pdf(doc_path, pdf_path):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(str(doc_path.resolve()))
        d.SaveAs(str(pdf_path.resolve()), FileFormat=17)
        d.Close()
        print(f"  Converted Word -> PDF: {pdf_path.name}")
    finally:
        word.Quit()

def convert_ppt_to_pdf(ppt_path, pdf_path):
    import win32com.client
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        p = ppt.Presentations.Open(str(ppt_path.resolve()), WithWindow=False)
        p.SaveAs(str(pdf_path.resolve()), 32)
        p.Close()
        print(f"  Converted PPT -> PDF: {pdf_path.name}")
    finally:
        ppt.Quit()

def convert_excel_to_pdf(xlsx_path, pdf_path):
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
        wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))
        wb.Close(False)
        print(f"  Converted Excel -> PDF: {pdf_path.name}")
    finally:
        excel.Quit()

# 5. ASSEMBLE FINAL SUBMISSION DIRECTORY
def build_final_submission_package():
    print(f"\n--- Assembling Clean 'final submission' Folder ---")
    if FINAL_DIR.exists():
        shutil.rmtree(FINAL_DIR)
    FINAL_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Project Report PDF
    src_rep_docx = DOCS_DIR / f"{PREFIX}_Project_Report.docx"
    dst_rep_pdf = FINAL_DIR / f"{PREFIX}_Project_Report.pdf"
    convert_doc_to_pdf(src_rep_docx, dst_rep_pdf)

    # 2. Presentation PDF
    src_ppt = DOCS_DIR / f"{PREFIX}_Presentation.pptx"
    dst_ppt_pdf = FINAL_DIR / f"{PREFIX}_Presentation.pdf"
    convert_ppt_to_pdf(src_ppt, dst_ppt_pdf)

    # 3. Source Code PDF
    src_code_docx = DOCS_DIR / f"{PREFIX}_Source_Code.docx"
    dst_code_pdf = FINAL_DIR / f"{PREFIX}_Source_Code.pdf"
    convert_doc_to_pdf(src_code_docx, dst_code_pdf)

    # 4. Database Schema and Project Database in XLSX format
    shutil.copy2(DOCS_DIR / "08_Database_Design.xlsx", FINAL_DIR / f"{PREFIX}_Database_Schema.xlsx")
    shutil.copy2(ROOT_DIR / "data" / "marketing_data.xlsx", FINAL_DIR / f"{PREFIX}_Project_Database.xlsx")

    # 5. GitHub and Streamlit Details PDF
    src_gh_docx = DOCS_DIR / f"{PREFIX}_GitHub_and_Streamlit_Details.docx"
    dst_gh_pdf = FINAL_DIR / f"{PREFIX}_GitHub_and_Streamlit_Details.pdf"
    convert_doc_to_pdf(src_gh_docx, dst_gh_pdf)

    # 6. Other Project Documents (Step order as PDFs, keep Excel as XLSX)
    other_docs = [
        ("01_Project_Understanding.docx", f"{PREFIX}_01_Project_Understanding.pdf", "word"),
        ("02_HMW_Statements.docx", f"{PREFIX}_02_HMW.pdf", "word"),
        ("03_Needs_Analysis.docx", f"{PREFIX}_03_Needs_Analysis.pdf", "word"),
        ("04_Traceability_Matrix.xlsx", f"{PREFIX}_04_Traceability_Matrix.xlsx", "copy_xlsx"),
        ("05_BRD.docx", f"{PREFIX}_05_BRD.pdf", "word"),
        ("07_TRD.docx", f"{PREFIX}_07_TRD.pdf", "word"),
        ("08_Database_Design.xlsx", f"{PREFIX}_08_Database_Design.xlsx", "copy_xlsx"),
        ("09_ER_Diagram.docx", f"{PREFIX}_09_ER_Diagram.pdf", "word"),
        (f"{PREFIX}_10_Data_Validation_Report.docx", f"{PREFIX}_10_Data_Validation_Report.pdf", "word"),
        ("11_User_Stories.docx", f"{PREFIX}_11_User_Stories.pdf", "word"),
        ("12_UI_Framework.docx", f"{PREFIX}_12_UI_Framework.pdf", "word"),
        ("13_Test_Cases.docx", f"{PREFIX}_13_Test_Cases.pdf", "word")
    ]

    for src_name, dst_name, kind in other_docs:
        src_path = DOCS_DIR / src_name
        dst_path = FINAL_DIR / dst_name
        if kind == "word" and src_path.exists():
            convert_doc_to_pdf(src_path, dst_path)
        elif kind == "copy_xlsx" and src_path.exists():
            shutil.copy2(src_path, dst_path)
            print(f"  Copied XLSX: {dst_name}")

    # 7. Project ZIP
    zip_path = FINAL_DIR / f"{PREFIX}_CWMP_Project.zip"
    print(f"\nCreating clean project ZIP archive at {zip_path.name}...")
    exclude_dirs = {".git", ".venv", "venv", "__pycache__", ".pytest_cache", "LMS_Upload_Package", "final submission"}
    exclude_ext = {".pyc", ".pyo"}

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(ROOT_DIR):
            dirs[:] = [d for d in dirs if d not in exclude_dirs]
            for file in files:
                if any(file.endswith(ext) for ext in exclude_ext):
                    continue
                file_full = Path(root) / file
                if "final submission" in file_full.parts or "LMS_Upload_Package" in file_full.parts:
                    continue
                arcname = file_full.relative_to(ROOT_DIR)
                zf.write(file_full, arcname)

    print(f"  ZIP Created! Size: {os.path.getsize(zip_path) / (1024*1024):.2f} MB")

def report_final_package():
    print("\n================ FINAL SUBMISSION MANIFEST ================")
    files = sorted(list(FINAL_DIR.glob("*")))
    total_size = sum(f.stat().st_size for f in files if f.is_file())
    for idx, f in enumerate(files, start=1):
        size_kb = f.stat().st_size / 1024
        size_str = f"{size_kb/1024:.2f} MB" if size_kb > 1024 else f"{size_kb:.1f} KB"
        print(f"{idx:2d}. {f.name:<60} | {size_str}")
    print("-----------------------------------------------------------")
    print(f"Total Files: {len(files)} (Target <= 20 files)")
    print(f"Total Size:  {total_size / (1024*1024):.2f} MB (Target < 100 MB)")
    print("===========================================================")

def main():
    print("Step 1: Generating GitHub & Streamlit Details Doc...")
    create_github_streamlit_details_doc()

    print("Step 2: Generating Data Validation Report Doc...")
    create_validation_report_doc()

    print("Step 3: Generating Source Code Doc...")
    create_source_code_doc()

    print("Step 4: Assembling Final Submission Package...")
    build_final_submission_package()

    print("Step 5: Verifying Final Submission Package...")
    report_final_package()

if __name__ == "__main__":
    main()
