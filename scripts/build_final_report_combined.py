"""
scripts/build_final_report_combined.py
Compiles all Step 8 documents into a single master PDF:
2500320140160_VaishnaviSingh_Final_Report.pdf
"""

import os
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import pypdf

ROOT_DIR = Path(r"C:\HCL_Project")
DOCS_DIR = ROOT_DIR / "docs"
FINAL_DIR = ROOT_DIR / "final submission"
ROLL_NO = "2500320140160"
NAME = "VaishnaviSingh"
PREFIX = f"{ROLL_NO}_{NAME}"

def create_cover_and_index():
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Inches(8.267)
        section.page_height = Inches(11.692)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        section.different_first_page_header_footer = True
        
        # Subsequent header / footer
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run("ABES Engineering College | Final Project Report Compilation | P084")
        hr.font.name = "Calibri"
        hr.font.size = Pt(8)
        hr.font.color.rgb = RGBColor(127, 127, 127)

        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run(f"Vaishnavi Singh | {ROLL_NO}     Final Project Report")
        fr.font.name = "Calibri"
        fr.font.size = Pt(8)
        fr.font.color.rgb = RGBColor(127, 127, 127)

    # 1. College Logo
    logo_path = DOCS_DIR / "abes_logo.jpg"
    if logo_path.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(6)
        p_logo.paragraph_format.space_after = Pt(4)
        p_logo.add_run().add_picture(str(logo_path), width=Inches(1.2))

    # 2. Institution Details
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(2)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("ABES ENGINEERING COLLEGE")
    r_inst.font.name = "Georgia"
    r_inst.font.size = Pt(15)
    r_inst.font.bold = True

    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.paragraph_format.space_before = Pt(0)
    p_city.paragraph_format.space_after = Pt(2)
    r_city = p_city.add_run("Ghaziabad, Uttar Pradesh")
    r_city.font.name = "Calibri"
    r_city.font.size = Pt(11)
    r_city.font.bold = True

    p_aff = doc.add_paragraph()
    p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_aff.paragraph_format.space_before = Pt(0)
    p_aff.paragraph_format.space_after = Pt(18)
    r_aff = p_aff.add_run("Affiliated to Dr. A.P.J. Abdul Kalam Technical University (AKTU), Uttar Pradesh")
    r_aff.font.name = "Calibri"
    r_aff.font.size = Pt(9.5)
    r_aff.font.color.rgb = RGBColor(100, 100, 100)

    # 3. Report Title
    p_rep = doc.add_paragraph()
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_rep.paragraph_format.space_before = Pt(4)
    p_rep.paragraph_format.space_after = Pt(3)
    r_rep = p_rep.add_run("FINAL PROJECT REPORT")
    r_rep.font.name = "Georgia"
    r_rep.font.size = Pt(20)
    r_rep.font.bold = True

    p_on = doc.add_paragraph()
    p_on.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_on.paragraph_format.space_before = Pt(0)
    p_on.paragraph_format.space_after = Pt(3)
    r_on = p_on.add_run("on")
    r_on.font.name = "Calibri"
    r_on.font.size = Pt(11)
    r_on.font.italic = True

    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj.paragraph_format.space_before = Pt(0)
    p_proj.paragraph_format.space_after = Pt(3)
    r_proj = p_proj.add_run("Channel Wise Marketing Performance Analysis")
    r_proj.font.name = "Georgia"
    r_proj.font.size = Pt(16)
    r_proj.font.bold = True

    p_subdesc = doc.add_paragraph()
    p_subdesc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subdesc.paragraph_format.space_before = Pt(0)
    p_subdesc.paragraph_format.space_after = Pt(20)
    r_subdesc = p_subdesc.add_run("Comprehensive Compilation of Project Documents (Step 01 through Step 13)")
    r_subdesc.font.name = "Calibri"
    r_subdesc.font.size = Pt(11)
    r_subdesc.font.color.rgb = RGBColor(70, 70, 70)

    # 4. Fulfilment Text
    p_subm = doc.add_paragraph()
    p_subm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subm.paragraph_format.space_before = Pt(0)
    p_subm.paragraph_format.space_after = Pt(3)
    r_subm = p_subm.add_run("Submitted in partial fulfilment of the requirements for the degree of")
    r_subm.font.name = "Calibri"
    r_subm.font.size = Pt(10.5)

    p_deg = doc.add_paragraph()
    p_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_deg.paragraph_format.space_before = Pt(0)
    p_deg.paragraph_format.space_after = Pt(2)
    r_deg = p_deg.add_run("Master of Computer Applications")
    r_deg.font.name = "Calibri"
    r_deg.font.size = Pt(13)
    r_deg.font.bold = True

    p_in = doc.add_paragraph()
    p_in.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_in.paragraph_format.space_before = Pt(0)
    p_in.paragraph_format.space_after = Pt(2)
    r_in = p_in.add_run("in")
    r_in.font.name = "Calibri"
    r_in.font.size = Pt(10.5)
    r_in.font.italic = True

    p_dept = doc.add_paragraph()
    p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dept.paragraph_format.space_before = Pt(0)
    p_dept.paragraph_format.space_after = Pt(6)
    r_dept = p_dept.add_run("Data Analytics")
    r_dept.font.name = "Calibri"
    r_dept.font.size = Pt(12)
    r_dept.font.bold = True

    p_prog = doc.add_paragraph()
    p_prog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_prog.paragraph_format.space_before = Pt(0)
    p_prog.paragraph_format.space_after = Pt(22)
    r_prog = p_prog.add_run("Training Programme: HCLTech Career Shaper Programme")
    r_prog.font.name = "Calibri"
    r_prog.font.size = Pt(10.5)

    # 5. Submission Details Block (Clean, Borderless)
    cover_table = doc.add_table(rows=6, cols=2)
    cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_data = [
        ("Submitted by:", "Vaishnavi Singh"),
        ("University Roll Number:", "2500320140160"),
        ("Project Guide:", "Mr. Dinesh"),
        ("Project ID / Batch:", "P084  |  ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Academic Session:", "2026–2027"),
        ("Submission Date:", "21 November 2026")
    ]
    for i, (k, v) in enumerate(cover_data):
        row = cover_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.6)
        c1.width = Inches(4.3)
        tcPr0 = c0._tc.get_or_add_tcPr()
        tcPr0.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))
        tcPr1 = c1._tc.get_or_add_tcPr()
        tcPr1.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(k)
        r0.font.name = "Calibri"
        r0.font.size = Pt(10.5)
        r0.font.bold = True
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(v)
        r1.font.name = "Calibri"
        r1.font.size = Pt(10.5)

    doc.add_page_break()

    # Page 2: Table of Contents / Index
    p_idx_h = doc.add_paragraph()
    p_idx_h.paragraph_format.space_before = Pt(10)
    p_idx_h.paragraph_format.space_after = Pt(8)
    r_ih = p_idx_h.add_run("Compendium Index — Project Documents (Step 01 to Step 13)")
    r_ih.font.name = "Georgia"
    r_ih.font.size = Pt(16)
    r_ih.font.bold = True

    index_items = [
        ("Step 01", "Project Understanding & Scope Definition", "01_Project_Understanding"),
        ("Step 02", "How Might We (HMW) Innovation Statements", "02_HMW"),
        ("Step 03", "Stakeholder Needs & Empathy Analysis", "03_Needs_Analysis"),
        ("Step 04", "Requirements Traceability Matrix (RTM)", "04_Traceability_Matrix"),
        ("Step 05", "Business Requirements Document (BRD)", "05_BRD"),
        ("Step 07", "Technical Requirements Document (TRD)", "07_TRD"),
        ("Step 08", "Database Architecture & Schema Design", "08_Database_Design"),
        ("Step 09", "Entity-Relationship (ER) Star Schema Diagram", "09_ER_Diagram"),
        ("Step 10", "Dataset Validation & Quality Governance Report", "10_Data_Validation_Report"),
        ("Step 11", "Agile User Stories & Acceptance Criteria", "11_User_Stories"),
        ("Step 12", "UI/UX Framework & Wireframe Design", "12_UI_Framework"),
        ("Step 13", "Functional & Technical Test Case Suite", "13_Test_Cases"),
    ]

    tbl_idx = doc.add_table(rows=len(index_items)+1, cols=3)
    tbl_idx.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Inches(1.2), Inches(3.6), Inches(2.2)]
    
    hdr = tbl_idx.rows[0]
    hdr.cells[0].paragraphs[0].add_run("Step No.").font.bold = True
    hdr.cells[1].paragraphs[0].add_run("Document Scope / Deliverable").font.bold = True
    hdr.cells[2].paragraphs[0].add_run("Section Identifier").font.bold = True
    for c in hdr.cells:
        c.paragraphs[0].runs[0].font.name = "Georgia"
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="A51C30"/>')
        c._tc.get_or_add_tcPr().append(shd)

    for i, (st, desc, sid) in enumerate(index_items):
        row = tbl_idx.rows[i+1]
        row.cells[0].paragraphs[0].add_run(st).font.bold = True
        row.cells[1].paragraphs[0].add_run(desc)
        row.cells[2].paragraphs[0].add_run(sid)
        for j, c in enumerate(row.cells):
            c.width = widths[j]
            c.paragraphs[0].runs[0].font.name = "Calibri"
            c.paragraphs[0].runs[0].font.size = Pt(9.0)
            if i % 2 == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
                c._tc.get_or_add_tcPr().append(shd)

    cover_docx = DOCS_DIR / "final_report_cover.docx"
    cover_pdf = DOCS_DIR / "final_report_cover.pdf"
    doc.save(str(cover_docx))

    # Convert to PDF
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(str(cover_docx.resolve()))
        d.SaveAs(str(cover_pdf.resolve()), FileFormat=17)
        d.Close()
    finally:
        word.Quit()
    print("Cover & Index PDF generated.")
    return cover_pdf

def merge_all_step8_pdfs(cover_pdf):
    merger = pypdf.PdfWriter()
    
    # 1. Add Cover & Index
    merger.append(str(cover_pdf))
    print("Appended Cover & Index")

    # 2. Add Step Documents in order
    steps_files = [
        FINAL_DIR / f"{PREFIX}_01_Project_Understanding.pdf",
        FINAL_DIR / f"{PREFIX}_02_HMW.pdf",
        FINAL_DIR / f"{PREFIX}_03_Needs_Analysis.pdf",
        DOCS_DIR / "04_Traceability_Matrix.pdf",
        FINAL_DIR / f"{PREFIX}_05_BRD.pdf",
        FINAL_DIR / f"{PREFIX}_07_TRD.pdf",
        DOCS_DIR / "08_Database_Design.pdf",
        FINAL_DIR / f"{PREFIX}_09_ER_Diagram.pdf",
        FINAL_DIR / f"{PREFIX}_10_Data_Validation_Report.pdf",
        FINAL_DIR / f"{PREFIX}_11_User_Stories.pdf",
        FINAL_DIR / f"{PREFIX}_12_UI_Framework.pdf",
        FINAL_DIR / f"{PREFIX}_13_Test_Cases.pdf",
    ]

    for pf in steps_files:
        if pf.exists():
            merger.append(str(pf))
            print(f"Appended: {pf.name}")
        else:
            print(f"WARNING: File not found: {pf}")

    # Output merged PDF in final submission
    out_final_pdf = FINAL_DIR / f"{PREFIX}_Final_Report.pdf"
    out_alias_pdf = FINAL_DIR / "Final_Report.pdf"
    
    with open(out_final_pdf, "wb") as f_out:
        merger.write(f_out)
    merger.close()
    
    import shutil
    shutil.copy2(out_final_pdf, out_alias_pdf)
    shutil.copy2(out_final_pdf, DOCS_DIR / f"{PREFIX}_Final_Report.pdf")

    size_mb = os.path.getsize(out_final_pdf) / (1024 * 1024)
    print(f"\nSUCCESS! Combined Final Report generated: {out_final_pdf.name} ({size_mb:.2f} MB)")

def main():
    cover_pdf = create_cover_and_index()
    merge_all_step8_pdfs(cover_pdf)

if __name__ == "__main__":
    main()
