"""
scripts/update_all_docs_metadata.py
Replaces learner roll number, batch ID, project ID, and mentor name across ALL Word documents in docs/
"""

import os
import glob
import docx

DOCS_DIR = r"C:\HCL_Project\docs"

REPLACEMENTS = [
    ("2503201450157", "2500320140160"),
    ("Prof. Kalpana Dwivedi, Department of Computer Applications", "Mr. Dinesh"),
    ("Prof. Kalpana Dwivedi", "Mr. Dinesh"),
    ("Kalpana Dwivedi, Department of Computer Applications", "Mr. Dinesh"),
    ("Kalpana Dwivedi", "Mr. Dinesh"),
    ("Mr. Dinesh (Data Analytics)", "Mr. Dinesh"),
    ("Mr. Dinesh (Assistant Professor, MCA Dept.)", "Mr. Dinesh"),
    ("Mr. Dinesh (ABES Engineering College)", "Mr. Dinesh"),
    ("Mr. Dinesh (Department of Computer Applications)", "Mr. Dinesh"),
    ("Mr. Dinesh, Department of Computer Applications", "Mr. Dinesh"),
    ("Mr. Dinesh, Dept of Computer Applications", "Mr. Dinesh"),
    ("Mr. Dinesh (Data Analyst)", "Mr. Dinesh"),
    ("ABES_OND26_MCA Batch 01", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
    ("CWMP-2026", "P084"),
]

def replace_in_text(text):
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text

def replace_in_paragraph(p):
    full_text = "".join(r.text for r in p.runs)
    has_match = any(old in full_text for old, _ in REPLACEMENTS)
    if has_match:
        # If match is in individual runs
        for r in p.runs:
            for old, new in REPLACEMENTS:
                if old in r.text:
                    r.text = r.text.replace(old, new)
        # If match spanned across runs, reset first run and clear others
        new_full = "".join(r.text for r in p.runs)
        if any(old in new_full for old, _ in REPLACEMENTS):
            for old, new in REPLACEMENTS:
                new_full = new_full.replace(old, new)
            if p.runs:
                p.runs[0].text = new_full
                for r in p.runs[1:]:
                    r.text = ""

def update_document(doc_path):
    print(f"Processing: {os.path.basename(doc_path)}")
    doc = docx.Document(doc_path)
    
    # 1. Body paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p)
        
    # 2. Tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)
                    
    # 3. Headers and Footers
    for section in doc.sections:
        for p in section.header.paragraphs:
            replace_in_paragraph(p)
        for p in section.footer.paragraphs:
            replace_in_paragraph(p)
            
    doc.save(doc_path)
    print(f"  Successfully updated {os.path.basename(doc_path)}")

def main():
    for f in sorted(glob.glob(os.path.join(DOCS_DIR, "*.docx"))):
        update_document(f)
    print("\nAll Word documents updated successfully!")

if __name__ == "__main__":
    main()
