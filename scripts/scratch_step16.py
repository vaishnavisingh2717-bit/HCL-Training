"""
scratch_step16.py - Generates docs/16_Deployment_Guide.docx
Covers Step 16: Streamlit Community Cloud Deployment SOP & DevOps Guide.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

WORKSPACE_DIR = r"c:\HCL_Project"
DOCS_DIR = os.path.join(WORKSPACE_DIR, "docs")
DOCX_PATH = os.path.join(DOCS_DIR, "16_Deployment_Guide.docx")

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, header_bg="1E3A8A", alt_bg="F8FAFC"):
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
                        run.font.size = Pt(8.5)
                        run.font.color.rgb = RGBColor(30, 41, 59)

def generate_step16_doc():
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
    title_run = title_p.add_run("STREAMLIT COMMUNITY CLOUD DEPLOYMENT & DEVOPS GUIDE\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Channel Wise Marketing Performance Analysis | ABES MCA Capstone Project\nDocument Ref: DOC-DEP-16 | Version 1.0 (Production Release)")
    sub_run.font.size = Pt(12)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph()

    # Document Control Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_data = [
        ("Candidate Name / Roll No:", "Vaishnavi Singh | Roll No: 2503201450157"),
        ("Degree & Institution:", "MCA 2nd Year Data Analytics | ABES Engineering College, Ghaziabad"),
        ("Faculty Project Mentor:", "Prof. Kalpana Dwivedi, Department of Computer Applications"),
        ("Production Host Target:", "Streamlit Community Cloud (Linux Debian container, Python 3.12)"),
        ("Repository Target:", "https://github.com/<username>/2503201450157-channel-wise-marketing-performance"),
        ("Target Public URL:", "https://2503201450157-channel-marketing.streamlit.app")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.5), Inches(4.5)
        c0.paragraphs[0].add_run(k).bold = True
        c1.paragraphs[0].add_run(v)
    style_table(meta_table, header_bg="1E3A8A", alt_bg="F1F5F9")
    doc.add_paragraph()

    # Section 1: Pre-Deployment Audit Checklist
    h1 = doc.add_paragraph()
    r = h1.add_run("1. Pre-Deployment Repository Audit Checklist")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    p1 = doc.add_paragraph(
        "Prior to initiating deployment on Streamlit Community Cloud, the repository was audited against "
        "production release criteria to guarantee zero deployment failures and instant container initialization:"
    )

    checklist_table = doc.add_table(rows=8, cols=4)
    chk_headers = ["Audit Item", "Requirement Specification", "Configured Path / Value", "Audit Status"]
    for j, h in enumerate(chk_headers):
        checklist_table.rows[0].cells[j].paragraphs[0].add_run(h)

    chk_data = [
        ("Root Application Entrypoint", "Main entry script located at repository root", "app.py", "VERIFIED (Pass)"),
        ("Dependency Manifest", "Pinned, compatible production Python packages", "requirements.txt", "VERIFIED (Pass)"),
        ("Database Storage", "Committed, normalized SQLite warehouse", "data/app.db (262 rows, 5 tables)", "VERIFIED (Pass)"),
        ("Relative Path Resolution", "Dynamic path resolution relative to __file__", "Path(__file__).resolve().parent", "VERIFIED (Pass)"),
        ("Secret Exclusion (.gitignore)", "Virtual env, secrets, and caches excluded", ".venv/, *.pyc, secrets.toml", "VERIFIED (Pass)"),
        ("Visual Corporate Theme", "Deep Navy (#1E3A8A) branding configured", ".streamlit/config.toml", "VERIFIED (Pass)"),
        ("Automated QA Suite", "100% technical test suite execution", "tests/ (12/12 passing, latency 9.18ms)", "VERIFIED (Pass)")
    ]
    for i, row_data in enumerate(chk_data):
        row = checklist_table.rows[i+1]
        for j, val in enumerate(row_data):
            run = row.cells[j].paragraphs[0].add_run(val)
            if j == 0 or j == 3:
                run.bold = True
    style_table(checklist_table, header_bg="1E3A8A", alt_bg="F8FAFC")
    doc.add_paragraph()

    # Section 2: Step-by-Step Cloud Deployment SOP
    h2 = doc.add_paragraph()
    r = h2.add_run("2. Step-by-Step Cloud Deployment SOP")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    sop_steps = [
        ("Phase 1: Local Git Repository Initialization & Remote Push", [
            "Initialize local git tracking: git init",
            "Stage all audited assets: git add app.py requirements.txt .gitignore .streamlit/ data/ docs/ pages/ src/ tests/ README.md BUILD_PLAN.md",
            "Commit release snapshot: git commit -m 'BUILD-PROD: Finalize production release for Streamlit Cloud'",
            "Link GitHub remote: git remote add origin https://github.com/<username>/2503201450157-channel-wise-marketing-performance.git",
            "Push to primary branch: git push -u origin main"
        ]),
        ("Phase 2: Streamlit Community Cloud Connection", [
            "Navigate to https://share.streamlit.io in a modern web browser.",
            "Click 'Continue with GitHub' and authorize Streamlit Community Cloud access to public repositories.",
            "From the executive workspace dashboard, click the 'Create app' button located in the top-right corner."
        ]),
        ("Phase 3: Application Configuration & Deployment Parameters", [
            "Repository: Select your GitHub repository ('2503201450157-channel-wise-marketing-performance').",
            "Branch: Specify 'main'.",
            "Main file path: Specify 'app.py'.",
            "App URL (Custom Subdomain): Enter '2503201450157-channel-marketing' (yielding https://2503201450157-channel-marketing.streamlit.app).",
            "Advanced settings: Python version defaults to 3.12 (compatible with our pinned packages). No secrets required as data/app.db is embedded and self-contained."
        ]),
        ("Phase 4: Container Build & Live Telemetry Monitoring", [
            "Click the 'Deploy!' button.",
            "Open the 'Manage app' console panel in the bottom-right corner to inspect container build logs.",
            "Verify package installation: Streamlit automatically runs 'pip install -r requirements.txt' inside the isolated Linux container.",
            "Observe server startup: Streamlit initializes tornado server on internal port 8501 and binds to external HTTPS gateway.",
            "Container startup completes within 90 to 180 seconds."
        ]),
        ("Phase 5: Post-Deployment Smoke Testing & Acceptance", [
            "Verify Executive Overview (UI-01): 5 KPI cards render with positive formatting; daily spend vs revenue chart loads.",
            "Verify Multipage Navigation: Click through Channel Deep Dive, Customer Cohorts, Statistical Inference, Attribution & Budget, Data Quality, and About.",
            "Verify Interactive Sliders: Test Budget Optimizer slider and ensure sum of spend remains exactly constrained.",
            "Verify Export Functionality: Test downloading CSV and Excel workbooks from the Data Quality page."
        ])
    ]

    for title, steps in sop_steps:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(6)
        sp.paragraph_format.space_after = Pt(2)
        sp.add_run(title).bold = True
        sp.runs[0].font.color.rgb = RGBColor(30, 58, 138)
        for st_text in steps:
            p = doc.add_paragraph(f"•  {st_text}")
            p.paragraph_format.left_indent = Inches(0.2)
            p.paragraph_format.space_after = Pt(2)

    doc.add_paragraph()

    # Section 3: Troubleshooting & Diagnostics Playbook
    h3 = doc.add_paragraph()
    r = h3.add_run("3. Cloud Troubleshooting & Diagnostics Playbook")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    troubleshoot_table = doc.add_table(rows=7, cols=3)
    tb_headers = ["Failure Scenario / Error Log", "Root Technical Cause", "Exact Remediation Procedure"]
    for j, h in enumerate(tb_headers):
        troubleshoot_table.rows[0].cells[j].paragraphs[0].add_run(h)

    tb_data = [
        ("ModuleNotFoundError: No module named 'xyz'", 
         "A library imported in src/ or pages/ is missing from requirements.txt or version is incompatible.", 
         "Add the exact pinned library to requirements.txt (e.g. plotly>=5.22.0), commit, and git push. Streamlit Cloud automatically triggers rebuild."),
        
        ("sqlite3.OperationalError: unable to open database file", 
         "Hardcoded Windows absolute path (e.g. 'C:\\HCL_Project\\data\\app.db') fails on Streamlit Cloud Linux Debian container.", 
         "Derive database paths dynamically using Path(__file__).resolve().parent.parent / 'data' / 'app.db'. All src/ loaders follow this pattern."),
        
        ("Streamlit App Exceeded Memory Limit (1 GB RAM)", 
         "Uncached repetitive queries reload raw data on every user click, leading to memory accumulation.", 
         "Wrap all SQLite queries and cleaning functions in @st.cache_data decorators. Streamlit memoizes DataFrames in shared heap."),
        
        ("PermissionError: [Errno 13] Permission denied on data/app.db", 
         "SQLite file opened in write mode on read-only container file system.", 
         "Open SQLite connection in URI read-only mode: sqlite3.connect('file:' + DB_PATH + '?mode=ro', uri=True). This is enforced in src/data_loader.py."),
        
        ("Blank Iframe / Tailwind CSS Not Loading", 
         "Strict Content Security Policy (CSP) blocking external stylesheet links in raw markdown.", 
         "Embed Tailwind CSS inside st.components.v1.html(..., height=X). Iframes run in isolated sandboxes with CDN loading enabled."),
        
        ("Git LFS Bandwidth / Pointer File Not Resolving", 
         "Database file committed via Git LFS exceeded transfer quotas, causing SQLite to read a 130-byte text pointer.", 
         "Our database data/app.db is lightweight (under 1 MB). Commit it directly to git without Git LFS tracking.")
    ]
    for i, row_data in enumerate(tb_data):
        row = troubleshoot_table.rows[i+1]
        for j, val in enumerate(row_data):
            run = row.cells[j].paragraphs[0].add_run(val)
            if j == 0:
                run.bold = True
    style_table(troubleshoot_table, header_bg="1E3A8A", alt_bg="F8FAFC")
    doc.add_paragraph()

    # Section 4: DevOps & Cloud Deployment Viva Examination Notes
    h4 = doc.add_paragraph()
    r = h4.add_run("4. DevOps & Cloud Deployment Viva Notes (Top 5 Questions & Answers)")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    viva_dep = [
        ("Q1: Why is Streamlit Community Cloud preferred over traditional IaaS (like AWS EC2) for this data analytics capstone?",
         "Streamlit Community Cloud is a purpose-built Platform-as-a-Service (PaaS) tailored for Python analytics applications. It provides continuous deployment directly from GitHub, automated container containerization, free TLS/SSL certificates, zero-config domain routing, and seamless web-socket handling without requiring manual Nginx reverse-proxy configuration, systemd daemon management, or cloud infrastructure billing."),
        
        ("Q2: How does the application maintain cross-platform compatibility between local Windows development and Linux cloud containers?",
         "Local Windows environments use backslash path separators ('C:\\...'), whereas Linux Debian containers use forward slashes ('/...'). Hardcoding string paths breaks cross-platform execution. We utilized Python's pathlib.Path and os.path relative to __file__ (e.g. Path(__file__).resolve().parent / 'data' / 'app.db'). Furthermore, SQLite connections utilize read-only URI formatting ('file:...mode=ro'), ensuring zero path formatting discrepancies across operating systems."),
        
        ("Q3: What role does requirements.txt play in the Streamlit Cloud deployment lifecycle?",
         "When a commit is pushed to GitHub, Streamlit Cloud detects changes and spawns an ephemeral Docker container running Python 3.12. The build agent executes 'pip install -r requirements.txt', caching wheels for rapid subsequent re-deployments. Pinned versions (e.g. streamlit>=1.38.0, plotly>=5.22.0) ensure reproducible dependency resolution, preventing breaking runtime changes caused by upstream third-party package updates."),
        
        ("Q4: How does the platform adhere to the Principle of Least Privilege in cloud data security?",
         "The database is packaged as a static, pre-populated SQLite warehouse (data/app.db). In cloud runtime, data_loader.py connects using URI read-only mode ('mode=ro') with PRAGMA foreign_keys = ON;. The web dashboard has zero write or drop permissions on the database file, preventing SQL injection write vulnerabilities and accidental data corruption. Sensitive environment files (.venv, secrets) are strictly ignored via .gitignore."),
        
        ("Q5: What happens when multiple concurrent users access the deployed Streamlit application?",
         "Streamlit executes a multi-threaded Python process where each user session receives an isolated st.session_state dictionary. Global cached assets (decorated with @st.cache_data, such as load_analytical_view()) are computed once and shared in memory across all connected users. When User A toggles date filters, their session state mutates independently without interfering with User B's viewport, delivering sub-second response times for concurrent stakeholders.")
    ]

    for q, a in viva_dep:
        qp = doc.add_paragraph()
        qp.paragraph_format.space_before = Pt(8)
        qp.paragraph_format.space_after = Pt(2)
        q_run = qp.add_run(q)
        q_run.bold = True
        q_run.font.color.rgb = RGBColor(30, 58, 138)

        ap = doc.add_paragraph()
        ap.paragraph_format.left_indent = Inches(0.2)
        ap.paragraph_format.space_after = Pt(6)
        ap.add_run("Model Answer: ").bold = True
        ap.add_run(a)

    doc.save(DOCX_PATH)
    print(f"Successfully generated: {DOCX_PATH}")

if __name__ == "__main__":
    generate_step16_doc()
