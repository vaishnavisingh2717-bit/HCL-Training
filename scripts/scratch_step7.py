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

def generate_step7_trd():
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
    title_run = title_p.add_run("TECHNICAL REQUIREMENTS DOCUMENT (TRD / TBD)\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nTechnical Blueprint Specification | Fixed Stack Architecture\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Identifier", "DOC-TRD-V1.0 (Step 7 Technical Blueprint)"),
        ("Author / Solution Architect", "Vaishnavi Singh (Roll No: 2503201450157, MCA 2nd Yr DA, ABES)"),
        ("Project Mentor", "Prof. Kalpana Dwivedi (ABES Engineering College)"),
        ("Core Technology Stack", "Python 3.12, Streamlit Multipage, SQLite (app.db), Plotly, Tailwind UI, pytest"),
        ("Target Cloud Environment", "Streamlit Community Cloud (Public Git Deployment)"),
        ("Architectural Paradigm", "Modular 3-Tier Layered Architecture (Data -> Analysis -> Presentation)")
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

    # 1. Architecture (Data -> Load -> Clean -> Analyse -> Pages)
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Technical System Architecture & End-to-End Flow").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "The application is engineered as a high-performance, decoupled three-tier analytical pipeline. "
        "The architecture guarantees deterministic data ingestion, in-memory caching, vectorized analytical processing, "
        "and reactive UI rendering without unnecessary database round-trips."
    )

    doc.add_paragraph(
        "A. Architecture Flow Definition:\n"
        "1. Data Layer: Normalized relational database stored in SQLite (data/app.db) containing dimensional masters and granular performance fact tables.\n"
        "2. Ingestion & Cache Layer: src/data_loader.py executes parameterized SQL queries wrapped in Streamlit's @st.cache_data decorator.\n"
        "3. Cleaning & Validation Layer: src/cleaner.py and src/validator.py enforce strict datatypes, impute missing values, tag outliers, and assert relational integrity.\n"
        "4. Analytical & Statistical Engine: Modular compute packages (kpi_metrics.py, statistical_tests.py, attribution.py, optimizer.py) process core formulas and inferential tests.\n"
        "5. Presentation Layer: Multi-page Streamlit application (app.py and pages/) integrating Tailwind CSS components and interactive Plotly figures."
    )

    doc.add_paragraph("B. Mermaid Architecture Blueprint (Text Specification):")
    arch_code = (
        "graph TD\n"
        "  subgraph Data_Tier ['1. Persistent Data Layer']\n"
        "    DB[('SQLite data/app.db')] -->|Read SQL| DL['src/data_loader.py']\n"
        "  end\n"
        "  subgraph Processing_Tier ['2. Data Cleaning & Validation Layer']\n"
        "    DL -->|Cached DataFrames| CL['src/cleaner.py']\n"
        "    CL -->|Validated Telemetry| VAL['src/validator.py']\n"
        "  end\n"
        "  subgraph Analytical_Tier ['3. Modular Analytics & Statistics Engine']\n"
        "    VAL --> KPI['src/kpi_metrics.py (ROAS, CAC, CTR, CVR, LTV)']\n"
        "    VAL --> STAT['src/statistical_tests.py (Z-Test, ANOVA, Chi-Sq)']\n"
        "    VAL --> ATTR['src/attribution.py (First, Last, Linear)']\n"
        "    VAL --> OPT['src/optimizer.py (Marginal Returns Heuristic)']\n"
        "  end\n"
        "  subgraph Presentation_Tier ['4. Streamlit UI & Visual Dashboard']\n"
        "    KPI & STAT & ATTR & OPT --> UI['src/ui_components.py (Tailwind CDN)']\n"
        "    KPI & STAT & ATTR & OPT --> VIZ['src/visualizations.py (Plotly)']\n"
        "    UI & VIZ --> P0['app.py (Overview & Executive KPIs)']\n"
        "    UI & VIZ --> P1['pages/1_Channel_Deep_Dive.py']\n"
        "    UI & VIZ --> P2['pages/2_Statistical_Inference.py']\n"
        "    UI & VIZ --> P3['pages/3_Attribution_&_Budget.py']\n"
        "    UI & VIZ --> P4['pages/4_Data_Quality_&_Export.py']\n"
        "  end"
    )
    p_arch = doc.add_paragraph()
    r_arch = p_arch.add_run(arch_code)
    r_arch.font.name = 'Consolas'
    r_arch.font.size = Pt(8.5)
    p_arch.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. Folder Structure
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Production Directory & Codebase Structure").font.color.rgb = RGBColor(30, 58, 138)

    folder_text = (
        "c:\\HCL_Project\\\n"
        "├── .streamlit/\n"
        "│   └── config.toml               # Streamlit UI theme, colors, font & layout configuration\n"
        "├── app.py                        # Application entry point: Executive Overview & KPI Dashboard\n"
        "├── pages/                        # Multipage Streamlit application screens\n"
        "│   ├── 1_Channel_Deep_Dive.py    # CTR/CVR funnels, temporal trends & ad fatigue analysis\n"
        "│   ├── 2_Statistical_Inference.py# Inferential hypothesis testing (Two-sample Z-test, ANOVA)\n"
        "│   ├── 3_Attribution_&_Budget.py # Multi-touch attribution comparison & budget reallocation simulator\n"
        "│   └── 4_Data_Quality_&_Export.py# Schema audits, anomaly detection & CSV/Excel export tools\n"
        "├── src/                          # Reusable Python business logic modules\n"
        "│   ├── __init__.py\n"
        "│   ├── data_loader.py            # SQLite database connector & @st.cache_data query functions\n"
        "│   ├── cleaner.py                # Datatype coercion, missing value handling, outlier tagging\n"
        "│   ├── validator.py              # Foreign key assertions, null checks, schema rules validation\n"
        "│   ├── kpi_metrics.py            # Financial & performance metrics math (ROAS, CAC, LTV:CAC)\n"
        "│   ├── statistical_tests.py      # SciPy / statsmodels hypothesis testing functions\n"
        "│   ├── attribution.py            # Attribution models (First-Touch, Last-Touch, Linear)\n"
        "│   ├── optimizer.py              # Marginal return curves & budget shift optimization\n"
        "│   ├── ui_components.py          # Tailwind CSS component renderers via st.components.v1.html\n"
        "│   └── visualizations.py         # Plotly Express and Graph Objects interactive chart generators\n"
        "├── data/                         # Persistent database & generated synthetic files\n"
        "│   ├── app.db                    # SQLite relational database containing all tables\n"
        "│   ├── marketing_data.xlsx       # Multi-sheet Excel workbook of raw source tables\n"
        "│   └── schema.sql                # SQLite DDL creation script\n"
        "├── tests/                        # Automated testing suite (pytest)\n"
        "│   ├── __init__.py\n"
        "│   ├── test_data.py              # Database integrity, schema constraints & data cleaning tests\n"
        "│   ├── test_analysis.py          # Unit tests for KPI calculations and statistical test validity\n"
        "│   └── test_app.py               # Streamlit AppTest integration tests for UI page rendering\n"
        "├── docs/                         # Official project documentation and academic artifacts\n"
        "├── requirements.txt              # Pinned Python package dependencies\n"
        "├── README.md                     # Setup instructions, architecture overview & live deployment URL\n"
        "└── .gitignore                    # Git version control exclusions (.venv, __pycache__, .DS_Store)"
    )
    p_fold = doc.add_paragraph()
    r_fold = p_fold.add_run(folder_text)
    r_fold.font.name = 'Consolas'
    r_fold.font.size = Pt(8.5)
    p_fold.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Functional Requirements (FR-01 to FR-12) & Modules
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Functional Requirements Specification (FR-01 to FR-12)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Technical functional requirements translate business requirements (BR-01 to BR-12) into specific software modules:"
    )

    fr_table = doc.add_table(rows=13, cols=5)
    fr_headers = ["FR ID", "Functional Requirement Description", "Linked BR", "Primary Python Module(s)", "Delivery Target Screen"]
    for j, fh in enumerate(fr_headers):
        fr_table.rows[0].cells[j].text = fh

    fr_data = [
        ("FR-01", "Query SQLite fact tables and calculate aggregated spend, revenue, impressions, clicks, conversions, blended ROAS, and CAC.", "BR-01", "src/kpi_metrics.py, src/data_loader.py", "app.py (Overview)"),
        ("FR-02", "Filter campaigns violating benchmark thresholds (ROAS < 2.0x or CAC > ₹1,500) and render an interactive alert table.", "BR-02", "src/kpi_metrics.py, src/ui_components.py", "app.py (Overview)"),
        ("FR-03", "Compute CTR (Clicks/Impr) and CVR (Conversions/Clicks) over time and generate interactive Plotly multi-stage funnel and line charts.", "BR-03", "src/kpi_metrics.py, src/visualizations.py", "pages/1_Channel_Deep_Dive.py"),
        ("FR-04", "Calculate customer repeat purchase frequency, cumulative revenue, and LTV:CAC ratios per acquisition channel.", "BR-04", "src/kpi_metrics.py, src/visualizations.py", "pages/1_Channel_Deep_Dive.py"),
        ("FR-05", "Decompose daily time-series data into day-of-week and monthly patterns with 7-day rolling centered moving averages.", "BR-05", "src/cleaner.py, src/visualizations.py", "pages/1_Channel_Deep_Dive.py"),
        ("FR-06", "Segment conversion volumes and average order values across age groups, gender, and city tiers using stacked bar and sunburst charts.", "BR-06", "src/visualizations.py", "pages/1_Channel_Deep_Dive.py"),
        ("FR-07", "Execute Two-sample proportion Z-tests, One-way ANOVA, and Chi-Square tests of independence, outputting test stats, p-values, and 95% CI.", "BR-07", "src/statistical_tests.py", "pages/2_Statistical_Inference.py"),
        ("FR-08", "Ingest touchpoint logs and compute conversion credit under First-Touch, Last-Touch, and Equal Linear models, displaying comparative bars.", "BR-08", "src/attribution.py, src/visualizations.py", "pages/3_Attribution_&_Budget.py"),
        ("FR-09", "Simulate cross-channel budget reallocations under fixed spend constraints using diminishing returns curves and dynamic sliders.", "BR-09", "src/optimizer.py, src/ui_components.py", "pages/3_Attribution_&_Budget.py"),
        ("FR-10", "Audit data completeness, detect missing UTM parameters, flag duplicate transaction IDs, and detect Z-score spend outliers (Z > 3).", "BR-10", "src/validator.py", "pages/4_Data_Quality_&_Export.py"),
        ("FR-11", "Manage global filter state (date ranges, channels, campaign types) across all multipage screens using st.session_state.", "BR-11", "src/data_loader.py, app.py", "Global Sidebar"),
        ("FR-12", "Export active filtered datasets, campaign summaries, and quality audit logs into downloadable CSV and formatted Excel files.", "BR-12", "pages/4_Data_Quality_&_Export.py", "pages/4_Data_Quality_&_Export.py")
    ]
    for r_idx, r_data in enumerate(fr_data, start=1):
        for c_idx, val in enumerate(r_data):
            fr_table.rows[r_idx].cells[c_idx].text = val
    style_table(fr_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. Data Layer Architecture
    h4 = doc.add_heading(level=1)
    h4.add_run("4. Data Layer Architecture & In-Memory Caching").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "A. Database Connection & Mode:\n"
        "The application interfaces with SQLite (data/app.db) using standard Python sqlite3 with URI read-only flags "
        "('file:data/app.db?mode=ro') during production dashboard execution. Foreign key constraints are enforced on every connection "
        "via 'PRAGMA foreign_keys = ON;'. Parameterized SQL queries ensure sub-second analytical execution."
    )
    doc.add_paragraph(
        "B. Streamlit Caching Optimization (@st.cache_data):\n"
        "To achieve sub-second filter responsiveness, data loading and heavy aggregations are decorated with @st.cache_data(ttl=3600). "
        "The caching key hashes the SQL query string and parameter tuples. When a user toggles sidebar filters, Streamlit filters "
        "the in-memory pandas DataFrame in microsecond time rather than re-querying disk storage."
    )
    doc.add_paragraph(
        "C. Schema Validation & Datatype Enforcement:\n"
        "During ingestion in src/validator.py, schema contracts are verified: non-null primary keys, valid foreign key references, "
        "positive numeric bounds for spend and revenue, and standardized datetime parsing (pd.to_datetime with UTC alignment)."
    )

    # 5. Statistical Methods Mapped to Every Operational Question
    h5 = doc.add_heading(level=1)
    h5.add_run("5. Statistical Methodology Mapped to Operational Questions (OQ-01 to OQ-10)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Every operational question is underpinned by a formal mathematical or statistical algorithm:"
    )

    stat_table = doc.add_table(rows=11, cols=4)
    stat_headers = ["OQ ID", "Operational Inquiry", "Core Statistical / Mathematical Formulation", "Implementation Library"]
    for j, sh in enumerate(stat_headers):
        stat_table.rows[0].cells[j].text = sh

    stat_data = [
        ("OQ-01", "Channel ROAS & CAC", "ROAS = Sum(Revenue) / Sum(Spend); CAC = Sum(Spend) / Sum(Conversions); 95% Confidence Intervals", "NumPy / pandas"),
        ("OQ-02", "CTR & CVR Funnel Dynamics", "CTR = Clicks / Impressions; CVR = Conversions / Clicks; Standard Error of Proportions SE = sqrt(p(1-p)/n)", "SciPy (scipy.stats)"),
        ("OQ-03", "LTV:CAC Ratio by Channel", "Cumulative Cohort Revenue = Sum(Order_Value); LTV = Avg Revenue * Purchase Frequency; LTV:CAC Ratio", "pandas / NumPy"),
        ("OQ-04", "Temporal Seasonality Trends", "Additive Time-Series Decomposition: Y_t = Trend_t + Seasonal_t + Residual_t; 7-Day Centered Moving Average", "statsmodels.tsa"),
        ("OQ-05", "Demographic Segment Affinities", "Two-Way Contingency Cross-Tabulation; Chi-Square Test of Independence: Chi2 = Sum((O - E)^2 / E)", "scipy.stats.chi2_contingency"),
        ("OQ-06", "Statistical Significance of Conversion Gaps", "Two-Sample Pooled Z-Test for Proportions: Z = (p1 - p2) / sqrt(p_pool * (1 - p_pool) * (1/n1 + 1/n2)); One-Way ANOVA F-Test", "scipy.stats.ttest_ind / f_oneway"),
        ("OQ-07", "Underperforming Ad Outliers", "Z-Score Outlier Detection: Z = (X - Mean) / StdDev (Z > 3 flagged); Interquartile Range: IQR = Q3 - Q1", "scipy.stats.zscore"),
        ("OQ-08", "Multi-Touch Journey Attribution", "Rule-Based Allocation: First-Touch (w_first=1.0), Last-Touch (w_last=1.0), Linear (w_i = 1/N for i in 1..N)", "Custom vectorized src/attribution.py"),
        ("OQ-09", "Constrained Budget Optimization", "Diminishing Returns Power Law: Revenue = alpha * Spend^beta (0 < beta < 1); Marginal ROAS Equimarginal Allocation", "scipy.optimize / Heuristic"),
        ("OQ-10", "Data Quality & Anomaly Auditing", "Missing Rate = Null_Count / Total_Rows; Duplicate Key Flagging: df.duplicated(subset=['ID']).sum()", "Custom src/validator.py")
    ]
    for r_idx, r_data in enumerate(stat_data, start=1):
        for c_idx, val in enumerate(r_data):
            stat_table.rows[r_idx].cells[c_idx].text = val
    style_table(stat_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 6. UI Notes: Tailwind, Theme & Layout
    h6 = doc.add_heading(level=1)
    h6.add_run("6. UI/UX Design System, Tailwind CSS & Theme Specification").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "A. Tailwind CSS CDN in st.components.v1.html:\n"
        "To deliver a modern UI, high-impact visual components (executive headers, metric cards, status badges, and alert banners) "
        "are rendered using Tailwind CSS loaded via CDN (<script src='https://cdn.tailwindcss.com'></script>) inside isolated Streamlit components. "
        "This approach bypasses Streamlit's rigid default container styling while guaranteeing zero CSS bleed."
    )
    doc.add_paragraph(
        "B. Theme Palette (.streamlit/config.toml):\n"
        "• Primary Accent: #1E3A8A (Deep Navy Blue)\n"
        "• Background: #F8FAFC (Light Slate / Slate 50)\n"
        "• Secondary Background: #FFFFFF (Clean Card Surface)\n"
        "• Text Color: #0F172A (Slate 900 High Contrast)\n"
        "• Success / Warning / Danger Accents: #10B981 (Emerald Green), #F59E0B (Amber), #EF4444 (Crimson Red)\n"
        "• Font Family: Inter / Segoe UI / sans-serif"
    )

    # 7. Performance, Error Handling & Security
    h7 = doc.add_heading(level=1)
    h7.add_run("7. Performance Optimization, Error Handling & Security").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "• Query Latency Optimization: Compound B-tree indexes on SQLite fact tables (Date, Campaign_ID, Channel_ID) ensure queries execute in < 20 ms.\n"
        "• Memory Management: Streamlit session state stores only lightweight filter parameters. DataFrames are cached in memory and sliced via zero-copy views.\n"
        "• Error Boundaries: Every data loading and analytical calculation is wrapped in try-except blocks. Missing parameters or empty filter results render user-friendly guidance messages (st.info / st.warning) rather than raw Python tracebacks.\n"
        "• Security Invariants: Zero live credentials or secrets required. Parameterized SQL prevents SQL injection. Zero real personally identifiable information (100% synthetic dataset)."
    )

    # 8. requirements.txt (Pinned)
    h8 = doc.add_heading(level=1)
    h8.add_run("8. Pinned Package Dependencies (requirements.txt)").font.color.rgb = RGBColor(30, 58, 138)

    req_text = (
        "streamlit==1.40.1\n"
        "pandas==2.2.3\n"
        "numpy==1.26.4\n"
        "scipy==1.14.1\n"
        "statsmodels==0.14.4\n"
        "plotly==5.24.1\n"
        "openpyxl==3.1.5\n"
        "python-docx==1.1.2\n"
        "pytest==8.3.3"
    )
    p_req = doc.add_paragraph()
    r_req = p_req.add_run(req_text)
    r_req.font.name = 'Consolas'
    r_req.font.size = Pt(9.5)
    p_req.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 9. Deployment Requirements (Streamlit Community Cloud)
    h9 = doc.add_heading(level=1)
    h9.add_run("9. Deployment Blueprint (Streamlit Community Cloud)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "• Git Repository Standards: Public GitHub repository with root-level app.py, requirements.txt, and .streamlit/config.toml.\n"
        "• Database Bundling: SQLite database data/app.db committed directly into the Git repository, guaranteeing zero external database hosting fees or connection latency.\n"
        "• Filesystem Path Portability: All file paths reference pathlib.Path(__file__).parent / 'data' / 'app.db', ensuring seamless resolution across Windows local development and Linux Streamlit Cloud containers.\n"
        "• Cloud Resource Caps: App designed to operate well within Streamlit Community Cloud limits (1 vCPU, 1 GB RAM)."
    )

    # 10. Technical Risks & Mitigation
    h10 = doc.add_heading(level=1)
    h10.add_run("10. Technical Risk Management Matrix").font.color.rgb = RGBColor(30, 58, 138)

    tech_risk_table = doc.add_table(rows=5, cols=4)
    tr_headers = ["Risk ID", "Technical Risk Description", "Severity", "Architectural Mitigation Strategy"]
    for j, th in enumerate(tr_headers):
        tech_risk_table.rows[0].cells[j].text = th

    tr_data = [
        ("TR-01", "SQLite concurrency lock errors during concurrent dashboard reads.", "High", "Open database connections in explicit read-only URI mode ('file:data/app.db?mode=ro') with timeout=10."),
        ("TR-02", "Large Plotly chart rendering degrading browser DOM performance.", "Medium", "Downsample dense time-series data using 7-day rolling aggregates and enable WebGL rendering in Plotly."),
        ("TR-03", "Streamlit Cloud container rebooting due to out-of-memory (OOM) spikes.", "High", "Avoid storing heavy duplicate DataFrames in st.session_state; rely strictly on @st.cache_data."),
        ("TR-04", "Cross-platform path incompatibilities between Windows (backslashes) and Linux (forward slashes).", "Medium", "Enforce pathlib.Path across all Python scripts, guaranteeing operating-system-agnostic path resolution.")
    ]
    for r_idx, r_data in enumerate(tr_data, start=1):
        for c_idx, val in enumerate(r_data):
            tech_risk_table.rows[r_idx].cells[c_idx].text = val
    style_table(tech_risk_table)

    # Save document
    path1 = os.path.abspath("docs/07_TRD.docx")
    path2 = os.path.abspath("docs/07_Technical_Blueprint_Document_TBD.docx")
    doc.save(path1)
    doc.save(path2)
    print(f"Saved: {path1}")
    print(f"Saved: {path2}")

if __name__ == "__main__":
    generate_step7_trd()
