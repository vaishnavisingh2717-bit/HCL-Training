"""
scratch_step14.py - Generates BUILD_PLAN.md and docs/14_BUILD_PLAN.docx
Covers Step 14 of the Channel Wise Marketing Performance Analysis capstone project.
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
MD_PATH = os.path.join(WORKSPACE_DIR, "BUILD_PLAN.md")
DOCX_PATH = os.path.join(DOCS_DIR, "14_BUILD_PLAN.docx")

# Styling Helpers for python-docx
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

# 1. Generate BUILD_PLAN.md
def generate_markdown():
    content = r"""# BUILD PLAN & TECHNICAL IMPLEMENTATION ROADMAP
## Channel Wise Marketing Performance Analysis
**Learner:** Vaishnavi Singh (Roll No: `2503201450157`)  
**Program:** MCA 2nd Year Data Analytics, ABES Engineering College  
**Faculty Mentor:** Prof. Kalpana Dwivedi  
**Date:** November 20, 2026 | **Version:** 1.0 (Production Master)

---

## 1. Project Summary & Traceability Links

1. **Platform Mission:** "Channel Wise Marketing Performance Analysis" is an enterprise-grade analytics decision platform developed by Vaishnavi Singh to systematically eliminate marketing spend waste across 5 core channels.
2. **Channel Telemetry Scope:** Ingests granular advertising data across Google Search Ads, Meta Performance Ads, Email Marketing, Influencer Partnerships, and Organic Search to resolve fragmented performance visibility.
3. **Relational Data Warehouse:** Backed by a normalized SQLite 3 data store (`data/app.db`) comprising 5 core entities and an integrated reporting view (`vw_marketing_analysis`) with 262 records and 38 analytical dimensions.
4. **Deterministic Math Engine:** Implements exact, audited formulas for Return on Ad Spend (ROAS), Customer Acquisition Cost (CAC), CTR, CPC, CVR, AOV, and Net Profit with automated threshold alerts for bleeding campaigns.
5. **Statistical Rigor:** Features a two-sample proportion Z-test engine proving weekend conversion surges (+34.3% uplift, Z=4.21, p < 0.0001), alongside OLS multi-variate regression and Pearson correlation matrices.
6. **Multi-Touch Attribution:** Solves last-click attribution bias by modeling First-Touch, Last-Touch, and Linear credit distributions across 450 customer touchpoint journeys.
7. **Budget Reallocation Optimizer:** Features an algorithmic capital simulator that enforces a strict 100% total budget constraint while projecting a +22.4% blended ROAS improvement.
8. **Executive UI Framework:** Delivered as a Streamlit multipage web application styled with isolated Tailwind CSS component iframes via CDN, delivering ultra-low query latency (< 10 ms).
9. **Traceability & QA:** Backed by a 100% finalized Requirements Traceability Matrix linking all 12 Business Requirements across 14 Functional Tests (`TC-F-01` to `TC-F-14`) and 12 Automated Pytest Suites (`TC-T-01` to `TC-T-12`).
10. **Deployment Readiness:** Fully packaged with pinned dependencies (`requirements.txt`), comprehensive theme definitions (`.streamlit/config.toml`), and production deployment guides for Streamlit Community Cloud.

### Project Artifact & Document Directory Links
- [01_Project_Understanding.docx](file:///c:/HCL_Project/docs/01_Project_Understanding.docx) — Business Problem Framing & Operational Questions (OQ-01 to OQ-10)
- [02_HMW_Statements.docx](file:///c:/HCL_Project/docs/02_HMW_Statements.docx) — Design Thinking & "How Might We" Innovation Framework (HMW-01 to HMW-10)
- [03_Needs_Analysis.docx](file:///c:/HCL_Project/docs/03_Needs_Analysis.docx) — MoSCoW Prioritized Business Requirements (BR-01 to BR-12)
- [04_Traceability_Matrix.xlsx](file:///c:/HCL_Project/docs/04_Traceability_Matrix.xlsx) — Master Requirements Traceability Matrix (100% Audited)
- [05_BRD.docx](file:///c:/HCL_Project/docs/05_BRD.docx) — Formal Business Requirements Document & RACI Governance
- [07_TRD.docx](file:///c:/HCL_Project/docs/07_TRD.docx) — Technical Requirements Document & 3-Tier System Architecture
- [08_Database_Design.xlsx](file:///c:/HCL_Project/docs/08_Database_Design.xlsx) — Relational Data Dictionary, Schema Specifications & DDL
- [09_ER_Diagram.docx](file:///c:/HCL_Project/docs/09_ER_Diagram.docx) / [09_ER_Diagram.png](file:///c:/HCL_Project/docs/09_ER_Diagram.png) — Visual Entity-Relationship Diagram & Business Rules
- [10_Dataset_Generation.docx](file:///c:/HCL_Project/docs/10_Dataset_Generation.docx) — Synthetic Dataset Generation Specifications & Quality Audit Log
- [11_User_Stories.docx](file:///c:/HCL_Project/docs/11_User_Stories.docx) — Agile User Stories (US-01 to US-14) with Given/When/Then Acceptance Criteria
- [12_UI_Framework.docx](file:///c:/HCL_Project/docs/12_UI_Framework.docx) — Tailwind CSS Dashboard Architecture, ASCII Wireframes & Chart Mappings
- [13_Test_Cases.docx](file:///c:/HCL_Project/docs/13_Test_Cases.docx) — Comprehensive Test Specification (TC-F-01..14 and TC-T-01..12)

---

## 2. Final Folder Structure

```text
c:\\HCL_Project\\
├── .streamlit\\
│   └── config.toml                  # Streamlit corporate theme (Deep Navy #1E3A8A) & port config
├── data\\
│   ├── app.db                       # Production SQLite 3 database with FK & CHECK constraints
│   ├── marketing_data.xlsx          # Master multi-tab Excel workbook with data dictionary
│   ├── dim_channel.csv              # Channel master table (5 records)
│   ├── dim_campaign.csv             # Campaign master table (25 records)
│   ├── dim_customer.csv             # Customer demographic master (150 records)
│   ├── fact_daily_campaign_performance.csv # Daily campaign performance fact table (262 records)
│   ├── fact_customer_touchpoint.csv # Multi-touch customer journey fact table (450 records)
│   ├── vw_marketing_analysis.csv    # Flattened analytical reporting view (262 rows, 38 columns)
│   └── validation_report.txt        # Automated dataset validation & sanity report
├── docs\\
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   ├── 07_TRD.docx
│   ├── 08_Database_Design.xlsx
│   ├── 09_ER_Diagram.docx
│   ├── 09_ER_Diagram.png
│   ├── 10_Dataset_Generation.docx
│   ├── 11_User_Stories.docx
│   ├── 12_UI_Framework.docx
│   ├── 13_Test_Cases.docx
│   ├── 14_BUILD_PLAN.docx           # Executive Word version of implementation roadmap
│   ├── 15_Build_Execution_Verification_Report.docx # Step 15 execution report
│   ├── 16_Deployment_Guide.docx     # Step 16 deployment SOP
│   ├── 17_Final_Project_Document.docx # Step 17 ABES master academic thesis
│   └── 18_Presentation_Deck.pptx    # Step 18 final viva presentation slide deck
├── pages\\
│   ├── 1_Channel_Deep_Dive.py       # UI-02: Channel comparative analytics & campaign scatter plots
│   ├── 2_Customer_Cohorts.py        # UI-03: Customer tier segments & demographic behavior heatmaps
│   ├── 3_Statistical_Inference.py   # UI-04: Z-test hypothesis testing, regression & correlations
│   ├── 4_Attribution_&_Budget.py    # UI-05: Multi-touch attribution modeling & budget simulator
│   ├── 5_Data_Quality_&_Export.py   # UI-06: Data health audit, anomaly logs & CSV/Excel exporter
│   └── 6_About.py                   # UI-07: Project architecture, student info & RTM overview
├── src\\
│   ├── __init__.py                  # Package initializer
│   ├── generate_data.py             # Seed-42 synthetic dataset generator & SQLite loader
│   ├── data_loader.py               # SQLite data ingestion, caching & query management
│   ├── cleaner.py                   # Data cleaning, null handling, deduplication & IQR imputation
│   ├── analysis.py                  # Core marketing & financial KPI computation engine
│   ├── statistics.py                # Hypothesis testing (Z-test), regression & ANOVA math
│   ├── attribution.py               # First-touch, last-touch & linear attribution models
│   ├── optimizer.py                 # Algorithmic budget reallocation optimization engine
│   └── ui.py                        # Tailwind CSS iframe components & Plotly visual themes
├── tests\\
│   ├── __init__.py                  # Test package initializer
│   ├── test_data.py                 # Automated data layer tests (TC-T-01 to TC-T-06)
│   ├── test_analysis.py             # Mathematical & statistical logic tests (TC-T-07 to TC-T-11)
│   └── test_app.py                  # UI components & query latency tests (TC-T-12)
├── app.py                           # UI-01: Main entrypoint & Executive Overview dashboard
├── BUILD_PLAN.md                    # Technical AI build input & developer execution guide
├── README.md                        # Production project documentation & execution guide
└── requirements.txt                 # Pinned Python package dependencies
```

---

## 3. Granular Build Tasks (BUILD-01 to BUILD-17)

### BUILD-01: Environment Setup, Dependencies & Theme Configuration
- **Goal:** Pin exact production dependencies, verify directory layout, and establish Streamlit corporate theme.
- **Files Modified/Created:** `requirements.txt`, `.streamlit/config.toml`
- **User Stories & Features:** `US-01` | `FR-01`, `F-01`
- **Exact AI Prompt:**
  ```text
  Create requirements.txt pinning Python 3.12 compatible libraries:
  streamlit>=1.38.0, pandas>=2.2.0, numpy>=1.26.0, scipy>=1.13.0, statsmodels>=0.14.0, plotly>=5.22.0, openpyxl>=3.1.2, pytest>=8.0.0.
  Verify .streamlit/config.toml contains primaryColor="#1E3A8A", backgroundColor="#F8FAFC", secondaryBackgroundColor="#F1F5F9", textColor="#0F172A", font="sans serif".
  ```
- **How to Check:** Verify file contents and ensure `pytest tests/test_app.py::test_config_toml_theme_file` passes.
- **Tests Passing:** `TC-T-12` (subtest: theme configuration).

---

### BUILD-02: SQLite Data Access Layer & Caching Engine
- **Goal:** Implement an abstraction layer for querying SQLite `data/app.db` with `@st.cache_data` caching and fallback for headless test runners.
- **Files Modified/Created:** `src/data_loader.py`, `src/__init__.py`
- **User Stories & Features:** `US-01`, `US-02` | `FR-01`, `FR-02`, `F-01`, `F-02`
- **Exact AI Prompt:**
  ```text
  Implement src/data_loader.py with functions:
  1. get_connection(read_only=True): returns sqlite3 connection with PRAGMA foreign_keys = ON.
  2. load_master_data(): returns dict of DataFrames for dim_channel, dim_campaign, dim_customer.
  3. load_analytical_view(): returns DataFrame for vw_marketing_analysis.
  4. load_touchpoints(): returns DataFrame for fact_customer_touchpoint.
  Wrap query functions in @st.cache_data when streamlit is imported, falling back gracefully to direct calls under headless pytest.
  ```
- **How to Check:** Execute `python -c "import src.data_loader as dl; df = dl.load_analytical_view(); print(df.shape)"`. Must return 262 rows, 38 columns.
- **Tests Passing:** `TC-T-01`, `TC-T-02`, `TC-T-03`, `TC-T-04`, `TC-T-12`.

---

### BUILD-03: Data Cleaning, Validation & Anomaly Tagging Engine
- **Goal:** Build data transformation, null imputation, deduplication, and IQR outlier identification functions.
- **Files Modified/Created:** `src/cleaner.py`
- **User Stories & Features:** `US-02`, `US-12` | `FR-02`, `FR-12`, `F-02`, `F-12`
- **Exact AI Prompt:**
  ```text
  Implement src/cleaner.py containing:
  1. clean_marketing_data(df):
     - Impute missing numeric values using channel-level medians.
     - Drop exact duplicate rows.
     - Detect spend and conversion outliers via IQR method (Q1 - 1.5*IQR to Q3 + 1.5*IQR).
     - Add boolean column 'is_outlier' and capped metrics for charting without dropping records.
     - Return tuple (cleaned_df, audit_metrics_dict).
  2. validate_data_integrity(df): asserts zero negative spend/conversions and clicks >= conversions.
  ```
- **How to Check:** Run `python -c "import src.data_loader as dl, src.cleaner as cl; df = dl.load_analytical_view(); c_df, audit = cl.clean_marketing_data(df); print(audit)"`.
- **Tests Passing:** `TC-T-05`, `TC-T-06`.

---

### BUILD-04: Core Analytical Metrics & Threshold Alert Engine
- **Goal:** Implement deterministic formulas for marketing/financial KPIs and campaign performance classification.
- **Files Modified/Created:** `src/analysis.py`
- **User Stories & Features:** `US-03`, `US-05` | `FR-03`, `FR-05`, `F-03`, `F-05`
- **Exact AI Prompt:**
  ```text
  Implement src/analysis.py containing:
  1. compute_kpis(df): returns dictionary of CTR, CPC, CVR, CAC, ROAS, AOV, Total Spend, Total Revenue, Net Profit.
  2. identify_underperforming_campaigns(df, roas_thresh=2.0, cac_thresh=1500.0): flags campaigns with ROAS < 2.0x or CAC > 1500.
  3. get_channel_summary(df): groups data by channel_name and aggregates spend, revenue, clicks, conversions, calculating aggregate ROAS and CAC.
  ```
- **How to Check:** Verify calculations against tiny deterministic dataset in `test_analysis.py`.
- **Tests Passing:** `TC-T-07`, `TC-T-08`.

---

### BUILD-05: Advanced Statistical Inference & Econometric Modeling
- **Goal:** Implement two-sample Z-test for conversion rate uplift, OLS multi-variate regression, and correlation matrix math.
- **Files Modified/Created:** `src/statistics.py`
- **User Stories & Features:** `US-08`, `US-09` | `FR-08`, `FR-09`, `F-08`, `F-09`
- **Exact AI Prompt:**
  ```text
  Implement src/statistics.py containing:
  1. two_sample_z_test_proportions(p1, n1, p2, n2): calculates pooled standard error, Z-score, and two-tailed p-value.
  2. calculate_weekend_uplift(df): divides data into weekend (Fri-Sun) and weekday (Mon-Thu), computes CVRs, and calls Z-test.
  3. run_ols_revenue_regression(df): fits OLS model (Revenue ~ Spend + Clicks + Impressions) returning coefficients, R2, and p-values.
  4. compute_correlation_matrix(df): returns Pearson correlation matrix for numeric marketing variables.
  ```
- **How to Check:** Run `python -c "import src.statistics as st; print(st.two_sample_z_test_proportions(0.038, 11400, 0.0283, 14800))"`.
- **Tests Passing:** `TC-T-09`.

---

### BUILD-06: Multi-Touch Attribution & Budget Optimization Simulator
- **Goal:** Implement First-Touch, Last-Touch, and Linear attribution models and constrained budget reallocation optimizer.
- **Files Modified/Created:** `src/attribution.py`, `src/optimizer.py`
- **User Stories & Features:** `US-10`, `US-11` | `FR-10`, `FR-11`, `F-10`, `F-11`
- **Exact AI Prompt:**
  ```text
  1. Implement src/attribution.py:
     - compute_attribution(touchpoints_df): aggregates conversions by customer journey under First-Touch, Last-Touch, and Linear models.
     - Assert sum of credited conversions matches total conversions across all models.
  2. Implement src/optimizer.py:
     - optimize_budget(channel_summary_df, total_budget, roas_weights):
       Reallocates budget prioritizing high-ROAS channels subject to min 5% and max 40% channel constraints, ensuring sum equals total_budget.
  ```
- **How to Check:** Run attribution on `fact_customer_touchpoint` and assert credit total equals conversion count.
- **Tests Passing:** `TC-T-10`, `TC-T-11`.

---

### BUILD-07: Tailwind CSS Component Library & Plotly Visual Theme Engine
- **Goal:** Deliver reusable Tailwind CSS UI iframe components and Plotly chart configuration templates.
- **Files Modified/Created:** `src/ui.py`
- **User Stories & Features:** `US-03`, `US-04` | `FR-03`, `FR-04`, `F-03`, `F-04`
- **Exact AI Prompt:**
  ```text
  Enhance src/ui.py to support:
  1. render_tailwind_header(title, subtitle, badge_text) via st.components.v1.html and CDN.
  2. render_kpi_card(title, value, delta, delta_type, subtitle).
  3. render_kpi_row(cards): displays responsive multi-card flex row.
  4. render_insight_box(title, text, type): renders color-coded alert banner (success, warning, info).
  5. apply_plotly_theme(fig): sets transparent background, navy accents (#1E3A8A), clean gridlines, Inter/Calibri font.
  ```
- **How to Check:** Run `pytest tests/test_app.py::test_ui_component_html_generation`.
- **Tests Passing:** `TC-T-12`.

---

### BUILD-08: UI Screen 1 — Executive Overview Dashboard (`app.py`)
- **Goal:** Implement the primary entrypoint dashboard with global filters, session state, 5 executive KPI cards, and trend/share charts.
- **Files Modified/Created:** `app.py`
- **User Stories & Features:** `US-03`, `US-04` | `FR-03`, `FR-04`, `F-03`, `F-04` | `UI-01`
- **Exact AI Prompt:**
  ```text
  Build app.py:
  1. Configure page (title="Marketing Intelligence Dashboard", layout="wide", initial_sidebar_state="expanded").
  2. Sidebar: Date Range slider/picker, Channel multiselect, Campaign Category multiselect. Store in st.session_state.
  3. Render Tailwind executive header with live status badge.
  4. Display 5 KPI cards in responsive row: Total Spend, Total Revenue, Net Profit, Blended ROAS, Blended CAC.
  5. Plotly Dual-Axis Line/Bar Chart: Daily Spend vs Revenue with 7-day rolling average.
  6. Plotly Donut Chart: Revenue share by channel.
  ```
- **How to Check:** Run `streamlit run app.py` and interact with sidebar filters.
- **Tests Passing:** `TC-F-01`, `TC-F-02`, `TC-F-03`, `TC-F-04`, `TC-T-12`.

---

### BUILD-09: UI Screen 2 — Channel Deep Dive & Campaign Analytics (`pages/1_Channel_Deep_Dive.py`)
- **Goal:** Build channel comparative bar charts, ROAS vs CAC scatter plot with threshold lines, and campaign table.
- **Files Modified/Created:** `pages/1_Channel_Deep_Dive.py`
- **User Stories & Features:** `US-05`, `US-06` | `FR-05`, `FR-06`, `F-05`, `F-06` | `UI-02`
- **Exact AI Prompt:**
  ```text
  Implement pages/1_Channel_Deep_Dive.py:
  1. Read filtered data from st.session_state or data loader.
  2. Channel Performance Summary Table with colored indicators.
  3. Plotly Scatter Plot: X=CAC, Y=ROAS, Size=Spend, Color=Channel.
     Add dashed horizontal threshold at ROAS=2.0x, dashed vertical threshold at CAC=₹1500.
  4. Filterable Campaign Data Table with visual tags for 'Underperforming' campaigns.
  ```
- **How to Check:** Select "Channel Deep Dive" in Streamlit navigation; check scatter plot quadrant lines.
- **Tests Passing:** `TC-F-05`, `TC-F-06`, `TC-T-08`.

---

### BUILD-10: UI Screen 3 — Customer Cohort & Demographic Analytics (`pages/2_Customer_Cohorts.py`)
- **Goal:** Implement customer tier breakdown (VIP, Repeat, One-Time), touchpoint journey length histogram, and demographic heatmap.
- **Files Modified/Created:** `pages/2_Customer_Cohorts.py`
- **User Stories & Features:** `US-07` | `FR-07`, `F-07` | `UI-03`
- **Exact AI Prompt:**
  ```text
  Implement pages/2_Customer_Cohorts.py:
  1. Customer Segment Metrics: Total Customers, VIP Share, Repeat Share, Average Customer LTV.
  2. Plotly Bar Chart: Average Order Value (AOV) and Revenue across Customer Tiers.
  3. Plotly Histogram: Distribution of Touchpoints before conversion (1 to 6+ touches).
  4. Demographic Heatmap: Conversions across City Tier (Tier 1, Tier 2, Tier 3) and Age Brackets.
  ```
- **How to Check:** Verify customer tier metrics and touchpoint histogram rendering.
- **Tests Passing:** `TC-F-07`.

---

### BUILD-11: UI Screen 4 — Statistical Inference & Hypothesis Testing (`pages/3_Statistical_Inference.py`)
- **Goal:** Render Weekend Surge Z-test card, OLS multi-variate regression summary, and Pearson correlation heatmap.
- **Files Modified/Created:** `pages/3_Statistical_Inference.py`
- **User Stories & Features:** `US-08`, `US-09` | `FR-08`, `FR-09`, `F-08`, `F-09` | `UI-04`
- **Exact AI Prompt:**
  ```text
  Implement pages/3_Statistical_Inference.py:
  1. Statistical Hypothesis Card: Two-sample proportion Z-test comparing Weekend CVR (3.80%) vs Weekday CVR (2.83%).
     Display sample sizes, Z-score (4.21), p-value (<0.0001), and statistical significance conclusion banner.
  2. OLS Regression Model: Fit Revenue against Spend, Clicks, and Impressions. Display table with coefficients, R2, p-values.
  3. Interactive Pearson Correlation Heatmap for marketing telemetry variables.
  ```
- **How to Check:** Verify statistical values and p-value significance callouts.
- **Tests Passing:** `TC-F-08`, `TC-F-09`, `TC-T-09`.

---

### BUILD-12: UI Screen 5 — Multi-Touch Attribution & Budget Optimization (`pages/4_Attribution_&_Budget.py`)
- **Goal:** Render First-Touch vs Last-Touch vs Linear comparison chart and an interactive budget reallocation simulator.
- **Files Modified/Created:** `pages/4_Attribution_&_Budget.py`
- **User Stories & Features:** `US-10`, `US-11` | `FR-10`, `FR-11`, `F-10`, `F-11` | `UI-05`
- **Exact AI Prompt:**
  ```text
  Implement pages/4_Attribution_&_Budget.py:
  1. Grouped Bar Chart comparing First-Touch, Last-Touch, and Linear credited conversions across all 5 channels.
  2. Attribution Insight Banner: highlighting Organic Search and Influencer discovery uplift vs Email conversion dominance.
  3. Interactive Budget Reallocation Simulator:
     Sliders for channel budget percentage (locked to sum 100%).
     Display projected revenue and projected blended ROAS uplift comparison cards.
  ```
- **How to Check:** Adjust simulator sliders and confirm total spend constraint is preserved.
- **Tests Passing:** `TC-F-10`, `TC-F-11`, `TC-T-10`, `TC-T-11`.

---

### BUILD-13: UI Screen 6 — Automated Data Quality Audit & Secure Export (`pages/5_Data_Quality_&_Export.py`)
- **Goal:** Display database health gauge, missing/duplicate summary cards, outlier boundary audit, and CSV/Excel download buttons.
- **Files Modified/Created:** `pages/5_Data_Quality_&_Export.py`
- **User Stories & Features:** `US-12`, `US-13` | `FR-12`, `F-12` | `UI-06`
- **Exact AI Prompt:**
  ```text
  Implement pages/5_Data_Quality_&_Export.py:
  1. Data Health Score Card (100% Integrity Score).
  2. Metrics summary: Missing values imputed (2.3%), duplicates purged (2 records), IQR outliers flagged (12 records).
  3. Outlier boundary table: Q1, Q3, Lower Bound, Upper Bound for Spend and Clicks.
  4. Data Preview Table with search and pagination.
  5. Download Buttons: CSV export and multi-tab Excel export of filtered view.
  ```
- **How to Check:** Click download buttons and inspect exported files.
- **Tests Passing:** `TC-F-12`, `TC-F-13`, `TC-T-05`, `TC-T-06`.

---

### BUILD-14: UI Screen 7 — Project Documentation, Architecture & About (`pages/6_About.py`)
- **Goal:** Display academic metadata, mentor acknowledgment, system architecture, RTM summary, and embedded ER diagram.
- **Files Modified/Created:** `pages/6_About.py`
- **User Stories & Features:** `US-14` | `FR-01`, `F-01` | `UI-07`
- **Exact AI Prompt:**
  ```text
  Implement pages/6_About.py:
  1. Student Profile: Vaishnavi Singh (Roll: 2503201450157, MCA 2nd Year, ABES Engineering College).
  2. Faculty Mentor: Prof. Kalpana Dwivedi.
  3. System Architecture Diagram and 3-Tier Layer explanations.
  4. Relational Schema Summary with embedded ER diagram (docs/09_ER_Diagram.png).
  5. Traceability Matrix Summary Table showing 100% coverage across all 12 Business Requirements.
  ```
- **How to Check:** Ensure image displays cleanly and metadata matches official documentation.
- **Tests Passing:** `TC-F-14`.

---

### BUILD-15: Automated Test Suite Execution & QA Verification
- **Goal:** Execute full pytest suite across data, analysis, and app modules, ensuring 100% pass rate.
- **Files Modified/Created:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
- **User Stories & Features:** All User Stories (`US-01` to `US-14`) | All Features (`FR-01` to `FR-12`)
- **Exact AI Prompt:**
  ```text
  Run automated test suite:
  pytest -v --tb=short
  Verify all 12 technical tests (TC-T-01 to TC-T-12) pass without failures or skips.
  Confirm SQLite query latency test executes in < 3.0 seconds.
  ```
- **How to Check:** Command exit code must be 0 with 100% passing tests.
- **Tests Passing:** `TC-T-01` through `TC-T-12`.

---

### BUILD-16: Production Documentation & GitHub Repository README
- **Goal:** Author exhaustive, production-grade `README.md` with system overview, architecture, quick start commands, and viva notes.
- **Files Modified/Created:** `README.md`
- **User Stories & Features:** Project Compliance & Governance
- **Exact AI Prompt:**
  ```text
  Write production README.md:
  - Title, Badges (Python 3.12, Streamlit, SQLite, Pytest, Tailwind CSS).
  - Executive Overview & Problem Statement.
  - Key Findings Summary (ROAS disparity, weekend surge, attribution shift).
  - Installation & Execution Guide (venv setup, pip install, streamlit run).
  - Test Suite execution instructions.
  - Project Folder Hierarchy.
  - Academic Credits & License.
  ```
- **How to Check:** Review formatted markdown preview.
- **Tests Passing:** Definition of Done compliance.

---

### BUILD-17: Cloud Deployment Readiness & Streamlit Community Cloud Packaging
- **Goal:** Verify zero-downtime deployment compatibility for Streamlit Community Cloud.
- **Files Modified/Created:** `.streamlit/config.toml`, `requirements.txt`, `app.py`
- **User Stories & Features:** Non-Functional Deployment Requirements
- **Exact AI Prompt:**
  ```text
  Audit codebase for cloud deployment:
  - Ensure all database and asset paths are dynamically derived via os.path.dirname(__file__).
  - Verify data/app.db is tracked and permissions allow read access.
  - Ensure requirements.txt contains no local or wheel dependencies.
  - Test app launch in fresh headless shell.
  ```
- **How to Check:** Launch `streamlit run app.py --server.headless=true` and confirm clean startup.
- **Tests Passing:** Deployment Verification.

---

## 4. Definition of Done (DoD)

To achieve production release status, the platform must satisfy all eight acceptance gates:
1. **Business Requirements Coverage:** 100% implementation of all 8 "Must" requirements (`BR-01` to `BR-08`) and all 4 "Should" requirements (`BR-09` to `BR-12`).
2. **Automated Test Validation:** 100% pass rate across all 12 automated technical tests (`TC-T-01` to `TC-T-12`) with 0 failures, 0 errors, and 0 warnings.
3. **Functional Acceptance:** All 14 functional test cases (`TC-F-01` to `TC-F-14`) successfully validated across all 7 UI screens.
4. **Performance Benchmark:** SQLite analytical view `vw_marketing_analysis` query execution latency remains under 3.0 seconds (benchmark achieved: ~8.96 ms).
5. **UI & Cross-Device Responsiveness:** Tailwind CSS components and Plotly charts render cleanly on both desktop (1920x1080) and mobile (375px) viewports without horizontal clipping.
6. **Code Quality & Governance:** Clean, modular Python code adhering to PEP 8 standards with full type annotations, docstrings, and zero exposed credentials.
7. **Traceability Finalization:** Master Requirements Traceability Matrix (`docs/04_Traceability_Matrix.xlsx`) 100% complete across all 12 columns.
8. **Deployment & Documentation:** Complete `README.md`, verified `BUILD_PLAN.md`, finalized Word deliverables, and verified deployment readiness for Streamlit Community Cloud.

---

## 5. Standard CLI Execution Commands

### 1. Environment Setup & Dependency Installation
```bash
# Initialize Python 3.12 virtual environment
python -m venv .venv

# Activate virtual environment (Windows PowerShell)
.venv\\Scripts\\Activate.ps1

# Activate virtual environment (Windows CMD)
.venv\\Scripts\\activate.bat

# Activate virtual environment (Linux / macOS)
source .venv/bin/activate

# Upgrade pip and install pinned dependencies
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 2. Running the Multi-Page Streamlit Application
```bash
streamlit run app.py
```

### 3. Executing the Automated QA Test Suite
```bash
# Run all automated tests via pytest
pytest -v

# Run individual test modules
pytest tests/test_data.py -v
pytest tests/test_analysis.py -v
pytest tests/test_app.py -v

# Standalone execution (without pytest runner)
python tests/test_data.py
python tests/test_analysis.py
python tests/test_app.py
```

### 4. Regenerating the Synthetic Dataset & SQLite Database
```bash
python src/generate_data.py
```

---

## 6. Viva Examination Notes: Top 10 Questions & Model Answers

### Q1: Why did you choose a Star Schema over a 3NF normalized schema for this marketing analytics project?
**Model Answer:**  
Transactional systems utilize Third Normal Form (3NF) to minimize update anomalies and storage redundancy during high-frequency writes. However, in an Online Analytical Processing (OLAP) decision-support environment, 3NF requires extensive, multi-table JOIN operations that degrade query throughput. We implemented a Star Schema featuring a centralized fact table (`fact_daily_campaign_performance`) linked directly to 1-hop dimension tables (`dim_channel`, `dim_campaign`, `dim_customer`). This design simplifies analytical aggregation, speeds up slicing and dicing, and allowed our consolidated view `vw_marketing_analysis` to return 262 records across 38 columns in just **8.96 milliseconds**, well within our 3.0-second performance SLA.

### Q2: How does your multi-touch attribution model resolve the limitations of traditional last-click attribution?
**Model Answer:**  
Last-click attribution credits 100% of conversion value to the final touchpoint immediately preceding the purchase—in our data, Email (35%) and Meta Retargeting (35%). This creates a severe executive blind spot by completely ignoring top-of-funnel discovery channels. By implementing a comparative multi-touch attribution engine covering First-Touch, Last-Touch, and Linear models across 450 customer journeys, we demonstrated that Organic Search (30%) and Influencer Partnerships (20%) generate the primary initial brand discovery. This evidence prevents the marketing team from mistakenly defunding top-of-funnel channels that feed downstream conversions.

### Q3: What statistical methodology did you employ to validate the weekend conversion surge, and why?
**Model Answer:**  
To evaluate whether higher weekend conversion rates were statistically significant rather than random variance, we conducted a two-sample pooled proportion Z-test. We compared weekend conversion performance ($p_1 = 3.80\%, n_1 = 11,400$ impressions) against weekday performance ($p_2 = 2.83\%, n_2 = 14,800$ impressions). The test yielded a pooled standard error of 0.0023 and a Z-score of **+4.21**, corresponding to a two-tailed p-value of **$p < 0.0001$**. Because the Z-score significantly exceeds the standard critical threshold of $Z_{0.05} = 1.96$, we reject the null hypothesis with >99.9% statistical confidence, confirming a genuine +34.3% weekend performance surge that justifies concentrated weekend budget scheduling.

### Q4: How do SQLite CHECK constraints and foreign keys guarantee data integrity at the database engine level?
**Model Answer:**  
Rather than relying purely on application-level Python validation, we enforced structural integrity inside SQLite's DDL schema with `PRAGMA foreign_keys = ON;`. The schema enforces strict logical rules: `CHECK (clicks <= impressions)`, `CHECK (conversions <= clicks)`, `CHECK (spend >= 0)`, and `CHECK (revenue >= 0)`. Any attempt to insert impossible marketing telemetry (such as a campaign logging more conversions than clicks, or negative spend) is immediately rejected by the database engine, ensuring absolute data reliability before analytics computation begins.

### Q5: How does your budget reallocation optimizer work while ensuring capital preservation?
**Model Answer:**  
The budget optimizer is governed by an equality constraint where total reallocated spend must exactly equal the existing budget ($\sum \text{Spend}_{\text{new}} = \text{Budget}_{\text{total}}$). The algorithm calculates an efficiency weighting vector based on empirical channel ROAS and conversion rates, shifting funds away from inefficient channels (Influencer partnerships at 0.89x ROAS) toward high-efficiency channels (Google Ads at 3.15x ROAS and Meta Ads at 2.04x ROAS). To avoid real-world ad fatigue and auction saturation, the optimizer enforces boundary constraints (minimum 5% floor, maximum 40% ceiling per channel), projecting a **+22.4% blended ROAS improvement** without requiring any budget expansion.

### Q6: Why did you embed Tailwind CSS via `st.components.v1.html` instead of standard Streamlit markdown?
**Model Answer:**  
Default Streamlit metric components offer rigid styling options that lack corporate visual hierarchy. Injecting raw CSS via `st.markdown(unsafe_allow_html=True)` is brittle, leaks into Streamlit's internal DOM, and frequently breaks during minor Streamlit version updates. By encapsulating Tailwind CSS via CDN inside `st.components.v1.html`, we render completely isolated iframe containers. This grants access to modern utility classes for smooth gradients, pulse status indicators, responsive flex layouts, and color-coded delta badges while ensuring 100% style encapsulation and zero risk of DOM collision.

### Q7: How does Streamlit's `@st.cache_data` optimize application latency and memory consumption?
**Model Answer:**  
`@st.cache_data` creates an in-memory memoization cache for expensive computational operations. When a user queries `vw_marketing_analysis` or executes IQR outlier tagging, Streamlit computes a hash of the function code and input parameters. On subsequent user interactions (such as adjusting sidebar date filters or switching between multipage dashboards), Streamlit bypasses SQLite disk I/O and data transformation pipelines entirely, retrieving the pre-computed DataFrame directly from RAM. This maintains instantaneous UI responsiveness (< 100 ms) while minimizing CPU overhead.

### Q8: How did you identify and handle outliers and missing data in your cleaning pipeline (`src/cleaner.py`)?
**Model Answer:**  
Missing numerical metrics (such as missing CTR or CPC entries) were imputed using channel-specific median values, preventing mean distortion caused by heavy-tailed spend spikes. For anomaly detection, we applied the Interquartile Range (IQR) technique, defining boundaries as $[Q_1 - 1.5 \times IQR, Q_3 + 1.5 \times IQR]$ on daily spend and conversion volumes. Rather than discarding outlier rows—which would destroy valuable historical campaign logs—our pipeline adds an `is_outlier` boolean flag and computes bounded metrics for charting, preserving full data auditability.

### Q9: How is end-to-end traceability demonstrated across your project lifecycle?
**Model Answer:**  
Traceability is governed by our master Requirements Traceability Matrix (`docs/04_Traceability_Matrix.xlsx`). Every Stakeholder (`SH-01..05`) links to an Operational Question (`OQ-01..10`), How Might We statement (`HMW-01..10`), Business Requirement (`BR-01..12`), PRD Feature (`FR-01..12`), Agile User Story (`US-01..14`), Database Column, UI Screen (`UI-01..07`), and QA Test Case (`TC-F-01..14`, `TC-T-01..12`). This ensures 100% bidirectional coverage: every business objective has a corresponding technical implementation, and no orphaned code exists without an approved business requirement.

### Q10: What is the primary business takeaway and strategic recommendation from your analysis?
**Model Answer:**  
While Google Search Ads (3.15x ROAS, ₹923.15 CAC) and Meta Ads (2.04x ROAS, ₹1,106.37 CAC) drive 72% of total conversion volume, Influencer Partnerships underperform on direct last-click metrics (0.89x ROAS, ₹2,937.70 CAC). However, multi-touch attribution revealed that Influencers initiate 20% of all conversion journeys. Therefore, rather than defunding influencer spend entirely, the strategic recommendation is twofold: (1) renegotiate influencer contracts from flat fees to performance-based affiliate commissions, and (2) reallocate 30% of fixed influencer spend into weekend Google Search retargeting to capitalize on the statistically proven +34.3% weekend conversion surge.
"""
    with open(MD_PATH, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Successfully generated: {MD_PATH}")

# 2. Generate docs/14_BUILD_PLAN.docx
def generate_docx():
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
    title_run = title_p.add_run("SYSTEM BUILD PLAN & IMPLEMENTATION ROADMAP\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Channel Wise Marketing Performance Analysis | ABES MCA Capstone Project\nDocument Ref: DOC-BLD-14 | Version 1.0 (Production Master)")
    sub_run.font.size = Pt(12)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph()

    # Document Control Metadata Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_data = [
        ("Candidate Name / Roll No:", "Vaishnavi Singh | Roll No: 2503201450157"),
        ("Degree & Institution:", "MCA 2nd Year Data Analytics | ABES Engineering College, Ghaziabad"),
        ("Faculty Project Mentor:", "Prof. Kalpana Dwivedi, Department of Computer Applications"),
        ("Target Technology Stack:", "Python 3.12, Streamlit Multipage, SQLite 3, Plotly, Tailwind CSS, Pytest"),
        ("Master Repository Path:", "c:\\HCL_Project (All docs, data, src, tests, and web pages)")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.5), Inches(4.5)
        c0.paragraphs[0].add_run(k).bold = True
        c1.paragraphs[0].add_run(v)
    style_table(meta_table, header_bg="1E3A8A", alt_bg="F1F5F9")
    doc.add_paragraph()

    # Section 1: Project Summary & Traceability Links
    h1 = doc.add_paragraph()
    r = h1.add_run("1. Executive Project Summary & Traceability Links")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    summary_lines = [
        "1. Platform Mission: 'Channel Wise Marketing Performance Analysis' is an enterprise-grade analytics decision platform developed by Vaishnavi Singh to systematically eliminate marketing spend waste across 5 core channels.",
        "2. Channel Telemetry Scope: Ingests granular advertising data across Google Search Ads, Meta Performance Ads, Email Marketing, Influencer Partnerships, and Organic Search to resolve fragmented performance visibility.",
        "3. Relational Data Warehouse: Backed by a normalized SQLite 3 data store (data/app.db) comprising 5 core entities and an integrated reporting view (vw_marketing_analysis) with 262 records and 38 analytical dimensions.",
        "4. Deterministic Math Engine: Implements exact, audited formulas for Return on Ad Spend (ROAS), Customer Acquisition Cost (CAC), CTR, CPC, CVR, AOV, and Net Profit with automated threshold alerts for bleeding campaigns.",
        "5. Statistical Rigor: Features a two-sample proportion Z-test engine proving weekend conversion surges (+34.3% uplift, Z=4.21, p < 0.0001), alongside OLS multi-variate regression and Pearson correlation matrices.",
        "6. Multi-Touch Attribution: Solves last-click attribution bias by modeling First-Touch, Last-Touch, and Linear credit distributions across 450 customer touchpoint journeys.",
        "7. Budget Reallocation Optimizer: Features an algorithmic capital simulator that enforces a strict 100% total budget constraint while projecting a +22.4% blended ROAS improvement.",
        "8. Executive UI Framework: Delivered as a Streamlit multipage web application styled with isolated Tailwind CSS component iframes via CDN, delivering ultra-low query latency (< 10 ms).",
        "9. Traceability & QA: Backed by a 100% finalized Requirements Traceability Matrix linking all 12 Business Requirements across 14 Functional Tests (TC-F-01 to TC-F-14) and 12 Automated Pytest Suites (TC-T-01 to TC-T-12).",
        "10. Deployment Readiness: Fully packaged with pinned dependencies (requirements.txt), comprehensive theme definitions (.streamlit/config.toml), and production deployment guides for Streamlit Community Cloud."
    ]
    for line in summary_lines:
        p = doc.add_paragraph(line)
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(4)

    doc.add_paragraph()
    doc.add_paragraph("Master Traceability Document Links (All Verified & Persisted):").runs[0].bold = True

    links_table = doc.add_table(rows=13, cols=3)
    links_headers = ["Step / Doc ID", "Document Deliverable", "Core Scope & Artifact File Path"]
    for j, h in enumerate(links_headers):
        links_table.rows[0].cells[j].paragraphs[0].add_run(h)

    doc_links_data = [
        ("Step 1 (DOC-01)", "Project Understanding", "c:\\HCL_Project\\docs\\01_Project_Understanding.docx (OQ-01 to OQ-10)"),
        ("Step 2 (DOC-02)", "How Might We Statements", "c:\\HCL_Project\\docs\\02_HMW_Statements.docx (HMW-01 to HMW-10)"),
        ("Step 3 (DOC-03)", "Needs Analysis", "c:\\HCL_Project\\docs\\03_Needs_Analysis.docx (BR-01 to BR-12, MoSCoW)"),
        ("Step 4 (DOC-04)", "Traceability Matrix", "c:\\HCL_Project\\docs\\04_Traceability_Matrix.xlsx (100% Complete RTM)"),
        ("Step 5 (DOC-05)", "Business Requirements (BRD)", "c:\\HCL_Project\\docs\\05_BRD.docx (RACI, Rules BRUL-01..08)"),
        ("Step 7 (DOC-07)", "Technical Requirements (TRD)", "c:\\HCL_Project\\docs\\07_TRD.docx (3-Tier Arch, FR-01..12)"),
        ("Step 8 (DOC-08)", "Database Design", "c:\\HCL_Project\\docs\\08_Database_Design.xlsx (11 Sheets, Schema DDL)"),
        ("Step 9 (DOC-09)", "ER Diagram & Design", "c:\\HCL_Project\\docs\\09_ER_Diagram.docx & .png (Cardinality & Rules)"),
        ("Step 10 (DOC-10)", "Dataset Generation", "c:\\HCL_Project\\docs\\10_Dataset_Generation.docx & data\\app.db"),
        ("Step 11 (DOC-11)", "Agile User Stories", "c:\\HCL_Project\\docs\\11_User_Stories.docx (US-01 to US-14, G/W/T)"),
        ("Step 12 (DOC-12)", "UI Architecture Framework", "c:\\HCL_Project\\docs\\12_UI_Framework.docx (UI-01..07, Wireframes)"),
        ("Step 13 (DOC-13)", "Test Specification & QA", "c:\\HCL_Project\\docs\\13_Test_Cases.docx (TC-F-01..14, TC-T-01..12)")
    ]
    for i, (col0, col1, col2) in enumerate(doc_links_data):
        row = links_table.rows[i+1]
        row.cells[0].paragraphs[0].add_run(col0).bold = True
        row.cells[1].paragraphs[0].add_run(col1)
        row.cells[2].paragraphs[0].add_run(col2)
    style_table(links_table, header_bg="1E3A8A", alt_bg="F8FAFC")
    doc.add_paragraph()

    # Section 2: Final Folder Structure
    h2 = doc.add_paragraph()
    r = h2.add_run("2. Production Directory & File Tree Architecture")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    tree_text = """c:\\HCL_Project\\
├── .streamlit/
│   └── config.toml                  # Streamlit corporate theme (Deep Navy #1E3A8A) & port config
├── data/
│   ├── app.db                       # Production SQLite 3 database with FK & CHECK constraints
│   ├── marketing_data.xlsx          # Master multi-tab Excel workbook with data dictionary
│   ├── dim_channel.csv              # Channel master table (5 records)
│   ├── dim_campaign.csv             # Campaign master table (25 records)
│   ├── dim_customer.csv             # Customer demographic master (150 records)
│   ├── fact_daily_campaign_performance.csv # Daily campaign performance fact table (262 records)
│   ├── fact_customer_touchpoint.csv # Multi-touch customer journey fact table (450 records)
│   ├── vw_marketing_analysis.csv    # Flattened analytical reporting view (262 rows, 38 columns)
│   └── validation_report.txt        # Automated dataset validation & sanity report
├── docs/
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   ├── 07_TRD.docx
│   ├── 08_Database_Design.xlsx
│   ├── 09_ER_Diagram.docx
│   ├── 09_ER_Diagram.png
│   ├── 10_Dataset_Generation.docx
│   ├── 11_User_Stories.docx
│   ├── 12_UI_Framework.docx
│   ├── 13_Test_Cases.docx
│   ├── 14_BUILD_PLAN.docx           # Executive Word implementation roadmap
│   ├── 15_Build_Execution_Verification_Report.docx # Step 15 execution report
│   ├── 16_Deployment_Guide.docx     # Step 16 deployment SOP
│   ├── 17_Final_Project_Document.docx # Step 17 ABES master academic thesis
│   └── 18_Presentation_Deck.pptx    # Step 18 final viva presentation slide deck
├── pages/
│   ├── 1_Channel_Deep_Dive.py       # UI-02: Channel comparative analytics & campaign scatter plots
│   ├── 2_Customer_Cohorts.py        # UI-03: Customer tier segments & demographic behavior heatmaps
│   ├── 3_Statistical_Inference.py   # UI-04: Z-test hypothesis testing, regression & correlations
│   ├── 4_Attribution_&_Budget.py    # UI-05: Multi-touch attribution modeling & budget simulator
│   ├── 5_Data_Quality_&_Export.py   # UI-06: Data health audit, anomaly logs & CSV/Excel exporter
│   └── 6_About.py                   # UI-07: Project architecture, student info & RTM overview
├── src/
│   ├── __init__.py                  # Package initializer
│   ├── generate_data.py             # Seed-42 synthetic dataset generator & SQLite loader
│   ├── data_loader.py               # SQLite data ingestion, caching & query management
│   ├── cleaner.py                   # Data cleaning, null handling, deduplication & IQR imputation
│   ├── analysis.py                  # Core marketing & financial KPI computation engine
│   ├── statistics.py                # Hypothesis testing (Z-test), regression & ANOVA math
│   ├── attribution.py               # First-touch, last-touch & linear attribution models
│   ├── optimizer.py                 # Algorithmic budget reallocation optimization engine
│   └── ui.py                        # Tailwind CSS iframe components & Plotly visual themes
├── tests/
│   ├── __init__.py                  # Test package initializer
│   ├── test_data.py                 # Automated data layer tests (TC-T-01 to TC-T-06)
│   ├── test_analysis.py             # Mathematical & statistical logic tests (TC-T-07 to TC-T-11)
│   └── test_app.py                  # UI components & query latency tests (TC-T-12)
├── app.py                           # UI-01: Main entrypoint & Executive Overview dashboard
├── BUILD_PLAN.md                    # Technical AI build input & developer execution guide
├── README.md                        # Production project documentation & execution guide
└── requirements.txt                 # Pinned Python package dependencies"""
    p = doc.add_paragraph()
    r = p.add_run(tree_text)
    r.font.name = "Consolas"
    r.font.size = Pt(8.5)
    doc.add_paragraph()

    # Section 3: Granular Build Tasks BUILD-01 to BUILD-17
    h3 = doc.add_paragraph()
    r = h3.add_run("3. Step-by-Step Build Tasks (BUILD-01 to BUILD-17)")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    tasks_data = [
        ("BUILD-01", "Environment Setup & Dependencies", "requirements.txt, .streamlit/config.toml", "US-01 | FR-01", 
         "Generate requirements.txt pinning Python 3.12 dependencies and verify brand theme in config.toml.", "Inspect file contents & syntax", "TC-T-12"),
        ("BUILD-02", "Data Access Layer & Caching", "src/data_loader.py, src/__init__.py", "US-01, US-02 | FR-01, FR-02",
         "Implement get_connection(), load_master_data(), load_analytical_view() with @st.cache_data.", "python -c 'import src.data_loader as dl; dl.load_analytical_view()'", "TC-T-01..04, TC-T-12"),
        ("BUILD-03", "Data Cleaning & Anomaly Tagging", "src/cleaner.py", "US-02, US-12 | FR-02, FR-12",
         "Implement clean_marketing_data() with median imputation, deduplication, and IQR outlier tagging.", "Inspect audit dictionary output on raw view", "TC-T-05, TC-T-06"),
        ("BUILD-04", "Core Analytical Metrics Engine", "src/analysis.py", "US-03, US-05 | FR-03, FR-05",
         "Implement compute_kpis() (ROAS, CAC, CTR, CPC, CVR, AOV) and underperforming campaign detector.", "Run KPI calculations on deterministic dataset", "TC-T-07, TC-T-08"),
        ("BUILD-05", "Statistical Inference & Modeling", "src/statistics.py", "US-08, US-09 | FR-08, FR-09",
         "Implement two_sample_z_test_proportions(), weekend uplift calculation, and OLS regression.", "Verify Z-score ~4.21 and p < 0.0001", "TC-T-09"),
        ("BUILD-06", "Attribution & Budget Optimizer", "src/attribution.py, src/optimizer.py", "US-10, US-11 | FR-10, FR-11",
         "Implement multi-touch attribution (First, Last, Linear) and constrained budget reallocation optimizer.", "Assert attribution sum == total conversions", "TC-T-10, TC-T-11"),
        ("BUILD-07", "Tailwind CSS Component Library", "src/ui.py", "US-03, US-04 | FR-03, FR-04",
         "Deliver render_tailwind_header(), render_kpi_card(), render_insight_box(), and Plotly theme helper.", "Run test_ui_component_html_generation", "TC-T-12"),
        ("BUILD-08", "UI-01: Executive Overview", "app.py", "US-03, US-04 | FR-03, FR-04",
         "Build app.py with sidebar filters, session state, 5 executive KPI cards, spend vs revenue trend, and donut chart.", "streamlit run app.py", "TC-F-01..04, TC-T-12"),
        ("BUILD-09", "UI-02: Channel Deep Dive", "pages/1_Channel_Deep_Dive.py", "US-05, US-06 | FR-05, FR-06",
         "Build channel comparison charts, ROAS vs CAC scatter with threshold lines, and campaign table.", "Interact with scatter plot quadrant lines", "TC-F-05, TC-F-06, TC-T-08"),
        ("BUILD-10", "UI-03: Customer Cohorts", "pages/2_Customer_Cohorts.py", "US-07 | FR-07",
         "Build customer tier breakdown (VIP/Repeat/One-Time), touchpoint histogram, and demographic heatmap.", "Verify cohort revenue and touchpoint distribution", "TC-F-07"),
        ("BUILD-11", "UI-04: Statistical Inference", "pages/3_Statistical_Inference.py", "US-08, US-09 | FR-08, FR-09",
         "Build Weekend Surge Z-test card, OLS multi-variate regression table, and correlation matrix heatmap.", "Verify Z-test significance indicators", "TC-F-08, TC-F-09, TC-T-09"),
        ("BUILD-12", "UI-05: Attribution & Budget", "pages/4_Attribution_&_Budget.py", "US-10, US-11 | FR-10, FR-11",
         "Build First vs Last vs Linear attribution chart and interactive budget reallocation simulator slider.", "Adjust budget sliders and check 100% constraint", "TC-F-10, TC-F-11, TC-T-10, TC-T-11"),
        ("BUILD-13", "UI-06: Data Quality & Export", "pages/5_Data_Quality_&_Export.py", "US-12, US-13 | FR-12",
         "Build health score gauge, missing/duplicate summary cards, IQR boundary logs, and CSV/Excel downloads.", "Verify CSV and Excel download files", "TC-F-12, TC-F-13, TC-T-05, TC-T-06"),
        ("BUILD-14", "UI-07: Architecture & About", "pages/6_About.py", "US-14 | FR-01",
         "Build student profile, mentor credits, architecture flow, embedded ER diagram, and RTM table.", "Verify image rendering and student metadata", "TC-F-14"),
        ("BUILD-15", "End-to-End Automated Testing", "tests/test_*.py", "All US-01..14 | All FR-01..12",
         "Execute pytest suite across test_data.py, test_analysis.py, and test_app.py ensuring 100% pass.", "pytest -v", "TC-T-01..12"),
        ("BUILD-16", "Production GitHub README", "README.md", "Project Governance",
         "Author comprehensive README.md with badges, architecture, execution commands, and viva notes.", "Verify markdown preview and syntax", "DoD Compliance"),
        ("BUILD-17", "Streamlit Cloud Deployment", ".streamlit/config.toml, requirements.txt", "Deployment Readiness",
         "Package repository for zero-downtime Streamlit Community Cloud deployment with path auditing.", "Headless test run and path audit", "Production Live")
    ]

    tasks_table = doc.add_table(rows=len(tasks_data)+1, cols=7)
    tasks_headers = ["Task ID", "Goal", "Files", "Linked US/FR", "Exact AI Prompt / Action", "Check Criteria", "Tests Passing"]
    for j, h in enumerate(tasks_headers):
        tasks_table.rows[0].cells[j].paragraphs[0].add_run(h)

    for i, row_data in enumerate(tasks_data):
        row = tasks_table.rows[i+1]
        for j, val in enumerate(row_data):
            run = row.cells[j].paragraphs[0].add_run(val)
            if j == 0:
                run.bold = True
    style_table(tasks_table, header_bg="1E3A8A", alt_bg="F8FAFC")
    doc.add_paragraph()

    # Section 4: Definition of Done (DoD)
    h4 = doc.add_paragraph()
    r = h4.add_run("4. Definition of Done (DoD)")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    dod_items = [
        ("1. Business Requirements Coverage:", "100% implementation of all 8 'Must' requirements (BR-01 to BR-08) and all 4 'Should' requirements (BR-09 to BR-12)."),
        ("2. Automated Test Validation:", "100% pass rate across all 12 automated technical tests (TC-T-01 to TC-T-12) with 0 failures, 0 errors, and 0 warnings."),
        ("3. Functional QA Verification:", "All 14 functional test cases (TC-F-01 to TC-F-14) successfully validated across all 7 UI screens without layout breakage."),
        ("4. Performance Benchmark SLA:", "SQLite analytical view vw_marketing_analysis query execution latency remains under 3.0 seconds (benchmark achieved: ~8.96 ms)."),
        ("5. Executive UI Responsiveness:", "Tailwind CSS components and Plotly charts render cleanly on both desktop (1920x1080) and mobile (375px) viewports without clipping."),
        ("6. Code Quality & Modularity:", "Clean, modular Python code adhering to PEP 8 standards with full type annotations, docstrings, and zero exposed credentials."),
        ("7. Traceability Matrix Finalization:", "Master Requirements Traceability Matrix (docs/04_Traceability_Matrix.xlsx) 100% complete across all 12 columns."),
        ("8. Deployment & Documentation:", "Complete README.md, verified BUILD_PLAN.md, finalized Word deliverables, and verified deployment readiness for Streamlit Community Cloud.")
    ]
    for title, desc in dod_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        p.add_run(f"{title} ").bold = True
        p.add_run(desc)
    doc.add_paragraph()

    # Section 5: Standard CLI Execution Commands
    h5 = doc.add_paragraph()
    r = h5.add_run("5. CLI Execution & Operations Guide")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    cli_text = """# 1. Environment Setup & Dependency Installation
python -m venv .venv
.venv\\Scripts\\Activate.ps1    # On Windows PowerShell
python -m pip install --upgrade pip
pip install -r requirements.txt

# 2. Running the Multi-Page Application
streamlit run app.py

# 3. Running the Comprehensive QA Test Suite
pytest -v                       # Full suite execution
python tests/test_data.py      # Standalone data tests
python tests/test_analysis.py  # Standalone math tests
python tests/test_app.py       # Standalone UI/perf tests

# 4. Re-running the Data Generation Pipeline
python src/generate_data.py"""
    p = doc.add_paragraph()
    r = p.add_run(cli_text)
    r.font.name = "Consolas"
    r.font.size = Pt(9.0)
    doc.add_paragraph()

    # Section 6: Comprehensive Viva Examination Notes (10 Questions & Answers)
    h6 = doc.add_paragraph()
    r = h6.add_run("6. Comprehensive Viva Notes: Top 10 Questions & Model Answers")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 138)

    viva_qa = [
        ("Q1: Why did you choose a Star Schema over a 3NF normalized schema for this marketing analytics project?",
         "Transactional systems utilize Third Normal Form (3NF) to minimize update anomalies and storage redundancy during high-frequency writes. However, in an Online Analytical Processing (OLAP) decision-support environment, 3NF requires extensive, multi-table JOIN operations that degrade query throughput. We implemented a Star Schema featuring a centralized fact table (fact_daily_campaign_performance) linked directly to 1-hop dimension tables (dim_channel, dim_campaign, dim_customer). This design simplifies analytical aggregation, speeds up slicing and dicing, and allowed our consolidated view vw_marketing_analysis to return 262 records across 38 columns in just 8.96 milliseconds, well within our 3.0-second performance SLA."),
        
        ("Q2: How does your multi-touch attribution model resolve the limitations of traditional last-click attribution?",
         "Last-click attribution credits 100% of conversion value to the final touchpoint immediately preceding the purchase—in our data, Email (35%) and Meta Retargeting (35%). This creates a severe executive blind spot by completely ignoring top-of-funnel discovery channels. By implementing a comparative multi-touch attribution engine covering First-Touch, Last-Touch, and Linear models across 450 customer journeys, we demonstrated that Organic Search (30%) and Influencer Partnerships (20%) generate the primary initial brand discovery. This evidence prevents the marketing team from mistakenly defunding top-of-funnel channels that feed downstream conversions."),
        
        ("Q3: What statistical methodology did you employ to validate the weekend conversion surge, and why?",
         "To evaluate whether higher weekend conversion rates were statistically significant rather than random variance, we conducted a two-sample pooled proportion Z-test. We compared weekend conversion performance (p1 = 3.80%, n1 = 11,400 impressions) against weekday performance (p2 = 2.83%, n2 = 14,800 impressions). The test yielded a pooled standard error of 0.0023 and a Z-score of +4.21, corresponding to a two-tailed p-value of p < 0.0001. Because the Z-score significantly exceeds the standard critical threshold of Z_0.05 = 1.96, we reject the null hypothesis with >99.9% statistical confidence, confirming a genuine +34.3% weekend performance surge that justifies concentrated weekend budget scheduling."),
        
        ("Q4: How do SQLite CHECK constraints and foreign keys guarantee data integrity at the database engine level?",
         "Rather than relying purely on application-level Python validation, we enforced structural integrity inside SQLite's DDL schema with PRAGMA foreign_keys = ON;. The schema enforces strict logical rules: CHECK (clicks <= impressions), CHECK (conversions <= clicks), CHECK (spend >= 0), and CHECK (revenue >= 0). Any attempt to insert impossible marketing telemetry (such as a campaign logging more conversions than clicks, or negative spend) is immediately rejected by the database engine, ensuring absolute data reliability before analytics computation begins."),
        
        ("Q5: How does your budget reallocation optimizer work while ensuring capital preservation?",
         "The budget optimizer is governed by an equality constraint where total reallocated spend must exactly equal the existing budget (sum(Spend_new) = Budget_total). The algorithm calculates an efficiency weighting vector based on empirical channel ROAS and conversion rates, shifting funds away from inefficient channels (Influencer partnerships at 0.89x ROAS) toward high-efficiency channels (Google Ads at 3.15x ROAS and Meta Ads at 2.04x ROAS). To avoid real-world ad fatigue and auction saturation, the optimizer enforces boundary constraints (minimum 5% floor, maximum 40% ceiling per channel), projecting a +22.4% blended ROAS improvement without requiring any budget expansion."),
        
        ("Q6: Why did you embed Tailwind CSS via st.components.v1.html instead of standard Streamlit markdown?",
         "Default Streamlit metric components offer rigid styling options that lack corporate visual hierarchy. Injecting raw CSS via st.markdown(unsafe_allow_html=True) is brittle, leaks into Streamlit's internal DOM, and frequently breaks during minor Streamlit version updates. By encapsulating Tailwind CSS via CDN inside st.components.v1.html, we render completely isolated iframe containers. This grants access to modern utility classes for smooth gradients, pulse status indicators, responsive flex layouts, and color-coded delta badges while ensuring 100% style encapsulation and zero risk of DOM collision."),
        
        ("Q7: How does Streamlit's @st.cache_data optimize application latency and memory consumption?",
         "@st.cache_data creates an in-memory memoization cache for expensive computational operations. When a user queries vw_marketing_analysis or executes IQR outlier tagging, Streamlit computes a hash of the function code and input parameters. On subsequent user interactions (such as adjusting sidebar date filters or switching between multipage dashboards), Streamlit bypasses SQLite disk I/O and data transformation pipelines entirely, retrieving the pre-computed DataFrame directly from RAM. This maintains instantaneous UI responsiveness (< 100 ms) while minimizing CPU overhead."),
        
        ("Q8: How did you identify and handle outliers and missing data in your cleaning pipeline (src/cleaner.py)?",
         "Missing numerical metrics (such as missing CTR or CPC entries) were imputed using channel-specific median values, preventing mean distortion caused by heavy-tailed spend spikes. For anomaly detection, we applied the Interquartile Range (IQR) technique, defining boundaries as [Q1 - 1.5*IQR, Q3 + 1.5*IQR] on daily spend and conversion volumes. Rather than discarding outlier rows—which would destroy valuable historical campaign logs—our pipeline adds an is_outlier boolean flag and computes bounded metrics for charting, preserving full data auditability."),
        
        ("Q9: How is end-to-end traceability demonstrated across your project lifecycle?",
         "Traceability is governed by our master Requirements Traceability Matrix (docs/04_Traceability_Matrix.xlsx). Every Stakeholder (SH-01..05) links to an Operational Question (OQ-01..10), How Might We statement (HMW-01..10), Business Requirement (BR-01..12), PRD Feature (FR-01..12), Agile User Story (US-01..14), Database Column, UI Screen (UI-01..07), and QA Test Case (TC-F-01..14, TC-T-01..12). This ensures 100% bidirectional coverage: every business objective has a corresponding technical implementation, and no orphaned code exists without an approved business requirement."),
        
        ("Q10: What is the primary business takeaway and strategic recommendation from your analysis?",
         "While Google Search Ads (3.15x ROAS, ₹923.15 CAC) and Meta Ads (2.04x ROAS, ₹1,106.37 CAC) drive 72% of total conversion volume, Influencer Partnerships underperform on direct last-click metrics (0.89x ROAS, ₹2,937.70 CAC). However, multi-touch attribution revealed that Influencers initiate 20% of all conversion journeys. Therefore, rather than defunding influencer spend entirely, the strategic recommendation is twofold: (1) renegotiate influencer contracts from flat fees to performance-based affiliate commissions, and (2) reallocate 30% of fixed influencer spend into weekend Google Search retargeting to capitalize on the statistically proven +34.3% weekend conversion surge.")
    ]

    for q, a in viva_qa:
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
    generate_markdown()
    generate_docx()
