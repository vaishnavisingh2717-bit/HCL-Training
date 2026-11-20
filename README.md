# Channel Wise Marketing Performance Analysis
### Enterprise Decision Intelligence & Acquisition Performance Platform

![Python 3.12](https://img.shields.io/badge/Python-3.12-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-Multipage-FF4B4B.svg)
![SQLite 3](https://img.shields.io/badge/SQLite-Star%20Schema-003B57.svg)
![Tailwind CSS](https://img.shields.io/badge/Tailwind-CDN%20Iframe-38B2AC.svg)
![Pytest](https://img.shields.io/badge/QA%20Tests-12%2F12%20Passing-brightgreen.svg)
![RTM](https://img.shields.io/badge/RTM%20Coverage-100%25-success.svg)

---

## 1. Executive Summary & Problem Statement

Modern direct-to-consumer and B2B enterprises deploy marketing capital across multiple digital channels—including Google Search Ads, Meta Performance Ads, Email Marketing, Influencer Partnerships, and Organic Search. However, marketing teams routinely suffer from:
1. **Spend Inefficiency:** Inability to flag campaigns failing unit economic thresholds (e.g., ROAS < 2.0x or CAC > ₹1,500).
2. **Last-Click Attribution Bias:** Crediting 100% of conversion value to final touchpoints (Email, Meta) while mistakenly defunding top-of-funnel discovery catalysts (Organic Search, Influencers).
3. **Suboptimal Day-Parting:** Failing to capitalize on statistically significant weekend conversion surges (+34.3% uplift).

This platform delivers an end-to-end, reproducible analytics platform that unifies relational data warehousing, deterministic KPI computing, econometric regression, multi-touch attribution, and algorithmic budget reallocation into an executive Tailwind-styled Streamlit multipage application.

---

## 2. Key Analytical & Business Findings

- **Disparity in Channel ROAS:** Organic Search (106.91x ROAS) and Email Marketing (33.99x ROAS) deliver exceptionally high margins, while Google Ads (3.15x ROAS) and Meta Ads (2.04x ROAS) represent scalable growth engines. Influencer Partnerships underperform on direct last-click ROAS (0.89x) with a high CAC (₹2,937.70).
- **Statistically Significant Weekend Surge:** A two-sample pooled proportion Z-test confirms that weekend conversion rate ($p_1 = 3.65\%$) significantly outperforms weekday conversion rate ($p_2 = 3.12\%$) with $Z = +3.93$ and $p < 0.0001$.
- **Attribution Transformation:** Under Last-Touch modeling, Email (44) and Meta (40) dominate conversions. Under First-Touch modeling, Influencer Partnerships (56 conversions) and Organic Search (31 conversions) represent the primary awareness catalysts.
- **Budget Reallocation Impact:** Reallocating ₹340,000 away from fixed influencer fees into high-performing weekend Google Search and Meta retargeting projects an **+8.4% to +22.4% blended revenue uplift** under strict total budget preservation.

---

## 3. System Architecture & Relational Schema

The platform implements a decoupled 3-tier architecture:
- **Data Tier:** Embedded SQLite 3 database (`data/app.db`) enforcing referential integrity (`PRAGMA foreign_keys = ON;`) and database CHECK constraints (`clicks <= impressions`, `conversions <= clicks`, `spend >= 0`, `revenue >= 0`). Query latency for the denormalized reporting view (`vw_marketing_analysis`) is **8.96 ms** (< 3.0s SLA).
- **Compute Tier:** Pure Python analytical modules in `src/` providing data cleaning, IQR outlier bounds, KPI math, Z-tests, OLS regression, multi-touch attribution, and capital optimization.
- **Presentation Tier:** Streamlit multipage web interface styled with isolated Tailwind CSS components via CDN iframes and responsive Plotly visual charts.

### Repository File Architecture
```text
C:\HCL_Project\
├── .streamlit/                      # Corporate theme styling (Navy #1E3A8A) & port config
├── data/                            # SQLite warehouse (app.db), CSVs & validation logs
├── docs/                            # Formal capstone documents, RTM, BRD, TRD, thesis & PPTX
├── pages/                           # 6 Streamlit sub-dashboards (UI-02 to UI-07)
├── scripts/                         # Automated artifact generation & data scripts
├── src/                             # Core Python packages (analysis, attribution, optimizer, ui)
├── tests/                           # 12 automated pytest suites & performance benchmarks
├── .gitignore                       # Git exclusion rules (.venv, caches, secrets)
├── app.py                           # UI-01 Main executive dashboard entrypoint
├── BUILD_PLAN.md                    # 17-task step-by-step developer build roadmap
├── README.md                        # Master production documentation
└── requirements.txt                 # Pinned Python package dependencies
```

---

## 4. Quick Start & CLI Execution Guide

### 1. Environment Setup & Installation
```bash
# Clone repository
git clone https://github.com/<username>/2503201450157-channel-wise-marketing-performance.git
cd 2503201450157-channel-wise-marketing-performance

# Create and activate virtual environment
python -m venv .venv

# On Windows PowerShell:
.venv\Scripts\Activate.ps1
# On Linux / macOS:
source .venv/bin/activate

# Install pinned dependencies
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 2. Launch the Application
```bash
streamlit run app.py
```
Open your browser at `http://localhost:8501`.

### 3. Run the Automated QA Test Suite
```bash
# Run all unit and integration tests via pytest
pytest -v

# Standalone execution
python tests/test_data.py
python tests/test_analysis.py
python tests/test_app.py
```

### 4. Re-generate Synthetic Dataset
```bash
python src/generate_data.py
```

---

## 5. Deployment to Streamlit Community Cloud

The repository is pre-configured for zero-configuration deployment to [Streamlit Community Cloud](https://share.streamlit.io):
1. Push repository to GitHub.
2. Sign in to Streamlit Community Cloud and click **"Create app"**.
3. Select repository, branch (`main`), and main file path (`app.py`).
4. Set App URL: `2503201450157-channel-wise-marketing-performance`.
5. Click **"Deploy"**. The app builds and deploys within 2–3 minutes.

---

## 6. Project Credentials & Governance

- **Learner:** Vaishnavi Singh
- **Roll Number:** `2503201450157`
- **Program:** Master of Computer Applications (MCA) – 2nd Year, Data Analytics
- **Institution:** ABES Engineering College, Ghaziabad
- **Faculty Mentor:** Prof. Kalpana Dwivedi, Department of Computer Applications
- **Academic Year:** 2026–2027
