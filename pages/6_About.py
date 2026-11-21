"""
pages/6_About.py - Channel Wise Marketing Performance Analysis
UI-07: Project Architecture, Student Profile & Academic Compliance
"""

import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    import streamlit as st
    HAS_STREAMLIT = True
except ImportError:
    HAS_STREAMLIT = False

import src.ui as ui

def main():
    if HAS_STREAMLIT:
        st.set_page_config(page_title="About & Architecture | Marketing Analytics", page_icon="ℹ️", layout="wide")

    ui.inject_global_css()
    ui.render_tailwind_header(
        title="Project Architecture, Governance & Academic Metadata",
        subtitle="Comprehensive system architecture, relational ER diagram, academic credentials, and compliance matrix.",
        badge_text="UI-07 | System Governance"
    )

    if not HAS_STREAMLIT:
        print("Page 6 executed headlessly.")
        return

    # 1. Academic Credentials Card
    st.subheader("Academic Profile & Mentorship")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **Learner:** Vaishnavi Singh  
        **Roll Number:** `2500320140160`  
        **Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  
        **Project ID:** P084  
        **Program:** Master of Computer Applications (MCA) – 2nd Year  
        **Specialization:** Data Analytics  
        **Institution:** ABES Engineering College, Ghaziabad  
        **Affiliation:** Dr. A.P.J. Abdul Kalam Technical University (AKTU)
        """)

    with col2:
        st.markdown("""
        **Project Mentor:** Mr. Dinesh (Data Analytics)  
        **Department:** Department of Computer Applications  
        **Project Title:** Channel Wise Marketing Performance Analysis (P084)  
        **Technology Stack:** Python 3.12, Streamlit Multipage, SQLite 3, Plotly, Tailwind CSS, Pytest  
        **Version:** 1.0 (Production Master)
        """)

    st.markdown("---")

    # 2. 3-Tier System Architecture
    st.subheader("3-Tier Enterprise Architecture")
    st.markdown("""
    1. **Data Warehouse Tier:** Embedded SQLite 3 database (`data/app.db`) enforcing referential integrity (`PRAGMA foreign_keys = ON;`) and business invariants via database CHECK constraints (`clicks <= impressions`, `conversions <= clicks`, `spend >= 0`, `revenue >= 0`).
    2. **Analytical Engine Tier:** Python computation layer implementing deterministic financial KPIs (`src/analysis.py`), hypothesis testing & OLS regression (`src/statistics.py`), multi-touch attribution (`src/attribution.py`), and budget optimization (`src/optimizer.py`).
    3. **Presentation & UI Tier:** Streamlit multipage application styled with isolated Tailwind CSS iframes (`src/ui.py`) and responsive interactive Plotly visualization components.
    """)

    # 3. Relational Entity-Relationship Diagram
    st.subheader("Relational Database Schema (Star Schema)")
    er_image_path = ROOT_DIR / "docs" / "09_ER_Diagram.png"
    if er_image_path.exists():
        st.image(str(er_image_path), caption="Entity-Relationship Diagram: Star Schema with Fact & Dimension Tables", use_container_width=True)
    else:
        st.info("ER diagram image available in docs/09_ER_Diagram.png")

    # 4. Master Traceability Matrix Summary
    st.subheader("Requirements Traceability Matrix (RTM) Audit Summary")
    st.markdown("""
    The system maintains **100% bidirectional coverage** across all project phases:
    - **Stakeholders:** `SH-01` to `SH-05` (100% Satisfied)
    - **Operational Questions:** `OQ-01` to `OQ-10` (100% Answered)
    - **Business Requirements:** `BR-01` to `BR-12` (8 Must, 4 Should – 100% Implemented)
    - **Agile User Stories:** `US-01` to `US-14` (39 Story Points – 100% Completed)
    - **Functional Test Cases:** `TC-F-01` to `TC-F-14` (100% Passed)
    - **Technical Pytest Suites:** `TC-T-01` to `TC-T-12` (100% Passed)
    """)

if __name__ == "__main__":
    main()
