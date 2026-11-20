"""
tests/test_app.py - Automated Technical Tests for Application & UI Architecture
Covers: TC-T-12, Performance Benchmarks & Component Rendering
"""

import os
import time
import sqlite3
try:
    import pytest
except ImportError:
    pytest = None

DB_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data", "app.db"))
CONFIG_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".streamlit", "config.toml"))

def test_tc_t_12_query_latency_under_3_seconds():
    """TC-T-12: Verify that querying the SQLite analytical view executes in < 3.0 seconds."""
    assert os.path.exists(DB_PATH), f"Database not found at {DB_PATH}"
    
    start_time = time.time()
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    cur = conn.cursor()
    cur.execute("SELECT * FROM vw_marketing_analysis;")
    rows = cur.fetchall()
    elapsed_time = time.time() - start_time
    conn.close()

    assert len(rows) >= 150, f"Expected at least 150 rows, got {len(rows)}"
    assert elapsed_time < 3.0, f"Query latency {elapsed_time:.3f}s exceeded target 3.0 seconds"
    print(f"Query latency: {elapsed_time*1000:.2f} ms for {len(rows)} rows (PASS)")

def test_config_toml_theme_file():
    """Verify .streamlit/config.toml exists and contains brand theme definitions."""
    assert os.path.exists(CONFIG_PATH), f"Config file not found at {CONFIG_PATH}"
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    assert "[theme]" in content
    assert "primaryColor" in content
    assert "#1E3A8A" in content

def test_ui_component_html_generation():
    """Verify src/ui.py generates non-empty Tailwind HTML components."""
    import src.ui as ui
    
    header_html = ui.render_tailwind_header("Test Title", "Test Subtitle")
    assert "<!DOCTYPE html>" in header_html
    assert "Test Title" in header_html
    assert "tailwindcss.com" in header_html

    kpi_html = ui.render_kpi_card("ROAS", "3.18x", "+12%", "positive", "Target 2.0x")
    assert "ROAS" in kpi_html
    assert "3.18x" in kpi_html

    insight_html = ui.render_insight_box("Key Finding", "Weekend surge detected", "success")
    assert "Key Finding" in insight_html
    assert "Weekend surge detected" in insight_html
