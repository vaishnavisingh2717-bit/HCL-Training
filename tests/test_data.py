"""
tests/test_data.py - Automated Technical Tests for Database & Data Layer
Covers: TC-T-01, TC-T-02, TC-T-03, TC-T-04, TC-T-05, TC-T-06
"""

import os
import sqlite3
try:
    import pytest
except ImportError:
    pytest = None

DB_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data", "app.db"))

def get_db_connection():
    assert os.path.exists(DB_PATH), f"Database file not found at {DB_PATH}"
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

if pytest is not None:
    @pytest.fixture(scope="module")
    def db_conn():
        conn = get_db_connection()
        yield conn
        conn.close()

def test_tc_t_01_database_connection_and_tables(db_conn=None):
    """TC-T-01: Verify SQLite database connects and contains all planned tables and views."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view');")
    entities = {row[0]: row[1] for row in cur.fetchall()}
    
    expected_tables = [
        "dim_channel",
        "dim_campaign",
        "dim_customer",
        "fact_daily_campaign_performance",
        "fact_customer_touchpoint"
    ]
    for tbl in expected_tables:
        assert tbl in entities, f"Missing table: {tbl}"
        assert entities[tbl] == "table"
        
    assert "vw_marketing_analysis" in entities, "Missing view: vw_marketing_analysis"
    assert entities["vw_marketing_analysis"] == "view"
    if db_conn is None:
        conn.close()

def test_tc_t_02_row_and_column_counts(db_conn=None):
    """TC-T-02: Verify planned row and column count thresholds across all entities."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM dim_channel;")
    assert cur.fetchone()[0] == 5, "dim_channel must have exactly 5 rows"
    
    cur.execute("SELECT COUNT(*) FROM dim_campaign;")
    n_cmp = cur.fetchone()[0]
    assert 20 <= n_cmp <= 35, f"dim_campaign count {n_cmp} outside expected range [20, 35]"
    
    cur.execute("SELECT COUNT(*) FROM dim_customer;")
    n_cust = cur.fetchone()[0]
    assert 100 <= n_cust <= 200, f"dim_customer count {n_cust} outside expected range [100, 200]"
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance;")
    n_fact = cur.fetchone()[0]
    assert n_fact >= 150, f"fact_daily_campaign_performance count {n_fact} must be >= 150"
    
    cur.execute("SELECT COUNT(*) FROM fact_customer_touchpoint;")
    n_tp = cur.fetchone()[0]
    assert n_tp >= 300, f"fact_customer_touchpoint count {n_tp} must be >= 300"
    
    cur.execute("SELECT * FROM vw_marketing_analysis LIMIT 1;")
    col_count = len(cur.description)
    assert 30 <= col_count <= 42, f"vw_marketing_analysis column count {col_count} outside target range [30, 42]"
    if db_conn is None:
        conn.close()

def test_tc_t_03_primary_key_uniqueness(db_conn=None):
    """TC-T-03: Assert 100% Primary Key uniqueness across master and fact tables."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    
    pk_checks = [
        ("dim_channel", "channel_id"),
        ("dim_campaign", "campaign_id"),
        ("dim_customer", "customer_id"),
        ("fact_daily_campaign_performance", "fact_id"),
        ("fact_customer_touchpoint", "touchpoint_id")
    ]
    for tbl, pk in pk_checks:
        cur.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {tbl};")
        total_pks, distinct_pks = cur.fetchone()
        assert total_pks == distinct_pks, f"Duplicate Primary Keys detected in table {tbl} on column {pk}"
    if db_conn is None:
        conn.close()

def test_tc_t_04_foreign_key_referential_integrity(db_conn=None):
    """TC-T-04: Execute PRAGMA foreign_key_check to verify zero orphan records."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_key_check;")
    violations = cur.fetchall()
    assert len(violations) == 0, f"Referential integrity violations found: {violations}"
    if db_conn is None:
        conn.close()

def test_tc_t_05_physical_range_and_bound_invariants(db_conn=None):
    """TC-T-05: Verify domain invariants (Clicks <= Impressions, Conv <= Clicks, Spend/Revenue >= 0)."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance WHERE clicks > impressions;")
    assert cur.fetchone()[0] == 0, "Found rows where clicks > impressions"
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance WHERE conversions > clicks;")
    assert cur.fetchone()[0] == 0, "Found rows where conversions > clicks"
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance WHERE spend < 0 OR revenue < 0;")
    assert cur.fetchone()[0] == 0, "Found rows with negative financial values"
    if db_conn is None:
        conn.close()

def test_tc_t_06_data_cleaning_and_outlier_detection(db_conn=None):
    """TC-T-06: Verify presence and detection of injected data quality issues (outliers, null UTMs)."""
    conn = db_conn or get_db_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance WHERE is_outlier = 1;")
    outlier_count = cur.fetchone()[0]
    assert outlier_count >= 1, "Expected injected outliers flagged with is_outlier=1"
    
    cur.execute("SELECT COUNT(*) FROM fact_daily_campaign_performance WHERE utm_medium IS NULL OR utm_source IS NULL;")
    null_utm_count = cur.fetchone()[0]
    assert null_utm_count >= 1, "Expected injected null UTM records for cleaning audit"
    if db_conn is None:
        conn.close()
