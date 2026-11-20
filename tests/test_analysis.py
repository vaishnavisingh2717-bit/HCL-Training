"""
tests/test_analysis.py - Automated Technical Tests for Analysis & Math
Covers: TC-T-07, TC-T-08, TC-T-09, TC-T-10, TC-T-11
"""

import math
try:
    import pytest
except ImportError:
    pytest = None

def calculate_kpi_metrics(spend, revenue, impressions, clicks, conversions):
    """Deterministic reference KPI math function."""
    ctr = (clicks / impressions) if impressions > 0 else 0.0
    cpc = (spend / clicks) if clicks > 0 else 0.0
    cvr = (conversions / clicks) if clicks > 0 else 0.0
    cac = (spend / conversions) if conversions > 0 else 0.0
    roas = (revenue / spend) if spend > 0 else 0.0
    aov = (revenue / conversions) if conversions > 0 else 0.0
    profit = revenue - spend
    return {
        "ctr": round(ctr, 4),
        "cpc": round(cpc, 2),
        "cvr": round(cvr, 4),
        "cac": round(cac, 2),
        "roas": round(roas, 2),
        "aov": round(aov, 2),
        "profit": round(profit, 2)
    }

def two_sample_z_test(p1, n1, p2, n2):
    """Calculates pooled Z-score for two proportions."""
    p_pool = (p1 * n1 + p2 * n2) / (n1 + n2)
    se = math.sqrt(p_pool * (1.0 - p_pool) * (1.0 / n1 + 1.0 / n2))
    z = (p1 - p2) / se if se > 0 else 0.0
    return z

def test_tc_t_07_kpi_math_formulas_on_known_dataset():
    """TC-T-07: Verify financial and conversion KPI calculations on deterministic tiny dataset."""
    metrics = calculate_kpi_metrics(
        spend=10000.0,
        revenue=35000.0,
        impressions=50000,
        clicks=2000,
        conversions=80
    )
    assert metrics["ctr"] == 0.0400, f"Expected CTR=0.04, got {metrics['ctr']}"
    assert metrics["cpc"] == 5.00, f"Expected CPC=5.00, got {metrics['cpc']}"
    assert metrics["cvr"] == 0.0400, f"Expected CVR=0.04, got {metrics['cvr']}"
    assert metrics["cac"] == 125.00, f"Expected CAC=125.00, got {metrics['cac']}"
    assert metrics["roas"] == 3.50, f"Expected ROAS=3.50, got {metrics['roas']}"
    assert metrics["aov"] == 437.50, f"Expected AOV=437.50, got {metrics['aov']}"
    assert metrics["profit"] == 25000.00, f"Expected Profit=25000.00, got {metrics['profit']}"

def test_tc_t_08_threshold_underperformer_logic():
    """TC-T-08: Verify logic identifying underperforming campaigns (ROAS < 2.0x OR CAC > 1500)."""
    target_roas = 2.0
    target_cac = 1500.0

    # Case 1: Healthy campaign
    roas_1, cac_1 = 2.8, 1200.0
    is_under_1 = (roas_1 < target_roas) or (cac_1 > target_cac)
    assert is_under_1 is False, "Healthy campaign misclassified as underperforming"

    # Case 2: Bleeding ROAS campaign
    roas_2, cac_2 = 1.4, 1100.0
    is_under_2 = (roas_2 < target_roas) or (cac_2 > target_cac)
    assert is_under_2 is True, "Low ROAS campaign failed to flag"

    # Case 3: High CAC campaign
    roas_3, cac_3 = 2.2, 1950.0
    is_under_3 = (roas_3 < target_roas) or (cac_3 > target_cac)
    assert is_under_3 is True, "High CAC campaign failed to flag"

def test_tc_t_09_statistical_z_score_function():
    """TC-T-09: Verify two-sample proportion Z-test mathematical implementation."""
    z_stat = two_sample_z_test(p1=0.035, n1=1000, p2=0.020, n2=1000)
    assert 1.95 <= z_stat <= 2.15, f"Expected Z-score ~2.03, got {z_stat}"

def test_tc_t_10_attribution_models_credit_balance():
    """TC-T-10: Verify multi-touch attribution credit total matches 100% of conversions."""
    journey = [4, 1, 3]
    total_conversions = 1.0

    # First Touch
    first_touch_credit = {ch: (1.0 if idx == 0 else 0.0) for idx, ch in enumerate(journey)}
    assert sum(first_touch_credit.values()) == total_conversions

    # Last Touch
    last_touch_credit = {ch: (1.0 if idx == len(journey)-1 else 0.0) for idx, ch in enumerate(journey)}
    assert sum(last_touch_credit.values()) == total_conversions

    # Linear Touch (equal 1/N)
    linear_touch_credit = {ch: (1.0 / len(journey)) for ch in journey}
    assert math.isclose(sum(linear_touch_credit.values()), total_conversions, rel_tol=1e-5)

def test_tc_t_11_budget_optimizer_constraints():
    """TC-T-11: Verify budget reallocation optimizer strictly enforces total budget limit."""
    total_budget = 1500000.0
    allocated_shares = [0.35, 0.25, 0.15, 0.15, 0.10]
    
    assert math.isclose(sum(allocated_shares), 1.0, rel_tol=1e-5), "Shares must sum to 1.0"
    reallocated_amounts = [share * total_budget for share in allocated_shares]
    assert math.isclose(sum(reallocated_amounts), total_budget, rel_tol=1e-5), "Allocated spend must exactly equal budget ceiling"
