"""
src/analysis.py - Core Analytical & Financial KPI Computation Engine
Calculates ROAS, CAC, CTR, CPC, CVR, AOV, Net Profit and campaign anomaly thresholds.
"""

def calculate_kpi_metrics(spend: float, revenue: float, impressions: int, clicks: int, conversions: int) -> dict:
    """Computes exact mathematical KPI formulas."""
    ctr = (clicks / impressions) if impressions > 0 else 0.0
    cpc = (spend / clicks) if clicks > 0 else 0.0
    cvr = (conversions / clicks) if clicks > 0 else 0.0
    cac = (spend / conversions) if conversions > 0 else 0.0
    roas = (revenue / spend) if spend > 0 else 0.0
    aov = (revenue / conversions) if conversions > 0 else 0.0
    profit = revenue - spend
    return {
        "spend": round(spend, 2),
        "revenue": round(revenue, 2),
        "profit": round(profit, 2),
        "impressions": int(impressions),
        "clicks": int(clicks),
        "conversions": int(conversions),
        "ctr": round(ctr, 4),
        "cpc": round(cpc, 2),
        "cvr": round(cvr, 4),
        "cac": round(cac, 2),
        "roas": round(roas, 2),
        "aov": round(aov, 2)
    }

def compute_kpis(data) -> dict:
    """Aggregates a dataset (DataFrame or list of dicts) into high-level KPIs."""
    records = data.to_dict(orient="records") if hasattr(data, "to_dict") else data
    tot_spend = sum(float(r.get("spend", 0.0)) for r in records)
    tot_rev = sum(float(r.get("revenue", 0.0)) for r in records)
    tot_imp = sum(int(r.get("impressions", 0)) for r in records)
    tot_clk = sum(int(r.get("clicks", 0)) for r in records)
    tot_conv = sum(int(r.get("conversions", 0)) for r in records)
    return calculate_kpi_metrics(tot_spend, tot_rev, tot_imp, tot_clk, tot_conv)

def identify_underperforming_campaigns(data, roas_threshold: float = 2.0, cac_threshold: float = 1500.0) -> list:
    """Flags campaigns failing financial health thresholds (ROAS < 2.0x OR CAC > 1500)."""
    records = data.to_dict(orient="records") if hasattr(data, "to_dict") else data
    
    # Aggregate by campaign
    campaign_map = {}
    for r in records:
        cid = r.get("campaign_id")
        cname = r.get("campaign_name", f"Campaign {cid}")
        chname = r.get("channel_name", "Unknown")
        if cid not in campaign_map:
            campaign_map[cid] = {
                "campaign_id": cid,
                "campaign_name": cname,
                "channel_name": chname,
                "spend": 0.0,
                "revenue": 0.0,
                "clicks": 0,
                "conversions": 0
            }
        campaign_map[cid]["spend"] += float(r.get("spend", 0.0))
        campaign_map[cid]["revenue"] += float(r.get("revenue", 0.0))
        campaign_map[cid]["clicks"] += int(r.get("clicks", 0))
        campaign_map[cid]["conversions"] += int(r.get("conversions", 0))
        
    underperformers = []
    for cid, stats in campaign_map.items():
        metrics = calculate_kpi_metrics(
            stats["spend"], stats["revenue"], 
            stats["clicks"] * 20, stats["clicks"], stats["conversions"]
        )
        is_bleeding = (metrics["roas"] < roas_threshold) or (metrics["cac"] > cac_threshold)
        if is_bleeding:
            reason = []
            if metrics["roas"] < roas_threshold:
                reason.append(f"Sub-target ROAS ({metrics['roas']}x < {roas_threshold}x)")
            if metrics["cac"] > cac_threshold:
                reason.append(f"Excessive CAC (₹{metrics['cac']} > ₹{cac_threshold})")
            underperformers.append({
                **stats,
                **metrics,
                "status": "Underperforming",
                "reason": " & ".join(reason)
            })
    return underperformers

def channel_performance_summary(data) -> list:
    """Aggregates metrics by channel name."""
    records = data.to_dict(orient="records") if hasattr(data, "to_dict") else data
    ch_map = {}
    for r in records:
        ch = r.get("channel_name", "Unknown")
        if ch not in ch_map:
            ch_map[ch] = {"channel_name": ch, "spend": 0.0, "revenue": 0.0, "impressions": 0, "clicks": 0, "conversions": 0}
        ch_map[ch]["spend"] += float(r.get("spend", 0.0))
        ch_map[ch]["revenue"] += float(r.get("revenue", 0.0))
        ch_map[ch]["impressions"] += int(r.get("impressions", 0))
        ch_map[ch]["clicks"] += int(r.get("clicks", 0))
        ch_map[ch]["conversions"] += int(r.get("conversions", 0))

    summary = []
    for ch, stats in ch_map.items():
        kpis = calculate_kpi_metrics(stats["spend"], stats["revenue"], stats["impressions"], stats["clicks"], stats["conversions"])
        summary.append({**stats, **kpis})
    return sorted(summary, key=lambda x: x["revenue"], reverse=True)
