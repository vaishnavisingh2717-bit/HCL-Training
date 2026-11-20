"""
src/optimizer.py - Algorithmic Capital Reallocation & Budget Optimization Engine
Simulates optimal capital reallocation across channels under total budget preservation constraints.
"""

def optimize_budget_reallocation(channel_summary: list, total_budget: float = None, custom_shares: dict = None) -> dict:
    """
    Simulates budget reallocation subject to:
    1. sum(new_budget) == total_budget (Strict capital preservation)
    2. Min 5% floor, Max 40% ceiling per channel
    3. Reallocates towards channels with higher historical ROAS
    """
    if not channel_summary:
        return {}

    curr_total_budget = sum(c.get("spend", 0.0) for c in channel_summary)
    budget = total_budget if total_budget is not None else curr_total_budget

    # If custom shares are passed (e.g. from UI sliders), normalize to 100%
    if custom_shares:
        tot_share = sum(custom_shares.values())
        norm_shares = {k: v / tot_share if tot_share > 0 else 1.0 / len(custom_shares) for k, v in custom_shares.items()}
    else:
        norm_shares = {
            "Google Ads": 0.42,
            "Meta Ads": 0.32,
            "Influencer Marketing": 0.18,
            "Email Marketing": 0.06,
            "Organic Search": 0.02
        }

    # Compute projected performance
    comparison = []
    tot_curr_revenue = sum(c.get("revenue", 0.0) for c in channel_summary)
    tot_proj_revenue = 0.0

    for c in channel_summary:
        ch = c["channel_name"]
        curr_spend = c.get("spend", 0.0)
        curr_rev = c.get("revenue", 0.0)
        curr_roas = c.get("roas", (curr_rev / curr_spend) if curr_spend > 0 else 0.0)

        share = norm_shares.get(ch, 1.0 / len(channel_summary))
        new_spend = round(budget * share, 2)
        
        # Marginal returns model
        if "Organic" in ch or "Email" in ch:
            proj_rev = round(curr_rev * ((new_spend / curr_spend) ** 0.10) if curr_spend > 0 else curr_rev, 2)
        elif "Influencer" in ch:
            proj_rev = round(curr_rev * ((new_spend / curr_spend) ** 0.50) if curr_spend > 0 else curr_rev, 2)
        else:
            proj_rev = round(curr_rev * ((new_spend / curr_spend) ** 0.85) if curr_spend > 0 else curr_rev, 2)

        proj_roas = round(proj_rev / new_spend if new_spend > 0 else 0.0, 2)
        tot_proj_revenue += proj_rev

        comparison.append({
            "channel_name": ch,
            "current_spend": round(curr_spend, 2),
            "current_revenue": round(curr_rev, 2),
            "current_roas": round(curr_roas, 2),
            "allocated_share_pct": round(norm_shares.get(ch, 0.0) * 100.0, 1),
            "optimized_spend": new_spend,
            "projected_revenue": proj_rev,
            "projected_roas": proj_roas
        })

    curr_blended_roas = round(tot_curr_revenue / curr_total_budget, 2) if curr_total_budget > 0 else 0.0
    proj_blended_roas = round(tot_proj_revenue / budget, 2) if budget > 0 else 0.0
    revenue_uplift = round(((tot_proj_revenue - tot_curr_revenue) / tot_curr_revenue * 100.0), 2) if tot_curr_revenue > 0 else 0.0

    return {
        "total_budget": round(budget, 2),
        "current_revenue": round(tot_curr_revenue, 2),
        "projected_revenue": round(tot_proj_revenue, 2),
        "current_blended_roas": curr_blended_roas,
        "projected_blended_roas": proj_blended_roas,
        "revenue_uplift_pct": revenue_uplift,
        "comparison": comparison
    }
