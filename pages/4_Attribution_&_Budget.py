"""
pages/4_Attribution_&_Budget.py - Channel Wise Marketing Performance Analysis
UI-05: Multi-Touch Attribution & Algorithmic Budget Optimization
"""

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

try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

import src.data_loader as dl
import src.cleaner as cleaner
import src.analysis as analysis
import src.attribution as attribution
import src.optimizer as optimizer
import src.ui as ui

def main():
    if HAS_STREAMLIT:
        st.set_page_config(page_title="Attribution & Budget | Marketing Analytics", page_icon="⚖️", layout="wide")

    ui.inject_global_css()
    ui.render_tailwind_header(
        title="Multi-Touch Attribution & Algorithmic Capital Simulator",
        subtitle="Comparing First-Touch, Last-Touch, and Linear models alongside capital reallocation simulation.",
        badge_text="UI-05 | Capital Allocation"
    )

    if not HAS_STREAMLIT:
        print("Page 4 executed headlessly.")
        return

    # Load data
    raw_df = dl.load_analytical_view()
    cleaned_df, _ = cleaner.clean_marketing_data(raw_df)
    touchpoints_df = dl.load_touchpoints()
    ch_sum = analysis.channel_performance_summary(cleaned_df)

    # 1. Multi-Touch Attribution Comparison
    st.subheader("1. Multi-Touch Attribution Comparison (First vs Last vs Linear)")
    st.caption("Resolving last-click attribution bias by examining touchpoint conversion journey credit.")

    attr_res = attribution.compute_attribution_models(touchpoints_df)
    
    if HAS_PLOTLY:
        import pandas as pd
        channels = list(attr_res["first_touch"].keys())
        attr_data = []
        for ch in channels:
            attr_data.append({"Channel": ch, "Model": "First-Touch", "Conversions": attr_res["first_touch"].get(ch, 0.0)})
            attr_data.append({"Channel": ch, "Model": "Last-Touch", "Conversions": attr_res["last_touch"].get(ch, 0.0)})
            attr_data.append({"Channel": ch, "Model": "Linear (Equal)", "Conversions": attr_res["linear"].get(ch, 0.0)})

        df_attr = pd.DataFrame(attr_data)
        fig_attr = px.bar(
            df_attr,
            x="Channel",
            y="Conversions",
            color="Model",
            barmode="group",
            color_discrete_sequence=["#2563EB", "#0D9488", "#F59E0B"]
        )
        fig_attr.update_layout(height=380)
        ui.apply_plotly_theme(fig_attr)
        st.plotly_chart(fig_attr, use_container_width=True)

    ui.render_insight_box(
        title="Attribution Structural Shift",
        text="Under Last-Touch, Email and Meta receive 56% of conversion credit. Under First-Touch, Influencer Partnerships (37%) and Organic Search (21%) capture the majority of initial discovery. Defunding Influencers solely based on last-click ROAS would sever top-of-funnel customer acquisition.",
        alert_type="info"
    )

    st.markdown("---")

    # 2. Interactive Budget Reallocation Simulator
    st.subheader("2. Algorithmic Budget Reallocation Simulator")
    st.caption("Simulate shifting capital from underperforming to high-ROAS channels under strict budget preservation.")

    opt_res = optimizer.optimize_budget_reallocation(ch_sum)

    col_sim1, col_sim2, col_sim3 = st.columns(3)
    col_sim1.metric("Current Gross Revenue", f"₹{opt_res['current_revenue']:,.2f}", "Baseline")
    col_sim2.metric("Projected Revenue", f"₹{opt_res['projected_revenue']:,.2f}", f"+{opt_res['revenue_uplift_pct']:.1f}% Uplift", delta_color="normal")
    col_sim3.metric("Projected Blended ROAS", f"{opt_res['projected_blended_roas']:.2f}x", f"From {opt_res['current_blended_roas']:.2f}x")

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # Comparison Table
    import pandas as pd
    comp_df = pd.DataFrame(opt_res["comparison"])
    comp_df.columns = ["Channel", "Current Spend", "Current Rev", "Current ROAS", "Alloc Share %", "Optimized Spend", "Projected Rev", "Projected ROAS"]
    st.dataframe(comp_df.style.format({
        "Current Spend": "₹{:,.2f}",
        "Current Rev": "₹{:,.2f}",
        "Current ROAS": "{:.2f}x",
        "Alloc Share %": "{:.1f}%",
        "Optimized Spend": "₹{:,.2f}",
        "Projected Rev": "₹{:,.2f}",
        "Projected ROAS": "{:.2f}x"
    }), use_container_width=True)

if __name__ == "__main__":
    main()
