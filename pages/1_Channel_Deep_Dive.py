"""
pages/1_Channel_Deep_Dive.py - Channel Wise Marketing Performance Analysis
UI-02: Channel Comparative Analytics & Campaign Quadrants
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
import src.ui as ui

def main():
    if HAS_STREAMLIT:
        st.set_page_config(page_title="Channel Deep Dive | Marketing Analytics", page_icon="📊", layout="wide")

    ui.inject_global_css()
    ui.render_tailwind_header(
        title="Channel Deep Dive & Campaign Performance Matrix",
        subtitle="Granular evaluation of acquisition channels, campaign unit economics, and efficiency quadrants.",
        badge_text="UI-02 | Deep Dive Analytics"
    )

    if not HAS_STREAMLIT:
        print("Page 1 executed headlessly.")
        return

    raw_df = dl.load_analytical_view()
    cleaned_df, _ = cleaner.clean_marketing_data(raw_df)

    # 1. Top Channel Metric Cards
    ch_sum = analysis.channel_performance_summary(cleaned_df)
    
    st.subheader("Acquisition Channel Performance Matrix")
    cols = st.columns(len(ch_sum))
    for i, ch in enumerate(ch_sum):
        with cols[i]:
            st.metric(
                label=ch["channel_name"],
                value=f"{ch['roas']:.2f}x ROAS",
                delta=f"₹{ch['cac']:,.0f} CAC",
                delta_color="inverse" if ch["cac"] > 1500 else "normal"
            )

    st.markdown("---")

    # 2. Campaign Efficiency Scatter Plot (ROAS vs CAC Quadrants)
    st.subheader("Campaign Unit Economics & Quadrant Analysis")
    st.caption("Visualizing campaigns across ROAS (efficiency) and CAC (acquisition cost). Target threshold: ROAS >= 2.0x, CAC <= ₹1,500.")

    underperformers = analysis.identify_underperforming_campaigns(cleaned_df, roas_threshold=2.0, cac_threshold=1500.0)
    
    if HAS_PLOTLY and hasattr(cleaned_df, "groupby"):
        cmp_stats = cleaned_df.groupby(["campaign_id", "campaign_name", "channel_name"]).agg({
            "spend": "sum",
            "revenue": "sum",
            "conversions": "sum",
            "clicks": "sum"
        }).reset_index()

        cmp_stats["roas"] = cmp_stats.apply(lambda r: round(r["revenue"] / r["spend"], 2) if r["spend"] > 0 else 0.0, axis=1)
        cmp_stats["cac"] = cmp_stats.apply(lambda r: round(r["spend"] / r["conversions"], 2) if r["conversions"] > 0 else 0.0, axis=1)

        fig_scatter = px.scatter(
            cmp_stats,
            x="cac",
            y="roas",
            size="spend",
            color="channel_name",
            hover_name="campaign_name",
            labels={"cac": "Customer Acquisition Cost (₹)", "roas": "Return on Ad Spend (ROAS)", "spend": "Total Spend"},
            color_discrete_sequence=["#1E3A8A", "#2563EB", "#0D9488", "#F59E0B", "#EF4444"]
        )
        
        # Add benchmark threshold lines with clear non-overlapping annotations
        fig_scatter.add_hline(y=2.0, line_dash="dash", line_color="#EF4444", annotation_text="Target ROAS (2.0x)", annotation_position="top left")
        fig_scatter.add_vline(x=1500.0, line_dash="dash", line_color="#EF4444", annotation_text="Max Acceptable CAC (₹1,500)", annotation_position="top right")
        fig_scatter.update_layout(height=420)
        
        ui.apply_plotly_theme(fig_scatter)
        st.plotly_chart(fig_scatter, use_container_width=True)

    # 3. Anomaly Callout Box
    n_under = len(underperformers)
    ui.render_insight_box(
        title=f"Anomaly Alert: {n_under} Campaigns Requiring Immediate Intervention",
        text=f"{n_under} campaigns currently violate capital efficiency guidelines by generating sub-2.0x ROAS or exceeding ₹1,500 CAC. These campaigns are predominantly concentrated in fixed-fee Influencer Partnerships and broad unsegmented Meta lookalikes.",
        alert_type="warning" if n_under > 0 else "success"
    )

    # 4. Underperforming Campaigns Table
    st.subheader("Campaign Action Register")
    if underperformers:
        import pandas as pd
        df_under = pd.DataFrame(underperformers)[["campaign_name", "channel_name", "spend", "revenue", "roas", "cac", "reason"]]
        df_under.columns = ["Campaign Name", "Channel", "Spend (₹)", "Revenue (₹)", "ROAS", "CAC (₹)", "Deficiency Reason"]
        st.dataframe(df_under.style.format({
            "Spend (₹)": "₹{:,.2f}",
            "Revenue (₹)": "₹{:,.2f}",
            "ROAS": "{:.2f}x",
            "CAC (₹)": "₹{:,.2f}"
        }), use_container_width=True)
    else:
        st.success("All campaigns meet target profitability thresholds.")

if __name__ == "__main__":
    main()
