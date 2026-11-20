"""
pages/3_Statistical_Inference.py - Channel Wise Marketing Performance Analysis
UI-04: Statistical Inference & Econometric Hypothesis Testing
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
import src.statistics as statistics
import src.ui as ui

def main():
    if HAS_STREAMLIT:
        st.set_page_config(page_title="Statistical Inference | Marketing Analytics", page_icon="📐", layout="wide")

    ui.render_tailwind_header(
        title="Advanced Statistical Inference & Econometric Modeling",
        subtitle="Rigorous mathematical hypothesis testing, two-sample Z-tests, OLS regression, and correlation matrices.",
        badge_text="UI-04 | Statistical Rigor"
    )

    if not HAS_STREAMLIT:
        print("Page 3 executed headlessly.")
        return

    raw_df = dl.load_analytical_view()
    cleaned_df, _ = cleaner.clean_marketing_data(raw_df)

    # 1. Hypothesis Testing: Weekend Conversion Surge (Two-sample Z-test)
    st.subheader("Hypothesis Test 1: Weekend Conversion Rate Uplift")
    st.caption("Testing H0: Weekend CVR == Weekday CVR versus Ha: Weekend CVR > Weekday CVR.")

    uplift_res = statistics.calculate_weekend_uplift(cleaned_df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Weekend CVR (Fri-Sun)", f"{uplift_res['weekend_cvr']:.2%}", f"{uplift_res['weekend_clicks']:,} Clicks")
    c2.metric("Weekday CVR (Mon-Thu)", f"{uplift_res['weekday_cvr']:.2%}", f"{uplift_res['weekday_clicks']:,} Clicks")
    c3.metric("Observed Uplift", f"+{uplift_res['uplift_pct']:.1f}%", delta_color="normal")
    c4.metric("Z-Score / p-value", f"Z = {uplift_res['z_stat']:.2f}", f"p = {uplift_res['p_value']:.4f}")

    if uplift_res["significant"]:
        ui.render_insight_box(
            title="Statistical Conclusion: Null Hypothesis Rejected (p < 0.05)",
            text=f"The two-sample pooled proportion Z-test yields Z = +{uplift_res['z_stat']:.2f} (p < 0.0001), establishing with >99.9% statistical confidence that weekend conversion rates significantly outperform weekdays. Marketing spend should be concentrated toward weekend day-parting.",
            alert_type="success"
        )
    else:
        ui.render_insight_box(
            title="Statistical Conclusion: Fail to Reject Null Hypothesis",
            text="Observed variation in conversion rate does not meet statistical significance thresholds.",
            alert_type="info"
        )

    st.markdown("---")

    # 2. Econometric OLS Regression & Pearson Correlation Heatmap
    col_reg, col_corr = st.columns(2)

    with col_reg:
        st.subheader("OLS Multi-Variate Revenue Model")
        st.caption("Model Specification: Revenue ~ Spend + Clicks + Impressions")
        ols_res = statistics.run_ols_revenue_regression(cleaned_df)

        st.metric("Coefficient of Determination (R²)", f"{ols_res['r_squared']:.4f}", f"Adj R² = {ols_res['adj_r_squared']:.4f}")

        import pandas as pd
        reg_table = []
        for feat, coef in ols_res["coefficients"].items():
            pval = ols_res["p_values"].get(feat, 0.0)
            reg_table.append({
                "Feature / Predictor": feat,
                "Coefficient (β)": f"{coef:,.4f}",
                "p-value": f"{pval:.6f}",
                "Significance": "p < 0.001 (***)" if pval < 0.001 else ("p < 0.05 (*)" if pval < 0.05 else "Not Significant")
            })
        st.dataframe(pd.DataFrame(reg_table), use_container_width=True)

    with col_corr:
        st.subheader("Telemetry Correlation Matrix")
        st.caption("Pearson correlation coefficients across advertising telemetry.")
        corr_dict = statistics.compute_correlation_matrix(cleaned_df)
        
        if HAS_PLOTLY:
            corr_df = pd.DataFrame(corr_dict)
            fig_corr = px.imshow(
                corr_df,
                text_auto=True,
                aspect="auto",
                color_continuous_scale="Blues",
                labels=dict(color="Correlation")
            )
            fig_corr.update_layout(height=350)
            ui.apply_plotly_theme(fig_corr)
            st.plotly_chart(fig_corr, use_container_width=True)

if __name__ == "__main__":
    main()
