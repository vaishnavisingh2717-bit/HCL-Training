import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_rtm_excel():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Styling Definitions
    # -------------------------------------------------------------
    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="475569")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Segoe UI", size=9.5, color="1E293B")
    font_bold_body = Font(name="Segoe UI", size=9.5, bold=True, color="1E293B")
    font_future = Font(name="Segoe UI", size=9, italic=True, color="64748B")
    font_kpi_num = Font(name="Segoe UI", size=16, bold=True, color="1E3A8A")
    font_kpi_label = Font(name="Segoe UI", size=9, bold=True, color="475569")

    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_future_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_must = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid") # Red accent
    fill_must_bg = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_should_bg = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_could_bg = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    fill_status_bg = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_card = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_card = Border(left=Side(border_style="medium", color="94A3B8"),
                         right=Side(border_style="medium", color="94A3B8"),
                         top=Side(border_style="medium", color="94A3B8"),
                         bottom=Side(border_style="medium", color="94A3B8"))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # =============================================================
    # SHEET 1: Main RTM
    # =============================================================
    ws_rtm = wb.active
    ws_rtm.title = "Traceability_Matrix"
    ws_rtm.views.sheetView[0].showGridLines = True

    headers = [
        "OQ", "HMW", "Req ID (BR)", "Requirement", "Stakeholder (SH)",
        "Priority", "Status", "Feature", "DB table.column", "User Story",
        "UI Screen", "Test Case (TC-F / TC-T)"
    ]

    for col_idx, h in enumerate(headers, start=1):
        cell = ws_rtm.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.alignment = align_header
        cell.border = border_cell
        if col_idx <= 7:
            cell.fill = fill_header
        else:
            cell.fill = fill_future_header # Greyed navy for future step columns
    ws_rtm.row_dimensions[1].height = 28

    rows_data = [
        ("OQ-01", "HMW-01", "BR-01", "The business needs to consolidate and compute cross-channel Return on Ad Spend (ROAS) and Customer Acquisition Cost (CAC) across Google Ads, Meta Ads, Email, SEO, and Influencer marketing so that marketing leadership can evaluate blended and channel-specific capital efficiency.", "SH-01 (CMO)", "Must", "Identified"),
        ("OQ-01", "HMW-01", "BR-11", "The business needs to dynamically filter performance data by date ranges, acquisition channels, and campaign objectives so that stakeholders can isolate specific operational periods and campaign cohorts without altering baseline data.", "SH-01, SH-02, SH-03", "Must", "Identified"),
        ("OQ-01", "HMW-01", "BR-12", "The business needs to export filtered analytical summaries, raw campaign records, and data audit logs into standardized tabular formats (CSV/Excel) so that analysts can conduct offline reporting, executive briefing, and compliance documentation.", "SH-03, SH-05", "Should", "Identified"),
        ("OQ-07", "HMW-02", "BR-02", "The business needs to automatically identify and flag campaigns and ad sets that breach defined performance thresholds (ROAS < 2.0x or CAC > ₹1,500) so that campaign managers can immediately halt budget wastage on underperforming ad sets.", "SH-02 (Paid Ads Lead)", "Must", "Identified"),
        ("OQ-02", "HMW-03", "BR-03", "The business needs to track and decompose Click-Through Rates (CTR) and Conversion Rates (CVR) across marketing channels and campaign types over time so that growth teams can pinpoint whether funnel drop-offs are driven by ad creative fatigue or landing page friction.", "SH-02 (Paid Ads Lead)", "Must", "Identified"),
        ("OQ-02", "HMW-03", "BR-11", "The business needs to dynamically filter performance data by date ranges, acquisition channels, and campaign objectives so that stakeholders can isolate specific operational periods and campaign cohorts without altering baseline data.", "SH-01, SH-02, SH-03", "Must", "Identified"),
        ("OQ-03", "HMW-04", "BR-04", "The business needs to analyze customer repeat order history and calculate Customer Lifetime Value relative to acquisition cost (LTV:CAC ratio) per acquisition channel so that retention teams can prioritize channels that attract loyal, high-value customers.", "SH-04 (Retention Lead)", "Should", "Identified"),
        ("OQ-04", "HMW-05", "BR-05", "The business needs to evaluate day-of-week and monthly performance patterns in spend, conversion volume, and ROAS so that media planners can schedule ad delivery and adjust day-parting bids to align with peak buyer conversion windows.", "SH-02 (Paid Ads Lead)", "Should", "Identified"),
        ("OQ-05", "HMW-06", "BR-06", "The business needs to segment channel conversion volume and average order values by customer demographics (Age Group, Gender, City Tier) so that marketing teams can tailor channel copy, targeting, and promotional offers to high-performing audience cohorts.", "SH-04 (Retention Lead)", "Could", "Identified"),
        ("OQ-06", "HMW-07", "BR-07", "The business needs to perform inferential hypothesis testing (Two-sample proportion tests, ANOVA, Chi-Square) on channel conversion rates so that marketing analysts and leadership can confirm whether channel efficiency gaps are statistically significant (p < 0.05) rather than random sampling noise.", "SH-03, SH-05", "Must", "Identified"),
        ("OQ-08", "HMW-08", "BR-08", "The business needs to compare multi-touch attribution distribution (First-Touch, Last-Touch, and Equal Linear models) across customer conversion journeys so that marketing leadership can properly value top-of-funnel discovery channels alongside closing channels.", "SH-01, SH-03", "Must", "Identified"),
        ("OQ-09", "HMW-09", "BR-09", "The business needs to simulate cross-channel budget reallocation scenarios under a fixed total marketing spend based on historical marginal returns so that executive leadership can project potential revenue and ROAS uplifts before committing capital.", "SH-01 (CMO)", "Must", "Identified"),
        ("OQ-10", "HMW-10", "BR-10", "The business needs to audit marketing data quality by automatically detecting untracked UTM parameters, duplicate transaction IDs, missing foreign keys, and statistical spend outliers so that all downstream marketing analytics are based on reliable and auditable data.", "SH-03, SH-05", "Must", "Identified"),
        ("OQ-10", "HMW-10", "BR-12", "The business needs to export filtered analytical summaries, raw campaign records, and data audit logs into standardized tabular formats (CSV/Excel) so that analysts can conduct offline reporting, executive briefing, and compliance documentation.", "SH-03, SH-05", "Should", "Identified")
    ]

    for r_idx, r_data in enumerate(rows_data, start=2):
        ws_rtm.row_dimensions[r_idx].height = 45
        is_alt = (r_idx % 2 == 1)
        base_fill = fill_alt if is_alt else fill_white

        # 1. OQ
        c_oq = ws_rtm.cell(row=r_idx, column=1, value=r_data[0])
        c_oq.alignment = align_center
        c_oq.font = font_bold_body
        c_oq.border = border_cell
        c_oq.fill = base_fill

        # 2. HMW
        c_hmw = ws_rtm.cell(row=r_idx, column=2, value=r_data[1])
        c_hmw.alignment = align_center
        c_hmw.font = font_bold_body
        c_hmw.border = border_cell
        c_hmw.fill = base_fill

        # 3. Req ID (BR)
        c_br = ws_rtm.cell(row=r_idx, column=3, value=r_data[2])
        c_br.alignment = align_center
        c_br.font = font_bold_body
        c_br.border = border_cell
        c_br.fill = base_fill

        # 4. Requirement
        c_req = ws_rtm.cell(row=r_idx, column=4, value=r_data[3])
        c_req.alignment = align_left
        c_req.font = font_body
        c_req.border = border_cell
        c_req.fill = base_fill

        # 5. Stakeholder (SH)
        c_sh = ws_rtm.cell(row=r_idx, column=5, value=r_data[4])
        c_sh.alignment = align_center
        c_sh.font = font_body
        c_sh.border = border_cell
        c_sh.fill = base_fill

        # 6. Priority
        c_pri = ws_rtm.cell(row=r_idx, column=6, value=r_data[5])
        c_pri.alignment = align_center
        c_pri.font = font_bold_body
        c_pri.border = border_cell
        if r_data[5] == "Must":
            c_pri.fill = fill_must_bg
        elif r_data[5] == "Should":
            c_pri.fill = fill_should_bg
        else:
            c_pri.fill = fill_could_bg

        # 7. Status
        c_stat = ws_rtm.cell(row=r_idx, column=7, value=r_data[6])
        c_stat.alignment = align_center
        c_stat.font = font_body
        c_stat.border = border_cell
        c_stat.fill = fill_status_bg

        # 8-12. Future Step Placeholders
        for future_col in range(8, 13):
            c_fut = ws_rtm.cell(row=r_idx, column=future_col, value="")
            c_fut.alignment = align_center
            c_fut.font = font_future
            c_fut.border = border_cell
            c_fut.fill = base_fill

    # Apply Auto-filter
    ws_rtm.auto_filter.ref = f"A1:L{len(rows_data)+1}"
    # Freeze header row
    ws_rtm.freeze_panes = "A2"

    col_widths = {
        "A": 10, # OQ
        "B": 12, # HMW
        "C": 14, # Req ID
        "D": 48, # Requirement
        "E": 20, # Stakeholder
        "F": 12, # Priority
        "G": 14, # Status
        "H": 16, # Feature (Step 6)
        "I": 18, # DB table.column (Step 8)
        "J": 16, # User Story (Step 11)
        "K": 16, # UI Screen (Step 12)
        "L": 22  # Test Case (Step 13)
    }
    for col_letter, width in col_widths.items():
        ws_rtm.column_dimensions[col_letter].width = width

    # =============================================================
    # SHEET 2: Coverage Analysis Sheet
    # =============================================================
    ws_cov = wb.create_sheet(title="Coverage_Analysis")
    ws_cov.views.sheetView[0].showGridLines = True

    # Title
    ws_cov.cell(row=1, column=1, value="Traceability & Coverage Audit Dashboard").font = font_title
    ws_cov.cell(row=2, column=1, value="Project: Channel Wise Marketing Performance | Step 4 SDLC Milestone").font = font_subtitle

    # KPI Summary Cards (Row 4-5)
    kpis = [
        (2, "Operational Questions", "10 / 10 (100%)", "OQ-01 to OQ-10"),
        (4, "How Might We (HMW)", "10 / 10 (100%)", "HMW-01 to HMW-10"),
        (6, "Business Requirements", "12 / 12 (100%)", "BR-01 to BR-12"),
        (8, "Coverage Gaps", "0 Gaps (0%)", "Complete Alignment")
    ]
    for col_pos, label, val, sub in kpis:
        # Merge 2 columns for card
        ws_cov.merge_cells(start_row=4, start_column=col_pos, end_row=5, end_column=col_pos+1)
        card_cell = ws_cov.cell(row=4, column=col_pos)
        card_cell.value = f"{label}\n{val}\n({sub})"
        card_cell.alignment = align_center
        card_cell.font = font_bold_body
        card_cell.fill = fill_card
        
        # Apply borders to the merged block
        for r in range(4, 6):
            for c in range(col_pos, col_pos+2):
                ws_cov.cell(row=r, column=c).border = border_card

    # Section 1: Coverage Summary Table (Row 7)
    ws_cov.cell(row=7, column=1, value="1. Traceability Dimension Gap Verification").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")

    gap_headers = ["Traceability Dimension", "Total Count", "Mapped Count", "Unmapped / Gap Count", "Compliance Status"]
    for j, gh in enumerate(gap_headers, start=1):
        c = ws_cov.cell(row=8, column=j, value=gh)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_header
        c.border = border_cell

    gap_data = [
        ("Operational Questions (OQs)", 10, 10, 0, "100% Covered — Zero Gaps"),
        ("How Might We Statements (HMWs)", 10, 10, 0, "100% Covered — Zero Gaps"),
        ("Business Requirements (BRs)", 12, 12, 0, "100% Covered — Zero Gaps"),
        ("HMWs without BR", 10, 0, 0, "Zero Orphan HMWs (Pass)"),
        ("BRs without HMW", 12, 0, 0, "Zero Orphan BRs (Pass)")
    ]

    for r_idx, g_row in enumerate(gap_data, start=9):
        is_alt = (r_idx % 2 == 1)
        base_fill = fill_alt if is_alt else fill_white
        for c_idx, val in enumerate(g_row, start=1):
            c = ws_cov.cell(row=r_idx, column=c_idx, value=val)
            c.font = font_bold_body if c_idx == 1 or c_idx == 5 else font_body
            c.alignment = align_center if c_idx > 1 else align_left
            c.border = border_cell
            c.fill = base_fill

    # Section 2: OQ Mapping Breakdown Table (Row 16)
    ws_cov.cell(row=16, column=1, value="2. Detailed Operational Question (OQ) Mapping Breakdown").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")

    oq_map_headers = ["OQ ID", "Question Summary", "Associated HMW(s)", "Associated BR(s)", "Primary Stakeholder", "Coverage Status"]
    for j, oh in enumerate(oq_map_headers, start=1):
        c = ws_cov.cell(row=17, column=j, value=oh)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_header
        c.border = border_cell

    oq_map_data = [
        ("OQ-01", "Overall & Channel ROAS and CAC", "HMW-01", "BR-01, BR-11, BR-12", "SH-01 (CMO)", "Complete (100%)"),
        ("OQ-02", "CTR & CVR Funnel Variations Over Time", "HMW-03", "BR-03, BR-11", "SH-02 (Paid Lead)", "Complete (100%)"),
        ("OQ-03", "Customer LTV:CAC Ratio by Channel", "HMW-04", "BR-04", "SH-04 (Retention Lead)", "Complete (100%)"),
        ("OQ-04", "Temporal & Day-of-Week Seasonality Patterns", "HMW-05", "BR-05", "SH-02 (Paid Lead)", "Complete (100%)"),
        ("OQ-05", "Demographic Segment Conversion Affinities", "HMW-06", "BR-06", "SH-04 (Retention Lead)", "Complete (100%)"),
        ("OQ-06", "Statistical Significance of Conversion Gaps", "HMW-07", "BR-07", "SH-03 (Analyst) / SH-05 (Mentor)", "Complete (100%)"),
        ("OQ-07", "Underperforming Ad Sets Against Thresholds", "HMW-02", "BR-02", "SH-02 (Paid Lead)", "Complete (100%)"),
        ("OQ-08", "Multi-Touch Customer Journey Attribution", "HMW-08", "BR-08", "SH-01 (CMO) / SH-03 (Analyst)", "Complete (100%)"),
        ("OQ-09", "Optimal Budget Reallocation under Fixed Spend", "HMW-09", "BR-09", "SH-01 (CMO)", "Complete (100%)"),
        ("OQ-10", "Data Quality Audits, UTMs & Spend Outliers", "HMW-10", "BR-10, BR-12", "SH-03 (Analyst) / SH-05 (Mentor)", "Complete (100%)")
    ]

    for r_idx, oq_row in enumerate(oq_map_data, start=18):
        is_alt = (r_idx % 2 == 1)
        base_fill = fill_alt if is_alt else fill_white
        for c_idx, val in enumerate(oq_row, start=1):
            c = ws_cov.cell(row=r_idx, column=c_idx, value=val)
            c.font = font_bold_body if c_idx == 1 or c_idx == 6 else font_body
            c.alignment = align_center if c_idx in [1, 3, 4, 6] else align_left
            c.border = border_cell
            c.fill = base_fill

    # Adjust Coverage Column Widths
    cov_widths = {"A": 12, "B": 36, "C": 20, "D": 22, "E": 28, "F": 24, "G": 18, "H": 18, "I": 18}
    for c_letter, w in cov_widths.items():
        ws_cov.column_dimensions[c_letter].width = w

    # Save to file
    out_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    build_rtm_excel()
