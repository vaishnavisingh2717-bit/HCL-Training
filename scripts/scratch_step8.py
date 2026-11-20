import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_database_design_excel():
    wb = openpyxl.Workbook()
    
    # Common styles
    font_title = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="475569")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Segoe UI", size=9.5, color="1E293B")
    font_code = Font(name="Consolas", size=9, color="0F172A")
    font_bold_body = Font(name="Segoe UI", size=9.5, bold=True, color="1E293B")
    font_pk = Font(name="Segoe UI", size=9.5, bold=True, color="B91C1C") # Crimson for PK
    font_fk = Font(name="Segoe UI", size=9.5, bold=True, color="1D4ED8") # Blue for FK

    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_card = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_sql = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_card = Border(left=Side(border_style="medium", color="94A3B8"),
                         right=Side(border_style="medium", color="94A3B8"),
                         top=Side(border_style="medium", color="94A3B8"),
                         bottom=Side(border_style="medium", color="94A3B8"))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def apply_table_styles(ws, start_row, headers, data, col_widths, code_col=None):
        ws.row_dimensions[start_row].height = 26
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=start_row, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_header
            cell.border = border_cell

        for r_idx, row_values in enumerate(data, start=start_row+1):
            ws.row_dimensions[r_idx].height = 24
            is_alt = (r_idx % 2 == 1)
            row_fill = fill_alt if is_alt else fill_white
            for c_idx, val in enumerate(row_values, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = border_cell
                cell.fill = row_fill
                
                # Column formatting
                if code_col and c_idx == code_col:
                    cell.font = font_code
                    cell.alignment = align_left
                elif str(val) == "PK":
                    cell.font = font_pk
                    cell.alignment = align_center
                elif str(val) == "FK":
                    cell.font = font_fk
                    cell.alignment = align_center
                elif c_idx == 1:
                    cell.font = font_bold_body
                    cell.alignment = align_left if isinstance(val, str) and len(str(val)) > 8 else align_center
                else:
                    cell.font = font_body
                    cell.alignment = align_center if len(str(val)) < 15 and not isinstance(val, str) or str(val) in ["INTEGER", "TEXT", "REAL", "YES", "NO", "1:N", "1:1"] else align_left

        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    # =============================================================
    # SHEET 1: Overview
    # =============================================================
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ov.views.sheetView[0].showGridLines = True

    ws_ov.cell(row=1, column=1, value="Relational Database Design & Architecture").font = font_title
    ws_ov.cell(row=2, column=1, value="Project: Channel Wise Marketing Performance | SQLite Schema Specification").font = font_subtitle

    meta_items = [
        ("Database Engine", "SQLite 3.x (Embedded file: data/app.db)"),
        ("Architectural Schema", "Modified Star Schema (Normalized Dimension Masters + Granular Fact Tables + Joined Analytical View)"),
        ("Primary Target Stack", "Python 3.12, pandas, sqlite3, Streamlit Community Cloud"),
        ("Dimensional Masters (3)", "dim_channel (5 rows), dim_campaign (25-30 rows), dim_customer (150 rows)"),
        ("Fact Tables (2)", "fact_daily_campaign_performance (250-300 rows), fact_customer_touchpoint (450 rows)"),
        ("Joined Analysis View", "vw_marketing_analysis (38 unified columns supporting all OQs and KPIs)"),
        ("Referential Integrity", "Enforced via foreign keys ('PRAGMA foreign_keys = ON;') with RESTRICT / CASCADE rules"),
        ("Data Quality Controls", "CHECK constraints (Spend >= 0, Clicks <= Impressions), NOT NULL constraints, Unique Keys")
    ]
    for idx, (k, v) in enumerate(meta_items, start=4):
        ws_ov.cell(row=idx, column=1, value=k).font = font_bold_body
        ws_ov.cell(row=idx, column=1).fill = fill_card
        ws_ov.cell(row=idx, column=1).border = border_cell
        ws_ov.cell(row=idx, column=2, value=v).font = font_body
        ws_ov.cell(row=idx, column=2).fill = fill_white
        ws_ov.cell(row=idx, column=2).border = border_cell
        ws_ov.row_dimensions[idx].height = 22

    ws_ov.column_dimensions["A"].width = 28
    ws_ov.column_dimensions["B"].width = 80

    # =============================================================
    # SHEET 2: Tables_Summary
    # =============================================================
    ws_ts = wb.create_sheet(title="Tables_Summary")
    ws_ts.views.sheetView[0].showGridLines = True
    ws_ts.cell(row=1, column=1, value="Database Tables & Entities Architectural Summary").font = font_title
    ws_ts.cell(row=2, column=1, value="Grains, Primary/Foreign Keys, Column Counts & Planned Row Volumes").font = font_subtitle

    ts_headers = ["Table Name", "Table Type", "Table Grain (1 Row Represents)", "Primary Key", "Foreign Keys", "Column Count", "Planned Rows"]
    ts_data = [
        ("dim_channel", "Dimension (Master)", "One marketing acquisition channel", "channel_id", "None", 6, 5),
        ("dim_campaign", "Dimension (Master)", "One distinct marketing campaign", "campaign_id", "channel_id -> dim_channel", 8, "25 – 30"),
        ("dim_customer", "Dimension (Master)", "One unique acquired customer", "customer_id", "acquisition_channel_id -> dim_channel", 8, "120 – 150"),
        ("fact_daily_campaign_performance", "Fact (Performance)", "Performance telemetry for one campaign on one day", "fact_id", "campaign_id -> dim_campaign", 15, "250 – 300"),
        ("fact_customer_touchpoint", "Fact (Attribution)", "One touchpoint interaction in a customer purchase journey", "touchpoint_id", "customer_id -> dim_customer, channel_id -> dim_channel", 8, "400 – 500"),
        ("vw_marketing_analysis", "Analytical View", "Joined analytical record per campaign per day with computed KPIs", "fact_id", "N/A (Joined View)", 38, "250 – 300")
    ]
    apply_table_styles(ws_ts, start_row=4, headers=ts_headers, data=ts_data, col_widths={
        "A": 32, "B": 20, "C": 45, "D": 16, "E": 32, "F": 15, "G": 18
    })

    # =============================================================
    # DATA DICTIONARY SHEETS
    # =============================================================
    dict_headers = ["Column Name", "Data Type", "Key", "Nullable", "Default / Allowed Values", "Description & Business Semantics", "OQ / BR Supported"]

    # 3. dict_dim_channel
    ws_dch = wb.create_sheet(title="dict_dim_channel")
    ws_dch.views.sheetView[0].showGridLines = True
    ws_dch.cell(row=1, column=1, value="Data Dictionary: dim_channel (Marketing Channel Master)").font = font_title
    dch_data = [
        ("channel_id", "INTEGER", "PK", "NO", "1, 2, 3, 4, 5", "Surrogate unique identifier for the marketing channel", "BR-01, OQ-01"),
        ("channel_name", "TEXT", "UNIQUE", "NO", "'Google Ads', 'Meta Ads', 'Email Marketing', 'Organic Search', 'Influencer Marketing'", "Standardized corporate channel name", "BR-01, BR-11, OQ-01"),
        ("channel_category", "TEXT", "-", "NO", "'Paid Search', 'Paid Social', 'Owned / Direct', 'Organic SEO', 'Partnerships'", "Strategic classification of acquisition media", "BR-01, OQ-01"),
        ("default_cpc_benchmark", "REAL", "-", "NO", "> 0 (e.g., 15.0 - 45.0)", "Baseline Expected Cost Per Click in INR (₹)", "BR-02, OQ-07"),
        ("benchmark_roas", "REAL", "-", "NO", "2.00 (Standard)", "Minimum acceptable target Return on Ad Spend (2.0x)", "BR-02, OQ-07"),
        ("benchmark_cac", "REAL", "-", "NO", "1500.00 (Standard)", "Maximum acceptable target Customer Acquisition Cost in INR (₹1,500)", "BR-02, OQ-07")
    ]
    apply_table_styles(ws_dch, start_row=3, headers=dict_headers, data=dch_data, col_widths={"A": 22, "B": 12, "C": 10, "D": 10, "E": 28, "F": 45, "G": 20})

    # 4. dict_dim_campaign
    ws_dcmp = wb.create_sheet(title="dict_dim_campaign")
    ws_dcmp.views.sheetView[0].showGridLines = True
    ws_dcmp.cell(row=1, column=1, value="Data Dictionary: dim_campaign (Marketing Campaign Master)").font = font_title
    dcmp_data = [
        ("campaign_id", "INTEGER", "PK", "NO", "101, 102, ...", "Unique identifier for individual marketing campaigns", "BR-02, OQ-07"),
        ("channel_id", "INTEGER", "FK", "NO", "References dim_channel(channel_id)", "Foreign key mapping campaign to parent acquisition channel", "BR-01, OQ-01"),
        ("campaign_name", "TEXT", "UNIQUE", "NO", "Fictional campaign naming (e.g., 'Diwali_Festive_Search_2026')", "Human-readable campaign tracking label", "BR-02, BR-11, OQ-07"),
        ("campaign_type", "TEXT", "-", "NO", "'Search', 'Shopping', 'Retargeting', 'Prospecting', 'Newsletter', 'Sponsored'", "Tactical format and objective of the campaign", "BR-03, OQ-02"),
        ("target_roas", "REAL", "-", "NO", "1.80 – 3.50", "Specific ROAS performance benchmark configured for this campaign", "BR-02, OQ-07"),
        ("target_cac", "REAL", "-", "NO", "1000.0 – 2200.0", "Specific CAC budget ceiling configured for this campaign in INR", "BR-02, OQ-07"),
        ("start_date", "TEXT", "-", "NO", "ISO-8601 Date ('YYYY-MM-DD')", "Official campaign activation launch date", "BR-05, OQ-04"),
        ("status", "TEXT", "-", "NO", "'Active', 'Paused', 'Completed'", "Current operational campaign status", "BR-02, OQ-07")
    ]
    apply_table_styles(ws_dcmp, start_row=3, headers=dict_headers, data=dcmp_data, col_widths={"A": 22, "B": 12, "C": 10, "D": 10, "E": 30, "F": 45, "G": 20})

    # 5. dict_dim_customer
    ws_dcust = wb.create_sheet(title="dict_dim_customer")
    ws_dcust.views.sheetView[0].showGridLines = True
    ws_dcust.cell(row=1, column=1, value="Data Dictionary: dim_customer (Customer Demographics Master)").font = font_title
    dcust_data = [
        ("customer_id", "INTEGER", "PK", "NO", "1001, 1002, ...", "Unique identifier assigned upon initial customer acquisition", "BR-04, OQ-03"),
        ("customer_name", "TEXT", "-", "NO", "Synthetic Indian names (e.g., 'Aarav Sharma', 'Priya Patel')", "Fictional customer representative name (Zero PII)", "Compliance"),
        ("age_group", "TEXT", "-", "NO", "'18-24', '25-34', '35-44', '45+'", "Demographic age cohort classification", "BR-06, OQ-05"),
        ("gender", "TEXT", "-", "NO", "'Male', 'Female', 'Other'", "Customer reported gender demographic", "BR-06, OQ-05"),
        ("city_tier", "TEXT", "-", "NO", "'Tier 1', 'Tier 2', 'Tier 3'", "Geographic classification (Metro Tier 1 vs. Emerging Tiers)", "BR-06, OQ-05"),
        ("city_name", "TEXT", "-", "NO", "Indian cities: 'Bengaluru', 'Mumbai', 'Lucknow', 'Jaipur', etc.", "Customer resident city", "BR-06, OQ-05"),
        ("acquisition_channel_id", "INTEGER", "FK", "NO", "References dim_channel(channel_id)", "First-touch channel responsible for customer initial discovery", "BR-04, OQ-03"),
        ("signup_date", "TEXT", "-", "NO", "ISO-8601 Date ('YYYY-MM-DD')", "Date when the customer created their account", "BR-04, OQ-03")
    ]
    apply_table_styles(ws_dcust, start_row=3, headers=dict_headers, data=dcust_data, col_widths={"A": 24, "B": 12, "C": 10, "D": 10, "E": 30, "F": 45, "G": 20})

    # 6. dict_fact_daily_perf
    ws_dfact = wb.create_sheet(title="dict_fact_daily_perf")
    ws_dfact.views.sheetView[0].showGridLines = True
    ws_dfact.cell(row=1, column=1, value="Data Dictionary: fact_daily_campaign_performance (Core Daily Fact)").font = font_title
    dfact_data = [
        ("fact_id", "INTEGER", "PK", "NO", "1, 2, 3, ... (Auto-increment)", "Unique surrogate primary key for the performance fact row", "Schema PK"),
        ("date", "TEXT", "-", "NO", "ISO-8601 Date ('YYYY-MM-DD')", "Telemetry observation date (spanning 12 months)", "BR-05, OQ-04"),
        ("campaign_id", "INTEGER", "FK", "NO", "References dim_campaign(campaign_id)", "Foreign key linking daily metrics to campaign master", "BR-01, OQ-01"),
        ("day_of_week", "INTEGER", "-", "NO", "0 (Mon) to 6 (Sun)", "Numerical day of week for cyclical seasonality analysis", "BR-05, OQ-04"),
        ("month", "INTEGER", "-", "NO", "1 to 12", "Calendar month for quarterly and seasonal evaluations", "BR-05, OQ-04"),
        ("spend", "REAL", "-", "NO", ">= 0.0 (in INR ₹)", "Total ad capital spent by the campaign on this day", "BR-01, OQ-01"),
        ("impressions", "INTEGER", "-", "NO", ">= 0", "Total ad views delivered across ad inventory", "BR-03, OQ-02"),
        ("clicks", "INTEGER", "-", "NO", ">= 0 (CHECK: clicks <= impressions)", "Total user clicks on ad creatives", "BR-03, OQ-02"),
        ("conversions", "INTEGER", "-", "NO", ">= 0 (CHECK: conversions <= clicks)", "Total completed conversion transactions attributed to campaign", "BR-01, OQ-01"),
        ("revenue", "REAL", "-", "NO", ">= 0.0 (in INR ₹)", "Gross revenue generated from conversions on this day", "BR-01, OQ-01"),
        ("utm_source", "TEXT", "-", "YES", "'google', 'meta', 'klaviyo', 'organic', 'instagram'", "UTM tracking source query parameter", "BR-10, OQ-10"),
        ("utm_medium", "TEXT", "-", "YES", "'cpc', 'paid_social', 'email', 'organic', 'influencer'", "UTM tracking medium parameter", "BR-10, OQ-10"),
        ("utm_campaign", "TEXT", "-", "YES", "Campaign slug string matching campaign_name", "UTM tracking campaign parameter", "BR-10, OQ-10"),
        ("is_outlier", "INTEGER", "-", "NO", "0 (Normal) or 1 (Statistical Anomaly)", "Flag indicating extreme spend/revenue anomaly (Z > 3)", "BR-10, OQ-10"),
        ("outlier_reason", "TEXT", "-", "YES", "'Spend Spike', 'Zero Impression Glitch', 'Revenue Outlier'", "Explanatory diagnostics for data quality audits", "BR-10, OQ-10")
    ]
    apply_table_styles(ws_dfact, start_row=3, headers=dict_headers, data=dfact_data, col_widths={"A": 22, "B": 12, "C": 10, "D": 10, "E": 28, "F": 45, "G": 20})

    # 7. dict_fact_touchpoint
    ws_dtouch = wb.create_sheet(title="dict_fact_touchpoint")
    ws_dtouch.views.sheetView[0].showGridLines = True
    ws_dtouch.cell(row=1, column=1, value="Data Dictionary: fact_customer_touchpoint (Multi-Touch Fact)").font = font_title
    dtouch_data = [
        ("touchpoint_id", "INTEGER", "PK", "NO", "1, 2, 3, ... (Auto-increment)", "Unique surrogate key for individual customer journey interactions", "Schema PK"),
        ("journey_id", "INTEGER", "-", "NO", "501, 502, ...", "Identifier grouping touchpoints belonging to the same purchase journey", "BR-08, OQ-08"),
        ("customer_id", "INTEGER", "FK", "NO", "References dim_customer(customer_id)", "Foreign key identifying the interacting customer", "BR-04, OQ-03"),
        ("channel_id", "INTEGER", "FK", "NO", "References dim_channel(channel_id)", "Marketing channel engaged during this specific touchpoint", "BR-08, OQ-08"),
        ("touchpoint_sequence", "INTEGER", "-", "NO", "1 (First Touch), 2, 3, ... N (Last Touch)", "Chronological sequence index of the touchpoint within the journey", "BR-08, OQ-08"),
        ("touchpoint_timestamp", "TEXT", "-", "NO", "ISO-8601 Timestamp ('YYYY-MM-DD HH:MM:SS')", "Exact date and time of customer interaction", "BR-08, OQ-08"),
        ("is_conversion", "INTEGER", "-", "NO", "0 (Touchpoint Only), 1 (Completed Order)", "Flag indicating whether this touchpoint finalized a purchase", "BR-08, OQ-08"),
        ("order_value", "REAL", "-", "NO", ">= 0.0 (in INR ₹)", "Monetary value of order if is_conversion=1; 0.0 otherwise", "BR-04, BR-08, OQ-03")
    ]
    apply_table_styles(ws_dtouch, start_row=3, headers=dict_headers, data=dtouch_data, col_widths={"A": 22, "B": 12, "C": 10, "D": 10, "E": 30, "F": 45, "G": 20})

    # 8. dict_view_marketing_analysis
    ws_dview = wb.create_sheet(title="dict_view_marketing_analysis")
    ws_dview.views.sheetView[0].showGridLines = True
    ws_dview.cell(row=1, column=1, value="Data Dictionary: vw_marketing_analysis (Joined Analytical View - 38 Columns)").font = font_title
    
    view_data = [
        ("fact_id", "INTEGER", "PK", "NO", "From fact table", "Primary key identifying the observation record", "Schema PK"),
        ("date", "TEXT", "-", "NO", "'YYYY-MM-DD'", "Observation date", "BR-05, OQ-04"),
        ("day_of_week", "INTEGER", "-", "NO", "0 to 6", "Day of week integer index (0=Monday)", "BR-05, OQ-04"),
        ("day_name", "TEXT", "-", "NO", "'Monday' - 'Sunday'", "Human-readable day name for seasonal grouping", "BR-05, OQ-04"),
        ("month", "INTEGER", "-", "NO", "1 to 12", "Calendar month integer index", "BR-05, OQ-04"),
        ("month_name", "TEXT", "-", "NO", "'January' - 'December'", "Calendar month name for trend visualization", "BR-05, OQ-04"),
        ("quarter", "TEXT", "-", "NO", "'Q1', 'Q2', 'Q3', 'Q4'", "Financial quarter classification", "BR-05, OQ-04"),
        ("campaign_id", "INTEGER", "FK", "NO", "101, 102...", "Campaign unique identifier", "BR-02, OQ-07"),
        ("campaign_name", "TEXT", "-", "NO", "Campaign name", "Human-readable campaign title", "BR-02, OQ-07"),
        ("campaign_type", "TEXT", "-", "NO", "Search, Retargeting...", "Campaign operational format", "BR-03, OQ-02"),
        ("campaign_status", "TEXT", "-", "NO", "Active, Paused", "Operational campaign state", "BR-02, OQ-07"),
        ("channel_id", "INTEGER", "FK", "NO", "1 to 5", "Channel identifier", "BR-01, OQ-01"),
        ("channel_name", "TEXT", "-", "NO", "Google Ads, Meta...", "Channel name", "BR-01, OQ-01"),
        ("channel_category", "TEXT", "-", "NO", "Paid Search, Paid Social...", "Strategic media category", "BR-01, OQ-01"),
        ("spend", "REAL", "-", "NO", ">= 0.0", "Marketing expenditure in INR", "BR-01, OQ-01"),
        ("impressions", "INTEGER", "-", "NO", ">= 0", "Ad impressions delivered", "BR-03, OQ-02"),
        ("clicks", "INTEGER", "-", "NO", ">= 0", "User clicks recorded", "BR-03, OQ-02"),
        ("conversions", "INTEGER", "-", "NO", ">= 0", "Transactions completed", "BR-01, OQ-01"),
        ("revenue", "REAL", "-", "NO", ">= 0.0", "Attributed gross revenue in INR", "BR-01, OQ-01"),
        ("ctr", "REAL", "-", "NO", "clicks / impressions", "Click-Through Rate decimal ratio", "BR-03, OQ-02"),
        ("cpc", "REAL", "-", "NO", "spend / clicks", "Cost Per Click in INR", "BR-03, OQ-02"),
        ("cvr", "REAL", "-", "NO", "conversions / clicks", "Conversion Rate decimal ratio", "BR-03, OQ-02"),
        ("cpa_cac", "REAL", "-", "NO", "spend / conversions", "Customer Acquisition Cost in INR", "BR-01, OQ-01"),
        ("roas", "REAL", "-", "NO", "revenue / spend", "Return on Ad Spend ratio (e.g. 2.45x)", "BR-01, OQ-01"),
        ("profit", "REAL", "-", "NO", "revenue - spend", "Net marketing contribution profit in INR", "BR-01, OQ-01"),
        ("target_roas", "REAL", "-", "NO", "From campaign master", "Target ROAS benchmark", "BR-02, OQ-07"),
        ("target_cac", "REAL", "-", "NO", "From campaign master", "Target CAC threshold", "BR-02, OQ-07"),
        ("roas_gap", "REAL", "-", "NO", "roas - target_roas", "Variance against target ROAS (+/-)", "BR-02, OQ-07"),
        ("cac_gap", "REAL", "-", "NO", "cpa_cac - target_cac", "Variance against target CAC (+/-)", "BR-02, OQ-07"),
        ("is_underperforming", "INTEGER", "-", "NO", "0 or 1", "Flag: 1 if roas < target_roas OR cac > target_cac", "BR-02, OQ-07"),
        ("benchmark_roas", "REAL", "-", "NO", "From channel master", "Default channel ROAS threshold", "BR-02, OQ-07"),
        ("benchmark_cac", "REAL", "-", "NO", "From channel master", "Default channel CAC ceiling", "BR-02, OQ-07"),
        ("utm_source", "TEXT", "-", "YES", "UTM source string", "Tracking parameter for source audit", "BR-10, OQ-10"),
        ("utm_medium", "TEXT", "-", "YES", "UTM medium string", "Tracking parameter for medium audit", "BR-10, OQ-10"),
        ("utm_campaign", "TEXT", "-", "YES", "UTM campaign string", "Tracking parameter for campaign audit", "BR-10, OQ-10"),
        ("is_outlier", "INTEGER", "-", "NO", "0 or 1", "Flag: 1 if spend/revenue violates 3-sigma bounds", "BR-10, OQ-10"),
        ("outlier_reason", "TEXT", "-", "YES", "Diagnostic string", "Description of data anomaly", "BR-10, OQ-10"),
        ("aov", "REAL", "-", "NO", "revenue / conversions", "Average Order Value in INR", "BR-04, OQ-03")
    ]
    apply_table_styles(ws_dview, start_row=3, headers=dict_headers, data=view_data, col_widths={"A": 22, "B": 12, "C": 10, "D": 10, "E": 28, "F": 45, "G": 20})

    # =============================================================
    # SHEET 9: Relationships
    # =============================================================
    ws_rel = wb.create_sheet(title="Relationships")
    ws_rel.views.sheetView[0].showGridLines = True
    ws_rel.cell(row=1, column=1, value="Entity Relational Integrity & Foreign Key Mappings").font = font_title
    
    rel_headers = ["Relationship ID", "Parent Table (Primary Key)", "Child Table (Foreign Key)", "Join Key Column", "Cardinality", "Business Rule & Constraint Behavior"]
    rel_data = [
        ("REL-01", "dim_channel (channel_id)", "dim_campaign (channel_id)", "channel_id", "1 : N (One-to-Many)", "One channel funds many marketing campaigns. ON DELETE RESTRICT."),
        ("REL-02", "dim_channel (channel_id)", "dim_customer (acquisition_channel_id)", "acquisition_channel_id", "1 : N (One-to-Many)", "One channel is the first-touch discovery source for many customers. ON DELETE RESTRICT."),
        ("REL-03", "dim_campaign (campaign_id)", "fact_daily_campaign_performance (campaign_id)", "campaign_id", "1 : N (One-to-Many)", "One campaign generates daily performance telemetry across multiple dates. ON DELETE RESTRICT."),
        ("REL-04", "dim_customer (customer_id)", "fact_customer_touchpoint (customer_id)", "customer_id", "1 : N (One-to-Many)", "One customer executes multiple sequential touchpoints across their journey. ON DELETE CASCADE."),
        ("REL-05", "dim_channel (channel_id)", "fact_customer_touchpoint (channel_id)", "channel_id", "1 : N (One-to-Many)", "One channel is visited across multiple touchpoints in different journeys. ON DELETE RESTRICT.")
    ]
    apply_table_styles(ws_rel, start_row=3, headers=rel_headers, data=rel_data, col_widths={"A": 16, "B": 28, "C": 36, "D": 22, "E": 18, "F": 50})

    # =============================================================
    # SHEET 10: KPIs_Measures
    # =============================================================
    ws_kpi = wb.create_sheet(title="KPIs_Measures")
    ws_kpi.views.sheetView[0].showGridLines = True
    ws_kpi.cell(row=1, column=1, value="Key Performance Indicators (KPIs) & Mathematical Formulations").font = font_title
    
    kpi_headers = ["Measure ID", "KPI / Metric Name", "Mathematical / SQL Formulation", "Unit / Format", "Target Threshold", "Supported OQ"]
    kpi_sheet_data = [
        ("M-01", "Return on Ad Spend (ROAS)", "SUM(revenue) / NULLIF(SUM(spend), 0)", "Ratio (x.xx X)", ">= 2.00x", "OQ-01"),
        ("M-02", "Customer Acquisition Cost (CAC)", "SUM(spend) / NULLIF(SUM(conversions), 0)", "Currency (₹)", "<= ₹1,500.00", "OQ-01"),
        ("M-03", "Click-Through Rate (CTR)", "(CAST(SUM(clicks) AS REAL) / NULLIF(SUM(impressions), 0)) * 100", "Percentage (%)", "1.50% – 4.50%", "OQ-02"),
        ("M-04", "Conversion Rate (CVR)", "(CAST(SUM(conversions) AS REAL) / NULLIF(SUM(clicks), 0)) * 100", "Percentage (%)", "2.00% – 4.00%", "OQ-02"),
        ("M-05", "Cost Per Click (CPC)", "SUM(spend) / NULLIF(SUM(clicks), 0)", "Currency (₹)", "<= ₹35.00", "OQ-02"),
        ("M-06", "Average Order Value (AOV)", "SUM(revenue) / NULLIF(SUM(conversions), 0)", "Currency (₹)", ">= ₹1,800.00", "OQ-03"),
        ("M-07", "Customer Lifetime Value (LTV)", "AOV * Purchase_Frequency (over customer lifetime)", "Currency (₹)", ">= ₹4,500.00", "OQ-03"),
        ("M-08", "LTV : CAC Ratio", "LTV / NULLIF(CAC, 0)", "Ratio (:1)", ">= 3.0 : 1", "OQ-03"),
        ("M-09", "Net Marketing Contribution (Profit)", "SUM(revenue) - SUM(spend)", "Currency (₹)", "> 0 (Positive)", "OQ-01"),
        ("M-10", "Underperforming Ad Flag", "CASE WHEN roas < target_roas OR cpa_cac > target_cac THEN 1 ELSE 0 END", "Binary Flag", "0 (Zero Waste)", "OQ-07"),
        ("M-11", "Two-Sample Conversion Z-Score", "(p1 - p2) / SQRT(p_pool * (1 - p_pool) * (1/n1 + 1/n2))", "Z-Statistic", "p-value < 0.05", "OQ-06"),
        ("M-12", "Data Completeness Ratio", "(1.0 - (CAST(null_count AS REAL) / total_count)) * 100", "Percentage (%)", ">= 98.0%", "OQ-10")
    ]
    apply_table_styles(ws_kpi, start_row=3, headers=kpi_headers, data=kpi_sheet_data, col_widths={"A": 14, "B": 28, "C": 52, "D": 16, "E": 18, "F": 15})

    # =============================================================
    # SHEET 11: SQLite_DDL
    # =============================================================
    ws_ddl = wb.create_sheet(title="SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    ws_ddl.cell(row=1, column=1, value="Production SQLite DDL Script (schema.sql)").font = font_title
    ws_ddl.cell(row=2, column=1, value="Table DDL, Primary/Foreign Keys, Check Constraints, Indexes & Joined View").font = font_subtitle

    ddl_script = (
        "-- =============================================================\n"
        "-- CHANNEL WISE MARKETING PERFORMANCE - PRODUCTION DDL\n"
        "-- Target Engine: SQLite 3.x (data/app.db)\n"
        "-- =============================================================\n\n"
        "PRAGMA foreign_keys = ON;\n\n"
        "-- 1. MARKETING CHANNEL MASTER\n"
        "CREATE TABLE IF NOT EXISTS dim_channel (\n"
        "    channel_id INTEGER PRIMARY KEY,\n"
        "    channel_name TEXT NOT NULL UNIQUE,\n"
        "    channel_category TEXT NOT NULL,\n"
        "    default_cpc_benchmark REAL NOT NULL CHECK (default_cpc_benchmark > 0),\n"
        "    benchmark_roas REAL NOT NULL DEFAULT 2.0 CHECK (benchmark_roas > 0),\n"
        "    benchmark_cac REAL NOT NULL DEFAULT 1500.0 CHECK (benchmark_cac > 0)\n"
        ");\n\n"
        "-- 2. CAMPAIGN MASTER\n"
        "CREATE TABLE IF NOT EXISTS dim_campaign (\n"
        "    campaign_id INTEGER PRIMARY KEY,\n"
        "    channel_id INTEGER NOT NULL,\n"
        "    campaign_name TEXT NOT NULL UNIQUE,\n"
        "    campaign_type TEXT NOT NULL CHECK (campaign_type IN ('Search', 'Shopping', 'Retargeting', 'Prospecting', 'Newsletter', 'Sponsored')),\n"
        "    target_roas REAL NOT NULL CHECK (target_roas > 0),\n"
        "    target_cac REAL NOT NULL CHECK (target_cac > 0),\n"
        "    start_date TEXT NOT NULL,\n"
        "    status TEXT NOT NULL DEFAULT 'Active' CHECK (status IN ('Active', 'Paused', 'Completed')),\n"
        "    FOREIGN KEY (channel_id) REFERENCES dim_channel(channel_id) ON DELETE RESTRICT\n"
        ");\n\n"
        "-- 3. CUSTOMER DEMOGRAPHICS MASTER\n"
        "CREATE TABLE IF NOT EXISTS dim_customer (\n"
        "    customer_id INTEGER PRIMARY KEY,\n"
        "    customer_name TEXT NOT NULL,\n"
        "    age_group TEXT NOT NULL CHECK (age_group IN ('18-24', '25-34', '35-44', '45+')),\n"
        "    gender TEXT NOT NULL CHECK (gender IN ('Male', 'Female', 'Other')),\n"
        "    city_tier TEXT NOT NULL CHECK (city_tier IN ('Tier 1', 'Tier 2', 'Tier 3')),\n"
        "    city_name TEXT NOT NULL,\n"
        "    acquisition_channel_id INTEGER NOT NULL,\n"
        "    signup_date TEXT NOT NULL,\n"
        "    FOREIGN KEY (acquisition_channel_id) REFERENCES dim_channel(channel_id) ON DELETE RESTRICT\n"
        ");\n\n"
        "-- 4. DAILY CAMPAIGN PERFORMANCE FACT TABLE\n"
        "CREATE TABLE IF NOT EXISTS fact_daily_campaign_performance (\n"
        "    fact_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    date TEXT NOT NULL,\n"
        "    campaign_id INTEGER NOT NULL,\n"
        "    day_of_week INTEGER NOT NULL CHECK (day_of_week BETWEEN 0 AND 6),\n"
        "    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),\n"
        "    spend REAL NOT NULL CHECK (spend >= 0),\n"
        "    impressions INTEGER NOT NULL CHECK (impressions >= 0),\n"
        "    clicks INTEGER NOT NULL CHECK (clicks >= 0 AND clicks <= impressions),\n"
        "    conversions INTEGER NOT NULL CHECK (conversions >= 0 AND conversions <= clicks),\n"
        "    revenue REAL NOT NULL CHECK (revenue >= 0),\n"
        "    utm_source TEXT,\n"
        "    utm_medium TEXT,\n"
        "    utm_campaign TEXT,\n"
        "    is_outlier INTEGER NOT NULL DEFAULT 0 CHECK (is_outlier IN (0, 1)),\n"
        "    outlier_reason TEXT,\n"
        "    FOREIGN KEY (campaign_id) REFERENCES dim_campaign(campaign_id) ON DELETE RESTRICT\n"
        ");\n\n"
        "-- 5. CUSTOMER TOUCHPOINTS FACT TABLE (ATTRIBUTION & ORDERS)\n"
        "CREATE TABLE IF NOT EXISTS fact_customer_touchpoint (\n"
        "    touchpoint_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    journey_id INTEGER NOT NULL,\n"
        "    customer_id INTEGER NOT NULL,\n"
        "    channel_id INTEGER NOT NULL,\n"
        "    touchpoint_sequence INTEGER NOT NULL CHECK (touchpoint_sequence >= 1),\n"
        "    touchpoint_timestamp TEXT NOT NULL,\n"
        "    is_conversion INTEGER NOT NULL CHECK (is_conversion IN (0, 1)),\n"
        "    order_value REAL NOT NULL DEFAULT 0.0 CHECK (order_value >= 0),\n"
        "    FOREIGN KEY (customer_id) REFERENCES dim_customer(customer_id) ON DELETE CASCADE,\n"
        "    FOREIGN KEY (channel_id) REFERENCES dim_channel(channel_id) ON DELETE RESTRICT\n"
        ");\n\n"
        "-- 6. PERFORMANCE B-TREE INDEXES\n"
        "CREATE INDEX IF NOT EXISTS idx_fact_perf_date ON fact_daily_campaign_performance(date);\n"
        "CREATE INDEX IF NOT EXISTS idx_fact_perf_campaign ON fact_daily_campaign_performance(campaign_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_fact_touchpoint_cust ON fact_customer_touchpoint(customer_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_fact_touchpoint_journey ON fact_customer_touchpoint(journey_id);\n\n"
        "-- 7. JOINED ANALYSIS VIEW (38 COLUMNS)\n"
        "CREATE VIEW IF NOT EXISTS vw_marketing_analysis AS\n"
        "SELECT \n"
        "    f.fact_id,\n"
        "    f.date,\n"
        "    f.day_of_week,\n"
        "    CASE f.day_of_week \n"
        "        WHEN 0 THEN 'Monday' WHEN 1 THEN 'Tuesday' WHEN 2 THEN 'Wednesday'\n"
        "        WHEN 3 THEN 'Thursday' WHEN 4 THEN 'Friday' WHEN 5 THEN 'Saturday' ELSE 'Sunday' \n"
        "    END AS day_name,\n"
        "    f.month,\n"
        "    CASE f.month \n"
        "        WHEN 1 THEN 'January' WHEN 2 THEN 'February' WHEN 3 THEN 'March' WHEN 4 THEN 'April'\n"
        "        WHEN 5 THEN 'May' WHEN 6 THEN 'June' WHEN 7 THEN 'July' WHEN 8 THEN 'August'\n"
        "        WHEN 9 THEN 'September' WHEN 10 THEN 'October' WHEN 11 THEN 'November' ELSE 'December' \n"
        "    END AS month_name,\n"
        "    CASE \n"
        "        WHEN f.month IN (1,2,3) THEN 'Q1' WHEN f.month IN (4,5,6) THEN 'Q2'\n"
        "        WHEN f.month IN (7,8,9) THEN 'Q3' ELSE 'Q4' \n"
        "    END AS quarter,\n"
        "    c.campaign_id,\n"
        "    c.campaign_name,\n"
        "    c.campaign_type,\n"
        "    c.status AS campaign_status,\n"
        "    ch.channel_id,\n"
        "    ch.channel_name,\n"
        "    ch.channel_category,\n"
        "    f.spend,\n"
        "    f.impressions,\n"
        "    f.clicks,\n"
        "    f.conversions,\n"
        "    f.revenue,\n"
        "    ROUND(CAST(f.clicks AS REAL) / NULLIF(f.impressions, 0), 4) AS ctr,\n"
        "    ROUND(f.spend / NULLIF(f.clicks, 0), 2) AS cpc,\n"
        "    ROUND(CAST(f.conversions AS REAL) / NULLIF(f.clicks, 0), 4) AS cvr,\n"
        "    ROUND(f.spend / NULLIF(f.conversions, 0), 2) AS cpa_cac,\n"
        "    ROUND(f.revenue / NULLIF(f.spend, 0), 2) AS roas,\n"
        "    ROUND(f.revenue - f.spend, 2) AS profit,\n"
        "    c.target_roas,\n"
        "    c.target_cac,\n"
        "    ROUND((f.revenue / NULLIF(f.spend, 0)) - c.target_roas, 2) AS roas_gap,\n"
        "    ROUND((f.spend / NULLIF(f.conversions, 0)) - c.target_cac, 2) AS cac_gap,\n"
        "    CASE \n"
        "        WHEN (f.revenue / NULLIF(f.spend, 0)) < c.target_roas \n"
        "          OR (f.spend / NULLIF(f.conversions, 0)) > c.target_cac THEN 1 \n"
        "        ELSE 0 \n"
        "    END AS is_underperforming,\n"
        "    ch.benchmark_roas,\n"
        "    ch.benchmark_cac,\n"
        "    f.utm_source,\n"
        "    f.utm_medium,\n"
        "    f.utm_campaign,\n"
        "    f.is_outlier,\n"
        "    f.outlier_reason,\n"
        "    ROUND(f.revenue / NULLIF(f.conversions, 0), 2) AS aov\n"
        "FROM fact_daily_campaign_performance f\n"
        "JOIN dim_campaign c ON f.campaign_id = c.campaign_id\n"
        "JOIN dim_channel ch ON c.channel_id = ch.channel_id;\n"
    )

    ws_ddl.cell(row=4, column=1, value=ddl_script).font = font_code
    ws_ddl.cell(row=4, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws_ddl.row_dimensions[4].height = 650
    ws_ddl.column_dimensions["A"].width = 110

    # Save Database Design Excel
    out_db_path = os.path.abspath("docs/08_Database_Design.xlsx")
    wb.save(out_db_path)
    print(f"Saved: {out_db_path}")

    # =============================================================
    # UPDATE RTM EXCEL (docs/04_Traceability_Matrix.xlsx)
    # =============================================================
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    if os.path.exists(rtm_path):
        wb_rtm = openpyxl.load_workbook(rtm_path)
        ws_r = wb_rtm["Traceability_Matrix"]
        
        # Mapping of (OQ, HMW, BR) to DB table.column
        db_mappings = {
            ("OQ-01", "HMW-01", "BR-01"): "fact_daily_campaign_performance.spend, revenue, conversions; vw_marketing_analysis.roas, cpa_cac",
            ("OQ-01", "HMW-01", "BR-11"): "fact_daily_campaign_performance.date; dim_channel.channel_name; dim_campaign.campaign_type",
            ("OQ-01", "HMW-01", "BR-12"): "vw_marketing_analysis.* (38 columns)",
            ("OQ-07", "HMW-02", "BR-02"): "dim_campaign.target_roas, target_cac; vw_marketing_analysis.is_underperforming, roas_gap, cac_gap",
            ("OQ-02", "HMW-03", "BR-03"): "fact_daily_campaign_performance.impressions, clicks, conversions; vw_marketing_analysis.ctr, cvr",
            ("OQ-02", "HMW-03", "BR-11"): "fact_daily_campaign_performance.date; dim_channel.channel_name",
            ("OQ-03", "HMW-04", "BR-04"): "dim_customer.customer_id, acquisition_channel_id; fact_customer_touchpoint.order_value, is_conversion",
            ("OQ-04", "HMW-05", "BR-05"): "fact_daily_campaign_performance.date, day_of_week, month; vw_marketing_analysis.day_name, month_name",
            ("OQ-05", "HMW-06", "BR-06"): "dim_customer.age_group, gender, city_tier; fact_customer_touchpoint.channel_id, order_value",
            ("OQ-06", "HMW-07", "BR-07"): "dim_channel.channel_name; fact_daily_campaign_performance.clicks, conversions; vw_marketing_analysis.cvr",
            ("OQ-08", "HMW-08", "BR-08"): "fact_customer_touchpoint.journey_id, touchpoint_sequence, channel_id, is_conversion, order_value",
            ("OQ-09", "HMW-09", "BR-09"): "fact_daily_campaign_performance.spend, revenue; dim_channel.channel_name; vw_marketing_analysis.roas",
            ("OQ-10", "HMW-10", "BR-10"): "fact_daily_campaign_performance.utm_source, utm_medium, utm_campaign, is_outlier, outlier_reason",
            ("OQ-10", "HMW-10", "BR-12"): "fact_daily_campaign_performance.is_outlier, utm_source; vw_marketing_analysis.*"
        }

        # Column 9 is "DB table.column"
        for row_idx in range(2, ws_r.max_row + 1):
            oq_val = ws_r.cell(row=row_idx, column=1).value
            hmw_val = ws_r.cell(row=row_idx, column=2).value
            br_val = ws_r.cell(row=row_idx, column=3).value
            
            key = (oq_val, hmw_val, br_val)
            if key in db_mappings:
                cell_db = ws_r.cell(row=row_idx, column=9)
                cell_db.value = db_mappings[key]
                cell_db.font = Font(name="Segoe UI", size=9, color="1E3A8A")
                cell_db.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        wb_rtm.save(rtm_path)
        print(f"Updated RTM with DB mappings: {rtm_path}")

if __name__ == "__main__":
    build_database_design_excel()
