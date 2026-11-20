"""
UI Component Library for Channel Wise Marketing Performance Analysis
Provides Tailwind CSS-styled headers, KPI cards, and insight banners
Rendered via st.components.v1.html with isolated Tailwind CDN and dynamic sizing
"""

try:
    import streamlit as st
    import streamlit.components.v1 as components
    HAS_STREAMLIT = True
except ImportError:
    HAS_STREAMLIT = False

def inject_global_css():
    """
    Injects custom CSS into the Streamlit app to prevent text overlapping,
    improve metric readability, and make layouts responsive on all screens.
    """
    if not HAS_STREAMLIT:
        return
    st.markdown("""
        <style>
        /* Prevent metric label and value overlapping */
        [data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            padding: 12px 16px;
            border-radius: 10px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            min-height: 95px;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.82rem !important;
            font-weight: 600 !important;
            color: #475569 !important;
            white-space: normal !important;
            word-break: break-word !important;
            line-height: 1.3 !important;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.35rem !important;
            font-weight: 700 !important;
            color: #0F172A !important;
            white-space: nowrap !important;
            line-height: 1.2 !important;
        }
        [data-testid="stMetricDelta"] {
            font-size: 0.8rem !important;
            white-space: nowrap !important;
        }
        /* Clean table and data presentation */
        .stDataFrame {
            width: 100% !important;
            margin-top: 8px !important;
        }
        /* Remove nested iframe margins */
        iframe {
            display: block;
            margin: 0 auto;
        }
        </style>
    """, unsafe_allow_html=True)

def render_tailwind_header(title: str, subtitle: str, badge_text: str = "MCA Capstone | ABES"):
    """
    Renders an executive, modern dashboard header using Tailwind CSS via CDN.
    Calculates height dynamically to ensure long titles/subtitles never overlap or get clipped.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
      <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white rounded-xl py-4 px-6 shadow-md border border-slate-700">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div class="flex-1">
            <div class="flex items-center gap-2 mb-1">
              <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500 bg-opacity-30 text-blue-200 border border-blue-400 border-opacity-30">
                {badge_text}
              </span>
              <span class="text-xs text-slate-400">Live Telemetry Pipeline</span>
            </div>
            <h1 class="text-xl md:text-2xl font-bold tracking-tight text-white mb-1 leading-snug">{title}</h1>
            <p class="text-xs md:text-sm text-slate-300 max-w-3xl leading-relaxed">{subtitle}</p>
          </div>
          <div class="flex items-center gap-3 shrink-0">
            <div class="bg-slate-800 bg-opacity-70 px-4 py-2 rounded-lg border border-slate-700 text-right">
              <div class="text-xs text-slate-400">System Status</div>
              <div class="text-sm font-semibold text-emerald-400 flex items-center gap-1.5 justify-end">
                <span class="h-2 w-2 rounded-full bg-emerald-400 animate-pulse"></span>
                Database Connected
              </div>
            </div>
          </div>
        </div>
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        # Dynamic height calculation to prevent vertical truncation or overlapping
        dynamic_height = 160 if len(title) > 40 else 140
        components.html(html_code, height=dynamic_height)
    return html_code

def render_kpi_card(title: str, value: str, delta: str = None, delta_type: str = "positive", subtitle: str = ""):
    """
    Renders a single Tailwind CSS metric card.
    delta_type: 'positive' (green), 'negative' (red), 'warning' (amber), 'neutral' (slate)
    """
    badge_colors = {
        "positive": "bg-emerald-50 text-emerald-700 border-emerald-200",
        "negative": "bg-rose-50 text-rose-700 border-rose-200",
        "warning": "bg-amber-50 text-amber-700 border-amber-200",
        "neutral": "bg-slate-100 text-slate-700 border-slate-200"
    }
    badge_cls = badge_colors.get(delta_type, badge_colors["neutral"])

    delta_html = ""
    if delta:
        delta_html = f"""
        <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold border {badge_cls} shrink-0">
          {delta}
        </span>
        """

    html_code = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-1 font-sans antialiased">
      <div class="bg-white rounded-xl p-4 shadow-sm border border-slate-200 hover:shadow-md transition-shadow">
        <div class="flex items-center justify-between gap-2 mb-2">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500 truncate">{title}</span>
          {delta_html}
        </div>
        <div class="text-2xl font-bold text-slate-900 tracking-tight leading-tight">{value}</div>
        <div class="text-xs text-slate-400 mt-1 truncate">{subtitle}</div>
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        components.html(html_code, height=120)
    return html_code

def render_kpi_row(cards: list):
    """
    Renders 4 KPI cards side by side in a single responsive container with ample padding
    to guarantee no value or label overlapping.
    """
    cards_html = ""
    for c in cards:
        badge_colors = {
            "positive": "bg-emerald-50 text-emerald-700 border-emerald-200",
            "negative": "bg-rose-50 text-rose-700 border-rose-200",
            "warning": "bg-amber-50 text-amber-700 border-amber-200",
            "neutral": "bg-slate-100 text-slate-700 border-slate-200"
        }
        badge_cls = badge_colors.get(c.get("delta_type", "neutral"), badge_colors["neutral"])
        delta_tag = ""
        if c.get("delta"):
            delta_tag = f"""<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold border {badge_cls} shrink-0">{c['delta']}</span>"""
        
        cards_html += f"""
        <div class="bg-white rounded-xl p-3.5 shadow-sm border border-slate-200 hover:border-blue-400 transition-colors flex flex-col justify-between">
          <div class="flex items-center justify-between gap-2 mb-1.5">
            <span class="text-xs font-semibold uppercase tracking-wider text-slate-500 truncate">{c['title']}</span>
            {delta_tag}
          </div>
          <div class="text-xl md:text-2xl font-bold text-slate-900 tracking-tight leading-tight my-0.5">{c['value']}</div>
          <div class="text-xs text-slate-400 truncate">{c.get('subtitle', '')}</div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-1 font-sans antialiased">
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        {cards_html}
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        components.html(full_html, height=125)
    return full_html

def render_insight_box(title: str, text: str, alert_type: str = "info"):
    """
    Renders a styled callout box for executive takeaways or statistical conclusions.
    Dynamically adjusts height based on text length so long takeaways never get cut off.
    """
    styles = {
        "info": ("border-indigo-500 bg-indigo-50 text-indigo-900", "text-indigo-700"),
        "success": ("border-emerald-500 bg-emerald-50 text-emerald-900", "text-emerald-700"),
        "warning": ("border-amber-500 bg-amber-50 text-amber-900", "text-amber-700"),
        "danger": ("border-rose-500 bg-rose-50 text-rose-900", "text-rose-700")
    }
    box_cls, title_cls = styles.get(alert_type, styles["info"])

    html_code = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-1 font-sans antialiased">
      <div class="border-l-4 {box_cls} py-3 px-4 rounded-r-xl shadow-sm mb-1">
        <h4 class="text-sm font-bold {title_cls} mb-1 leading-snug">{title}</h4>
        <p class="text-xs text-slate-700 leading-relaxed m-0">{text}</p>
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        # Dynamic height calculation: base height + extra lines based on character length
        text_len = len(text) + len(title)
        calc_height = 105 if text_len < 160 else (125 if text_len < 280 else 145)
        components.html(html_code, height=calc_height)
    return html_code

def apply_plotly_theme(fig):
    """
    Applies unified corporate styling to a Plotly figure:
    Navy blue accent, clean background, readable typography, subtle gridlines.
    Ensures legends and titles never collide or overlap.
    """
    if fig is None:
        return fig
    try:
        fig.update_layout(
            font=dict(family="Inter, Calibri, sans-serif", size=11, color="#0F172A"),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(248,250,252,0.6)",
            margin=dict(l=45, r=25, t=35, b=45),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.28,
                xanchor="center",
                x=0.5,
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#E2E8F0",
                borderwidth=1
            ),
            xaxis=dict(
                showgrid=True,
                gridcolor="#E2E8F0",
                linecolor="#CBD5E1",
                zeroline=False
            ),
            yaxis=dict(
                showgrid=True,
                gridcolor="#E2E8F0",
                linecolor="#CBD5E1",
                zeroline=False
            )
        )
    except Exception:
        pass
    return fig
