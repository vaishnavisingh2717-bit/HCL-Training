"""
UI Component Library for Channel Wise Marketing Performance Analysis
Provides Tailwind CSS-styled headers, KPI cards, and insight banners
Rendered via st.components.v1.html with isolated Tailwind CDN
"""

try:
    import streamlit as st
    import streamlit.components.v1 as components
    HAS_STREAMLIT = True
except ImportError:
    HAS_STREAMLIT = False

def render_tailwind_header(title: str, subtitle: str, badge_text: str = "MCA Capstone | ABES"):
    """
    Renders an executive, modern dashboard header using Tailwind CSS via CDN.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
      <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white rounded-xl p-6 shadow-md border border-slate-700">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div class="flex items-center gap-2 mb-1">
              <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500 bg-opacity-30 text-blue-200 border border-blue-400 border-opacity-30">
                {badge_text}
              </span>
              <span class="text-xs text-slate-400">Live Telemetry Pipeline</span>
            </div>
            <h1 class="text-2xl font-bold tracking-tight text-white mb-1">{title}</h1>
            <p class="text-sm text-slate-300 max-w-2xl">{subtitle}</p>
          </div>
          <div class="flex items-center gap-3">
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
        components.html(html_code, height=130)
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
        <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold border {badge_cls}">
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
        <div class="flex items-center justify-between mb-2">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</span>
          {delta_html}
        </div>
        <div class="text-2xl font-bold text-slate-900 tracking-tight">{value}</div>
        <div class="text-xs text-slate-400 mt-1 truncate">{subtitle}</div>
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        components.html(html_code, height=105)
    return html_code

def render_kpi_row(cards: list):
    """
    Renders 4 KPI cards side by side in a single responsive container.
    cards: list of dicts with keys: title, value, delta, delta_type, subtitle
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
            delta_tag = f"""<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold border {badge_cls}">{c['delta']}</span>"""
        
        cards_html += f"""
        <div class="bg-white rounded-xl p-4 shadow-sm border border-slate-200 hover:border-blue-400 transition-colors">
          <div class="flex items-center justify-between mb-2">
            <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{c['title']}</span>
            {delta_tag}
          </div>
          <div class="text-2xl font-bold text-slate-900 tracking-tight">{c['value']}</div>
          <div class="text-xs text-slate-400 mt-1 truncate">{c.get('subtitle', '')}</div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {cards_html}
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        components.html(full_html, height=115)
    return full_html

def render_insight_box(title: str, text: str, alert_type: str = "info"):
    """
    Renders a styled callout box for executive takeaways or statistical conclusions.
    alert_type: 'info' (indigo), 'success' (emerald), 'warning' (amber), 'danger' (rose)
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
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
      <div class="border-l-4 {box_cls} p-4 rounded-r-xl shadow-sm mb-2">
        <h4 class="text-sm font-bold {title_cls} mb-1">{title}</h4>
        <p class="text-xs text-slate-700 leading-relaxed m-0">{text}</p>
      </div>
    </body>
    </html>
    """
    if HAS_STREAMLIT:
        components.html(html_code, height=90)
    return html_code

def apply_plotly_theme(fig):
    """
    Applies unified corporate styling to a Plotly figure:
    Navy blue accent, clean background, readable typography, subtle gridlines.
    """
    if fig is None:
        return fig
    try:
        fig.update_layout(
            font=dict(family="Inter, Calibri, sans-serif", size=12, color="#0F172A"),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(248,250,252,0.6)",
            margin=dict(l=40, r=40, t=50, b=40),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                bgcolor="rgba(255,255,255,0.8)"
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
