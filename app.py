"""
Grandiose Bakery — Financial Performance & Cost Optimization Dashboard
GIP III | Parineeta Jain, Rajveer Singh, Tarang Gupta

Bold dark-theme redesign: near-black canvas, vibrant amber/coral/violet
accents, sparkline KPI cards, glowing trend charts, pill navigation.

Two tabs:
  Tab A - Performance tracker (KPIs, trend, category panels)
  Tab B - Scenario & resilience (4 modules)

Scope: bakery operations only (catering excluded per GIP III scope).
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import smtplib
import re
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
from anthropic import Anthropic
from pypdf import PdfReader
from docx import Document as DocxDocument
from supabase import create_client

# ----------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Grandiose Bakery Dashboard",
    page_icon="🍞",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------
# COLOR PALETTE — bold dark theme. Near-black canvas with vibrant,
# bakery-warm accents (amber/gold as primary, coral and violet as
# secondary/tertiary) instead of generic blue/purple.
# ----------------------------------------------------------------------
COLORS = {
    "bg":        "#0A0A0A",   # near-black — matches the portfolio site exactly
    "surface":   "#16141E",   # dark plum-tinted card surface
    "surface2":  "#1D1926",   # sidebar / secondary surface
    "border":    "#2E2838",   # muted plum-gray border
    "text":      "#F4F0F8",   # near-white, cool lilac-white
    "text_soft": "#9C93AC",   # muted lilac-gray secondary text
    "primary":   "#B9A0DC",   # soft lilac — main accent
    "primary_d": "#6B5789",   # deeper lilac, for text on light-lilac backgrounds
    "secondary": "#E8AECB",   # soft pink
    "tertiary":  "#8C76B0",   # deeper plum/violet, for chart variety
    "success":   "#C9B8E8",   # pale lilac — used for "on target / good" status
    "warning":   "#D9A8C4",   # dusty rose/mauve — used for "caution" status
    "danger":    "#9C6B8C",   # muted plum-wine — used for "risk" status (deliberately not red)
}

CATEGORICAL = [COLORS["primary"], COLORS["secondary"], "#F0DCEA", COLORS["tertiary"], COLORS["success"]]

def hex_to_rgba(hex_color, alpha):
    hex_color = hex_color.lstrip("#")
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Anton&family=Space+Grotesk:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)

st.markdown(f"""
<style>
    .stApp {{ background-color: {COLORS['bg']}; }}
    section[data-testid="stSidebar"] {{ background-color: {COLORS['surface2']}; border-right: 1px solid {COLORS['border']}; }}

    /* Top app header/toolbar band (Share, star, edit, GitHub icons) —
       white by default, doesn't inherit the dark theme on its own.
       Real testid is "stHeader" (className happens to be "stAppHeader",
       which is not the same thing — that mismatch was the bug last time). */
    [data-testid="stHeader"] {{
        background-color: {COLORS['bg']} !important;
        border-bottom: 1px solid {COLORS['border']} !important;
    }}
    [data-testid="stHeader"] * {{
        color: {COLORS['text']} !important;
        fill: {COLORS['text']} !important;
    }}
    [data-testid="stHeader"] [data-testid^="stBaseButton"] {{
        background-color: {COLORS['surface']} !important;
    }}
    [data-testid="stHeader"] [data-testid^="stBaseButton"] p,
    [data-testid="stHeader"] [data-testid^="stBaseButton"] span {{
        color: {COLORS['text']} !important;
    }}
    [data-testid="stMainMenu"], [data-testid="stStatusWidget"], [data-testid="stAppDeployButton"] {{
        background-color: transparent !important;
        color: {COLORS['text']} !important;
        fill: {COLORS['text']} !important;
    }}
    /* The actual clickable icon is a nested "stMainMenuButton", a different
       component from the other header buttons — this is the specific
       element that was still showing as an unstyled white block. */
    [data-testid="stMainMenuButton"] {{
        background-color: transparent !important;
        color: {COLORS['text']} !important;
    }}
    [data-testid="stMainMenuButton"] svg, [data-testid="stMainMenuButton"] path,
    [data-testid="stMainMenu"] svg, [data-testid="stStatusWidget"] svg {{
        fill: {COLORS['text']} !important;
        color: {COLORS['text']} !important;
    }}

    .stApp, .stApp p, .stApp span, .stApp li, .stApp label,
    div[data-testid="stMarkdownContainer"], div[data-testid="stMarkdownContainer"] p {{
        color: {COLORS['text']} !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }}
    div[data-testid="stCaptionContainer"], .stApp small {{ color: {COLORS['text_soft']} !important; }}
    h1, h2 {{
        color: {COLORS['text']} !important; font-family: 'Anton', sans-serif !important;
        font-weight: 400 !important; letter-spacing: 0.5px;
    }}
    h3, h4, h5 {{
        color: {COLORS['text']} !important; font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {COLORS['text']} !important; font-family: 'Space Grotesk', sans-serif !important; }}

    /* Restore Streamlit's icon-ligature font wherever our broad font-family
       rules above would otherwise break it — this is what renders expander/
       popover arrows correctly instead of as literal "expand_more" text. */
    span[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] * {{
        font-family: 'Material Symbols Rounded' !important;
    }}
    section[data-testid="stSidebar"] span[data-testid="stIconMaterial"] {{
        font-family: 'Material Symbols Rounded' !important;
    }}

    /* Popovers render in a separate layer that doesn't inherit .stApp's dark
       background — without this, forced near-white text sits on Streamlit's
       default white popover background and becomes unreadable. */
    div[data-testid="stPopoverBody"] {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
        border-radius: 14px !important;
    }}
    div[data-testid="stPopoverBody"] * {{
        color: {COLORS['text']} !important;
    }}
    div[data-testid="stPopoverBody"] input, div[data-testid="stPopoverBody"] textarea {{
        background-color: {COLORS['surface2']} !important;
        color: {COLORS['text']} !important;
        border: 1px solid {COLORS['border']} !important;
    }}
    div[data-testid="stPopoverBody"] ::placeholder {{
        color: {COLORS['text_soft']} !important;
        opacity: 1 !important;
    }}

    /* Every floating/portaled element (tooltips, selectbox dropdown menus,
       date pickers, multiselect lists) gets rendered by baseweb OUTSIDE
       .stApp, in a "baseweb-layer" portal attached to <body> — none of it
       inherits our dark theme by default, which is why hover tooltips and
       dropdown option lists have shown up as white-on-white. Fix them all
       at once rather than chasing each one individually. */
    .baseweb-layer, [data-testid="stTooltipContent"], [data-baseweb="menu"],
    [data-baseweb="popover"], [data-baseweb="calendar"] {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
    }}
    .baseweb-layer *, [data-testid="stTooltipContent"] *, [data-baseweb="menu"] *,
    [data-baseweb="popover"] *, [data-baseweb="calendar"] * {{
        color: {COLORS['text']} !important;
        background-color: transparent !important;
    }}
    [data-baseweb="menu"] li:hover, [data-baseweb="menu"] li[aria-selected="true"] {{
        background-color: {hex_to_rgba(COLORS['primary'], 0.18)} !important;
    }}

    /* Every button in the app (st.button, form_submit_button, download_button,
       popover triggers, etc.) — Streamlit's default "secondary" button style
       doesn't have enough contrast once our text-color override is applied,
       so every button gets a consistent solid lilac pill with dark text. */
    [data-testid^="stBaseButton"] {{
        background-color: {COLORS['primary']} !important;
        border: none !important;
        border-radius: 10px !important;
    }}
    [data-testid^="stBaseButton"] p, [data-testid^="stBaseButton"] span,
    [data-testid^="stBaseButton"] div {{
        color: #1A1424 !important;
        font-weight: 700 !important;
    }}
    [data-testid^="stBaseButton"]:hover {{
        background-color: {COLORS['secondary']} !important;
    }}
    [data-testid^="stBaseButton"]:disabled, [data-testid^="stBaseButton"]:disabled p {{
        background-color: {COLORS['border']} !important;
        color: {COLORS['text_soft']} !important;
    }}

    /* Expander headers — safety net so the label never fades to low-contrast */
    div[data-testid="stExpander"] summary p, div[data-testid="stExpander"] summary span {{
        color: {COLORS['text']} !important;
        opacity: 1 !important;
    }}


    /* Bordered containers -> dark elevated cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']} !important;
        border-radius: 18px !important;
        padding: 4px 6px;
    }}

    .kpi-icon {{
        display:inline-flex; align-items:center; justify-content:center;
        width:34px; height:34px; border-radius:10px; font-size:17px; margin-right:8px;
    }}
    .kpi-label {{ color:{COLORS['text_soft']}; font-size:0.82rem; font-weight:500; }}
    .kpi-value {{
        font-size:2.1rem; font-weight:600; margin: 4px 0 2px 0; letter-spacing:-0.5px;
        font-family: 'IBM Plex Mono', monospace !important;
    }}
    .pill {{
        display:inline-block; padding:3px 11px; border-radius:20px;
        font-size:0.76rem; font-weight:700;
    }}
    .pill-up-bad {{ background-color: rgba(156,107,140,0.20); color:{COLORS['danger']} !important; }}
    .pill-up-good {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-down-good {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-flat {{ background-color: rgba(156,147,172,0.15); color:{COLORS['text_soft']} !important; }}
    .pill-ok {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-warn {{ background-color: rgba(217,168,196,0.20); color:{COLORS['warning']} !important; }}
    .pill-risk {{ background-color: rgba(156,107,140,0.22); color:{COLORS['danger']} !important; }}

    .cat-row {{
        display:flex; justify-content:space-between; padding:7px 0;
        border-bottom:1px solid {COLORS['border']}; font-size:0.9rem;
    }}
    .cat-row span:first-child {{ color:{COLORS['text_soft']}; }}
    .cat-row span:last-child {{ font-weight:600; }}

    .info-row {{
        padding:8px 0; border-bottom:1px solid {COLORS['border']};
        font-size:0.92rem; line-height:1.5; display:flex; gap:8px;
    }}
    .info-row:last-child {{ border-bottom:none; }}
    .info-row::before {{ content:"—"; color:{COLORS['primary']}; flex-shrink:0; }}

    .hero-badge {{
        display:inline-block; padding:5px 14px; border-radius:20px; font-size:0.78rem;
        font-weight:700; background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
        color:{COLORS['text_soft']} !important; margin-right:8px;
    }}

    /* Pill-style tabs */
    .stTabs [data-baseweb="tab-list"] {{
        background-color: {COLORS['surface']}; padding:6px; border-radius:14px;
        border:1px solid {COLORS['border']}; gap:4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {COLORS['text_soft']} !important; border-radius:10px !important;
        padding: 8px 16px !important;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {COLORS['primary']} !important;
        color: #1A1424 !important; font-weight:700 !important;
    }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color: transparent !important; }}
    .stTabs [data-baseweb="tab-border"] {{ display:none !important; }}

    div[data-testid="stMetric"] {{ background-color: transparent; }}

    /* Sliders / inputs on dark */
    div[data-baseweb="slider"] div[role="slider"] {{ background-color: {COLORS['primary']} !important; }}
    .stSlider [data-testid="stTickBar"] {{ display:none; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_DARK = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=COLORS["text"], size=12),
    margin=dict(l=10, r=10, t=25, b=10),
)
GRID_COLOR = "rgba(255,255,255,0.06)"

def sparkline(values, color):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        y=values, mode="lines", line=dict(color=color, width=2.5),
        fill="tozeroy", fillcolor=color.replace(")", ",0.12)").replace("rgb", "rgba") if "rgb" in color else None,
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=2, b=2), height=48,
        xaxis=dict(visible=False), yaxis=dict(visible=False), showlegend=False,
    )
    return fig

# ----------------------------------------------------------------------
# SAMPLE / BENCHMARK DATA (placeholder until mentor-provided actuals)
# ----------------------------------------------------------------------
months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul"]
food_cost_trend = [29.8, 30.1, 30.6, 30.9, 31.0, 31.4]
target_line = [30.0] * 6
wastage_trend = [5.4, 5.2, 5.1, 4.9, 4.9, 4.8]
margin_trend = [42.3, 42.0, 41.8, 42.0, 42.2, 42.1]
cost_unit_trend = [3.55, 3.62, 3.68, 3.74, 3.80, 3.85]

baseline = {
    "food_cost_pct": 31.4,
    "wastage_pct": 4.8,
    "gross_margin_pct": 42.1,
    "cost_per_unit": 3.85,
    "standard_cost_per_unit": 3.60,
}

category_panels = {
    "🥖 Production": {
        "Batch efficiency": "88%",
        "Labor cost / unit": "AED 1.10",
        "Std vs actual cost": "+6.9%",
    },
    "🚚 Procurement": {
        "Purchase price variance": "+3.2%",
        "Supplier lead time": "2.4 days",
        "Top-3 flour supplier share": "76%",
    },
    "📦 Warehouse & inventory": {
        "Inventory turnover": "14.2x/yr",
        "Expired/damaged stock": "AED 8,400",
        "Slow-moving SKUs": "6",
    },
    "💰 Cost control & margin": {
        "Margin leakage (est.)": "AED 42,000/mo",
        "Packaging cost impact": "2.1% of sales",
        "Overhead allocation": "9.4% of sales",
    },
}

# ----------------------------------------------------------------------
# STATUS REPORT (EXCEL) + EMAIL DELIVERY
# Each "Email this ..." button builds a workbook tailored to the section
# it's called from, not a generic dashboard dump — Performance Tracker,
# Scenario & Resilience, a department comparison, or the Company Profile.
# ----------------------------------------------------------------------
XL_PRIMARY = "B9A0DC"
XL_DARK = "1A1424"
XL_HEADER_FONT = "FFFFFF"
XL_TEXT = "2B2733"
XL_SOFT_TEXT = "6B6577"
XL_LIGHT_FILL = "F3EFFA"
XL_BORDER_COLOR = "D9D2E9"

def _xl_border():
    side = Side(style="thin", color=XL_BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)

def _xl_title(ws, title, subtitle=None):
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=XL_DARK)
    row = 2
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = Font(name="Arial", size=10, italic=True, color=XL_SOFT_TEXT)
        row = 3
    return row + 1

def _xl_section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = Font(name="Arial", size=12, bold=True, color=XL_DARK)
    return row + 1

def _xl_header_row(ws, row, headers, start_col=1):
    border = _xl_border()
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color=XL_HEADER_FONT)
        c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = border
    return row + 1

def _xl_row(ws, row, values, start_col=1, formats=None, zebra=False):
    border = _xl_border()
    fill = PatternFill("solid", fgColor=XL_LIGHT_FILL) if zebra else None
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = Font(name="Arial", size=10, color=XL_TEXT)
        c.border = border
        if fill:
            c.fill = fill
        if formats and i < len(formats) and formats[i]:
            c.number_format = formats[i]
    return row + 1

def _xl_autosize(ws, max_width=64):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or hasattr(cell, "column_letter") is False:
                continue
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            widths[col] = max(widths.get(col, 8), min(len(str(cell.value)) + 3, max_width))
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def _xl_footer(ws, row):
    row += 1
    ws.cell(row=row, column=1,
             value=f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')} · GIP III · Bakery division only, catering excluded"
             ).font = Font(name="Arial", size=8, italic=True, color=XL_SOFT_TEXT)
    row += 1
    ws.cell(row=row, column=1,
             value="Prepared by Parineeta Jain, Rajveer Singh, and Tarang Gupta · Figures are illustrative "
                   "benchmarks pending Grandiose-provided actuals."
             ).font = Font(name="Arial", size=8, italic=True, color=XL_SOFT_TEXT)
    return row

PCT_FMT = '0.0"%"'
AED_FMT = '"AED "#,##0.00'
NUM_FMT = "#,##0"


def _build_performance_tracker_sheet(wb):
    ws = wb.create_sheet("Performance Tracker")
    row = _xl_title(ws, "Grandiose Bakery — Performance Tracker",
                     "Where the bakery division stands today")

    row = _xl_section(ws, row, "Headline KPIs")
    row = _xl_header_row(ws, row, ["Metric", "Value", "vs. target / last month"])
    kpi_rows = [
        ("Food cost %", baseline["food_cost_pct"], "1.2pt above 30% target"),
        ("Wastage %", baseline["wastage_pct"], "Down 0.3pt vs last month"),
        ("Gross margin %", baseline["gross_margin_pct"], "Flat vs last month"),
        ("Cost per unit (AED)", baseline["cost_per_unit"], f"Standard: AED {baseline['standard_cost_per_unit']:.2f}"),
    ]
    for i, (label, val, note) in enumerate(kpi_rows):
        fmt = AED_FMT if "unit" in label.lower() else PCT_FMT
        row = _xl_row(ws, row, [label, val, note], formats=[None, fmt, None], zebra=(i % 2 == 0))
    row += 1

    row = _xl_section(ws, row, "Food cost % — last 6 months vs. 30% target")
    row = _xl_header_row(ws, row, ["Month", "Food cost %", "Target %"])
    for i, (m, fc, tg) in enumerate(zip(months, food_cost_trend, target_line)):
        row = _xl_row(ws, row, [m, fc, tg], formats=[None, PCT_FMT, PCT_FMT], zebra=(i % 2 == 0))
    row += 1

    row = _xl_section(ws, row, "Category panels")
    row = _xl_header_row(ws, row, ["Category", "Metric", "Value"])
    zebra = False
    for cat_name, metrics in category_panels.items():
        clean_cat = cat_name.split(" ", 1)[-1]
        for k, v in metrics.items():
            row = _xl_row(ws, row, [clean_cat, k, v], zebra=zebra)
            zebra = not zebra
    row += 1

    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def _build_scenario_resilience_sheets(wb, context):
    infl = context.get("inflation", {})
    disr = context.get("disruption", {})
    pand = context.get("pandemic", {})
    supp = context.get("supplier", {})

    # --- Inflation sensitivity ---
    ws = wb.create_sheet("Inflation Sensitivity")
    row = _xl_title(ws, "Scenario & Resilience — Inflation Sensitivity",
                     "Dual-track inflation input, net of any subsidy/offset")
    row = _xl_section(ws, row, "Assumptions used")
    row = _xl_header_row(ws, row, ["Input", "Value"])
    row = _xl_row(ws, row, ["Headline inflation forecast", infl.get("headline_inf", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws, row, ["Actual food-input inflation", infl.get("food_inf", 0)], formats=[None, PCT_FMT])
    row = _xl_row(ws, row, ["Subsidy / price-cap offset (AED/unit)", infl.get("subsidy_offset", 0)], formats=[None, AED_FMT], zebra=True)
    row += 1
    row = _xl_section(ws, row, "Resulting cost per unit")
    row = _xl_header_row(ws, row, ["Scenario", "Cost per unit (AED)"])
    row = _xl_row(ws, row, ["Current", infl.get("base_cost", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws, row, ["Headline scenario", infl.get("adj_cost_headline", 0)], formats=[None, AED_FMT])
    row = _xl_row(ws, row, ["Food-inflation scenario", infl.get("adj_cost_food", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws, row, ["Food cost % under food-inflation scenario", infl.get("food_cost_pct_adj", 0)], formats=[None, PCT_FMT])
    _xl_footer(ws, row)
    _xl_autosize(ws)

    # --- Supply disruption ---
    ws2 = wb.create_sheet("Supply Disruption")
    row = _xl_title(ws2, "Scenario & Resilience — Supply Disruption Risk",
                     "UAE imports 80-90% of its food, so this is mainly an import-shock model")
    row = _xl_section(ws2, row, "Selected scenario")
    row = _xl_header_row(ws2, row, ["Field", "Value"])
    row = _xl_row(ws2, row, ["Scenario", disr.get("scenario", "—")], zebra=True)
    row = _xl_row(ws2, row, ["Lead-time extension (days)", disr.get("delay_days", 0)], formats=[None, NUM_FMT])
    row = _xl_row(ws2, row, ["Cost premium %", disr.get("cost_premium", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws2, row, ["Stockout probability", disr.get("stockout_prob", 0) * 100], formats=[None, PCT_FMT])
    row += 1
    row = _xl_section(ws2, row, "Buffer vs. expected stockout cost")
    row = _xl_header_row(ws2, row, ["Metric", "AED"])
    row = _xl_row(ws2, row, ["Cost of holding buffer stock", disr.get("buffer_stock_cost", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws2, row, ["Expected cost of stockout", disr.get("expected_stockout_cost", 0)], formats=[None, AED_FMT])
    _xl_footer(ws2, row)
    _xl_autosize(ws2)

    # --- Pandemic preparedness ---
    ws3 = wb.create_sheet("Pandemic Preparedness")
    row = _xl_title(ws3, "Scenario & Resilience — Pandemic Preparedness",
                     "Demand-side and supply-side resilience, tracked together")
    row = _xl_section(ws3, row, "Customer retention / attraction")
    row = _xl_header_row(ws3, row, ["Metric", "Value"])
    row = _xl_row(ws3, row, ["Repeat-purchase rate", pand.get("repeat_rate", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws3, row, ["Delivery/online-order revenue share", pand.get("delivery_share", 0)], formats=[None, PCT_FMT])
    row = _xl_row(ws3, row, ["Average basket size (AED)", pand.get("basket_size", 0)], formats=[None, AED_FMT], zebra=True)
    row += 1
    row = _xl_section(ws3, row, "Supply chain optimization")
    row = _xl_header_row(ws3, row, ["Metric", "Value"])
    row = _xl_row(ws3, row, ["Safety stock (days of cover)", pand.get("safety_days", 0)], formats=[None, NUM_FMT], zebra=True)
    row = _xl_row(ws3, row, ["Qualified alternate suppliers / key ingredient", pand.get("alt_suppliers", 0)], formats=[None, NUM_FMT])
    row = _xl_row(ws3, row, ["% ingredients single-sourced", pand.get("single_sourced", 0)], formats=[None, PCT_FMT], zebra=True)
    _xl_footer(ws3, row)
    _xl_autosize(ws3)

    # --- Supplier concentration ---
    ws4 = wb.create_sheet("Supplier Concentration")
    row = _xl_title(ws4, "Scenario & Resilience — Supplier Concentration",
                     "HHI-style concentration index, following Ali et al. (2022)'s UAE cereal-supply-risk methodology")
    row = _xl_section(ws4, row, "Concentration index")
    row = _xl_header_row(ws4, row, ["Metric", "Value"])
    row = _xl_row(ws4, row, ["HHI", supp.get("hhi", 0)], formats=[None, NUM_FMT], zebra=True)
    row = _xl_row(ws4, row, ["Risk level", supp.get("risk_label", "—")])
    row += 1
    shares_df = supp.get("shares_df")
    if shares_df is not None and not shares_df.empty:
        row = _xl_section(ws4, row, "Supplier spend mix")
        row = _xl_header_row(ws4, row, list(shares_df.columns))
        for i, (_, r) in enumerate(shares_df.iterrows()):
            row = _xl_row(ws4, row, list(r.values), zebra=(i % 2 == 0))
        row += 1
    alt_df = supp.get("alt_df")
    if alt_df is not None and not alt_df.empty:
        row = _xl_section(ws4, row, "Pre-identified alternate suppliers")
        row = _xl_header_row(ws4, row, list(alt_df.columns))
        for i, (_, r) in enumerate(alt_df.iterrows()):
            row = _xl_row(ws4, row, list(r.values), zebra=(i % 2 == 0))
    _xl_footer(ws4, row)
    _xl_autosize(ws4)


def _build_department_sheet(wb, context):
    dept_name = context.get("department", "Department")
    ws = wb.create_sheet(dept_name[:31] if dept_name else "Department")
    row = _xl_title(ws, f"Employee Portal — {dept_name} Comparison",
                     "Performance ranking within this department")
    headcount = context.get("headcount", "—")
    tag = context.get("tag_label")
    row = _xl_section(ws, row, f"Headcount on this line: {headcount}" + (f" · {tag}" if tag else ""))
    df = context.get("summary_df")
    if df is not None and not df.empty:
        row = _xl_header_row(ws, row, list(df.columns))
        for i, (_, r) in enumerate(df.iterrows()):
            row = _xl_row(ws, row, list(r.values), zebra=(i % 2 == 0))
    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def _build_company_profile_sheet(wb, context):
    ws = wb.create_sheet("Company Profile")
    row = _xl_title(ws, "Grandiose Bakery — Company Profile",
                     "Everything collected from the 27 Jul meeting with the GM")
    company_info = context.get("company_info", [])
    for icon, title, points in company_info:
        row = _xl_section(ws, row, f"{icon} {title}")
        for i, p in enumerate(points):
            row = _xl_row(ws, row, [p], zebra=(i % 2 == 0))
        row += 1
    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def _coerce_cell(value):
    """Try to turn AI-returned strings that look numeric into real numbers,
    so Excel treats them as numbers (right-aligned, sortable) not text.
    Leaves percentage strings ('5.0%') and currency strings ('AED 3.85')
    as readable text rather than stripping their meaning away."""
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if "%" in stripped or "AED" in stripped:
            return stripped
        cleaned = stripped.replace(",", "")
        try:
            if "." in cleaned:
                return float(cleaned)
            return int(cleaned)
        except (ValueError, TypeError):
            return value
    return value


def _build_smart_upload_sheets(wb, context):
    result = context.get("result", {})
    source_files = context.get("source_files", [])

    ws = wb.create_sheet("Summary")
    row = _xl_title(ws, "Grandiose Bakery — Processed Data Report",
                     "AI-interpreted from your uploaded file(s)")
    row = _xl_section(ws, row, "Source files")
    row = _xl_header_row(ws, row, ["File name"])
    for i, fname in enumerate(source_files):
        row = _xl_row(ws, row, [fname], zebra=(i % 2 == 0))
    row += 1
    row = _xl_section(ws, row, "Detected data type")
    ws.cell(row=row, column=1, value=result.get("detected_data_type", "—")).font = Font(name="Arial", size=10, color=XL_TEXT)
    row += 2
    row = _xl_section(ws, row, "Overall summary")
    for line in (result.get("summary", "") or "").split("\n"):
        if line.strip():
            ws.cell(row=row, column=1, value=line.strip()).font = Font(name="Arial", size=10, color=XL_TEXT)
            row += 1
    _xl_footer(ws, row)
    _xl_autosize(ws)

    for sheet_info in result.get("sheets", []):
        raw_name = sheet_info.get("sheet_name") or sheet_info.get("title") or "Sheet"
        safe_name = re.sub(r'[\\/*?:\[\]]', "-", raw_name)[:31] or "Sheet"
        base_name, n = safe_name, 1
        while safe_name in wb.sheetnames:
            n += 1
            safe_name = f"{base_name[:28]}_{n}"
        ws2 = wb.create_sheet(safe_name)
        row = _xl_title(ws2, sheet_info.get("title", raw_name))
        columns = sheet_info.get("columns", [])
        data_rows = sheet_info.get("rows", [])
        if columns:
            row = _xl_header_row(ws2, row, columns)
            for i, r in enumerate(data_rows):
                r = list(r) + [""] * (len(columns) - len(r)) if len(r) < len(columns) else r[:len(columns)]
                row = _xl_row(ws2, row, [_coerce_cell(v) for v in r], zebra=(i % 2 == 0))
            row += 1
        insights = sheet_info.get("insights", [])
        if insights:
            row = _xl_section(ws2, row, "Insights")
            for i, ins in enumerate(insights):
                row = _xl_row(ws2, row, [ins], zebra=(i % 2 == 0))
        _xl_footer(ws2, row)
        _xl_autosize(ws2)


def extract_pdf_text(file_bytes):
    try:
        reader = PdfReader(BytesIO(file_bytes))
        pages = [p.extract_text() or "" for p in reader.pages[:30]]
        return "\n".join(pages)
    except Exception as e:
        return f"[Could not read PDF: {e}]"


def extract_docx_text(file_bytes):
    try:
        doc = DocxDocument(BytesIO(file_bytes))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for r in table.rows:
                parts.append(" | ".join(c.text for c in r.cells))
        return "\n".join(parts)
    except Exception as e:
        return f"[Could not read DOCX: {e}]"


def extract_excel_text(file_bytes, filename):
    try:
        engine = "openpyxl" if filename.lower().endswith((".xlsx", ".xlsm")) else None
        sheets = pd.read_excel(BytesIO(file_bytes), sheet_name=None, engine=engine) if not filename.lower().endswith(".csv") \
            else {"Sheet1": pd.read_csv(BytesIO(file_bytes))}
        parts = []
        for name, df in sheets.items():
            parts.append(f"--- Sheet: {name} ---")
            parts.append(df.to_csv(index=False))
        return "\n".join(parts)
    except Exception as e:
        return f"[Could not read spreadsheet: {e}]"


MAX_UPLOAD_CHARS = 60000  # keeps API cost/latency bounded across a batch of files

@st.cache_resource
def get_anthropic_client():
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        return None
    try:
        return Anthropic(api_key=api_key)
    except Exception:
        return None


INTERPRETER_SYSTEM_PROMPT = """You are a financial and operations analyst working on a cost-optimization \
dashboard for Grandiose Bakery, a UAE bakery division (GIP III project). You will be given raw, rough, \
messy content extracted from one or more uploaded files (Excel, PDF, or Word) — it could be production logs, \
procurement records, wastage logs, sales data, inventory data, employee shift logs, or something else entirely.

Your job:
1. Identify what kind of business data this actually is.
2. Clean and interpret the rough numbers — infer column meaning even if headers are messy, missing, or \
inconsistent. Handle obvious typos and unit inconsistencies sensibly.
3. Where the data supports it, compute relevant KPIs using standard bakery cost-optimization metrics: \
wastage % (wasted / (produced + wasted) * 100), cost per unit, batch-time adherence %, quality pass rate, \
supplier concentration (HHI), productivity (value generated per employee per day), or straightforward \
sums/averages/trends — but ONLY compute what the data actually supports. Never fabricate numbers that \
aren't derivable from what was given.
4. Write 2-4 sentences of plain-English, management-readable insight per output sheet — what the numbers \
show and what it implies, not just a restatement of the table.

Respond with STRICT JSON only, no markdown fences, no commentary outside the JSON, matching exactly this \
shape:
{
  "detected_data_type": "short description of what this data is",
  "summary": "2-4 sentence overall summary of what was found across all files",
  "sheets": [
    {
      "sheet_name": "short sheet name, max 31 chars, no slashes or brackets",
      "title": "human-readable title for this sheet",
      "columns": ["Column A", "Column B", ...],
      "rows": [["value", "value", ...], ...],
      "insights": ["insight sentence 1", "insight sentence 2"]
    }
  ]
}

Produce 1-4 sheets depending on what's genuinely present in the data — don't invent sheets with no real \
content. If the uploaded content is unreadable or empty, still return valid JSON with an empty "sheets" \
list and explain why in "summary"."""


def interpret_uploaded_files(file_contents: dict):
    """file_contents: {filename: extracted_text}. Returns (result_dict, error_message)."""
    client = get_anthropic_client()
    if client is None:
        return None, ("AI processing isn't configured yet. Add ANTHROPIC_API_KEY in Streamlit Cloud's "
                       "Secrets settings (App settings -> Secrets) — never in the code or GitHub repo.")

    combined = ""
    for fname, text in file_contents.items():
        combined += f"\n\n===== FILE: {fname} =====\n{text}"
    truncated = len(combined) > MAX_UPLOAD_CHARS
    if truncated:
        combined = combined[:MAX_UPLOAD_CHARS]

    try:
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=8000,
            system=INTERPRETER_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": f"Here is the extracted content to interpret:{combined}"}],
        )
        raw_text = "".join(block.text for block in response.content if hasattr(block, "text"))
        cleaned = raw_text.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned.strip())
        result = json.loads(cleaned)
        if truncated:
            result["summary"] = (result.get("summary", "") +
                                  " (Note: input was long and was truncated before analysis — "
                                  "results reflect only the portion processed.)")
        return result, None
    except json.JSONDecodeError:
        return None, "The AI response couldn't be parsed as structured data. Try again, or with a smaller/simpler file."
    except Exception as e:
        return None, f"Could not process files: {e}"


def build_excel_report(section, context=None):
    """Builds a section-specific, formatted Excel workbook and returns raw bytes."""
    context = context or {}
    wb = Workbook()
    wb.remove(wb.active)

    if section == "scenario_resilience":
        _build_scenario_resilience_sheets(wb, context)
    elif section == "employee_department":
        _build_department_sheet(wb, context)
    elif section == "company_profile":
        _build_company_profile_sheet(wb, context)
    elif section == "smart_upload":
        _build_smart_upload_sheets(wb, context)
    else:
        _build_performance_tracker_sheet(wb)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


SECTION_LABELS = {
    "performance_tracker": ("Performance Tracker", "Performance_Tracker"),
    "scenario_resilience": ("Scenario & Resilience", "Scenario_Resilience"),
    "employee_department": ("Employee Department Comparison", "Department_Comparison"),
    "company_profile": ("Company Profile", "Company_Profile"),
    "smart_upload": ("Processed Data Report", "Processed_Data"),
}

def send_report_email(recipient_email, note="", section="performance_tracker", context=None):
    """Sends a section-specific Excel report to recipient_email using SMTP
    credentials stored in Streamlit secrets. Returns (success: bool, message: str)."""
    try:
        sender_email = st.secrets["EMAIL_ADDRESS"]
        sender_password = st.secrets["EMAIL_APP_PASSWORD"]
    except (KeyError, FileNotFoundError):
        return False, ("Email is not configured yet. Add EMAIL_ADDRESS and EMAIL_APP_PASSWORD "
                        "in Streamlit Cloud's Secrets settings (App settings -> Secrets) — "
                        "never in the code or GitHub repo.")

    smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    section_title, section_slug = SECTION_LABELS.get(section, SECTION_LABELS["performance_tracker"])

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = f"Grandiose Bakery — {section_title} Report"
    body = (f"Hi,\n\nPlease find attached the {section_title} report from the Grandiose bakery "
            f"dashboard, generated just now.\n")
    if note:
        body += f"\nNote from sender:\n{note}\n"
    body += "\n— Sent automatically from the Grandiose Bakery Dashboard (GIP III)"
    msg.attach(MIMEText(body, "plain"))

    xlsx_bytes = build_excel_report(section, context)
    attachment = MIMEApplication(
        xlsx_bytes, _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    filename = f"Grandiose_Bakery_{section_slug}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    attachment.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(attachment)

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=15) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True, f"Report sent to {recipient_email}."
    except Exception as e:
        return False, f"Could not send email: {e}"


EMAIL_PATTERN = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

def render_email_share(key_suffix, label="Email this report", section="performance_tracker", context=None):
    """Small 📧 popover button — usable right under any chart/section so
    people don't have to go back to the sidebar to share the dashboard.
    Sends a formatted Excel workbook tailored to the calling section."""
    section_title, _ = SECTION_LABELS.get(section, SECTION_LABELS["performance_tracker"])
    with st.popover("📧", help=label):
        st.markdown(f"**{label}**")
        recipient = st.text_input("Recipient email", placeholder="name@company.com", key=f"email_recipient_{key_suffix}")
        note = st.text_area("Add a note (optional)", key=f"email_note_{key_suffix}", height=68)
        if st.button("Send report", key=f"send_email_btn_{key_suffix}", width='stretch'):
            if not recipient or not re.match(EMAIL_PATTERN, recipient):
                st.error("Enter a valid email address.")
            else:
                with st.spinner("Building and sending Excel report..."):
                    success, message = send_report_email(recipient, note, section=section, context=context)
                if success:
                    st.success(message)
                else:
                    st.error(message)
        st.caption(f"Sends a formatted Excel workbook of the current {section_title} data.")


st.sidebar.markdown(f"""
<div style='background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
     border-radius:14px; padding:16px; margin-bottom:14px;'>
  <div style='font-size:1.5rem;'>🍞</div>
  <div style='font-weight:800; font-size:1.15rem; margin-top:4px;'>Grandiose Bakery</div>
  <div style='color:{COLORS['text_soft']}; font-size:0.8rem; margin-top:4px;'>
    GIP III — Financial performance &amp; cost optimization<br>Bakery division only · catering excluded
  </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown(f"""
<div style='background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
     border-radius:14px; padding:14px; margin-bottom:14px;'>
  <div style='color:{COLORS['text_soft']}; font-size:0.78rem; font-weight:700; letter-spacing:0.5px;'>PREPARED BY</div>
  <div style='margin-top:6px; font-size:0.9rem;'>Parineeta Jain</div>
  <div style='font-size:0.9rem;'>Rajveer Singh</div>
  <div style='font-size:0.9rem;'>Tarang Gupta</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown(f"""
<div style='background-color:{hex_to_rgba(COLORS['primary'], 0.1)}; border:1px solid {hex_to_rgba(COLORS['primary'], 0.3)};
     border-radius:14px; padding:14px; font-size:0.82rem; color:{COLORS['text']} !important;'>
  ⚠️ All figures shown are illustrative benchmarks pending Grandiose's actual bakery data.
</div>
""", unsafe_allow_html=True)

with st.sidebar.expander("📧  Email this report", expanded=False):
    recipient = st.text_input("Recipient email", placeholder="name@company.com", key="email_recipient")
    note = st.text_area("Add a note (optional)", key="email_note", height=70)
    if st.button("Send report", key="send_email_btn", width='stretch'):
        email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
        if not recipient or not re.match(email_pattern, recipient):
            st.error("Enter a valid email address.")
        else:
            with st.spinner("Building and sending Excel report..."):
                success, message = send_report_email(recipient, note, section="performance_tracker")
            if success:
                st.success(message)
            else:
                st.error(message)
    st.caption("Sends a formatted Excel workbook of the current Performance Tracker data.")


# ----------------------------------------------------------------------
# EMPLOYEE PORTAL — data layer
# Uses Supabase (Postgres) when configured via secrets; falls back to an
# in-session store so the portal is still fully clickable/demoable before
# Supabase is wired up. Data in fallback mode does not persist across
# app restarts — see README for the one-time Supabase setup.
# ----------------------------------------------------------------------

# Actual staff breakdown by line/department (approximate headcounts).
DEPARTMENT_HEADCOUNT = {
    "Baklava": 10,
    "French Bakery": 10,
    "Arabic Bread": 7,
    "Viennoiserie (Croissant/Danish)": 7,
    "Tahina": 5,
    "QC": 3,
    "Packing & Dispatch": 15,
    "Store": 2,
    "Maintenance": 4,
    "GM": 1,
    "Office": 10,
    "Drivers": 7,
    "Third-Party Cleaners": 10,
}

# Company-wide headcount context from the 27 Jul MoM. Note this figure (~110)
# is larger than the sum of the itemized bakery lines above (91) — likely
# because it includes catering/other roles not broken out per line. Shown
# for context only; the itemized breakdown above remains the bakery-scope
# source of truth for GIP III.
COMPANY_HEADCOUNT_CURRENT = 110
COMPANY_HEADCOUNT_TARGET = 200

# Illustrative daily capacity targets (units/day) for production lines only —
# PLACEHOLDER figures pending Grandiose's actual max-output numbers per line.
# The MoM gives an overall utilization figure (30-35%, target 100% within a
# year) but not a per-line breakdown, so utilization % shown against these
# targets should be treated as directional, not precise, until replaced.
DEPARTMENT_CAPACITY_TARGET = {
    "Baklava": 400,
    "French Bakery": 500,
    "Arabic Bread": 900,
    "Viennoiserie (Croissant/Danish)": 450,
    "Tahina": 250,
}

# Lines the MoM specifically called out — used to tag department cards/tables.
DEPARTMENT_PRIORITY_TAG = {
    "Arabic Bread": ("🔥 Top revenue driver", "pill-ok"),
    "Tahina": ("⚠️ Ramp-up focus (underutilized)", "pill-warn"),
    "Baklava": ("⚠️ Ramp-up focus (underutilized)", "pill-warn"),
}

# GM's stated productivity benchmark: each employee = AED 1,000/day value.
VALUE_PER_EMPLOYEE_DAY_TARGET = 1000
# Standard shift length per the MoM ("Standard shift is 9 hours, including a
# break") — used to normalize revenue/hour up to a full standard day.
STANDARD_SHIFT_HOURS = 9.0

# Company-wide wastage target from the MoM (current is 2-3%, target is 1%).
WASTAGE_TARGET_PCT = 1.0

# Illustrative sample employees per department (placeholders — swap in
# Grandiose's real staff names/IDs once available). Not every headcount
# above has a named entry; these exist so the portal and department
# comparisons are demoable today.
EMPLOYEES = [
    {"name": "Ahmad Yousef", "id": "BAK-01", "department": "Baklava"},
    {"name": "Rania Haddad", "id": "BAK-02", "department": "Baklava"},
    {"name": "Mohammed Iqbal", "id": "BAK-03", "department": "Baklava"},

    {"name": "Rahul Nair", "id": "FRB-01", "department": "French Bakery"},
    {"name": "Fatima Zahra", "id": "FRB-02", "department": "French Bakery"},
    {"name": "John Cruz", "id": "FRB-03", "department": "French Bakery"},

    {"name": "Khalid Obaid", "id": "ARB-01", "department": "Arabic Bread"},
    {"name": "Youssef Amer", "id": "ARB-02", "department": "Arabic Bread"},
    {"name": "Layla Mansour", "id": "ARB-03", "department": "Arabic Bread"},

    {"name": "Ana Reyes", "id": "VIE-01", "department": "Viennoiserie (Croissant/Danish)"},
    {"name": "Omar Saleh", "id": "VIE-02", "department": "Viennoiserie (Croissant/Danish)"},
    {"name": "Priya Sharma", "id": "VIE-03", "department": "Viennoiserie (Croissant/Danish)"},

    {"name": "Hassan Ali", "id": "TAH-01", "department": "Tahina"},
    {"name": "Meera Pillai", "id": "TAH-02", "department": "Tahina"},

    {"name": "Noura Saeed", "id": "QC-01", "department": "QC"},
    {"name": "Vikram Singh", "id": "QC-02", "department": "QC"},

    {"name": "Josie Santos", "id": "PKD-01", "department": "Packing & Dispatch"},
    {"name": "Ravi Kumar", "id": "PKD-02", "department": "Packing & Dispatch"},
    {"name": "Abdullah Nasser", "id": "PKD-03", "department": "Packing & Dispatch"},

    {"name": "Aisha Rahman", "id": "STR-01", "department": "Store"},
    {"name": "Manuel Cruz", "id": "STR-02", "department": "Store"},

    {"name": "Suresh Pillai", "id": "MNT-01", "department": "Maintenance"},
    {"name": "Ibrahim Al Farsi", "id": "MNT-02", "department": "Maintenance"},

    {"name": "Khalifa Al Marzooqi", "id": "GM-01", "department": "GM"},

    {"name": "Sara Al Ali", "id": "OFF-01", "department": "Office"},
    {"name": "Deepak Verma", "id": "OFF-02", "department": "Office"},
    {"name": "Nadia Haddad", "id": "OFF-03", "department": "Office"},

    {"name": "Mohammed Rafiq", "id": "DRV-01", "department": "Drivers"},
    {"name": "Carlos Dionisio", "id": "DRV-02", "department": "Drivers"},
    {"name": "Salim Bakhit", "id": "DRV-03", "department": "Drivers"},

    {"name": "Ganesh Kumar", "id": "CLN-01", "department": "Third-Party Cleaners"},
    {"name": "Rosa Villanueva", "id": "CLN-02", "department": "Third-Party Cleaners"},
]


@st.cache_resource
def get_supabase_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except (KeyError, FileNotFoundError):
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None

def portal_insert(table, row):
    client = get_supabase_client()
    if client is not None:
        try:
            client.table(table).insert(row).execute()
            return True, None
        except Exception as e:
            return False, str(e)
    else:
        row = {**row, "created_at": datetime.now().isoformat()}
        key = f"_local_{table}"
        st.session_state.setdefault(key, [])
        row["id"] = len(st.session_state[key]) + 1
        st.session_state[key].append(row)
        return True, None

def portal_fetch(table, filters=None):
    client = get_supabase_client()
    if client is not None:
        try:
            q = client.table(table).select("*")
            if filters:
                for k, v in filters.items():
                    q = q.eq(k, v)
            res = q.order("created_at", desc=True).execute()
            return pd.DataFrame(res.data), None
        except Exception as e:
            return pd.DataFrame(), str(e)
    else:
        rows = st.session_state.get(f"_local_{table}", [])
        df = pd.DataFrame(rows)
        if not df.empty:
            if filters:
                for k, v in filters.items():
                    df = df[df[k] == v]
            if "created_at" in df.columns:
                df = df.sort_values("created_at", ascending=False)
        return df, None

def portal_storage_mode():
    return "Supabase (persistent)" if get_supabase_client() is not None else "In-session only (not saved after restart)"


# ----------------------------------------------------------------------
# HERO HEADER
# ----------------------------------------------------------------------
st.markdown(f"""
<div style='margin-bottom:6px;'>
  <span class='hero-badge'>🇦🇪 UAE · Grandiose</span>
  <span class='hero-badge'>Bakery division</span>
  <span class='hero-badge'>GIP III</span>
</div>
""", unsafe_allow_html=True)
st.markdown("## Financial Performance Dashboard")

st.markdown(f"""
<style>
    div[data-baseweb="select"] > div {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
        border-radius: 12px !important;
    }}
</style>
""", unsafe_allow_html=True)

section = st.selectbox(
    "Dashboard section",
    ["📊  Performance tracker", "🧭  Scenario & resilience", "👤  Employee portal", "📋  Company profile",
     "📥  Data processor"],
    label_visibility="collapsed",
)

# ========================================================================
# SECTION A — PERFORMANCE TRACKER
# ========================================================================
if section == "📊  Performance tracker":
    st.caption("Where the bakery division stands today")

    kpis = [
        ("🍞", "Food cost %", f"{baseline['food_cost_pct']}%", "↑ 1.2pt vs target", "pill-up-bad",
         food_cost_trend, COLORS["primary"]),
        ("🗑️", "Wastage %", f"{baseline['wastage_pct']}%", "↓ 0.3pt vs last month", "pill-down-good",
         wastage_trend, COLORS["success"]),
        ("📈", "Gross margin", f"{baseline['gross_margin_pct']}%", "flat vs last month", "pill-flat",
         margin_trend, COLORS["tertiary"]),
        ("💵", "Cost / unit (AED)", f"{baseline['cost_per_unit']:.2f}", f"↑ vs std {baseline['standard_cost_per_unit']:.2f}", "pill-up-bad",
         cost_unit_trend, COLORS["secondary"]),
    ]
    cols = st.columns(4)
    for col, (icon, label, value, delta, pill_class, trend, color) in zip(cols, kpis):
        with col:
            with st.container(border=True):
                st.markdown(f"""
                <div style='padding:10px 14px 0 14px;'>
                    <span class='kpi-icon' style='background-color:{hex_to_rgba(color,0.15)};'>{icon}</span>
                    <span class='kpi-label'>{label}</span>
                    <div class='kpi-value'>{value}</div>
                    <span class='pill {pill_class}'>{delta}</span>
                </div>
                """, unsafe_allow_html=True)
                st.plotly_chart(sparkline(trend, color), width='stretch',
                                 config={"displayModeBar": False}, key=f"spark_{label}")

    st.markdown("#### Food cost % vs target — last 6 months")
    with st.container(border=True):
        fig = go.Figure()
        # glow effect: wide low-opacity trace behind crisp line
        fig.add_trace(go.Scatter(x=months, y=food_cost_trend, mode="lines", line=dict(color=COLORS["primary"], width=10),
                                  opacity=0.15, showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=months, y=food_cost_trend, mode="lines+markers", name="Food cost %",
                                  line=dict(color=COLORS["primary"], width=3),
                                  marker=dict(size=8, color=COLORS["bg"], line=dict(color=COLORS["primary"], width=2)),
                                  fill="tozeroy", fillcolor=hex_to_rgba(COLORS["primary"], 0.12)))
        fig.add_trace(go.Scatter(x=months, y=target_line, mode="lines", name="Target",
                                  line=dict(color=COLORS["text_soft"], width=1.5, dash="dot")))
        fig.update_layout(**PLOTLY_DARK, height=300,
                           yaxis=dict(ticksuffix="%", range=[28, 33], gridcolor=GRID_COLOR, zeroline=False),
                           xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                           legend=dict(orientation="h", yanchor="bottom", y=1.03, x=0,
                                       font=dict(color=COLORS["text_soft"])))
        st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    render_email_share("trend_chart", "Email this Performance Tracker report", section="performance_tracker")

    st.markdown("#### Category panels")
    cols2 = st.columns(4)
    for col, (name, metrics) in zip(cols2, category_panels.items()):
        with col:
            with st.container(border=True):
                rows = "".join([f"<div class='cat-row'><span>{k}</span><span>{v}</span></div>" for k, v in metrics.items()])
                st.markdown(f"<div style='padding:12px 16px;'><h4 style='margin:0 0 8px 0; font-size:1rem;'>{name}</h4>{rows}</div>",
                            unsafe_allow_html=True)
    render_email_share("category_panels", "Email this Performance Tracker report", section="performance_tracker")

# ========================================================================
# TAB B — SCENARIO & RESILIENCE
# ========================================================================
# ========================================================================
# SECTION B — SCENARIO & RESILIENCE
# ========================================================================
elif section == "🧭  Scenario & resilience":
    st.caption("What happens if — stress-tests the same cost and margin outputs under new assumptions")

    mod1, mod2, mod3, mod4 = st.tabs([
        "1 · Inflation sensitivity", "2 · Supply disruption risk",
        "3 · Pandemic preparedness", "4 · Supplier alternatives",
    ])

    # ---------------- MODULE 1: INFLATION-ADJUSTED COST SENSITIVITY ----------------
    with mod1:
        st.markdown("##### 🔥 Inflation-adjusted cost sensitivity")
        st.caption("Dual-track inflation input with an adjustable bar, net of any subsidy/offset deducted.")

        base_cost = baseline["cost_per_unit"]
        colL, colR = st.columns([1, 1.3])
        with colL:
            with st.container(border=True):
                headline_inf = st.slider("Headline inflation forecast (%)", 0.0, 10.0, 2.8, 0.1)
                food_inf = st.slider("Actual food-input inflation (%)", 0.0, 20.0, 5.3, 0.1,
                                      help="Flour, dairy, and packaging inflation can diverge sharply from headline CPI during shocks.")
                subsidy_offset = st.number_input("Subsidy / price-cap offset (AED per unit)", 0.0, 2.0, 0.05, 0.01)

                st.markdown(f"<div style='margin-top:6px; color:{COLORS['text_soft']}; font-size:0.82rem;'>"
                            f"Cost composition — updates live with the sliders above</div>", unsafe_allow_html=True)
                inflation_addon = round(base_cost * food_inf / 100, 3)
                net_after_subsidy = round(base_cost + inflation_addon - subsidy_offset, 2)
                donut = go.Figure(go.Pie(
                    labels=["Base cost", "Inflation add-on"],
                    values=[base_cost, max(inflation_addon, 0.001)],
                    hole=0.64, sort=False, textinfo="percent",
                    marker=dict(colors=[COLORS["text_soft"], COLORS["danger"]],
                                line=dict(color=COLORS["surface"], width=3)),
                ))
                donut.update_layout(
                    **PLOTLY_DARK, height=220,
                    legend=dict(orientation="h", y=-0.12, font=dict(color=COLORS["text_soft"])),
                    annotations=[dict(
                        text=f"<b>AED {net_after_subsidy:.2f}</b><br><span style='font-size:10px;color:{COLORS['text_soft']}'>net / unit</span>",
                        x=0.5, y=0.5, font=dict(size=18, color=COLORS['text']), showarrow=False)],
                )
                st.plotly_chart(donut, width='stretch', config={"displayModeBar": False}, key="inflation_donut")

        adj_cost_headline = round(base_cost * (1 + headline_inf / 100) - subsidy_offset, 3)
        adj_cost_food = round(base_cost * (1 + food_inf / 100) - subsidy_offset, 3)
        base_revenue_per_unit = base_cost / (baseline["food_cost_pct"] / 100)
        food_cost_pct_adj = round((adj_cost_food / base_revenue_per_unit) * 100, 1)

        with colR:
            m1, m2, m3 = st.columns(3)
            m1.metric("Cost/unit — headline scenario", f"AED {adj_cost_headline:.2f}")
            m2.metric("Cost/unit — food-inflation scenario", f"AED {adj_cost_food:.2f}",
                      f"{adj_cost_food - base_cost:+.2f} vs current")
            m3.metric("Food cost % — food-inflation scenario", f"{food_cost_pct_adj}%",
                      f"{food_cost_pct_adj - baseline['food_cost_pct']:+.1f}pt", delta_color="inverse")

            with st.container(border=True):
                fig1 = go.Figure(go.Bar(
                    x=["Current", "Headline scenario", "Food-inflation scenario"],
                    y=[base_cost, adj_cost_headline, adj_cost_food],
                    marker_color=[COLORS["text_soft"], COLORS["tertiary"], COLORS["danger"]],
                    marker_line_width=0,
                ))
                fig1.update_layout(**PLOTLY_DARK, height=260, yaxis_title="AED per unit",
                                    yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                st.plotly_chart(fig1, width='stretch', config={"displayModeBar": False})

        st.caption("Context: UAE food inflation has ranged from a 2022 peak near 9% to negative readings in 2021, "
                   "against headline forecasts around 2–3% for 2026 — a single flat assumption misses this swing.")

    # ---------------- MODULE 2: NATURAL CALAMITY / SUPPLY DISRUPTION ----------------
    with mod2:
        st.markdown("##### 🌪️ Natural calamity & supply disruption risk")
        st.caption("UAE imports 80–90% of its food, so 'natural calamity' risk for a bakery is mainly an import-shock risk.")

        scenarios = {
            "No disruption": (0, 0, 0.02),
            "Port congestion": (7, 6, 0.15),
            "Strait chokepoint closure": (18, 14, 0.35),
            "Extreme heat / sandstorm logistics delay": (3, 3, 0.10),
            "Custom": None,
        }
        choice = st.selectbox("Disruption scenario", list(scenarios.keys()), index=2)

        if choice == "Custom":
            colA, colB, colC = st.columns(3)
            delay_days = colA.slider("Lead-time extension (days)", 0, 30, 10)
            cost_premium = colB.slider("Cost premium (%)", 0, 40, 10)
            stockout_prob = colC.slider("Stockout probability", 0.0, 1.0, 0.2, 0.05)
        else:
            delay_days, cost_premium, stockout_prob = scenarios[choice]
            colA, colB, colC = st.columns(3)
            colA.metric("Lead-time extension", f"{delay_days} days")
            colB.metric("Cost premium", f"{cost_premium}%")
            colC.metric("Stockout probability", f"{stockout_prob:.0%}")

        avg_daily_cost = 4200
        buffer_stock_days = max(delay_days, 3)
        buffer_stock_cost = round(buffer_stock_days * avg_daily_cost * 0.02, 0)
        expected_stockout_cost = round(stockout_prob * avg_daily_cost * delay_days * 1.5, 0)

        with st.container(border=True):
            st.markdown(
                f"<div style='padding:6px 4px;'><h4 style='margin-top:0;'>Recommended buffer vs expected stockout cost</h4>"
                f"Holding <b style='color:{COLORS['primary']}'>{buffer_stock_days} days</b> of buffer stock costs approximately "
                f"<b style='color:{COLORS['primary']}'>AED {buffer_stock_cost:,.0f}</b>. The expected cost of not holding buffer, given this scenario's "
                f"probability, is approximately <b style='color:{COLORS['danger']}'>AED {expected_stockout_cost:,.0f}</b>.</div>",
                unsafe_allow_html=True)

            fig2 = go.Figure(go.Bar(
                x=["Cost of holding buffer stock", "Expected cost of stockout"],
                y=[buffer_stock_cost, expected_stockout_cost],
                marker_color=[COLORS["primary"], COLORS["danger"]], marker_line_width=0,
            ))
            fig2.update_layout(**PLOTLY_DARK, height=260, yaxis_title="AED",
                                yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig2, width='stretch', config={"displayModeBar": False})

    # ---------------- MODULE 3: PANDEMIC PREPAREDNESS ----------------
    with mod3:
        st.markdown("##### 🦠 Pandemic preparedness")
        st.caption("Two linked panels: customer retention/attraction (demand side) and supply chain optimization (supply side).")

        colX, colY = st.columns(2)
        with colX:
            with st.container(border=True):
                st.markdown(f"<h4 style='color:{COLORS['secondary']}'>👥 Customer retention / attraction</h4>", unsafe_allow_html=True)
                repeat_rate = st.slider("Repeat-purchase rate (%)", 0, 100, 62, key="repeat")
                delivery_share = st.slider("Delivery/online-order revenue share (%)", 0, 100, 28, key="delivery")
                basket_size = st.slider("Average basket size (AED)", 20, 200, 74, key="basket")
                flags = []
                if repeat_rate < 50: flags.append("Repeat-purchase rate below healthy baseline")
                if delivery_share < 15: flags.append("Delivery share low — limited resilience to footfall shocks")
                badge, label = ("pill-risk", "Needs attention") if flags else ("pill-ok", "Within healthy range")
                st.markdown(f"<span class='pill {badge}'>{label}</span>", unsafe_allow_html=True)
                for f in flags:
                    st.caption(f"⚠️ {f}")

        with colY:
            with st.container(border=True):
                st.markdown(f"<h4 style='color:{COLORS['primary']}'>🚚 Supply chain optimization</h4>", unsafe_allow_html=True)
                safety_days = st.slider("Safety stock (days of cover)", 0, 30, 9, key="safety")
                alt_suppliers = st.slider("Qualified alternate suppliers per key ingredient", 0, 5, 1, key="alt")
                single_sourced = st.slider("% ingredients single-sourced", 0, 100, 45, key="single")
                flags2 = []
                if safety_days < 7: flags2.append("Safety stock below 7-day resilience threshold")
                if alt_suppliers < 2: flags2.append("Fewer than 2 qualified alternates — concentration risk")
                if single_sourced > 40: flags2.append("Over 40% of ingredients single-sourced")
                badge2, label2 = ("pill-risk", "Needs attention") if flags2 else ("pill-ok", "Within healthy range")
                st.markdown(f"<span class='pill {badge2}'>{label2}</span>", unsafe_allow_html=True)
                for f in flags2:
                    st.caption(f"⚠️ {f}")

        st.caption("Rationale: pandemics create a dual shock — demand shifting from in-store to delivery, and "
                   "supply disruption at the same time — so both sides need to be tracked together as an early-warning system.")

    # ---------------- MODULE 4: SUPPLIER ALTERNATIVES / CONCENTRATION ----------------
    with mod4:
        st.markdown("##### 🌾 Raw material & supplier alternatives")
        st.caption("Herfindahl–Hirschman-style concentration index for core bakery inputs (e.g., flour), "
                   "following the measurement approach used in UAE cereal-supply-risk research (Ali et al., 2022).")

        default_suppliers = pd.DataFrame({
            "Supplier / origin": ["Russia", "Canada", "Australia", "India", "Other"],
            "Spend share (%)": [31, 28, 12, 20, 9],
        })
        st.caption("Edit the spend shares below to reflect current or hypothetical sourcing mix (must sum to ~100%).")
        edited = st.data_editor(default_suppliers, num_rows="fixed", width='stretch', hide_index=True)

        shares = edited["Spend share (%)"].astype(float)
        total = shares.sum()
        shares_norm = shares / total * 100 if total > 0 else shares
        hhi = round((shares_norm ** 2).sum(), 0)

        if hhi < 1500:
            risk_label, risk_badge = "Low concentration risk", "pill-ok"
        elif hhi < 2500:
            risk_label, risk_badge = "Moderate concentration risk", "pill-warn"
        else:
            risk_label, risk_badge = "High concentration risk", "pill-risk"

        colM, colN = st.columns([1, 2])
        with colM:
            with st.container(border=True):
                st.markdown(f"<div class='kpi-label'>Supplier concentration index (HHI)</div>"
                            f"<div class='kpi-value' style='font-size:2.4rem;'>{hhi:,.0f}</div>"
                            f"<span class='pill {risk_badge}'>{risk_label}</span>", unsafe_allow_html=True)
                st.caption("Reference bands: <1,500 low · 1,500–2,500 moderate · >2,500 high (standard HHI convention).")

        with colN:
            with st.container(border=True):
                fig3 = go.Figure(go.Pie(labels=edited["Supplier / origin"], values=shares_norm,
                                         hole=0.55, marker=dict(colors=CATEGORICAL, line=dict(color=COLORS["surface"], width=2))))
                fig3.update_layout(**PLOTLY_DARK, height=280, legend=dict(font=dict(color=COLORS["text_soft"])))
                st.plotly_chart(fig3, width='stretch', config={"displayModeBar": False})

        st.markdown("###### Pre-identified alternate suppliers (illustrative)")
        alt_df = pd.DataFrame({
            "Alternate supplier": ["Local UAE mill (blended flour)", "Turkey-origin supplier", "Egypt-origin supplier"],
            "Cost delta vs current": ["+4.5%", "+2.1%", "+1.8%"],
            "Lead-time delta": ["-3 days", "-1 day", "0 days"],
        })
        st.dataframe(alt_df, width='stretch', hide_index=True)

    scenario_context = {
        "inflation": {
            "headline_inf": headline_inf, "food_inf": food_inf, "subsidy_offset": subsidy_offset,
            "base_cost": base_cost, "adj_cost_headline": adj_cost_headline, "adj_cost_food": adj_cost_food,
            "food_cost_pct_adj": food_cost_pct_adj,
        },
        "disruption": {
            "scenario": choice, "delay_days": delay_days, "cost_premium": cost_premium,
            "stockout_prob": stockout_prob, "buffer_stock_cost": buffer_stock_cost,
            "expected_stockout_cost": expected_stockout_cost,
        },
        "pandemic": {
            "repeat_rate": repeat_rate, "delivery_share": delivery_share, "basket_size": basket_size,
            "safety_days": safety_days, "alt_suppliers": alt_suppliers, "single_sourced": single_sourced,
        },
        "supplier": {"hhi": hhi, "risk_label": risk_label, "shares_df": edited, "alt_df": alt_df},
    }
    render_email_share("scenario_resilience", "Email this Scenario & Resilience report",
                        section="scenario_resilience", context=scenario_context)

# ========================================================================
# SECTION C — EMPLOYEE PORTAL
# ========================================================================
elif section == "👤  Employee portal":
    st.caption("Employee self-service performance tracking — bakery division")

    storage_mode = portal_storage_mode()
    badge_cls = "pill-ok" if "Supabase" in storage_mode else "pill-warn"
    st.markdown(f"<span class='pill {badge_cls}'>💾 Storage: {storage_mode}</span>", unsafe_allow_html=True)
    if "Supabase" not in storage_mode:
        with st.expander("⚙️ Set up persistent storage (Supabase)"):
            st.markdown(
                "The portal works right now, but entries only last for this browser session. "
                "To make them persistent across visits and deployments, connect a free Supabase project — "
                "see the **Employee portal setup** section in the README for the exact SQL and secrets to add."
            )

    with st.expander("🏭 Staff breakdown by line", expanded=False):
        dep_df = pd.DataFrame(
            [{"Department": d, "Headcount": c,
              "Focus": DEPARTMENT_PRIORITY_TAG.get(d, ("", ""))[0]} for d, c in DEPARTMENT_HEADCOUNT.items()]
        )
        dep_df.loc[len(dep_df)] = ["Total (itemized bakery lines)", dep_df["Headcount"].sum(), ""]
        st.dataframe(dep_df, width='stretch', hide_index=True)
        cA, cB = st.columns(2)
        cA.metric("Company-wide headcount (current)", COMPANY_HEADCOUNT_CURRENT)
        cB.metric("Growth target (within 1 year)", COMPANY_HEADCOUNT_TARGET,
                   f"+{COMPANY_HEADCOUNT_TARGET - COMPANY_HEADCOUNT_CURRENT} needed")
        st.caption("Company-wide figures per the 27 Jul MoM — larger than the itemized bakery-line total above, "
                   "likely including catering/other roles not broken out per line. Approximate headcounts overall; "
                   "a small illustrative sample of named employees per department is used for the interactive "
                   "demo below — replace with the full roster once available.")

    st.markdown("#### Who are you?")
    col_d1, col_d2 = st.columns([1, 1.4])
    with col_d1:
        dept_choice = st.selectbox("Department", ["🧑‍💼 Manager / HR view"] + list(DEPARTMENT_HEADCOUNT.keys()))

    identity_employee = None
    if dept_choice != "🧑‍💼 Manager / HR view":
        dept_employees = [e for e in EMPLOYEES if e["department"] == dept_choice]
        emp_labels = [f"{e['name']} ({e['id']})" for e in dept_employees]
        with col_d2:
            emp_choice = st.selectbox("Employee", ["— Select —"] + emp_labels)
        if emp_choice != "— Select —":
            identity_employee = dept_employees[emp_labels.index(emp_choice)]

    # -------------------- EMPLOYEE SELF-SERVICE --------------------
    if identity_employee is not None:
        emp_name, emp_id, emp_dept = identity_employee["name"], identity_employee["id"], identity_employee["department"]
        tag_label, tag_cls = DEPARTMENT_PRIORITY_TAG.get(emp_dept, (None, None))
        tag_html = f"<span class='pill {tag_cls}' style='margin-left:6px;'>{tag_label}</span>" if tag_label else ""
        st.markdown(f"<span class='pill pill-ok'>{emp_dept}</span>{tag_html}", unsafe_allow_html=True)

        etab1, etab2, etab3 = st.tabs(["📊 My performance", "📝 Daily log", "🎯 Goals & feedback"])

        # --- My performance ---
        with etab1:
            perf_df, err = portal_fetch("employee_performance", {"employee_id": emp_id})
            if err:
                st.error(f"Could not load performance data: {err}")
            elif perf_df.empty:
                pass
            else:
                for col in ["units_produced", "wastage_pct", "hours_worked", "batch_time_adherence_pct",
                            "quality_pass_pct", "revenue_generated"]:
                    if col in perf_df.columns:
                        perf_df[col] = pd.to_numeric(perf_df[col], errors="coerce")

                avg_wastage = perf_df["wastage_pct"].mean()
                wastage_pill = ("pill-ok" if avg_wastage <= WASTAGE_TARGET_PCT
                                 else "pill-warn" if avg_wastage <= 3 else "pill-risk")

                c1, c2, c3, c4 = st.columns(4)
                with c1, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Avg units / shift</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['units_produced'].mean():.0f}</div>",
                                unsafe_allow_html=True)
                with c2, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Avg wastage %</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{avg_wastage:.1f}%</div>"
                                f"<span class='pill {wastage_pill}'>Target ≤{WASTAGE_TARGET_PCT:.0f}%</span>",
                                unsafe_allow_html=True)
                with c3, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Batch-time adherence</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['batch_time_adherence_pct'].mean():.0f}%</div>",
                                unsafe_allow_html=True)
                with c4, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Quality pass rate</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['quality_pass_pct'].mean():.0f}%</div>",
                                unsafe_allow_html=True)

                # --- Productivity benchmark (AED 1,000/day per GM's rule of thumb) ---
                c5, c6 = st.columns(2)
                with c5, st.container(border=True):
                    avg_hours = perf_df["hours_worked"].mean() if perf_df["hours_worked"].mean() > 0 else STANDARD_SHIFT_HOURS
                    avg_revenue = perf_df["revenue_generated"].mean() if "revenue_generated" in perf_df.columns else 0.0
                    value_per_day = (avg_revenue / avg_hours) * STANDARD_SHIFT_HOURS if avg_hours else avg_revenue
                    val_pill = "pill-ok" if value_per_day >= VALUE_PER_EMPLOYEE_DAY_TARGET else "pill-warn"
                    st.markdown(f"<div class='kpi-label'>Value generated / day</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>AED {value_per_day:,.0f}</div>"
                                f"<span class='pill {val_pill}'>GM benchmark: AED {VALUE_PER_EMPLOYEE_DAY_TARGET:,.0f}/day</span>",
                                unsafe_allow_html=True)
                with c6, st.container(border=True):
                    cap_target = DEPARTMENT_CAPACITY_TARGET.get(emp_dept)
                    if cap_target:
                        util_pct = (perf_df["units_produced"].mean() / cap_target) * 100
                        util_pill = "pill-ok" if util_pct >= 70 else "pill-warn" if util_pct >= 40 else "pill-risk"
                        st.markdown(f"<div class='kpi-label'>Capacity utilization ({emp_dept})</div>"
                                    f"<div class='kpi-value' style='font-size:1.7rem;'>{util_pct:.0f}%</div>"
                                    f"<span class='pill {util_pill}'>vs illustrative target {cap_target}/day</span>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div class='kpi-label'>Capacity utilization</div>"
                                    f"<div class='kpi-value' style='font-size:1rem; color:{COLORS['text_soft']};'>Not tracked for this line</div>",
                                    unsafe_allow_html=True)

                with st.container(border=True):
                    plot_df = perf_df.sort_values("log_date") if "log_date" in perf_df.columns else perf_df
                    fig_e = go.Figure()
                    fig_e.add_trace(go.Scatter(x=plot_df.get("log_date", plot_df.index), y=plot_df["units_produced"],
                                                mode="lines+markers", name="Units produced",
                                                line=dict(color=COLORS["primary"], width=2.5)))
                    fig_e.update_layout(**PLOTLY_DARK, height=260, yaxis=dict(gridcolor=GRID_COLOR),
                                        xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                                        legend=dict(font=dict(color=COLORS["text_soft"])))
                    st.plotly_chart(fig_e, width='stretch', config={"displayModeBar": False})

                # --- Shift breakdown (Arabic Bread runs a separate night shift) ---
                if emp_dept == "Arabic Bread" and "shift" in perf_df.columns and perf_df["shift"].nunique() > 1:
                    with st.container(border=True):
                        st.markdown("###### Morning vs. night shift")
                        shift_summary = perf_df.groupby("shift").agg(
                            avg_units=("units_produced", "mean"),
                            avg_wastage_pct=("wastage_pct", "mean"),
                            shifts_logged=("shift", "count"),
                        ).round(1).reset_index()
                        st.dataframe(shift_summary, width='stretch', hide_index=True)

                st.dataframe(perf_df.drop(columns=[c for c in ["id"] if c in perf_df.columns]),
                             width='stretch', hide_index=True)

        # --- Daily log ---
        with etab2:
            st.caption("Log today's (or a past) shift in plain counts — no percentages to work out yourself. "
                       "Your performance report will still show the percentages, calculated automatically.")
            with st.form(f"daily_log_form_{emp_id}", clear_on_submit=True):
                log_date = st.date_input("Date", value=datetime.now().date())
                if emp_dept == "Arabic Bread":
                    shift = st.selectbox("Shift", ["Morning", "Night"],
                                          help="Arabic Bread runs a separate night shift in addition to the standard morning shift.")
                else:
                    shift = "Morning"
                colf1, colf2 = st.columns(2)
                units = colf1.number_input("Units produced (good units)", 0, 5000, 250, 10)
                units_wasted = colf2.number_input("Units wasted / discarded", 0, 1000, 6, 1,
                                                   help="How many units were thrown away, spoiled, or damaged this shift.")
                hours = colf1.number_input("Hours worked", 0.0, 16.0, STANDARD_SHIFT_HOURS, 0.5)
                units_failed_qc = colf2.number_input("Units that failed QC check", 0, 1000, 5, 1,
                                                       help="How many units failed a quality check this shift.")
                colf3, colf4 = st.columns(2)
                batches_completed = colf3.number_input("Batches completed", 0, 200, 10, 1)
                batches_on_time = colf4.number_input("Batches finished on/within standard time", 0, 200, 9, 1,
                                                       help="Of the batches completed, how many finished within the standard time for that batch?")
                revenue = st.number_input(
                    "Revenue/value generated (AED, optional)", 0.0, 50000.0, 0.0, 50.0,
                    help="Feeds the productivity benchmark: GM's rule of thumb is each employee should "
                         "generate ~AED 1,000/day. Leave at 0 if not tracked yet for this shift."
                )
                notes = st.text_area("Notes (optional)", placeholder="Any incidents, equipment issues, etc.")
                submitted = st.form_submit_button("Submit shift log", width='stretch')
                if submitted:
                    total_output = units + units_wasted
                    wastage_pct = round((units_wasted / total_output) * 100, 1) if total_output > 0 else 0.0
                    batch_adh_pct = round((batches_on_time / batches_completed) * 100) if batches_completed > 0 else 0
                    quality_pct = round(max(units - units_failed_qc, 0) / units * 100) if units > 0 else 0
                    ok, err = portal_insert("employee_performance", {
                        "employee_name": emp_name, "employee_id": emp_id, "department": emp_dept,
                        "log_date": str(log_date), "shift": shift, "units_produced": int(units),
                        "units_wasted": int(units_wasted), "batches_completed": int(batches_completed),
                        "batches_on_time": int(batches_on_time), "units_failed_qc": int(units_failed_qc),
                        "wastage_pct": wastage_pct, "hours_worked": float(hours),
                        "batch_time_adherence_pct": batch_adh_pct, "quality_pass_pct": quality_pct,
                        "revenue_generated": float(revenue), "notes": notes,
                    })
                    if ok:
                        st.toast(f"Shift log saved — wastage {wastage_pct}%, batch adherence {batch_adh_pct}%, "
                                 f"quality pass {quality_pct}% (calculated automatically from your counts).", icon="✅")
                        st.rerun()
                    else:
                        st.error(f"Could not save: {err}")

        # --- Goals & feedback ---
        with etab3:
            goals_df, err = portal_fetch("employee_goals", {"employee_id": emp_id})
            if err:
                st.error(f"Could not load goals: {err}")
            elif not goals_df.empty:
                for _, row in goals_df.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**Goal:** {row.get('goal_text', '—')}")
                        st.caption(f"Self-assessment: {row.get('self_assessment') or '—'}")
                        if row.get("manager_feedback"):
                            st.markdown(f"<span class='pill pill-ok'>Manager feedback</span> {row['manager_feedback']}",
                                        unsafe_allow_html=True)
            else:
                pass

            st.markdown("###### Add a new goal / self-assessment")
            with st.form(f"goals_form_{emp_id}", clear_on_submit=True):
                goal_text = st.text_area("Current goal / focus area")
                self_assessment = st.text_area("Self-assessment for this period")
                submitted_g = st.form_submit_button("Save", width='stretch')
                if submitted_g:
                    ok, err = portal_insert("employee_goals", {
                        "employee_name": emp_name, "employee_id": emp_id, "department": emp_dept,
                        "goal_text": goal_text, "self_assessment": self_assessment,
                        "manager_feedback": None,
                    })
                    if ok:
                        st.toast("Goal saved.", icon="✅")
                        st.rerun()
                    else:
                        st.error(f"Could not save: {err}")

    # -------------------- MANAGER / HR VIEW --------------------
    elif dept_choice == "🧑‍💼 Manager / HR view":
        all_perf, err = portal_fetch("employee_performance")
        if err:
            st.error(f"Could not load team data: {err}")
        elif all_perf.empty:
            pass
        else:
            for col in ["units_produced", "wastage_pct", "batch_time_adherence_pct", "quality_pass_pct", "revenue_generated", "hours_worked"]:
                if col in all_perf.columns:
                    all_perf[col] = pd.to_numeric(all_perf[col], errors="coerce")

            st.markdown("#### Performance comparison within a department")
            available_depts = sorted(all_perf["department"].dropna().unique()) if "department" in all_perf.columns else []
            if not available_depts:
                pass
            else:
                dept_filter = st.selectbox("Choose a department to compare", available_depts, key="mgr_dept_filter")
                tag_label, tag_cls = DEPARTMENT_PRIORITY_TAG.get(dept_filter, (None, None))
                tag_html = f"<span class='pill {tag_cls}' style='margin-left:8px;'>{tag_label}</span>" if tag_label else ""

                dept_perf = all_perf[all_perf["department"] == dept_filter].copy()

                def _value_per_day(row_group):
                    hrs = row_group["hours_worked"].mean()
                    hrs = hrs if hrs and hrs > 0 else STANDARD_SHIFT_HOURS
                    rev = row_group["revenue_generated"].mean() if "revenue_generated" in row_group.columns else 0.0
                    return (rev / hrs) * STANDARD_SHIFT_HOURS

                dept_summary = dept_perf.groupby("employee_name").agg(
                    shifts_logged=("employee_name", "count"),
                    avg_units=("units_produced", "mean"),
                    avg_wastage_pct=("wastage_pct", "mean"),
                    avg_batch_adherence=("batch_time_adherence_pct", "mean"),
                    avg_quality_pass=("quality_pass_pct", "mean"),
                ).round(1).reset_index()
                dept_summary["value_per_day_aed"] = dept_summary["employee_name"].apply(
                    lambda n: round(_value_per_day(dept_perf[dept_perf["employee_name"] == n]), 0))
                dept_summary["wastage_status"] = dept_summary["avg_wastage_pct"].apply(
                    lambda w: "✅ On target" if w <= WASTAGE_TARGET_PCT else "⚠️ Above 1% target")

                # Composite performance score (0-100) so "who's doing best" has a
                # single clear answer instead of reading several charts separately.
                max_units = dept_summary["avg_units"].max()
                dept_summary["units_normalized"] = (dept_summary["avg_units"] / max_units * 100) if max_units > 0 else 0
                dept_summary["performance_score"] = (
                    dept_summary["avg_quality_pass"] * 0.30
                    + dept_summary["avg_batch_adherence"] * 0.30
                    + (100 - dept_summary["avg_wastage_pct"]).clip(lower=0) * 0.20
                    + dept_summary["units_normalized"] * 0.20
                ).round(1)
                dept_summary = dept_summary.sort_values("performance_score", ascending=False).reset_index(drop=True)
                dept_summary["rank"] = dept_summary.index + 1

                cap_target = DEPARTMENT_CAPACITY_TARGET.get(dept_filter)

                with st.container(border=True):
                    st.markdown(f"**{dept_filter}**{tag_html} — {DEPARTMENT_HEADCOUNT.get(dept_filter, '—')} total headcount "
                                f"on the line, {dept_perf['employee_name'].nunique()} with logged shifts",
                                unsafe_allow_html=True)
                    if cap_target:
                        util_pct = (dept_summary["avg_units"].mean() / cap_target) * 100
                        st.caption(f"Capacity utilization vs illustrative target of {cap_target} units/day/employee: **{util_pct:.0f}%**")
                    if len(dept_summary) > 0:
                        top = dept_summary.iloc[0]
                        st.markdown(f"<span class='pill pill-ok'>🏆 Top performer: {top['employee_name']} "
                                    f"(score {top['performance_score']:.0f}/100)</span>", unsafe_allow_html=True)
                    display_cols = ["rank", "employee_name", "performance_score", "shifts_logged", "avg_units",
                                     "avg_wastage_pct", "avg_batch_adherence", "avg_quality_pass",
                                     "value_per_day_aed", "wastage_status"]
                    st.dataframe(dept_summary[display_cols], width='stretch', hide_index=True)
                    st.caption("Performance score = 30% quality pass rate + 30% batch-time adherence + "
                               "20% (100 − wastage %) + 20% units produced (normalized to the department's top performer).")
                    render_email_share(
                        f"dept_{dept_filter}", f"Email the {dept_filter} comparison",
                        section="employee_department",
                        context={
                            "department": dept_filter,
                            "headcount": DEPARTMENT_HEADCOUNT.get(dept_filter, "—"),
                            "tag_label": tag_label,
                            "summary_df": dept_summary[display_cols],
                        },
                    )

                with st.container(border=True):
                    st.markdown("###### Comparative progress across employees in this department")
                    fig_cmp = go.Figure()
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_quality_pass"],
                                              name="Quality pass %", marker_color=COLORS["primary"], marker_line_width=0))
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_batch_adherence"],
                                              name="Batch-time adherence %", marker_color=COLORS["tertiary"], marker_line_width=0))
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"],
                                              y=(100 - dept_summary["avg_wastage_pct"]).clip(lower=0),
                                              name="100 − wastage %", marker_color=COLORS["secondary"], marker_line_width=0))
                    fig_cmp.update_layout(**PLOTLY_DARK, height=300, barmode="group", yaxis_title="%",
                                           yaxis=dict(gridcolor=GRID_COLOR, range=[0, 105]),
                                           xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                                           legend=dict(orientation="h", y=-0.15, font=dict(color=COLORS["text_soft"])))
                    st.plotly_chart(fig_cmp, width='stretch', config={"displayModeBar": False})

                with st.container(border=True):
                    fig_dept = go.Figure()
                    fig_dept.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_units"],
                                               name="Avg units/shift", marker_color=COLORS["primary"], marker_line_width=0))
                    fig_dept.update_layout(**PLOTLY_DARK, height=260, yaxis_title="Avg units / shift",
                                            yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                    st.plotly_chart(fig_dept, width='stretch', config={"displayModeBar": False})

                with st.container(border=True):
                    fig_dept2 = go.Figure(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_wastage_pct"],
                                                  marker_color=COLORS["secondary"], marker_line_width=0))
                    fig_dept2.add_hline(y=WASTAGE_TARGET_PCT, line_dash="dot", line_color=COLORS["text_soft"],
                                         annotation_text="1% target", annotation_font_color=COLORS["text_soft"])
                    fig_dept2.update_layout(**PLOTLY_DARK, height=260, yaxis_title="Avg wastage %",
                                             yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                    st.plotly_chart(fig_dept2, width='stretch', config={"displayModeBar": False})

                if dept_filter == "Arabic Bread" and "shift" in dept_perf.columns and dept_perf["shift"].nunique() > 1:
                    with st.container(border=True):
                        st.markdown("###### Morning vs. night shift — Arabic Bread")
                        shift_cmp = dept_perf.groupby("shift").agg(
                            avg_units=("units_produced", "mean"),
                            avg_wastage_pct=("wastage_pct", "mean"),
                            shifts_logged=("shift", "count"),
                        ).round(1).reset_index()
                        st.dataframe(shift_cmp, width='stretch', hide_index=True)

            st.markdown("#### All departments — overview")
            dept_overview = all_perf.groupby("department").agg(
                shifts_logged=("department", "count"),
                avg_units=("units_produced", "mean"),
                avg_wastage_pct=("wastage_pct", "mean"),
            ).round(1).reset_index() if "department" in all_perf.columns else pd.DataFrame()
            if not dept_overview.empty:
                dept_overview["headcount"] = dept_overview["department"].map(DEPARTMENT_HEADCOUNT)
                dept_overview["capacity_target"] = dept_overview["department"].map(DEPARTMENT_CAPACITY_TARGET)
                dept_overview["utilization_pct"] = (dept_overview["avg_units"] / dept_overview["capacity_target"] * 100).round(0)
                dept_overview["focus"] = dept_overview["department"].map(lambda d: DEPARTMENT_PRIORITY_TAG.get(d, ("", ""))[0])
                with st.container(border=True):
                    st.dataframe(dept_overview.sort_values("avg_units", ascending=False), width='stretch', hide_index=True)

        st.markdown("#### Recent goals & self-assessments")
        all_goals, err = portal_fetch("employee_goals")
        if err:
            st.error(f"Could not load goals: {err}")
        elif all_goals.empty:
            pass
        else:
            for _, row in all_goals.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{row.get('employee_name','—')}** ({row.get('department','—')}) — {row.get('goal_text','—')}")
                    st.caption(f"Self-assessment: {row.get('self_assessment') or '—'}")
                    if row.get("manager_feedback"):
                        st.markdown(f"<span class='pill pill-ok'>Feedback given</span> {row['manager_feedback']}",
                                    unsafe_allow_html=True)

        st.markdown("###### Give feedback on an employee's goal")
        with st.form("manager_feedback_form", clear_on_submit=True):
            fb_labels = [f"{e['department']} — {e['name']}" for e in EMPLOYEES]
            fb_choice = st.selectbox("Employee", fb_labels, key="fb_employee")
            fb_text = st.text_area("Feedback")
            fb_submit = st.form_submit_button("Submit feedback", width='stretch')
            if fb_submit:
                fb_emp = EMPLOYEES[fb_labels.index(fb_choice)]
                ok, err = portal_insert("employee_goals", {
                    "employee_name": fb_emp["name"], "employee_id": fb_emp["id"], "department": fb_emp["department"],
                    "goal_text": "(Manager feedback entry)", "self_assessment": None,
                    "manager_feedback": fb_text,
                })
                if ok:
                    st.toast("Feedback recorded.", icon="✅")
                    st.rerun()
                else:
                    st.error(f"Could not save: {err}")
    else:
        st.info("Select your department and name (or Manager / HR view) above to get started.")

# ========================================================================
# SECTION D — COMPANY PROFILE
# ========================================================================
elif section == "📋  Company profile":
    st.caption("Everything collected from the Grandiose team so far — from the 27 Jul personal meeting with the GM")

    COMPANY_INFO = [
        ("🏭", "Production capacity & growth", [
            "Flour Country Bakery is currently running at only 30-35% utilization — roughly 65% headroom still available.",
            "Target is 100% utilization within a year, with an estimated AED 10M/month production potential once reached.",
            "Granola Catering is running at approximately 50% utilization.",
            "Company-wide headcount is ~110 today, with a growth target of 200 within a year.",
        ]),
        ("⏰", "Shift structure", [
            "Standard shift is 9 hours, including a break.",
            "Arabic Bread is the one line that also runs a separate night shift, in addition to the standard morning shift.",
        ]),
        ("🛒", "Sales channel mix (B2B vs. B2C)", [
            "Flour Country Bakery is 100% B2B.",
            "Catering is roughly 35% B2C and 65% B2B.",
        ]),
        ("💰", "Cost structure & targets", [
            "Food cost: capped at a maximum of 30%.",
            "Labour cost: currently running at ~40% — target is 16-18%, with 20% as the absolute ceiling.",
            "Utilities: capped at a maximum of 10%.",
            "Rent: capped at a maximum of 15%.",
            "Net profit target: 20-25%.",
            "Total expense ceiling (all costs combined): 75-80% of revenue.",
        ]),
        ("🏷️", "Pricing approach", [
            "Material cost is scaled 35-50% depending on client volume.",
            "Supermarket supply is priced cost-to-cost, with no margin built in.",
            "B2B contract pricing varies by negotiation and account.",
        ]),
        ("📈", "Productivity benchmark", [
            "No formal labour productivity measurement system exists yet.",
            "The GM's working rule of thumb: each employee should generate approximately AED 1,000/day of value.",
            "This is the benchmark now built into the Employee Portal's productivity tracking.",
        ]),
        ("🔥", "Top & underutilized product lines", [
            "Top revenue/profit drivers: daily bread (Arabic bread, toasted bread, buns, Samurai sandwich) and pastry.",
            "Least-utilized lines: Tahina and Turkish baklava — both flagged as ramp-up focus areas.",
        ]),
        ("📦", "Inventory management", [
            "Stock is managed on a FIFO (first-in, first-out) basis.",
            "Ordering cycles: bi-monthly for shelf-stable/dry goods, weekly for short-shelf-life items, daily for fresh fruit and vegetables.",
            "Payment terms: ~95% of suppliers are on net 90 days (60 days from statement of account); a few critical items run 30-60 days.",
            "Approved vendor base: ~65-70 for bakery, ~160-170 for catering.",
            "Safety stock normally aligns to supplier lead times, extending to ~3 months during geopolitical disruptions.",
            "Packaging (sourced from China/Indonesia): ~1 month of stock held against a 15-day lead time.",
            "De-risking in progress: alternate vendors are being onboarded for critical ingredients.",
            "KPIs monitored: inventory movement (fast/slow movers), depletion cycles, and overstock/obsolescence — reviewed bi-weekly.",
        ]),
        ("🗑️", "Wastage management", [
            "Current wastage: 2-3% — target is 1%.",
            "Plan to close the gap: grow sales volume and push supermarket promotions on affected items.",
            "Near-expiry raw materials are redirected into production rather than discarded.",
            "Near-expiry finished goods (with shelf life over 30 days) move to clearance pricing.",
            "Fully expired stock is written off.",
            "Yield: catering runs ~90-95% (about 5% loss to spillage/damage); bakery yield loss is minimal; frozen raw meat yield loss is 20-25%.",
        ]),
        ("🗓️", "Dashboard & next steps (from this meeting)", [
            "ERP integration is planned so the dashboard can eventually auto-update from live company data.",
            "In-dashboard email sharing was requested — ✅ already built into this dashboard.",
            "An employee productivity module linking individual output to section-level sales was requested — ✅ already built into the Employee Portal.",
            "A separate, equivalent dashboard is planned for Granola Catering.",
            "The GM wants the dashboard link shared for internal review before Friday.",
            "Next steps: send Friday feedback slots, share the dashboard link, and continue refining with confirmed figures ahead of the mid-review.",
        ]),
    ]

    for icon, title, points in COMPANY_INFO:
        with st.container(border=True):
            rows = "".join([f"<div class='info-row'><span>{p}</span></div>" for p in points])
            st.markdown(f"<div style='padding:14px 18px;'><h4 style='margin:0 0 10px 0; font-size:1.05rem;'>"
                        f"{icon} {title}</h4>{rows}</div>", unsafe_allow_html=True)

    st.caption("Source: personal meeting with the Grandiose GM, 27 Jul. Figures here are as reported in that "
               "conversation and should be treated as the current source of truth pending any further updates.")
    render_email_share("company_profile", "Email the Company Profile", section="company_profile",
                        context={"company_info": COMPANY_INFO})

# ========================================================================
# SECTION E — DATA PROCESSOR (AI-powered rough-data cleanup)
# ========================================================================
elif section == "📥  Data processor":
    st.caption("Drop in rough Excel, PDF, or Word files — get back a cleaned, evaluated Excel report")

    ai_client = get_anthropic_client()
    badge_cls = "pill-ok" if ai_client is not None else "pill-warn"
    ai_status = "AI processing: connected" if ai_client is not None else "AI processing: not configured"
    st.markdown(f"<span class='pill {badge_cls}'>🤖 {ai_status}</span>", unsafe_allow_html=True)
    if ai_client is None:
        with st.expander("⚙️ Set up AI processing (Anthropic API key)"):
            st.markdown(
                "This feature reads whatever's in your uploaded files — even messy or inconsistently "
                "formatted data — and turns it into a cleaned, evaluated Excel report. It needs an "
                "Anthropic API key. See the **Data processor setup** section in the README for where "
                "to get one and exactly what to add to Streamlit Cloud's Secrets panel."
            )

    st.markdown("#### Upload files")
    uploaded_files = st.file_uploader(
        "Drop Excel, PDF, or Word files here",
        type=["xlsx", "xls", "csv", "pdf", "docx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        st.caption(f"{len(uploaded_files)} file(s) ready: " + ", ".join(f.name for f in uploaded_files))

    process_clicked = st.button("Process files", width='stretch', disabled=not uploaded_files)

    if process_clicked and uploaded_files:
        with st.spinner("Reading files and asking the AI to interpret and evaluate the data..."):
            file_contents = {}
            for f in uploaded_files:
                fbytes = f.getvalue()
                ext = f.name.lower().rsplit(".", 1)[-1]
                if ext == "pdf":
                    file_contents[f.name] = extract_pdf_text(fbytes)
                elif ext == "docx":
                    file_contents[f.name] = extract_docx_text(fbytes)
                else:
                    file_contents[f.name] = extract_excel_text(fbytes, f.name)
            result, err = interpret_uploaded_files(file_contents)
        if err:
            st.error(err)
            st.session_state.pop("smart_upload_result", None)
        else:
            st.session_state["smart_upload_result"] = result
            st.session_state["smart_upload_files"] = [f.name for f in uploaded_files]
            st.toast("Files processed.", icon="✅")

    result = st.session_state.get("smart_upload_result")
    if result:
        st.markdown("#### Results")
        with st.container(border=True):
            st.markdown(f"**Detected data type:** {result.get('detected_data_type', '—')}")
            st.caption(result.get("summary", ""))

        for sheet_info in result.get("sheets", []):
            with st.container(border=True):
                st.markdown(f"###### {sheet_info.get('title', sheet_info.get('sheet_name', 'Sheet'))}")
                columns = sheet_info.get("columns", [])
                rows = sheet_info.get("rows", [])
                if columns and rows:
                    preview_df = pd.DataFrame(
                        [r + [""] * (len(columns) - len(r)) if len(r) < len(columns) else r[:len(columns)] for r in rows],
                        columns=columns,
                    )
                    st.dataframe(preview_df, width='stretch', hide_index=True)
                for ins in sheet_info.get("insights", []):
                    st.caption(f"💡 {ins}")

        excel_bytes = build_excel_report("smart_upload", context={
            "result": result, "source_files": st.session_state.get("smart_upload_files", []),
        })
        st.download_button(
            "Download processed Excel report", data=excel_bytes,
            file_name=f"Grandiose_Bakery_Processed_Data_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )
        render_email_share("smart_upload", "Email this processed report", section="smart_upload",
                            context={"result": result, "source_files": st.session_state.get("smart_upload_files", [])})

    st.caption("Files are sent to Anthropic's API for interpretation and are not stored by this dashboard "
               "beyond the current session. Numbers are AI-interpreted from whatever was uploaded — "
               "always sanity-check before acting on them.")

st.markdown("---")
st.caption("Financial Performance and Cost Optimization for Grandiose Bakery Operations · GIP III · "
           "Bakery division only, catering excluded · Figures shown are illustrative benchmarks pending "
           "Grandiose-provided actuals.")
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------
# COLOR PALETTE — bold dark theme. Near-black canvas with vibrant,
# bakery-warm accents (amber/gold as primary, coral and violet as
# secondary/tertiary) instead of generic blue/purple.
# ----------------------------------------------------------------------
COLORS = {
    "bg":        "#0A0A0A",   # near-black — matches the portfolio site exactly
    "surface":   "#16141E",   # dark plum-tinted card surface
    "surface2":  "#1D1926",   # sidebar / secondary surface
    "border":    "#2E2838",   # muted plum-gray border
    "text":      "#F4F0F8",   # near-white, cool lilac-white
    "text_soft": "#9C93AC",   # muted lilac-gray secondary text
    "primary":   "#B9A0DC",   # soft lilac — main accent
    "primary_d": "#6B5789",   # deeper lilac, for text on light-lilac backgrounds
    "secondary": "#E8AECB",   # soft pink
    "tertiary":  "#8C76B0",   # deeper plum/violet, for chart variety
    "success":   "#C9B8E8",   # pale lilac — used for "on target / good" status
    "warning":   "#D9A8C4",   # dusty rose/mauve — used for "caution" status
    "danger":    "#9C6B8C",   # muted plum-wine — used for "risk" status (deliberately not red)
}

CATEGORICAL = [COLORS["primary"], COLORS["secondary"], "#F0DCEA", COLORS["tertiary"], COLORS["success"]]

def hex_to_rgba(hex_color, alpha):
    hex_color = hex_color.lstrip("#")
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Anton&family=Space+Grotesk:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)

st.markdown(f"""
<style>
    .stApp {{ background-color: {COLORS['bg']}; }}
    section[data-testid="stSidebar"] {{ background-color: {COLORS['surface2']}; border-right: 1px solid {COLORS['border']}; }}

    /* Top app header/toolbar band (Share, star, edit, GitHub icons) —
       white by default, doesn't inherit the dark theme on its own.
       Real testid is "stHeader" (className happens to be "stAppHeader",
       which is not the same thing — that mismatch was the bug last time). */
    [data-testid="stHeader"] {{
        background-color: {COLORS['bg']} !important;
        border-bottom: 1px solid {COLORS['border']} !important;
    }}
    [data-testid="stHeader"] * {{
        color: {COLORS['text']} !important;
        fill: {COLORS['text']} !important;
    }}
    [data-testid="stHeader"] [data-testid^="stBaseButton"] {{
        background-color: {COLORS['surface']} !important;
    }}
    [data-testid="stHeader"] [data-testid^="stBaseButton"] p,
    [data-testid="stHeader"] [data-testid^="stBaseButton"] span {{
        color: {COLORS['text']} !important;
    }}
    [data-testid="stMainMenu"], [data-testid="stStatusWidget"], [data-testid="stAppDeployButton"] {{
        background-color: transparent !important;
        color: {COLORS['text']} !important;
        fill: {COLORS['text']} !important;
    }}
    /* The actual clickable icon is a nested "stMainMenuButton", a different
       component from the other header buttons — this is the specific
       element that was still showing as an unstyled white block. */
    [data-testid="stMainMenuButton"] {{
        background-color: transparent !important;
        color: {COLORS['text']} !important;
    }}
    [data-testid="stMainMenuButton"] svg, [data-testid="stMainMenuButton"] path,
    [data-testid="stMainMenu"] svg, [data-testid="stStatusWidget"] svg {{
        fill: {COLORS['text']} !important;
        color: {COLORS['text']} !important;
    }}

    .stApp, .stApp p, .stApp span, .stApp li, .stApp label,
    div[data-testid="stMarkdownContainer"], div[data-testid="stMarkdownContainer"] p {{
        color: {COLORS['text']} !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }}
    div[data-testid="stCaptionContainer"], .stApp small {{ color: {COLORS['text_soft']} !important; }}
    h1, h2 {{
        color: {COLORS['text']} !important; font-family: 'Anton', sans-serif !important;
        font-weight: 400 !important; letter-spacing: 0.5px;
    }}
    h3, h4, h5 {{
        color: {COLORS['text']} !important; font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {COLORS['text']} !important; font-family: 'Space Grotesk', sans-serif !important; }}

    /* Restore Streamlit's icon-ligature font wherever our broad font-family
       rules above would otherwise break it — this is what renders expander/
       popover arrows correctly instead of as literal "expand_more" text. */
    span[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] * {{
        font-family: 'Material Symbols Rounded' !important;
    }}
    section[data-testid="stSidebar"] span[data-testid="stIconMaterial"] {{
        font-family: 'Material Symbols Rounded' !important;
    }}

    /* Popovers render in a separate layer that doesn't inherit .stApp's dark
       background — without this, forced near-white text sits on Streamlit's
       default white popover background and becomes unreadable. */
    div[data-testid="stPopoverBody"] {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
        border-radius: 14px !important;
    }}
    div[data-testid="stPopoverBody"] * {{
        color: {COLORS['text']} !important;
    }}
    div[data-testid="stPopoverBody"] input, div[data-testid="stPopoverBody"] textarea {{
        background-color: {COLORS['surface2']} !important;
        color: {COLORS['text']} !important;
        border: 1px solid {COLORS['border']} !important;
    }}
    div[data-testid="stPopoverBody"] ::placeholder {{
        color: {COLORS['text_soft']} !important;
        opacity: 1 !important;
    }}

    /* Every floating/portaled element (tooltips, selectbox dropdown menus,
       date pickers, multiselect lists) gets rendered by baseweb OUTSIDE
       .stApp, in a "baseweb-layer" portal attached to <body> — none of it
       inherits our dark theme by default, which is why hover tooltips and
       dropdown option lists have shown up as white-on-white. Fix them all
       at once rather than chasing each one individually. */
    .baseweb-layer, [data-testid="stTooltipContent"], [data-baseweb="menu"],
    [data-baseweb="popover"], [data-baseweb="calendar"] {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
    }}
    .baseweb-layer *, [data-testid="stTooltipContent"] *, [data-baseweb="menu"] *,
    [data-baseweb="popover"] *, [data-baseweb="calendar"] * {{
        color: {COLORS['text']} !important;
        background-color: transparent !important;
    }}
    [data-baseweb="menu"] li:hover, [data-baseweb="menu"] li[aria-selected="true"] {{
        background-color: {hex_to_rgba(COLORS['primary'], 0.18)} !important;
    }}

    /* Every button in the app (st.button, form_submit_button, download_button,
       popover triggers, etc.) — Streamlit's default "secondary" button style
       doesn't have enough contrast once our text-color override is applied,
       so every button gets a consistent solid lilac pill with dark text. */
    [data-testid^="stBaseButton"] {{
        background-color: {COLORS['primary']} !important;
        border: none !important;
        border-radius: 10px !important;
    }}
    [data-testid^="stBaseButton"] p, [data-testid^="stBaseButton"] span,
    [data-testid^="stBaseButton"] div {{
        color: #1A1424 !important;
        font-weight: 700 !important;
    }}
    [data-testid^="stBaseButton"]:hover {{
        background-color: {COLORS['secondary']} !important;
    }}
    [data-testid^="stBaseButton"]:disabled, [data-testid^="stBaseButton"]:disabled p {{
        background-color: {COLORS['border']} !important;
        color: {COLORS['text_soft']} !important;
    }}

    /* Expander headers — safety net so the label never fades to low-contrast */
    div[data-testid="stExpander"] summary p, div[data-testid="stExpander"] summary span {{
        color: {COLORS['text']} !important;
        opacity: 1 !important;
    }}


    /* Bordered containers -> dark elevated cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']} !important;
        border-radius: 18px !important;
        padding: 4px 6px;
    }}

    .kpi-icon {{
        display:inline-flex; align-items:center; justify-content:center;
        width:34px; height:34px; border-radius:10px; font-size:17px; margin-right:8px;
    }}
    .kpi-label {{ color:{COLORS['text_soft']}; font-size:0.82rem; font-weight:500; }}
    .kpi-value {{
        font-size:2.1rem; font-weight:600; margin: 4px 0 2px 0; letter-spacing:-0.5px;
        font-family: 'IBM Plex Mono', monospace !important;
    }}
    .pill {{
        display:inline-block; padding:3px 11px; border-radius:20px;
        font-size:0.76rem; font-weight:700;
    }}
    .pill-up-bad {{ background-color: rgba(156,107,140,0.20); color:{COLORS['danger']} !important; }}
    .pill-up-good {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-down-good {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-flat {{ background-color: rgba(156,147,172,0.15); color:{COLORS['text_soft']} !important; }}
    .pill-ok {{ background-color: rgba(201,184,232,0.18); color:{COLORS['success']} !important; }}
    .pill-warn {{ background-color: rgba(217,168,196,0.20); color:{COLORS['warning']} !important; }}
    .pill-risk {{ background-color: rgba(156,107,140,0.22); color:{COLORS['danger']} !important; }}

    .cat-row {{
        display:flex; justify-content:space-between; padding:7px 0;
        border-bottom:1px solid {COLORS['border']}; font-size:0.9rem;
    }}
    .cat-row span:first-child {{ color:{COLORS['text_soft']}; }}
    .cat-row span:last-child {{ font-weight:600; }}

    .info-row {{
        padding:8px 0; border-bottom:1px solid {COLORS['border']};
        font-size:0.92rem; line-height:1.5; display:flex; gap:8px;
    }}
    .info-row:last-child {{ border-bottom:none; }}
    .info-row::before {{ content:"—"; color:{COLORS['primary']}; flex-shrink:0; }}

    .hero-badge {{
        display:inline-block; padding:5px 14px; border-radius:20px; font-size:0.78rem;
        font-weight:700; background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
        color:{COLORS['text_soft']} !important; margin-right:8px;
    }}

    /* Pill-style tabs */
    .stTabs [data-baseweb="tab-list"] {{
        background-color: {COLORS['surface']}; padding:6px; border-radius:14px;
        border:1px solid {COLORS['border']}; gap:4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {COLORS['text_soft']} !important; border-radius:10px !important;
        padding: 8px 16px !important;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {COLORS['primary']} !important;
        color: #1A1424 !important; font-weight:700 !important;
    }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color: transparent !important; }}
    .stTabs [data-baseweb="tab-border"] {{ display:none !important; }}

    div[data-testid="stMetric"] {{ background-color: transparent; }}

    /* Sliders / inputs on dark */
    div[data-baseweb="slider"] div[role="slider"] {{ background-color: {COLORS['primary']} !important; }}
    .stSlider [data-testid="stTickBar"] {{ display:none; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_DARK = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=COLORS["text"], size=12),
    margin=dict(l=10, r=10, t=25, b=10),
)
GRID_COLOR = "rgba(255,255,255,0.06)"

def sparkline(values, color):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        y=values, mode="lines", line=dict(color=color, width=2.5),
        fill="tozeroy", fillcolor=color.replace(")", ",0.12)").replace("rgb", "rgba") if "rgb" in color else None,
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=2, b=2), height=48,
        xaxis=dict(visible=False), yaxis=dict(visible=False), showlegend=False,
    )
    return fig

# ----------------------------------------------------------------------
# SAMPLE / BENCHMARK DATA (placeholder until mentor-provided actuals)
# ----------------------------------------------------------------------
months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul"]
food_cost_trend = [29.8, 30.1, 30.6, 30.9, 31.0, 31.4]
target_line = [30.0] * 6
wastage_trend = [5.4, 5.2, 5.1, 4.9, 4.9, 4.8]
margin_trend = [42.3, 42.0, 41.8, 42.0, 42.2, 42.1]
cost_unit_trend = [3.55, 3.62, 3.68, 3.74, 3.80, 3.85]

baseline = {
    "food_cost_pct": 31.4,
    "wastage_pct": 4.8,
    "gross_margin_pct": 42.1,
    "cost_per_unit": 3.85,
    "standard_cost_per_unit": 3.60,
}

category_panels = {
    "🥖 Production": {
        "Batch efficiency": "88%",
        "Labor cost / unit": "AED 1.10",
        "Std vs actual cost": "+6.9%",
    },
    "🚚 Procurement": {
        "Purchase price variance": "+3.2%",
        "Supplier lead time": "2.4 days",
        "Top-3 flour supplier share": "76%",
    },
    "📦 Warehouse & inventory": {
        "Inventory turnover": "14.2x/yr",
        "Expired/damaged stock": "AED 8,400",
        "Slow-moving SKUs": "6",
    },
    "💰 Cost control & margin": {
        "Margin leakage (est.)": "AED 42,000/mo",
        "Packaging cost impact": "2.1% of sales",
        "Overhead allocation": "9.4% of sales",
    },
}

# ----------------------------------------------------------------------
# STATUS REPORT (EXCEL) + EMAIL DELIVERY
# Each "Email this ..." button builds a workbook tailored to the section
# it's called from, not a generic dashboard dump — Performance Tracker,
# Scenario & Resilience, a department comparison, or the Company Profile.
# ----------------------------------------------------------------------
XL_PRIMARY = "B9A0DC"
XL_DARK = "1A1424"
XL_HEADER_FONT = "FFFFFF"
XL_TEXT = "2B2733"
XL_SOFT_TEXT = "6B6577"
XL_LIGHT_FILL = "F3EFFA"
XL_BORDER_COLOR = "D9D2E9"

def _xl_border():
    side = Side(style="thin", color=XL_BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)

def _xl_title(ws, title, subtitle=None):
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=XL_DARK)
    row = 2
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = Font(name="Arial", size=10, italic=True, color=XL_SOFT_TEXT)
        row = 3
    return row + 1

def _xl_section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = Font(name="Arial", size=12, bold=True, color=XL_DARK)
    return row + 1

def _xl_header_row(ws, row, headers, start_col=1):
    border = _xl_border()
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color=XL_HEADER_FONT)
        c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = border
    return row + 1

def _xl_row(ws, row, values, start_col=1, formats=None, zebra=False):
    border = _xl_border()
    fill = PatternFill("solid", fgColor=XL_LIGHT_FILL) if zebra else None
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = Font(name="Arial", size=10, color=XL_TEXT)
        c.border = border
        if fill:
            c.fill = fill
        if formats and i < len(formats) and formats[i]:
            c.number_format = formats[i]
    return row + 1

def _xl_autosize(ws, max_width=64):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or hasattr(cell, "column_letter") is False:
                continue
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            widths[col] = max(widths.get(col, 8), min(len(str(cell.value)) + 3, max_width))
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def _xl_footer(ws, row):
    row += 1
    ws.cell(row=row, column=1,
             value=f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')} · GIP III · Bakery division only, catering excluded"
             ).font = Font(name="Arial", size=8, italic=True, color=XL_SOFT_TEXT)
    row += 1
    ws.cell(row=row, column=1,
             value="Prepared by Parineeta Jain, Rajveer Singh, and Tarang Gupta · Figures are illustrative "
                   "benchmarks pending Grandiose-provided actuals."
             ).font = Font(name="Arial", size=8, italic=True, color=XL_SOFT_TEXT)
    return row

PCT_FMT = '0.0"%"'
AED_FMT = '"AED "#,##0.00'
NUM_FMT = "#,##0"


def _build_performance_tracker_sheet(wb):
    ws = wb.create_sheet("Performance Tracker")
    row = _xl_title(ws, "Grandiose Bakery — Performance Tracker",
                     "Where the bakery division stands today")

    row = _xl_section(ws, row, "Headline KPIs")
    row = _xl_header_row(ws, row, ["Metric", "Value", "vs. target / last month"])
    kpi_rows = [
        ("Food cost %", baseline["food_cost_pct"], "1.2pt above 30% target"),
        ("Wastage %", baseline["wastage_pct"], "Down 0.3pt vs last month"),
        ("Gross margin %", baseline["gross_margin_pct"], "Flat vs last month"),
        ("Cost per unit (AED)", baseline["cost_per_unit"], f"Standard: AED {baseline['standard_cost_per_unit']:.2f}"),
    ]
    for i, (label, val, note) in enumerate(kpi_rows):
        fmt = AED_FMT if "unit" in label.lower() else PCT_FMT
        row = _xl_row(ws, row, [label, val, note], formats=[None, fmt, None], zebra=(i % 2 == 0))
    row += 1

    row = _xl_section(ws, row, "Food cost % — last 6 months vs. 30% target")
    row = _xl_header_row(ws, row, ["Month", "Food cost %", "Target %"])
    for i, (m, fc, tg) in enumerate(zip(months, food_cost_trend, target_line)):
        row = _xl_row(ws, row, [m, fc, tg], formats=[None, PCT_FMT, PCT_FMT], zebra=(i % 2 == 0))
    row += 1

    row = _xl_section(ws, row, "Category panels")
    row = _xl_header_row(ws, row, ["Category", "Metric", "Value"])
    zebra = False
    for cat_name, metrics in category_panels.items():
        clean_cat = cat_name.split(" ", 1)[-1]
        for k, v in metrics.items():
            row = _xl_row(ws, row, [clean_cat, k, v], zebra=zebra)
            zebra = not zebra
    row += 1

    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def _build_scenario_resilience_sheets(wb, context):
    infl = context.get("inflation", {})
    disr = context.get("disruption", {})
    pand = context.get("pandemic", {})
    supp = context.get("supplier", {})

    # --- Inflation sensitivity ---
    ws = wb.create_sheet("Inflation Sensitivity")
    row = _xl_title(ws, "Scenario & Resilience — Inflation Sensitivity",
                     "Dual-track inflation input, net of any subsidy/offset")
    row = _xl_section(ws, row, "Assumptions used")
    row = _xl_header_row(ws, row, ["Input", "Value"])
    row = _xl_row(ws, row, ["Headline inflation forecast", infl.get("headline_inf", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws, row, ["Actual food-input inflation", infl.get("food_inf", 0)], formats=[None, PCT_FMT])
    row = _xl_row(ws, row, ["Subsidy / price-cap offset (AED/unit)", infl.get("subsidy_offset", 0)], formats=[None, AED_FMT], zebra=True)
    row += 1
    row = _xl_section(ws, row, "Resulting cost per unit")
    row = _xl_header_row(ws, row, ["Scenario", "Cost per unit (AED)"])
    row = _xl_row(ws, row, ["Current", infl.get("base_cost", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws, row, ["Headline scenario", infl.get("adj_cost_headline", 0)], formats=[None, AED_FMT])
    row = _xl_row(ws, row, ["Food-inflation scenario", infl.get("adj_cost_food", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws, row, ["Food cost % under food-inflation scenario", infl.get("food_cost_pct_adj", 0)], formats=[None, PCT_FMT])
    _xl_footer(ws, row)
    _xl_autosize(ws)

    # --- Supply disruption ---
    ws2 = wb.create_sheet("Supply Disruption")
    row = _xl_title(ws2, "Scenario & Resilience — Supply Disruption Risk",
                     "UAE imports 80-90% of its food, so this is mainly an import-shock model")
    row = _xl_section(ws2, row, "Selected scenario")
    row = _xl_header_row(ws2, row, ["Field", "Value"])
    row = _xl_row(ws2, row, ["Scenario", disr.get("scenario", "—")], zebra=True)
    row = _xl_row(ws2, row, ["Lead-time extension (days)", disr.get("delay_days", 0)], formats=[None, NUM_FMT])
    row = _xl_row(ws2, row, ["Cost premium %", disr.get("cost_premium", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws2, row, ["Stockout probability", disr.get("stockout_prob", 0) * 100], formats=[None, PCT_FMT])
    row += 1
    row = _xl_section(ws2, row, "Buffer vs. expected stockout cost")
    row = _xl_header_row(ws2, row, ["Metric", "AED"])
    row = _xl_row(ws2, row, ["Cost of holding buffer stock", disr.get("buffer_stock_cost", 0)], formats=[None, AED_FMT], zebra=True)
    row = _xl_row(ws2, row, ["Expected cost of stockout", disr.get("expected_stockout_cost", 0)], formats=[None, AED_FMT])
    _xl_footer(ws2, row)
    _xl_autosize(ws2)

    # --- Pandemic preparedness ---
    ws3 = wb.create_sheet("Pandemic Preparedness")
    row = _xl_title(ws3, "Scenario & Resilience — Pandemic Preparedness",
                     "Demand-side and supply-side resilience, tracked together")
    row = _xl_section(ws3, row, "Customer retention / attraction")
    row = _xl_header_row(ws3, row, ["Metric", "Value"])
    row = _xl_row(ws3, row, ["Repeat-purchase rate", pand.get("repeat_rate", 0)], formats=[None, PCT_FMT], zebra=True)
    row = _xl_row(ws3, row, ["Delivery/online-order revenue share", pand.get("delivery_share", 0)], formats=[None, PCT_FMT])
    row = _xl_row(ws3, row, ["Average basket size (AED)", pand.get("basket_size", 0)], formats=[None, AED_FMT], zebra=True)
    row += 1
    row = _xl_section(ws3, row, "Supply chain optimization")
    row = _xl_header_row(ws3, row, ["Metric", "Value"])
    row = _xl_row(ws3, row, ["Safety stock (days of cover)", pand.get("safety_days", 0)], formats=[None, NUM_FMT], zebra=True)
    row = _xl_row(ws3, row, ["Qualified alternate suppliers / key ingredient", pand.get("alt_suppliers", 0)], formats=[None, NUM_FMT])
    row = _xl_row(ws3, row, ["% ingredients single-sourced", pand.get("single_sourced", 0)], formats=[None, PCT_FMT], zebra=True)
    _xl_footer(ws3, row)
    _xl_autosize(ws3)

    # --- Supplier concentration ---
    ws4 = wb.create_sheet("Supplier Concentration")
    row = _xl_title(ws4, "Scenario & Resilience — Supplier Concentration",
                     "HHI-style concentration index, following Ali et al. (2022)'s UAE cereal-supply-risk methodology")
    row = _xl_section(ws4, row, "Concentration index")
    row = _xl_header_row(ws4, row, ["Metric", "Value"])
    row = _xl_row(ws4, row, ["HHI", supp.get("hhi", 0)], formats=[None, NUM_FMT], zebra=True)
    row = _xl_row(ws4, row, ["Risk level", supp.get("risk_label", "—")])
    row += 1
    shares_df = supp.get("shares_df")
    if shares_df is not None and not shares_df.empty:
        row = _xl_section(ws4, row, "Supplier spend mix")
        row = _xl_header_row(ws4, row, list(shares_df.columns))
        for i, (_, r) in enumerate(shares_df.iterrows()):
            row = _xl_row(ws4, row, list(r.values), zebra=(i % 2 == 0))
        row += 1
    alt_df = supp.get("alt_df")
    if alt_df is not None and not alt_df.empty:
        row = _xl_section(ws4, row, "Pre-identified alternate suppliers")
        row = _xl_header_row(ws4, row, list(alt_df.columns))
        for i, (_, r) in enumerate(alt_df.iterrows()):
            row = _xl_row(ws4, row, list(r.values), zebra=(i % 2 == 0))
    _xl_footer(ws4, row)
    _xl_autosize(ws4)


def _build_department_sheet(wb, context):
    dept_name = context.get("department", "Department")
    ws = wb.create_sheet(dept_name[:31] if dept_name else "Department")
    row = _xl_title(ws, f"Employee Portal — {dept_name} Comparison",
                     "Performance ranking within this department")
    headcount = context.get("headcount", "—")
    tag = context.get("tag_label")
    row = _xl_section(ws, row, f"Headcount on this line: {headcount}" + (f" · {tag}" if tag else ""))
    df = context.get("summary_df")
    if df is not None and not df.empty:
        row = _xl_header_row(ws, row, list(df.columns))
        for i, (_, r) in enumerate(df.iterrows()):
            row = _xl_row(ws, row, list(r.values), zebra=(i % 2 == 0))
    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def _build_company_profile_sheet(wb, context):
    ws = wb.create_sheet("Company Profile")
    row = _xl_title(ws, "Grandiose Bakery — Company Profile",
                     "Everything collected from the 27 Jul meeting with the GM")
    company_info = context.get("company_info", [])
    for icon, title, points in company_info:
        row = _xl_section(ws, row, f"{icon} {title}")
        for i, p in enumerate(points):
            row = _xl_row(ws, row, [p], zebra=(i % 2 == 0))
        row += 1
    _xl_footer(ws, row)
    _xl_autosize(ws)
    return ws


def build_excel_report(section, context=None):
    """Builds a section-specific, formatted Excel workbook and returns raw bytes."""
    context = context or {}
    wb = Workbook()
    wb.remove(wb.active)

    if section == "scenario_resilience":
        _build_scenario_resilience_sheets(wb, context)
    elif section == "employee_department":
        _build_department_sheet(wb, context)
    elif section == "company_profile":
        _build_company_profile_sheet(wb, context)
    else:
        _build_performance_tracker_sheet(wb)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


SECTION_LABELS = {
    "performance_tracker": ("Performance Tracker", "Performance_Tracker"),
    "scenario_resilience": ("Scenario & Resilience", "Scenario_Resilience"),
    "employee_department": ("Employee Department Comparison", "Department_Comparison"),
    "company_profile": ("Company Profile", "Company_Profile"),
}

def send_report_email(recipient_email, note="", section="performance_tracker", context=None):
    """Sends a section-specific Excel report to recipient_email using SMTP
    credentials stored in Streamlit secrets. Returns (success: bool, message: str)."""
    try:
        sender_email = st.secrets["EMAIL_ADDRESS"]
        sender_password = st.secrets["EMAIL_APP_PASSWORD"]
    except (KeyError, FileNotFoundError):
        return False, ("Email is not configured yet. Add EMAIL_ADDRESS and EMAIL_APP_PASSWORD "
                        "in Streamlit Cloud's Secrets settings (App settings -> Secrets) — "
                        "never in the code or GitHub repo.")

    smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    section_title, section_slug = SECTION_LABELS.get(section, SECTION_LABELS["performance_tracker"])

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = f"Grandiose Bakery — {section_title} Report"
    body = (f"Hi,\n\nPlease find attached the {section_title} report from the Grandiose bakery "
            f"dashboard, generated just now.\n")
    if note:
        body += f"\nNote from sender:\n{note}\n"
    body += "\n— Sent automatically from the Grandiose Bakery Dashboard (GIP III)"
    msg.attach(MIMEText(body, "plain"))

    xlsx_bytes = build_excel_report(section, context)
    attachment = MIMEApplication(
        xlsx_bytes, _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    filename = f"Grandiose_Bakery_{section_slug}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    attachment.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(attachment)

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=15) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True, f"Report sent to {recipient_email}."
    except Exception as e:
        return False, f"Could not send email: {e}"


EMAIL_PATTERN = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

def render_email_share(key_suffix, label="Email this report", section="performance_tracker", context=None):
    """Small 📧 popover button — usable right under any chart/section so
    people don't have to go back to the sidebar to share the dashboard.
    Sends a formatted Excel workbook tailored to the calling section."""
    section_title, _ = SECTION_LABELS.get(section, SECTION_LABELS["performance_tracker"])
    with st.popover("📧", help=label):
        st.markdown(f"**{label}**")
        recipient = st.text_input("Recipient email", placeholder="name@company.com", key=f"email_recipient_{key_suffix}")
        note = st.text_area("Add a note (optional)", key=f"email_note_{key_suffix}", height=68)
        if st.button("Send report", key=f"send_email_btn_{key_suffix}", width='stretch'):
            if not recipient or not re.match(EMAIL_PATTERN, recipient):
                st.error("Enter a valid email address.")
            else:
                with st.spinner("Building and sending Excel report..."):
                    success, message = send_report_email(recipient, note, section=section, context=context)
                if success:
                    st.success(message)
                else:
                    st.error(message)
        st.caption(f"Sends a formatted Excel workbook of the current {section_title} data.")


st.sidebar.markdown(f"""
<div style='background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
     border-radius:14px; padding:16px; margin-bottom:14px;'>
  <div style='font-size:1.5rem;'>🍞</div>
  <div style='font-weight:800; font-size:1.15rem; margin-top:4px;'>Grandiose Bakery</div>
  <div style='color:{COLORS['text_soft']}; font-size:0.8rem; margin-top:4px;'>
    GIP III — Financial performance &amp; cost optimization<br>Bakery division only · catering excluded
  </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown(f"""
<div style='background-color:{COLORS['surface']}; border:1px solid {COLORS['border']};
     border-radius:14px; padding:14px; margin-bottom:14px;'>
  <div style='color:{COLORS['text_soft']}; font-size:0.78rem; font-weight:700; letter-spacing:0.5px;'>PREPARED BY</div>
  <div style='margin-top:6px; font-size:0.9rem;'>Parineeta Jain</div>
  <div style='font-size:0.9rem;'>Rajveer Singh</div>
  <div style='font-size:0.9rem;'>Tarang Gupta</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown(f"""
<div style='background-color:{hex_to_rgba(COLORS['primary'], 0.1)}; border:1px solid {hex_to_rgba(COLORS['primary'], 0.3)};
     border-radius:14px; padding:14px; font-size:0.82rem; color:{COLORS['text']} !important;'>
  ⚠️ All figures shown are illustrative benchmarks pending Grandiose's actual bakery data.
</div>
""", unsafe_allow_html=True)

with st.sidebar.expander("📧  Email this report", expanded=False):
    recipient = st.text_input("Recipient email", placeholder="name@company.com", key="email_recipient")
    note = st.text_area("Add a note (optional)", key="email_note", height=70)
    if st.button("Send report", key="send_email_btn", width='stretch'):
        email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
        if not recipient or not re.match(email_pattern, recipient):
            st.error("Enter a valid email address.")
        else:
            with st.spinner("Building and sending Excel report..."):
                success, message = send_report_email(recipient, note, section="performance_tracker")
            if success:
                st.success(message)
            else:
                st.error(message)
    st.caption("Sends a formatted Excel workbook of the current Performance Tracker data.")


# ----------------------------------------------------------------------
# EMPLOYEE PORTAL — data layer
# Uses Supabase (Postgres) when configured via secrets; falls back to an
# in-session store so the portal is still fully clickable/demoable before
# Supabase is wired up. Data in fallback mode does not persist across
# app restarts — see README for the one-time Supabase setup.
# ----------------------------------------------------------------------

# Actual staff breakdown by line/department (approximate headcounts).
DEPARTMENT_HEADCOUNT = {
    "Baklava": 10,
    "French Bakery": 10,
    "Arabic Bread": 7,
    "Viennoiserie (Croissant/Danish)": 7,
    "Tahina": 5,
    "QC": 3,
    "Packing & Dispatch": 15,
    "Store": 2,
    "Maintenance": 4,
    "GM": 1,
    "Office": 10,
    "Drivers": 7,
    "Third-Party Cleaners": 10,
}

# Company-wide headcount context from the 27 Jul MoM. Note this figure (~110)
# is larger than the sum of the itemized bakery lines above (91) — likely
# because it includes catering/other roles not broken out per line. Shown
# for context only; the itemized breakdown above remains the bakery-scope
# source of truth for GIP III.
COMPANY_HEADCOUNT_CURRENT = 110
COMPANY_HEADCOUNT_TARGET = 200

# Illustrative daily capacity targets (units/day) for production lines only —
# PLACEHOLDER figures pending Grandiose's actual max-output numbers per line.
# The MoM gives an overall utilization figure (30-35%, target 100% within a
# year) but not a per-line breakdown, so utilization % shown against these
# targets should be treated as directional, not precise, until replaced.
DEPARTMENT_CAPACITY_TARGET = {
    "Baklava": 400,
    "French Bakery": 500,
    "Arabic Bread": 900,
    "Viennoiserie (Croissant/Danish)": 450,
    "Tahina": 250,
}

# Lines the MoM specifically called out — used to tag department cards/tables.
DEPARTMENT_PRIORITY_TAG = {
    "Arabic Bread": ("🔥 Top revenue driver", "pill-ok"),
    "Tahina": ("⚠️ Ramp-up focus (underutilized)", "pill-warn"),
    "Baklava": ("⚠️ Ramp-up focus (underutilized)", "pill-warn"),
}

# GM's stated productivity benchmark: each employee = AED 1,000/day value.
VALUE_PER_EMPLOYEE_DAY_TARGET = 1000
# Standard shift length per the MoM ("Standard shift is 9 hours, including a
# break") — used to normalize revenue/hour up to a full standard day.
STANDARD_SHIFT_HOURS = 9.0

# Company-wide wastage target from the MoM (current is 2-3%, target is 1%).
WASTAGE_TARGET_PCT = 1.0

# Illustrative sample employees per department (placeholders — swap in
# Grandiose's real staff names/IDs once available). Not every headcount
# above has a named entry; these exist so the portal and department
# comparisons are demoable today.
EMPLOYEES = [
    {"name": "Ahmad Yousef", "id": "BAK-01", "department": "Baklava"},
    {"name": "Rania Haddad", "id": "BAK-02", "department": "Baklava"},
    {"name": "Mohammed Iqbal", "id": "BAK-03", "department": "Baklava"},

    {"name": "Rahul Nair", "id": "FRB-01", "department": "French Bakery"},
    {"name": "Fatima Zahra", "id": "FRB-02", "department": "French Bakery"},
    {"name": "John Cruz", "id": "FRB-03", "department": "French Bakery"},

    {"name": "Khalid Obaid", "id": "ARB-01", "department": "Arabic Bread"},
    {"name": "Youssef Amer", "id": "ARB-02", "department": "Arabic Bread"},
    {"name": "Layla Mansour", "id": "ARB-03", "department": "Arabic Bread"},

    {"name": "Ana Reyes", "id": "VIE-01", "department": "Viennoiserie (Croissant/Danish)"},
    {"name": "Omar Saleh", "id": "VIE-02", "department": "Viennoiserie (Croissant/Danish)"},
    {"name": "Priya Sharma", "id": "VIE-03", "department": "Viennoiserie (Croissant/Danish)"},

    {"name": "Hassan Ali", "id": "TAH-01", "department": "Tahina"},
    {"name": "Meera Pillai", "id": "TAH-02", "department": "Tahina"},

    {"name": "Noura Saeed", "id": "QC-01", "department": "QC"},
    {"name": "Vikram Singh", "id": "QC-02", "department": "QC"},

    {"name": "Josie Santos", "id": "PKD-01", "department": "Packing & Dispatch"},
    {"name": "Ravi Kumar", "id": "PKD-02", "department": "Packing & Dispatch"},
    {"name": "Abdullah Nasser", "id": "PKD-03", "department": "Packing & Dispatch"},

    {"name": "Aisha Rahman", "id": "STR-01", "department": "Store"},
    {"name": "Manuel Cruz", "id": "STR-02", "department": "Store"},

    {"name": "Suresh Pillai", "id": "MNT-01", "department": "Maintenance"},
    {"name": "Ibrahim Al Farsi", "id": "MNT-02", "department": "Maintenance"},

    {"name": "Khalifa Al Marzooqi", "id": "GM-01", "department": "GM"},

    {"name": "Sara Al Ali", "id": "OFF-01", "department": "Office"},
    {"name": "Deepak Verma", "id": "OFF-02", "department": "Office"},
    {"name": "Nadia Haddad", "id": "OFF-03", "department": "Office"},

    {"name": "Mohammed Rafiq", "id": "DRV-01", "department": "Drivers"},
    {"name": "Carlos Dionisio", "id": "DRV-02", "department": "Drivers"},
    {"name": "Salim Bakhit", "id": "DRV-03", "department": "Drivers"},

    {"name": "Ganesh Kumar", "id": "CLN-01", "department": "Third-Party Cleaners"},
    {"name": "Rosa Villanueva", "id": "CLN-02", "department": "Third-Party Cleaners"},
]


@st.cache_resource
def get_supabase_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except (KeyError, FileNotFoundError):
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None

def portal_insert(table, row):
    client = get_supabase_client()
    if client is not None:
        try:
            client.table(table).insert(row).execute()
            return True, None
        except Exception as e:
            return False, str(e)
    else:
        row = {**row, "created_at": datetime.now().isoformat()}
        key = f"_local_{table}"
        st.session_state.setdefault(key, [])
        row["id"] = len(st.session_state[key]) + 1
        st.session_state[key].append(row)
        return True, None

def portal_fetch(table, filters=None):
    client = get_supabase_client()
    if client is not None:
        try:
            q = client.table(table).select("*")
            if filters:
                for k, v in filters.items():
                    q = q.eq(k, v)
            res = q.order("created_at", desc=True).execute()
            return pd.DataFrame(res.data), None
        except Exception as e:
            return pd.DataFrame(), str(e)
    else:
        rows = st.session_state.get(f"_local_{table}", [])
        df = pd.DataFrame(rows)
        if not df.empty:
            if filters:
                for k, v in filters.items():
                    df = df[df[k] == v]
            if "created_at" in df.columns:
                df = df.sort_values("created_at", ascending=False)
        return df, None

def portal_storage_mode():
    return "Supabase (persistent)" if get_supabase_client() is not None else "In-session only (not saved after restart)"


# ----------------------------------------------------------------------
# HERO HEADER
# ----------------------------------------------------------------------
st.markdown(f"""
<div style='margin-bottom:6px;'>
  <span class='hero-badge'>🇦🇪 UAE · Grandiose</span>
  <span class='hero-badge'>Bakery division</span>
  <span class='hero-badge'>GIP III</span>
</div>
""", unsafe_allow_html=True)
st.markdown("## Financial Performance Dashboard")

st.markdown(f"""
<style>
    div[data-baseweb="select"] > div {{
        background-color: {COLORS['surface']} !important;
        border: 1px solid {COLORS['border']} !important;
        border-radius: 12px !important;
    }}
</style>
""", unsafe_allow_html=True)

section = st.selectbox(
    "Dashboard section",
    ["📊  Performance tracker", "🧭  Scenario & resilience", "👤  Employee portal", "📋  Company profile"],
    label_visibility="collapsed",
)

# ========================================================================
# SECTION A — PERFORMANCE TRACKER
# ========================================================================
if section == "📊  Performance tracker":
    st.caption("Where the bakery division stands today")

    kpis = [
        ("🍞", "Food cost %", f"{baseline['food_cost_pct']}%", "↑ 1.2pt vs target", "pill-up-bad",
         food_cost_trend, COLORS["primary"]),
        ("🗑️", "Wastage %", f"{baseline['wastage_pct']}%", "↓ 0.3pt vs last month", "pill-down-good",
         wastage_trend, COLORS["success"]),
        ("📈", "Gross margin", f"{baseline['gross_margin_pct']}%", "flat vs last month", "pill-flat",
         margin_trend, COLORS["tertiary"]),
        ("💵", "Cost / unit (AED)", f"{baseline['cost_per_unit']:.2f}", f"↑ vs std {baseline['standard_cost_per_unit']:.2f}", "pill-up-bad",
         cost_unit_trend, COLORS["secondary"]),
    ]
    cols = st.columns(4)
    for col, (icon, label, value, delta, pill_class, trend, color) in zip(cols, kpis):
        with col:
            with st.container(border=True):
                st.markdown(f"""
                <div style='padding:10px 14px 0 14px;'>
                    <span class='kpi-icon' style='background-color:{hex_to_rgba(color,0.15)};'>{icon}</span>
                    <span class='kpi-label'>{label}</span>
                    <div class='kpi-value'>{value}</div>
                    <span class='pill {pill_class}'>{delta}</span>
                </div>
                """, unsafe_allow_html=True)
                st.plotly_chart(sparkline(trend, color), width='stretch',
                                 config={"displayModeBar": False}, key=f"spark_{label}")

    st.markdown("#### Food cost % vs target — last 6 months")
    with st.container(border=True):
        fig = go.Figure()
        # glow effect: wide low-opacity trace behind crisp line
        fig.add_trace(go.Scatter(x=months, y=food_cost_trend, mode="lines", line=dict(color=COLORS["primary"], width=10),
                                  opacity=0.15, showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=months, y=food_cost_trend, mode="lines+markers", name="Food cost %",
                                  line=dict(color=COLORS["primary"], width=3),
                                  marker=dict(size=8, color=COLORS["bg"], line=dict(color=COLORS["primary"], width=2)),
                                  fill="tozeroy", fillcolor=hex_to_rgba(COLORS["primary"], 0.12)))
        fig.add_trace(go.Scatter(x=months, y=target_line, mode="lines", name="Target",
                                  line=dict(color=COLORS["text_soft"], width=1.5, dash="dot")))
        fig.update_layout(**PLOTLY_DARK, height=300,
                           yaxis=dict(ticksuffix="%", range=[28, 33], gridcolor=GRID_COLOR, zeroline=False),
                           xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                           legend=dict(orientation="h", yanchor="bottom", y=1.03, x=0,
                                       font=dict(color=COLORS["text_soft"])))
        st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    render_email_share("trend_chart", "Email this Performance Tracker report", section="performance_tracker")

    st.markdown("#### Category panels")
    cols2 = st.columns(4)
    for col, (name, metrics) in zip(cols2, category_panels.items()):
        with col:
            with st.container(border=True):
                rows = "".join([f"<div class='cat-row'><span>{k}</span><span>{v}</span></div>" for k, v in metrics.items()])
                st.markdown(f"<div style='padding:12px 16px;'><h4 style='margin:0 0 8px 0; font-size:1rem;'>{name}</h4>{rows}</div>",
                            unsafe_allow_html=True)
    render_email_share("category_panels", "Email this Performance Tracker report", section="performance_tracker")

# ========================================================================
# TAB B — SCENARIO & RESILIENCE
# ========================================================================
# ========================================================================
# SECTION B — SCENARIO & RESILIENCE
# ========================================================================
elif section == "🧭  Scenario & resilience":
    st.caption("What happens if — stress-tests the same cost and margin outputs under new assumptions")

    mod1, mod2, mod3, mod4 = st.tabs([
        "1 · Inflation sensitivity", "2 · Supply disruption risk",
        "3 · Pandemic preparedness", "4 · Supplier alternatives",
    ])

    # ---------------- MODULE 1: INFLATION-ADJUSTED COST SENSITIVITY ----------------
    with mod1:
        st.markdown("##### 🔥 Inflation-adjusted cost sensitivity")
        st.caption("Dual-track inflation input with an adjustable bar, net of any subsidy/offset deducted.")

        base_cost = baseline["cost_per_unit"]
        colL, colR = st.columns([1, 1.3])
        with colL:
            with st.container(border=True):
                headline_inf = st.slider("Headline inflation forecast (%)", 0.0, 10.0, 2.8, 0.1)
                food_inf = st.slider("Actual food-input inflation (%)", 0.0, 20.0, 5.3, 0.1,
                                      help="Flour, dairy, and packaging inflation can diverge sharply from headline CPI during shocks.")
                subsidy_offset = st.number_input("Subsidy / price-cap offset (AED per unit)", 0.0, 2.0, 0.05, 0.01)

                st.markdown(f"<div style='margin-top:6px; color:{COLORS['text_soft']}; font-size:0.82rem;'>"
                            f"Cost composition — updates live with the sliders above</div>", unsafe_allow_html=True)
                inflation_addon = round(base_cost * food_inf / 100, 3)
                net_after_subsidy = round(base_cost + inflation_addon - subsidy_offset, 2)
                donut = go.Figure(go.Pie(
                    labels=["Base cost", "Inflation add-on"],
                    values=[base_cost, max(inflation_addon, 0.001)],
                    hole=0.64, sort=False, textinfo="percent",
                    marker=dict(colors=[COLORS["text_soft"], COLORS["danger"]],
                                line=dict(color=COLORS["surface"], width=3)),
                ))
                donut.update_layout(
                    **PLOTLY_DARK, height=220,
                    legend=dict(orientation="h", y=-0.12, font=dict(color=COLORS["text_soft"])),
                    annotations=[dict(
                        text=f"<b>AED {net_after_subsidy:.2f}</b><br><span style='font-size:10px;color:{COLORS['text_soft']}'>net / unit</span>",
                        x=0.5, y=0.5, font=dict(size=18, color=COLORS['text']), showarrow=False)],
                )
                st.plotly_chart(donut, width='stretch', config={"displayModeBar": False}, key="inflation_donut")

        adj_cost_headline = round(base_cost * (1 + headline_inf / 100) - subsidy_offset, 3)
        adj_cost_food = round(base_cost * (1 + food_inf / 100) - subsidy_offset, 3)
        base_revenue_per_unit = base_cost / (baseline["food_cost_pct"] / 100)
        food_cost_pct_adj = round((adj_cost_food / base_revenue_per_unit) * 100, 1)

        with colR:
            m1, m2, m3 = st.columns(3)
            m1.metric("Cost/unit — headline scenario", f"AED {adj_cost_headline:.2f}")
            m2.metric("Cost/unit — food-inflation scenario", f"AED {adj_cost_food:.2f}",
                      f"{adj_cost_food - base_cost:+.2f} vs current")
            m3.metric("Food cost % — food-inflation scenario", f"{food_cost_pct_adj}%",
                      f"{food_cost_pct_adj - baseline['food_cost_pct']:+.1f}pt", delta_color="inverse")

            with st.container(border=True):
                fig1 = go.Figure(go.Bar(
                    x=["Current", "Headline scenario", "Food-inflation scenario"],
                    y=[base_cost, adj_cost_headline, adj_cost_food],
                    marker_color=[COLORS["text_soft"], COLORS["tertiary"], COLORS["danger"]],
                    marker_line_width=0,
                ))
                fig1.update_layout(**PLOTLY_DARK, height=260, yaxis_title="AED per unit",
                                    yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                st.plotly_chart(fig1, width='stretch', config={"displayModeBar": False})

        st.caption("Context: UAE food inflation has ranged from a 2022 peak near 9% to negative readings in 2021, "
                   "against headline forecasts around 2–3% for 2026 — a single flat assumption misses this swing.")

    # ---------------- MODULE 2: NATURAL CALAMITY / SUPPLY DISRUPTION ----------------
    with mod2:
        st.markdown("##### 🌪️ Natural calamity & supply disruption risk")
        st.caption("UAE imports 80–90% of its food, so 'natural calamity' risk for a bakery is mainly an import-shock risk.")

        scenarios = {
            "No disruption": (0, 0, 0.02),
            "Port congestion": (7, 6, 0.15),
            "Strait chokepoint closure": (18, 14, 0.35),
            "Extreme heat / sandstorm logistics delay": (3, 3, 0.10),
            "Custom": None,
        }
        choice = st.selectbox("Disruption scenario", list(scenarios.keys()), index=2)

        if choice == "Custom":
            colA, colB, colC = st.columns(3)
            delay_days = colA.slider("Lead-time extension (days)", 0, 30, 10)
            cost_premium = colB.slider("Cost premium (%)", 0, 40, 10)
            stockout_prob = colC.slider("Stockout probability", 0.0, 1.0, 0.2, 0.05)
        else:
            delay_days, cost_premium, stockout_prob = scenarios[choice]
            colA, colB, colC = st.columns(3)
            colA.metric("Lead-time extension", f"{delay_days} days")
            colB.metric("Cost premium", f"{cost_premium}%")
            colC.metric("Stockout probability", f"{stockout_prob:.0%}")

        avg_daily_cost = 4200
        buffer_stock_days = max(delay_days, 3)
        buffer_stock_cost = round(buffer_stock_days * avg_daily_cost * 0.02, 0)
        expected_stockout_cost = round(stockout_prob * avg_daily_cost * delay_days * 1.5, 0)

        with st.container(border=True):
            st.markdown(
                f"<div style='padding:6px 4px;'><h4 style='margin-top:0;'>Recommended buffer vs expected stockout cost</h4>"
                f"Holding <b style='color:{COLORS['primary']}'>{buffer_stock_days} days</b> of buffer stock costs approximately "
                f"<b style='color:{COLORS['primary']}'>AED {buffer_stock_cost:,.0f}</b>. The expected cost of not holding buffer, given this scenario's "
                f"probability, is approximately <b style='color:{COLORS['danger']}'>AED {expected_stockout_cost:,.0f}</b>.</div>",
                unsafe_allow_html=True)

            fig2 = go.Figure(go.Bar(
                x=["Cost of holding buffer stock", "Expected cost of stockout"],
                y=[buffer_stock_cost, expected_stockout_cost],
                marker_color=[COLORS["primary"], COLORS["danger"]], marker_line_width=0,
            ))
            fig2.update_layout(**PLOTLY_DARK, height=260, yaxis_title="AED",
                                yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig2, width='stretch', config={"displayModeBar": False})

    # ---------------- MODULE 3: PANDEMIC PREPAREDNESS ----------------
    with mod3:
        st.markdown("##### 🦠 Pandemic preparedness")
        st.caption("Two linked panels: customer retention/attraction (demand side) and supply chain optimization (supply side).")

        colX, colY = st.columns(2)
        with colX:
            with st.container(border=True):
                st.markdown(f"<h4 style='color:{COLORS['secondary']}'>👥 Customer retention / attraction</h4>", unsafe_allow_html=True)
                repeat_rate = st.slider("Repeat-purchase rate (%)", 0, 100, 62, key="repeat")
                delivery_share = st.slider("Delivery/online-order revenue share (%)", 0, 100, 28, key="delivery")
                basket_size = st.slider("Average basket size (AED)", 20, 200, 74, key="basket")
                flags = []
                if repeat_rate < 50: flags.append("Repeat-purchase rate below healthy baseline")
                if delivery_share < 15: flags.append("Delivery share low — limited resilience to footfall shocks")
                badge, label = ("pill-risk", "Needs attention") if flags else ("pill-ok", "Within healthy range")
                st.markdown(f"<span class='pill {badge}'>{label}</span>", unsafe_allow_html=True)
                for f in flags:
                    st.caption(f"⚠️ {f}")

        with colY:
            with st.container(border=True):
                st.markdown(f"<h4 style='color:{COLORS['primary']}'>🚚 Supply chain optimization</h4>", unsafe_allow_html=True)
                safety_days = st.slider("Safety stock (days of cover)", 0, 30, 9, key="safety")
                alt_suppliers = st.slider("Qualified alternate suppliers per key ingredient", 0, 5, 1, key="alt")
                single_sourced = st.slider("% ingredients single-sourced", 0, 100, 45, key="single")
                flags2 = []
                if safety_days < 7: flags2.append("Safety stock below 7-day resilience threshold")
                if alt_suppliers < 2: flags2.append("Fewer than 2 qualified alternates — concentration risk")
                if single_sourced > 40: flags2.append("Over 40% of ingredients single-sourced")
                badge2, label2 = ("pill-risk", "Needs attention") if flags2 else ("pill-ok", "Within healthy range")
                st.markdown(f"<span class='pill {badge2}'>{label2}</span>", unsafe_allow_html=True)
                for f in flags2:
                    st.caption(f"⚠️ {f}")

        st.caption("Rationale: pandemics create a dual shock — demand shifting from in-store to delivery, and "
                   "supply disruption at the same time — so both sides need to be tracked together as an early-warning system.")

    # ---------------- MODULE 4: SUPPLIER ALTERNATIVES / CONCENTRATION ----------------
    with mod4:
        st.markdown("##### 🌾 Raw material & supplier alternatives")
        st.caption("Herfindahl–Hirschman-style concentration index for core bakery inputs (e.g., flour), "
                   "following the measurement approach used in UAE cereal-supply-risk research (Ali et al., 2022).")

        default_suppliers = pd.DataFrame({
            "Supplier / origin": ["Russia", "Canada", "Australia", "India", "Other"],
            "Spend share (%)": [31, 28, 12, 20, 9],
        })
        st.caption("Edit the spend shares below to reflect current or hypothetical sourcing mix (must sum to ~100%).")
        edited = st.data_editor(default_suppliers, num_rows="fixed", width='stretch', hide_index=True)

        shares = edited["Spend share (%)"].astype(float)
        total = shares.sum()
        shares_norm = shares / total * 100 if total > 0 else shares
        hhi = round((shares_norm ** 2).sum(), 0)

        if hhi < 1500:
            risk_label, risk_badge = "Low concentration risk", "pill-ok"
        elif hhi < 2500:
            risk_label, risk_badge = "Moderate concentration risk", "pill-warn"
        else:
            risk_label, risk_badge = "High concentration risk", "pill-risk"

        colM, colN = st.columns([1, 2])
        with colM:
            with st.container(border=True):
                st.markdown(f"<div class='kpi-label'>Supplier concentration index (HHI)</div>"
                            f"<div class='kpi-value' style='font-size:2.4rem;'>{hhi:,.0f}</div>"
                            f"<span class='pill {risk_badge}'>{risk_label}</span>", unsafe_allow_html=True)
                st.caption("Reference bands: <1,500 low · 1,500–2,500 moderate · >2,500 high (standard HHI convention).")

        with colN:
            with st.container(border=True):
                fig3 = go.Figure(go.Pie(labels=edited["Supplier / origin"], values=shares_norm,
                                         hole=0.55, marker=dict(colors=CATEGORICAL, line=dict(color=COLORS["surface"], width=2))))
                fig3.update_layout(**PLOTLY_DARK, height=280, legend=dict(font=dict(color=COLORS["text_soft"])))
                st.plotly_chart(fig3, width='stretch', config={"displayModeBar": False})

        st.markdown("###### Pre-identified alternate suppliers (illustrative)")
        alt_df = pd.DataFrame({
            "Alternate supplier": ["Local UAE mill (blended flour)", "Turkey-origin supplier", "Egypt-origin supplier"],
            "Cost delta vs current": ["+4.5%", "+2.1%", "+1.8%"],
            "Lead-time delta": ["-3 days", "-1 day", "0 days"],
        })
        st.dataframe(alt_df, width='stretch', hide_index=True)

    scenario_context = {
        "inflation": {
            "headline_inf": headline_inf, "food_inf": food_inf, "subsidy_offset": subsidy_offset,
            "base_cost": base_cost, "adj_cost_headline": adj_cost_headline, "adj_cost_food": adj_cost_food,
            "food_cost_pct_adj": food_cost_pct_adj,
        },
        "disruption": {
            "scenario": choice, "delay_days": delay_days, "cost_premium": cost_premium,
            "stockout_prob": stockout_prob, "buffer_stock_cost": buffer_stock_cost,
            "expected_stockout_cost": expected_stockout_cost,
        },
        "pandemic": {
            "repeat_rate": repeat_rate, "delivery_share": delivery_share, "basket_size": basket_size,
            "safety_days": safety_days, "alt_suppliers": alt_suppliers, "single_sourced": single_sourced,
        },
        "supplier": {"hhi": hhi, "risk_label": risk_label, "shares_df": edited, "alt_df": alt_df},
    }
    render_email_share("scenario_resilience", "Email this Scenario & Resilience report",
                        section="scenario_resilience", context=scenario_context)

# ========================================================================
# SECTION C — EMPLOYEE PORTAL
# ========================================================================
elif section == "👤  Employee portal":
    st.caption("Employee self-service performance tracking — bakery division")

    storage_mode = portal_storage_mode()
    badge_cls = "pill-ok" if "Supabase" in storage_mode else "pill-warn"
    st.markdown(f"<span class='pill {badge_cls}'>💾 Storage: {storage_mode}</span>", unsafe_allow_html=True)
    if "Supabase" not in storage_mode:
        with st.expander("⚙️ Set up persistent storage (Supabase)"):
            st.markdown(
                "The portal works right now, but entries only last for this browser session. "
                "To make them persistent across visits and deployments, connect a free Supabase project — "
                "see the **Employee portal setup** section in the README for the exact SQL and secrets to add."
            )

    with st.expander("🏭 Staff breakdown by line", expanded=False):
        dep_df = pd.DataFrame(
            [{"Department": d, "Headcount": c,
              "Focus": DEPARTMENT_PRIORITY_TAG.get(d, ("", ""))[0]} for d, c in DEPARTMENT_HEADCOUNT.items()]
        )
        dep_df.loc[len(dep_df)] = ["Total (itemized bakery lines)", dep_df["Headcount"].sum(), ""]
        st.dataframe(dep_df, width='stretch', hide_index=True)
        cA, cB = st.columns(2)
        cA.metric("Company-wide headcount (current)", COMPANY_HEADCOUNT_CURRENT)
        cB.metric("Growth target (within 1 year)", COMPANY_HEADCOUNT_TARGET,
                   f"+{COMPANY_HEADCOUNT_TARGET - COMPANY_HEADCOUNT_CURRENT} needed")
        st.caption("Company-wide figures per the 27 Jul MoM — larger than the itemized bakery-line total above, "
                   "likely including catering/other roles not broken out per line. Approximate headcounts overall; "
                   "a small illustrative sample of named employees per department is used for the interactive "
                   "demo below — replace with the full roster once available.")

    st.markdown("#### Who are you?")
    col_d1, col_d2 = st.columns([1, 1.4])
    with col_d1:
        dept_choice = st.selectbox("Department", ["🧑‍💼 Manager / HR view"] + list(DEPARTMENT_HEADCOUNT.keys()))

    identity_employee = None
    if dept_choice != "🧑‍💼 Manager / HR view":
        dept_employees = [e for e in EMPLOYEES if e["department"] == dept_choice]
        emp_labels = [f"{e['name']} ({e['id']})" for e in dept_employees]
        with col_d2:
            emp_choice = st.selectbox("Employee", ["— Select —"] + emp_labels)
        if emp_choice != "— Select —":
            identity_employee = dept_employees[emp_labels.index(emp_choice)]

    # -------------------- EMPLOYEE SELF-SERVICE --------------------
    if identity_employee is not None:
        emp_name, emp_id, emp_dept = identity_employee["name"], identity_employee["id"], identity_employee["department"]
        tag_label, tag_cls = DEPARTMENT_PRIORITY_TAG.get(emp_dept, (None, None))
        tag_html = f"<span class='pill {tag_cls}' style='margin-left:6px;'>{tag_label}</span>" if tag_label else ""
        st.markdown(f"<span class='pill pill-ok'>{emp_dept}</span>{tag_html}", unsafe_allow_html=True)

        etab1, etab2, etab3 = st.tabs(["📊 My performance", "📝 Daily log", "🎯 Goals & feedback"])

        # --- My performance ---
        with etab1:
            perf_df, err = portal_fetch("employee_performance", {"employee_id": emp_id})
            if err:
                st.error(f"Could not load performance data: {err}")
            elif perf_df.empty:
                pass
            else:
                for col in ["units_produced", "wastage_pct", "hours_worked", "batch_time_adherence_pct",
                            "quality_pass_pct", "revenue_generated"]:
                    if col in perf_df.columns:
                        perf_df[col] = pd.to_numeric(perf_df[col], errors="coerce")

                avg_wastage = perf_df["wastage_pct"].mean()
                wastage_pill = ("pill-ok" if avg_wastage <= WASTAGE_TARGET_PCT
                                 else "pill-warn" if avg_wastage <= 3 else "pill-risk")

                c1, c2, c3, c4 = st.columns(4)
                with c1, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Avg units / shift</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['units_produced'].mean():.0f}</div>",
                                unsafe_allow_html=True)
                with c2, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Avg wastage %</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{avg_wastage:.1f}%</div>"
                                f"<span class='pill {wastage_pill}'>Target ≤{WASTAGE_TARGET_PCT:.0f}%</span>",
                                unsafe_allow_html=True)
                with c3, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Batch-time adherence</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['batch_time_adherence_pct'].mean():.0f}%</div>",
                                unsafe_allow_html=True)
                with c4, st.container(border=True):
                    st.markdown(f"<div class='kpi-label'>Quality pass rate</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>{perf_df['quality_pass_pct'].mean():.0f}%</div>",
                                unsafe_allow_html=True)

                # --- Productivity benchmark (AED 1,000/day per GM's rule of thumb) ---
                c5, c6 = st.columns(2)
                with c5, st.container(border=True):
                    avg_hours = perf_df["hours_worked"].mean() if perf_df["hours_worked"].mean() > 0 else STANDARD_SHIFT_HOURS
                    avg_revenue = perf_df["revenue_generated"].mean() if "revenue_generated" in perf_df.columns else 0.0
                    value_per_day = (avg_revenue / avg_hours) * STANDARD_SHIFT_HOURS if avg_hours else avg_revenue
                    val_pill = "pill-ok" if value_per_day >= VALUE_PER_EMPLOYEE_DAY_TARGET else "pill-warn"
                    st.markdown(f"<div class='kpi-label'>Value generated / day</div>"
                                f"<div class='kpi-value' style='font-size:1.7rem;'>AED {value_per_day:,.0f}</div>"
                                f"<span class='pill {val_pill}'>GM benchmark: AED {VALUE_PER_EMPLOYEE_DAY_TARGET:,.0f}/day</span>",
                                unsafe_allow_html=True)
                with c6, st.container(border=True):
                    cap_target = DEPARTMENT_CAPACITY_TARGET.get(emp_dept)
                    if cap_target:
                        util_pct = (perf_df["units_produced"].mean() / cap_target) * 100
                        util_pill = "pill-ok" if util_pct >= 70 else "pill-warn" if util_pct >= 40 else "pill-risk"
                        st.markdown(f"<div class='kpi-label'>Capacity utilization ({emp_dept})</div>"
                                    f"<div class='kpi-value' style='font-size:1.7rem;'>{util_pct:.0f}%</div>"
                                    f"<span class='pill {util_pill}'>vs illustrative target {cap_target}/day</span>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div class='kpi-label'>Capacity utilization</div>"
                                    f"<div class='kpi-value' style='font-size:1rem; color:{COLORS['text_soft']};'>Not tracked for this line</div>",
                                    unsafe_allow_html=True)

                with st.container(border=True):
                    plot_df = perf_df.sort_values("log_date") if "log_date" in perf_df.columns else perf_df
                    fig_e = go.Figure()
                    fig_e.add_trace(go.Scatter(x=plot_df.get("log_date", plot_df.index), y=plot_df["units_produced"],
                                                mode="lines+markers", name="Units produced",
                                                line=dict(color=COLORS["primary"], width=2.5)))
                    fig_e.update_layout(**PLOTLY_DARK, height=260, yaxis=dict(gridcolor=GRID_COLOR),
                                        xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                                        legend=dict(font=dict(color=COLORS["text_soft"])))
                    st.plotly_chart(fig_e, width='stretch', config={"displayModeBar": False})

                # --- Shift breakdown (Arabic Bread runs a separate night shift) ---
                if emp_dept == "Arabic Bread" and "shift" in perf_df.columns and perf_df["shift"].nunique() > 1:
                    with st.container(border=True):
                        st.markdown("###### Morning vs. night shift")
                        shift_summary = perf_df.groupby("shift").agg(
                            avg_units=("units_produced", "mean"),
                            avg_wastage_pct=("wastage_pct", "mean"),
                            shifts_logged=("shift", "count"),
                        ).round(1).reset_index()
                        st.dataframe(shift_summary, width='stretch', hide_index=True)

                st.dataframe(perf_df.drop(columns=[c for c in ["id"] if c in perf_df.columns]),
                             width='stretch', hide_index=True)

        # --- Daily log ---
        with etab2:
            st.caption("Log today's (or a past) shift in plain counts — no percentages to work out yourself. "
                       "Your performance report will still show the percentages, calculated automatically.")
            with st.form(f"daily_log_form_{emp_id}", clear_on_submit=True):
                log_date = st.date_input("Date", value=datetime.now().date())
                if emp_dept == "Arabic Bread":
                    shift = st.selectbox("Shift", ["Morning", "Night"],
                                          help="Arabic Bread runs a separate night shift in addition to the standard morning shift.")
                else:
                    shift = "Morning"
                colf1, colf2 = st.columns(2)
                units = colf1.number_input("Units produced (good units)", 0, 5000, 250, 10)
                units_wasted = colf2.number_input("Units wasted / discarded", 0, 1000, 6, 1,
                                                   help="How many units were thrown away, spoiled, or damaged this shift.")
                hours = colf1.number_input("Hours worked", 0.0, 16.0, STANDARD_SHIFT_HOURS, 0.5)
                units_failed_qc = colf2.number_input("Units that failed QC check", 0, 1000, 5, 1,
                                                       help="How many units failed a quality check this shift.")
                colf3, colf4 = st.columns(2)
                batches_completed = colf3.number_input("Batches completed", 0, 200, 10, 1)
                batches_on_time = colf4.number_input("Batches finished on/within standard time", 0, 200, 9, 1,
                                                       help="Of the batches completed, how many finished within the standard time for that batch?")
                revenue = st.number_input(
                    "Revenue/value generated (AED, optional)", 0.0, 50000.0, 0.0, 50.0,
                    help="Feeds the productivity benchmark: GM's rule of thumb is each employee should "
                         "generate ~AED 1,000/day. Leave at 0 if not tracked yet for this shift."
                )
                notes = st.text_area("Notes (optional)", placeholder="Any incidents, equipment issues, etc.")
                submitted = st.form_submit_button("Submit shift log", width='stretch')
                if submitted:
                    total_output = units + units_wasted
                    wastage_pct = round((units_wasted / total_output) * 100, 1) if total_output > 0 else 0.0
                    batch_adh_pct = round((batches_on_time / batches_completed) * 100) if batches_completed > 0 else 0
                    quality_pct = round(max(units - units_failed_qc, 0) / units * 100) if units > 0 else 0
                    ok, err = portal_insert("employee_performance", {
                        "employee_name": emp_name, "employee_id": emp_id, "department": emp_dept,
                        "log_date": str(log_date), "shift": shift, "units_produced": int(units),
                        "units_wasted": int(units_wasted), "batches_completed": int(batches_completed),
                        "batches_on_time": int(batches_on_time), "units_failed_qc": int(units_failed_qc),
                        "wastage_pct": wastage_pct, "hours_worked": float(hours),
                        "batch_time_adherence_pct": batch_adh_pct, "quality_pass_pct": quality_pct,
                        "revenue_generated": float(revenue), "notes": notes,
                    })
                    if ok:
                        st.toast(f"Shift log saved — wastage {wastage_pct}%, batch adherence {batch_adh_pct}%, "
                                 f"quality pass {quality_pct}% (calculated automatically from your counts).", icon="✅")
                        st.rerun()
                    else:
                        st.error(f"Could not save: {err}")

        # --- Goals & feedback ---
        with etab3:
            goals_df, err = portal_fetch("employee_goals", {"employee_id": emp_id})
            if err:
                st.error(f"Could not load goals: {err}")
            elif not goals_df.empty:
                for _, row in goals_df.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**Goal:** {row.get('goal_text', '—')}")
                        st.caption(f"Self-assessment: {row.get('self_assessment') or '—'}")
                        if row.get("manager_feedback"):
                            st.markdown(f"<span class='pill pill-ok'>Manager feedback</span> {row['manager_feedback']}",
                                        unsafe_allow_html=True)
            else:
                pass

            st.markdown("###### Add a new goal / self-assessment")
            with st.form(f"goals_form_{emp_id}", clear_on_submit=True):
                goal_text = st.text_area("Current goal / focus area")
                self_assessment = st.text_area("Self-assessment for this period")
                submitted_g = st.form_submit_button("Save", width='stretch')
                if submitted_g:
                    ok, err = portal_insert("employee_goals", {
                        "employee_name": emp_name, "employee_id": emp_id, "department": emp_dept,
                        "goal_text": goal_text, "self_assessment": self_assessment,
                        "manager_feedback": None,
                    })
                    if ok:
                        st.toast("Goal saved.", icon="✅")
                        st.rerun()
                    else:
                        st.error(f"Could not save: {err}")

    # -------------------- MANAGER / HR VIEW --------------------
    elif dept_choice == "🧑‍💼 Manager / HR view":
        all_perf, err = portal_fetch("employee_performance")
        if err:
            st.error(f"Could not load team data: {err}")
        elif all_perf.empty:
            pass
        else:
            for col in ["units_produced", "wastage_pct", "batch_time_adherence_pct", "quality_pass_pct", "revenue_generated", "hours_worked"]:
                if col in all_perf.columns:
                    all_perf[col] = pd.to_numeric(all_perf[col], errors="coerce")

            st.markdown("#### Performance comparison within a department")
            available_depts = sorted(all_perf["department"].dropna().unique()) if "department" in all_perf.columns else []
            if not available_depts:
                pass
            else:
                dept_filter = st.selectbox("Choose a department to compare", available_depts, key="mgr_dept_filter")
                tag_label, tag_cls = DEPARTMENT_PRIORITY_TAG.get(dept_filter, (None, None))
                tag_html = f"<span class='pill {tag_cls}' style='margin-left:8px;'>{tag_label}</span>" if tag_label else ""

                dept_perf = all_perf[all_perf["department"] == dept_filter].copy()

                def _value_per_day(row_group):
                    hrs = row_group["hours_worked"].mean()
                    hrs = hrs if hrs and hrs > 0 else STANDARD_SHIFT_HOURS
                    rev = row_group["revenue_generated"].mean() if "revenue_generated" in row_group.columns else 0.0
                    return (rev / hrs) * STANDARD_SHIFT_HOURS

                dept_summary = dept_perf.groupby("employee_name").agg(
                    shifts_logged=("employee_name", "count"),
                    avg_units=("units_produced", "mean"),
                    avg_wastage_pct=("wastage_pct", "mean"),
                    avg_batch_adherence=("batch_time_adherence_pct", "mean"),
                    avg_quality_pass=("quality_pass_pct", "mean"),
                ).round(1).reset_index()
                dept_summary["value_per_day_aed"] = dept_summary["employee_name"].apply(
                    lambda n: round(_value_per_day(dept_perf[dept_perf["employee_name"] == n]), 0))
                dept_summary["wastage_status"] = dept_summary["avg_wastage_pct"].apply(
                    lambda w: "✅ On target" if w <= WASTAGE_TARGET_PCT else "⚠️ Above 1% target")

                # Composite performance score (0-100) so "who's doing best" has a
                # single clear answer instead of reading several charts separately.
                max_units = dept_summary["avg_units"].max()
                dept_summary["units_normalized"] = (dept_summary["avg_units"] / max_units * 100) if max_units > 0 else 0
                dept_summary["performance_score"] = (
                    dept_summary["avg_quality_pass"] * 0.30
                    + dept_summary["avg_batch_adherence"] * 0.30
                    + (100 - dept_summary["avg_wastage_pct"]).clip(lower=0) * 0.20
                    + dept_summary["units_normalized"] * 0.20
                ).round(1)
                dept_summary = dept_summary.sort_values("performance_score", ascending=False).reset_index(drop=True)
                dept_summary["rank"] = dept_summary.index + 1

                cap_target = DEPARTMENT_CAPACITY_TARGET.get(dept_filter)

                with st.container(border=True):
                    st.markdown(f"**{dept_filter}**{tag_html} — {DEPARTMENT_HEADCOUNT.get(dept_filter, '—')} total headcount "
                                f"on the line, {dept_perf['employee_name'].nunique()} with logged shifts",
                                unsafe_allow_html=True)
                    if cap_target:
                        util_pct = (dept_summary["avg_units"].mean() / cap_target) * 100
                        st.caption(f"Capacity utilization vs illustrative target of {cap_target} units/day/employee: **{util_pct:.0f}%**")
                    if len(dept_summary) > 0:
                        top = dept_summary.iloc[0]
                        st.markdown(f"<span class='pill pill-ok'>🏆 Top performer: {top['employee_name']} "
                                    f"(score {top['performance_score']:.0f}/100)</span>", unsafe_allow_html=True)
                    display_cols = ["rank", "employee_name", "performance_score", "shifts_logged", "avg_units",
                                     "avg_wastage_pct", "avg_batch_adherence", "avg_quality_pass",
                                     "value_per_day_aed", "wastage_status"]
                    st.dataframe(dept_summary[display_cols], width='stretch', hide_index=True)
                    st.caption("Performance score = 30% quality pass rate + 30% batch-time adherence + "
                               "20% (100 − wastage %) + 20% units produced (normalized to the department's top performer).")
                    render_email_share(
                        f"dept_{dept_filter}", f"Email the {dept_filter} comparison",
                        section="employee_department",
                        context={
                            "department": dept_filter,
                            "headcount": DEPARTMENT_HEADCOUNT.get(dept_filter, "—"),
                            "tag_label": tag_label,
                            "summary_df": dept_summary[display_cols],
                        },
                    )

                with st.container(border=True):
                    st.markdown("###### Comparative progress across employees in this department")
                    fig_cmp = go.Figure()
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_quality_pass"],
                                              name="Quality pass %", marker_color=COLORS["primary"], marker_line_width=0))
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_batch_adherence"],
                                              name="Batch-time adherence %", marker_color=COLORS["tertiary"], marker_line_width=0))
                    fig_cmp.add_trace(go.Bar(x=dept_summary["employee_name"],
                                              y=(100 - dept_summary["avg_wastage_pct"]).clip(lower=0),
                                              name="100 − wastage %", marker_color=COLORS["secondary"], marker_line_width=0))
                    fig_cmp.update_layout(**PLOTLY_DARK, height=300, barmode="group", yaxis_title="%",
                                           yaxis=dict(gridcolor=GRID_COLOR, range=[0, 105]),
                                           xaxis=dict(gridcolor="rgba(0,0,0,0)"),
                                           legend=dict(orientation="h", y=-0.15, font=dict(color=COLORS["text_soft"])))
                    st.plotly_chart(fig_cmp, width='stretch', config={"displayModeBar": False})

                with st.container(border=True):
                    fig_dept = go.Figure()
                    fig_dept.add_trace(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_units"],
                                               name="Avg units/shift", marker_color=COLORS["primary"], marker_line_width=0))
                    fig_dept.update_layout(**PLOTLY_DARK, height=260, yaxis_title="Avg units / shift",
                                            yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                    st.plotly_chart(fig_dept, width='stretch', config={"displayModeBar": False})

                with st.container(border=True):
                    fig_dept2 = go.Figure(go.Bar(x=dept_summary["employee_name"], y=dept_summary["avg_wastage_pct"],
                                                  marker_color=COLORS["secondary"], marker_line_width=0))
                    fig_dept2.add_hline(y=WASTAGE_TARGET_PCT, line_dash="dot", line_color=COLORS["text_soft"],
                                         annotation_text="1% target", annotation_font_color=COLORS["text_soft"])
                    fig_dept2.update_layout(**PLOTLY_DARK, height=260, yaxis_title="Avg wastage %",
                                             yaxis=dict(gridcolor=GRID_COLOR), xaxis=dict(gridcolor="rgba(0,0,0,0)"))
                    st.plotly_chart(fig_dept2, width='stretch', config={"displayModeBar": False})

                if dept_filter == "Arabic Bread" and "shift" in dept_perf.columns and dept_perf["shift"].nunique() > 1:
                    with st.container(border=True):
                        st.markdown("###### Morning vs. night shift — Arabic Bread")
                        shift_cmp = dept_perf.groupby("shift").agg(
                            avg_units=("units_produced", "mean"),
                            avg_wastage_pct=("wastage_pct", "mean"),
                            shifts_logged=("shift", "count"),
                        ).round(1).reset_index()
                        st.dataframe(shift_cmp, width='stretch', hide_index=True)

            st.markdown("#### All departments — overview")
            dept_overview = all_perf.groupby("department").agg(
                shifts_logged=("department", "count"),
                avg_units=("units_produced", "mean"),
                avg_wastage_pct=("wastage_pct", "mean"),
            ).round(1).reset_index() if "department" in all_perf.columns else pd.DataFrame()
            if not dept_overview.empty:
                dept_overview["headcount"] = dept_overview["department"].map(DEPARTMENT_HEADCOUNT)
                dept_overview["capacity_target"] = dept_overview["department"].map(DEPARTMENT_CAPACITY_TARGET)
                dept_overview["utilization_pct"] = (dept_overview["avg_units"] / dept_overview["capacity_target"] * 100).round(0)
                dept_overview["focus"] = dept_overview["department"].map(lambda d: DEPARTMENT_PRIORITY_TAG.get(d, ("", ""))[0])
                with st.container(border=True):
                    st.dataframe(dept_overview.sort_values("avg_units", ascending=False), width='stretch', hide_index=True)

        st.markdown("#### Recent goals & self-assessments")
        all_goals, err = portal_fetch("employee_goals")
        if err:
            st.error(f"Could not load goals: {err}")
        elif all_goals.empty:
            pass
        else:
            for _, row in all_goals.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{row.get('employee_name','—')}** ({row.get('department','—')}) — {row.get('goal_text','—')}")
                    st.caption(f"Self-assessment: {row.get('self_assessment') or '—'}")
                    if row.get("manager_feedback"):
                        st.markdown(f"<span class='pill pill-ok'>Feedback given</span> {row['manager_feedback']}",
                                    unsafe_allow_html=True)

        st.markdown("###### Give feedback on an employee's goal")
        with st.form("manager_feedback_form", clear_on_submit=True):
            fb_labels = [f"{e['department']} — {e['name']}" for e in EMPLOYEES]
            fb_choice = st.selectbox("Employee", fb_labels, key="fb_employee")
            fb_text = st.text_area("Feedback")
            fb_submit = st.form_submit_button("Submit feedback", width='stretch')
            if fb_submit:
                fb_emp = EMPLOYEES[fb_labels.index(fb_choice)]
                ok, err = portal_insert("employee_goals", {
                    "employee_name": fb_emp["name"], "employee_id": fb_emp["id"], "department": fb_emp["department"],
                    "goal_text": "(Manager feedback entry)", "self_assessment": None,
                    "manager_feedback": fb_text,
                })
                if ok:
                    st.toast("Feedback recorded.", icon="✅")
                    st.rerun()
                else:
                    st.error(f"Could not save: {err}")
    else:
        st.info("Select your department and name (or Manager / HR view) above to get started.")

# ========================================================================
# SECTION D — COMPANY PROFILE
# ========================================================================
elif section == "📋  Company profile":
    st.caption("Everything collected from the Grandiose team so far — from the 27 Jul personal meeting with the GM")

    COMPANY_INFO = [
        ("🏭", "Production capacity & growth", [
            "Flour Country Bakery is currently running at only 30-35% utilization — roughly 65% headroom still available.",
            "Target is 100% utilization within a year, with an estimated AED 10M/month production potential once reached.",
            "Granola Catering is running at approximately 50% utilization.",
            "Company-wide headcount is ~110 today, with a growth target of 200 within a year.",
        ]),
        ("⏰", "Shift structure", [
            "Standard shift is 9 hours, including a break.",
            "Arabic Bread is the one line that also runs a separate night shift, in addition to the standard morning shift.",
        ]),
        ("🛒", "Sales channel mix (B2B vs. B2C)", [
            "Flour Country Bakery is 100% B2B.",
            "Catering is roughly 35% B2C and 65% B2B.",
        ]),
        ("💰", "Cost structure & targets", [
            "Food cost: capped at a maximum of 30%.",
            "Labour cost: currently running at ~40% — target is 16-18%, with 20% as the absolute ceiling.",
            "Utilities: capped at a maximum of 10%.",
            "Rent: capped at a maximum of 15%.",
            "Net profit target: 20-25%.",
            "Total expense ceiling (all costs combined): 75-80% of revenue.",
        ]),
        ("🏷️", "Pricing approach", [
            "Material cost is scaled 35-50% depending on client volume.",
            "Supermarket supply is priced cost-to-cost, with no margin built in.",
            "B2B contract pricing varies by negotiation and account.",
        ]),
        ("📈", "Productivity benchmark", [
            "No formal labour productivity measurement system exists yet.",
            "The GM's working rule of thumb: each employee should generate approximately AED 1,000/day of value.",
            "This is the benchmark now built into the Employee Portal's productivity tracking.",
        ]),
        ("🔥", "Top & underutilized product lines", [
            "Top revenue/profit drivers: daily bread (Arabic bread, toasted bread, buns, Samurai sandwich) and pastry.",
            "Least-utilized lines: Tahina and Turkish baklava — both flagged as ramp-up focus areas.",
        ]),
        ("📦", "Inventory management", [
            "Stock is managed on a FIFO (first-in, first-out) basis.",
            "Ordering cycles: bi-monthly for shelf-stable/dry goods, weekly for short-shelf-life items, daily for fresh fruit and vegetables.",
            "Payment terms: ~95% of suppliers are on net 90 days (60 days from statement of account); a few critical items run 30-60 days.",
            "Approved vendor base: ~65-70 for bakery, ~160-170 for catering.",
            "Safety stock normally aligns to supplier lead times, extending to ~3 months during geopolitical disruptions.",
            "Packaging (sourced from China/Indonesia): ~1 month of stock held against a 15-day lead time.",
            "De-risking in progress: alternate vendors are being onboarded for critical ingredients.",
            "KPIs monitored: inventory movement (fast/slow movers), depletion cycles, and overstock/obsolescence — reviewed bi-weekly.",
        ]),
        ("🗑️", "Wastage management", [
            "Current wastage: 2-3% — target is 1%.",
            "Plan to close the gap: grow sales volume and push supermarket promotions on affected items.",
            "Near-expiry raw materials are redirected into production rather than discarded.",
            "Near-expiry finished goods (with shelf life over 30 days) move to clearance pricing.",
            "Fully expired stock is written off.",
            "Yield: catering runs ~90-95% (about 5% loss to spillage/damage); bakery yield loss is minimal; frozen raw meat yield loss is 20-25%.",
        ]),
        ("🗓️", "Dashboard & next steps (from this meeting)", [
            "ERP integration is planned so the dashboard can eventually auto-update from live company data.",
            "In-dashboard email sharing was requested — ✅ already built into this dashboard.",
            "An employee productivity module linking individual output to section-level sales was requested — ✅ already built into the Employee Portal.",
            "A separate, equivalent dashboard is planned for Granola Catering.",
            "The GM wants the dashboard link shared for internal review before Friday.",
            "Next steps: send Friday feedback slots, share the dashboard link, and continue refining with confirmed figures ahead of the mid-review.",
        ]),
    ]

    for icon, title, points in COMPANY_INFO:
        with st.container(border=True):
            rows = "".join([f"<div class='info-row'><span>{p}</span></div>" for p in points])
            st.markdown(f"<div style='padding:14px 18px;'><h4 style='margin:0 0 10px 0; font-size:1.05rem;'>"
                        f"{icon} {title}</h4>{rows}</div>", unsafe_allow_html=True)

    st.caption("Source: personal meeting with the Grandiose GM, 27 Jul. Figures here are as reported in that "
               "conversation and should be treated as the current source of truth pending any further updates.")
    render_email_share("company_profile", "Email the Company Profile", section="company_profile",
                        context={"company_info": COMPANY_INFO})

st.markdown("---")
st.caption("Financial Performance and Cost Optimization for Grandiose Bakery Operations · GIP III · "
           "Bakery division only, catering excluded · Figures shown are illustrative benchmarks pending "
           "Grandiose-provided actuals.")
