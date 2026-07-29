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
from fpdf import FPDF

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
    "bg":        "#0B0B0F",   # near-black page background
    "surface":   "#16171D",   # card surface
    "surface2":  "#1E1F27",   # sidebar / secondary surface
    "border":    "#2A2B33",
    "text":      "#F5F5F7",   # near-white
    "text_soft": "#9B9BA5",   # muted gray secondary text
    "primary":   "#FFB020",   # vibrant amber/gold — bakery warmth
    "primary_d": "#7A5A16",
    "secondary": "#FF5C8A",   # vibrant coral/pink
    "tertiary":  "#7C6FF0",   # vibrant violet
    "success":   "#3ECF8E",
    "warning":   "#FFC94D",
    "danger":    "#FF5C5C",
}

CATEGORICAL = [COLORS["primary"], COLORS["secondary"], COLORS["tertiary"], COLORS["success"], "#4DD0E1"]

st.markdown(f"""
<style>
    .stApp {{ background-color: {COLORS['bg']}; }}
    section[data-testid="stSidebar"] {{ background-color: {COLORS['surface2']}; border-right: 1px solid {COLORS['border']}; }}

    .stApp, .stApp p, .stApp span, .stApp li, .stApp label,
    div[data-testid="stMarkdownContainer"], div[data-testid="stMarkdownContainer"] p {{
        color: {COLORS['text']} !important;
    }}
    div[data-testid="stCaptionContainer"], .stApp small {{ color: {COLORS['text_soft']} !important; }}
    h1, h2, h3, h4, h5 {{ color: {COLORS['text']} !important; font-weight: 700 !important; }}
    section[data-testid="stSidebar"] * {{ color: {COLORS['text']} !important; }}

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
    .kpi-value {{ font-size:2.1rem; font-weight:800; margin: 4px 0 2px 0; letter-spacing:-0.5px; }}
    .pill {{
        display:inline-block; padding:3px 11px; border-radius:20px;
        font-size:0.76rem; font-weight:700;
    }}
    .pill-up-bad {{ background-color: rgba(255,92,92,0.15); color:{COLORS['danger']} !important; }}
    .pill-up-good {{ background-color: rgba(62,207,142,0.15); color:{COLORS['success']} !important; }}
    .pill-down-good {{ background-color: rgba(62,207,142,0.15); color:{COLORS['success']} !important; }}
    .pill-flat {{ background-color: rgba(155,155,165,0.15); color:{COLORS['text_soft']} !important; }}
    .pill-ok {{ background-color: rgba(62,207,142,0.15); color:{COLORS['success']} !important; }}
    .pill-warn {{ background-color: rgba(255,201,77,0.15); color:{COLORS['warning']} !important; }}
    .pill-risk {{ background-color: rgba(255,92,92,0.15); color:{COLORS['danger']} !important; }}

    .cat-row {{
        display:flex; justify-content:space-between; padding:7px 0;
        border-bottom:1px solid {COLORS['border']}; font-size:0.9rem;
    }}
    .cat-row span:first-child {{ color:{COLORS['text_soft']}; }}
    .cat-row span:last-child {{ font-weight:600; }}

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
        color: #1A1300 !important; font-weight:700 !important;
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

def hex_to_rgba(hex_color, alpha):
    hex_color = hex_color.lstrip("#")
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"

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
# STATUS REPORT (PDF) + EMAIL DELIVERY
# ----------------------------------------------------------------------
def build_status_report_pdf():
    """Builds a one-page PDF snapshot of current KPIs and category panels."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Grandiose Bakery - Financial Status Report", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 6, f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')} | GIP III | Bakery division only", ln=True)
    pdf.ln(4)

    pdf.set_text_color(20, 20, 20)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Headline KPIs", ln=True)
    pdf.set_font("Helvetica", "", 11)
    kpi_lines = [
        f"Food cost %: {baseline['food_cost_pct']}%  (1.2pt above target)",
        f"Wastage %: {baseline['wastage_pct']}%  (down 0.3pt vs last month)",
        f"Gross margin: {baseline['gross_margin_pct']}%  (flat vs last month)",
        f"Cost per unit: AED {baseline['cost_per_unit']:.2f}  (standard: AED {baseline['standard_cost_per_unit']:.2f})",
    ]
    for line in kpi_lines:
        pdf.cell(0, 7, f"- {line}", ln=True)
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Category panels", ln=True)
    for name, metrics in category_panels.items():
        clean_name = name.split(" ", 1)[-1]  # drop emoji for PDF font compatibility
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, clean_name, ln=True)
        pdf.set_font("Helvetica", "", 10)
        for k, v in metrics.items():
            pdf.cell(0, 6, f"    {k}: {v}", ln=True)
        pdf.ln(1)

    pdf.ln(2)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(120, 120, 120)
    pdf.multi_cell(0, 5, "Prepared by Parineeta Jain, Rajveer Singh, and Tarang Gupta. "
                         "Figures shown are illustrative benchmarks pending Grandiose's actual bakery data. "
                         "Catering is excluded, per GIP III scope.")

    return bytes(pdf.output())


def send_report_email(recipient_email, note=""):
    """Sends the status-report PDF to recipient_email using SMTP credentials
    stored in Streamlit secrets. Returns (success: bool, message: str)."""
    try:
        sender_email = st.secrets["EMAIL_ADDRESS"]
        sender_password = st.secrets["EMAIL_APP_PASSWORD"]
    except (KeyError, FileNotFoundError):
        return False, ("Email is not configured yet. Add EMAIL_ADDRESS and EMAIL_APP_PASSWORD "
                        "in Streamlit Cloud's Secrets settings (App settings -> Secrets) — "
                        "never in the code or GitHub repo.")

    smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = "Grandiose Bakery — Financial Status Report"
    body = ("Hi,\n\nPlease find attached the latest Grandiose bakery financial status report, "
            "generated from the live dashboard.\n")
    if note:
        body += f"\nNote from sender:\n{note}\n"
    body += "\n— Sent automatically from the Grandiose Bakery Dashboard (GIP III)"
    msg.attach(MIMEText(body, "plain"))

    pdf_bytes = build_status_report_pdf()
    attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
    filename = f"Grandiose_Bakery_Status_Report_{datetime.now().strftime('%Y%m%d')}.pdf"
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
    if st.button("Send report", key="send_email_btn", use_container_width=True):
        email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
        if not recipient or not re.match(email_pattern, recipient):
            st.error("Enter a valid email address.")
        else:
            with st.spinner("Sending..."):
                success, message = send_report_email(recipient, note)
            if success:
                st.success(message)
            else:
                st.error(message)
    st.caption("Sends a one-page PDF snapshot of the current KPIs and category panels.")


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
    ["📊  Performance tracker", "🧭  Scenario & resilience"],
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
                st.plotly_chart(sparkline(trend, color), use_container_width=True,
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
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("#### Category panels")
    cols2 = st.columns(4)
    for col, (name, metrics) in zip(cols2, category_panels.items()):
        with col:
            with st.container(border=True):
                rows = "".join([f"<div class='cat-row'><span>{k}</span><span>{v}</span></div>" for k, v in metrics.items()])
                st.markdown(f"<div style='padding:12px 16px;'><h4 style='margin:0 0 8px 0; font-size:1rem;'>{name}</h4>{rows}</div>",
                            unsafe_allow_html=True)

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
                st.plotly_chart(donut, use_container_width=True, config={"displayModeBar": False}, key="inflation_donut")

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
                st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})

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
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

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
        edited = st.data_editor(default_suppliers, num_rows="fixed", use_container_width=True, hide_index=True)

        shares = edited["Spend share (%)"].astype(float)
        total = shares.sum()
        shares_norm = shares / total * 100 if total > 0 else shares
        hhi = round((shares_norm ** 2).sum() * 100, 0)

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
                st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

        st.markdown("###### Pre-identified alternate suppliers (illustrative)")
        alt_df = pd.DataFrame({
            "Alternate supplier": ["Local UAE mill (blended flour)", "Turkey-origin supplier", "Egypt-origin supplier"],
            "Cost delta vs current": ["+4.5%", "+2.1%", "+1.8%"],
            "Lead-time delta": ["-3 days", "-1 day", "0 days"],
        })
        st.dataframe(alt_df, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption("Financial Performance and Cost Optimization for Grandiose Bakery Operations · GIP III · "
           "Bakery division only, catering excluded · Figures shown are illustrative benchmarks pending "
           "Grandiose-provided actuals.")
