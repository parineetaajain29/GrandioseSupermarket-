"""
Grandiose Bakery — Financial Performance & Cost Optimization Dashboard
GIP III | Parineeta Jain, Rajveer Singh, Tarang Gupta

Two-tab dashboard:
  Tab A - Performance tracker (KPIs, trend, category panels)
  Tab B - Scenario & resilience (4 modules: inflation, disruption risk,
          pandemic preparedness, supplier concentration)

Scope: bakery operations only (catering excluded per GIP III scope).
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

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
# COLOR PALETTE — muted, desaturated tones echoing Grandiose's fresh
# neighbourhood-grocery identity (sage green / warm terracotta / cream),
# deliberately avoiding bold/saturated primary colours.
# ----------------------------------------------------------------------
COLORS = {
    "bg":        "#F7F4EE",   # warm cream page background
    "surface":   "#FFFFFF",   # card surface
    "surface2":  "#F0EBE1",   # secondary surface / table stripe
    "primary":   "#5B7F5E",   # muted sage green (brand-adjacent, desaturated)
    "primary_d": "#3F5C42",   # darker sage for text on light green
    "accent":    "#C08552",   # muted terracotta / warm bakery tone
    "accent_d":  "#8C5A34",   # darker terracotta for text
    "warn":      "#C9A227",   # muted amber (not bright yellow)
    "danger":    "#A6564B",   # muted brick red (not bright red)
    "text":      "#3A3532",   # warm dark neutral for body text
    "text_soft": "#6B655F",   # secondary text
    "border":    "#DDD6C9",
}

CATEGORICAL = [COLORS["primary"], COLORS["accent"], "#7C9A8E", "#B08968", "#95836E"]

st.markdown(f"""
<style>
    .stApp {{ background-color: {COLORS['bg']}; }}
    section[data-testid="stSidebar"] {{ background-color: {COLORS['surface2']}; }}
    div[data-testid="stMetric"] {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']};
        border-radius: 10px;
        padding: 14px 16px 10px 16px;
    }}
    div[data-testid="stMetricLabel"] {{ color: {COLORS['text_soft']}; }}
    div[data-testid="stMetricValue"] {{ color: {COLORS['text']}; }}
    .card {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']};
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 14px;
    }}
    .card h4 {{ color: {COLORS['primary_d']}; margin-top:0; margin-bottom: 10px; }}
    .badge {{
        display:inline-block; padding:3px 10px; border-radius:14px;
        font-size:0.78rem; font-weight:600;
    }}
    .badge-ok {{ background-color:#E7EEE3; color:{COLORS['primary_d']}; }}
    .badge-warn {{ background-color:#F4EBD3; color:#7A5E10; }}
    .badge-risk {{ background-color:#F1DEDA; color:{COLORS['danger']}; }}
    h1, h2, h3 {{ color: {COLORS['text']}; }}
    .stTabs [data-baseweb="tab"] {{ color: {COLORS['text_soft']}; }}
    .stTabs [aria-selected="true"] {{ color: {COLORS['primary_d']}; font-weight:600; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor=COLORS["surface"],
    plot_bgcolor=COLORS["surface"],
    font=dict(color=COLORS["text"], size=13),
    margin=dict(l=10, r=10, t=30, b=10),
)

# ----------------------------------------------------------------------
# SAMPLE / BENCHMARK DATA (placeholder until mentor-provided actuals)
# ----------------------------------------------------------------------
months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul"]
food_cost_trend = [29.8, 30.1, 30.6, 30.9, 31.0, 31.4]
target_line = [30.0] * 6

baseline = {
    "food_cost_pct": 31.4,
    "wastage_pct": 4.8,
    "gross_margin_pct": 42.1,
    "cost_per_unit": 3.85,
    "standard_cost_per_unit": 3.60,
}

category_panels = {
    "Production": {
        "Batch efficiency": "88%",
        "Labor cost / unit": "AED 1.10",
        "Std vs actual cost": "+6.9%",
    },
    "Procurement": {
        "Purchase price variance": "+3.2%",
        "Supplier lead time": "2.4 days",
        "Top-3 flour supplier share": "76%",
    },
    "Warehouse & inventory": {
        "Inventory turnover": "14.2x/yr",
        "Expired/damaged stock": "AED 8,400",
        "Slow-moving SKUs": "6",
    },
    "Cost control & margin": {
        "Margin leakage (est.)": "AED 42,000/mo",
        "Packaging cost impact": "2.1% of sales",
        "Overhead allocation": "9.4% of sales",
    },
}

st.sidebar.title("🍞 Grandiose Bakery")
st.sidebar.caption("GIP III — Financial performance & cost optimization\nBakery division only · catering excluded")
st.sidebar.markdown("---")
st.sidebar.markdown("**Prepared by**\n\nParineeta Jain · Rajveer Singh · Tarang Gupta")
st.sidebar.markdown("---")
st.sidebar.info("All figures below are illustrative benchmarks pending Grandiose's actual bakery data.")

st.title("Grandiose Bakery — Financial Performance Dashboard")

tab_track, tab_scn = st.tabs(["📊 Performance tracker", "🧭 Scenario & resilience"])

# ========================================================================
# TAB A — PERFORMANCE TRACKER
# ========================================================================
with tab_track:
    st.caption("Where the bakery division stands today")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Food cost %", f"{baseline['food_cost_pct']}%", "1.2pt vs target", delta_color="inverse")
    c2.metric("Wastage %", f"{baseline['wastage_pct']}%", "-0.3pt vs last month", delta_color="inverse")
    c3.metric("Gross margin", f"{baseline['gross_margin_pct']}%", "flat vs last month")
    c4.metric("Cost per unit (AED)", f"{baseline['cost_per_unit']:.2f}", f"vs std {baseline['standard_cost_per_unit']:.2f}", delta_color="inverse")

    st.markdown("#### Food cost % vs target — last 6 months")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=months, y=food_cost_trend, mode="lines+markers",
                              name="Food cost %", line=dict(color=COLORS["primary"], width=3),
                              fill="tozeroy", fillcolor="rgba(91,127,94,0.12)"))
    fig.add_trace(go.Scatter(x=months, y=target_line, mode="lines", name="Target",
                              line=dict(color=COLORS["text_soft"], width=1.5, dash="dash")))
    fig.update_layout(**PLOTLY_LAYOUT, yaxis=dict(ticksuffix="%", range=[28, 33]), height=300,
                       legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Category panels")
    cols = st.columns(4)
    for col, (name, metrics) in zip(cols, category_panels.items()):
        with col:
            rows = "".join([f"<div style='display:flex;justify-content:space-between;padding:4px 0;"
                             f"border-bottom:1px solid {COLORS['border']};font-size:0.88rem;'>"
                             f"<span style='color:{COLORS['text_soft']}'>{k}</span><span>{v}</span></div>"
                             for k, v in metrics.items()])
            st.markdown(f"<div class='card'><h4>{name}</h4>{rows}</div>", unsafe_allow_html=True)

# ========================================================================
# TAB B — SCENARIO & RESILIENCE
# ========================================================================
with tab_scn:
    st.caption("What happens if — stress-tests the same cost and margin outputs under new assumptions")

    mod1, mod2, mod3, mod4 = st.tabs([
        "1 · Inflation sensitivity", "2 · Supply disruption risk",
        "3 · Pandemic preparedness", "4 · Supplier alternatives",
    ])

    # ---------------- MODULE 1: INFLATION-ADJUSTED COST SENSITIVITY ----------------
    with mod1:
        st.markdown("##### Inflation-adjusted cost sensitivity")
        st.caption("Dual-track inflation input with an adjustable bar, net of any subsidy/offset deducted.")

        colL, colR = st.columns([1, 1.3])
        with colL:
            headline_inf = st.slider("Headline inflation forecast (%)", 0.0, 10.0, 2.8, 0.1)
            food_inf = st.slider("Actual food-input inflation (%)", 0.0, 20.0, 5.3, 0.1,
                                  help="Flour, dairy, and packaging inflation can diverge sharply from headline CPI during shocks.")
            subsidy_offset = st.number_input("Subsidy / price-cap offset (AED per unit)", 0.0, 2.0, 0.05, 0.01)

        base_cost = baseline["cost_per_unit"]
        adj_cost_headline = round(base_cost * (1 + headline_inf / 100) - subsidy_offset, 3)
        adj_cost_food = round(base_cost * (1 + food_inf / 100) - subsidy_offset, 3)
        base_revenue_per_unit = base_cost / (baseline["food_cost_pct"] / 100)
        food_cost_pct_adj = round((adj_cost_food / base_revenue_per_unit) * 100, 1)
        margin_adj = round(100 - food_cost_pct_adj - (100 - baseline["gross_margin_pct"] - baseline["food_cost_pct"]), 1)

        with colR:
            m1, m2, m3 = st.columns(3)
            m1.metric("Cost/unit — headline scenario", f"AED {adj_cost_headline:.2f}")
            m2.metric("Cost/unit — food-inflation scenario", f"AED {adj_cost_food:.2f}",
                      f"{adj_cost_food - base_cost:+.2f} vs current")
            m3.metric("Food cost % — food-inflation scenario", f"{food_cost_pct_adj}%",
                      f"{food_cost_pct_adj - baseline['food_cost_pct']:+.1f}pt", delta_color="inverse")

            fig1 = go.Figure(go.Bar(
                x=["Current", "Headline scenario", "Food-inflation scenario"],
                y=[base_cost, adj_cost_headline, adj_cost_food],
                marker_color=[COLORS["text_soft"], COLORS["accent"], COLORS["danger"]],
            ))
            fig1.update_layout(**PLOTLY_LAYOUT, height=260, yaxis_title="AED per unit")
            st.plotly_chart(fig1, use_container_width=True)

        st.caption("Context: UAE food inflation has ranged from a 2022 peak near 9% to negative readings in 2021, "
                   "against headline forecasts around 2–3% for 2026 — a single flat assumption misses this swing.")

    # ---------------- MODULE 2: NATURAL CALAMITY / SUPPLY DISRUPTION ----------------
    with mod2:
        st.markdown("##### Natural calamity & supply disruption risk")
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

        avg_daily_cost = 4200  # illustrative average daily flour/ingredient spend for bakery division
        buffer_stock_days = max(delay_days, 3)
        buffer_stock_cost = round(buffer_stock_days * avg_daily_cost * 0.02, 0)  # ~2% holding cost/day
        expected_stockout_cost = round(stockout_prob * avg_daily_cost * delay_days * 1.5, 0)

        st.markdown(
            f"<div class='card'><h4>Recommended buffer vs expected stockout cost</h4>"
            f"Holding <b>{buffer_stock_days} days</b> of buffer stock costs approximately "
            f"<b>AED {buffer_stock_cost:,.0f}</b>. The expected cost of not holding buffer, given this scenario's "
            f"probability, is approximately <b>AED {expected_stockout_cost:,.0f}</b>.</div>",
            unsafe_allow_html=True)

        fig2 = go.Figure(go.Bar(
            x=["Cost of holding buffer stock", "Expected cost of stockout"],
            y=[buffer_stock_cost, expected_stockout_cost],
            marker_color=[COLORS["primary"], COLORS["danger"]],
        ))
        fig2.update_layout(**PLOTLY_LAYOUT, height=260, yaxis_title="AED")
        st.plotly_chart(fig2, use_container_width=True)

    # ---------------- MODULE 3: PANDEMIC PREPAREDNESS ----------------
    with mod3:
        st.markdown("##### Pandemic preparedness")
        st.caption("Two linked panels: customer retention/attraction (demand side) and supply chain optimization (supply side).")

        colX, colY = st.columns(2)
        with colX:
            st.markdown(f"<div class='card'><h4 style='color:{COLORS['accent_d']}'>Customer retention / attraction</h4>", unsafe_allow_html=True)
            repeat_rate = st.slider("Repeat-purchase rate (%)", 0, 100, 62, key="repeat")
            delivery_share = st.slider("Delivery/online-order revenue share (%)", 0, 100, 28, key="delivery")
            basket_size = st.slider("Average basket size (AED)", 20, 200, 74, key="basket")
            flags = []
            if repeat_rate < 50: flags.append("Repeat-purchase rate below healthy baseline")
            if delivery_share < 15: flags.append("Delivery share low — limited resilience to footfall shocks")
            badge = "badge-risk" if flags else "badge-ok"
            label = "Needs attention" if flags else "Within healthy range"
            st.markdown(f"<span class='badge {badge}'>{label}</span></div>", unsafe_allow_html=True)
            for f in flags:
                st.caption(f"⚠️ {f}")

        with colY:
            st.markdown(f"<div class='card'><h4 style='color:{COLORS['primary_d']}'>Supply chain optimization</h4>", unsafe_allow_html=True)
            safety_days = st.slider("Safety stock (days of cover)", 0, 30, 9, key="safety")
            alt_suppliers = st.slider("Qualified alternate suppliers per key ingredient", 0, 5, 1, key="alt")
            single_sourced = st.slider("% ingredients single-sourced", 0, 100, 45, key="single")
            flags2 = []
            if safety_days < 7: flags2.append("Safety stock below 7-day resilience threshold")
            if alt_suppliers < 2: flags2.append("Fewer than 2 qualified alternates — concentration risk")
            if single_sourced > 40: flags2.append("Over 40% of ingredients single-sourced")
            badge2 = "badge-risk" if flags2 else "badge-ok"
            label2 = "Needs attention" if flags2 else "Within healthy range"
            st.markdown(f"<span class='badge {badge2}'>{label2}</span></div>", unsafe_allow_html=True)
            for f in flags2:
                st.caption(f"⚠️ {f}")

        st.caption("Rationale: pandemics create a dual shock — demand shifting from in-store to delivery, and "
                   "supply disruption at the same time — so both sides need to be tracked together as an early-warning system.")

    # ---------------- MODULE 4: SUPPLIER ALTERNATIVES / CONCENTRATION ----------------
    with mod4:
        st.markdown("##### Raw material & supplier alternatives")
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
        hhi = round((shares_norm ** 2).sum() * 100, 0)  # scaled to conventional 0-10,000 HHI

        if hhi < 1500:
            risk_label, risk_badge = "Low concentration risk", "badge-ok"
        elif hhi < 2500:
            risk_label, risk_badge = "Moderate concentration risk", "badge-warn"
        else:
            risk_label, risk_badge = "High concentration risk", "badge-risk"

        colM, colN = st.columns([1, 2])
        with colM:
            st.metric("Supplier concentration index (HHI)", f"{hhi:,.0f}")
            st.markdown(f"<span class='badge {risk_badge}'>{risk_label}</span>", unsafe_allow_html=True)
            st.caption("Reference bands: <1,500 low · 1,500–2,500 moderate · >2,500 high (standard HHI convention).")

        with colN:
            fig3 = go.Figure(go.Pie(labels=edited["Supplier / origin"], values=shares_norm,
                                     hole=0.45, marker=dict(colors=CATEGORICAL)))
            fig3.update_layout(**PLOTLY_LAYOUT, height=280)
            st.plotly_chart(fig3, use_container_width=True)

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
