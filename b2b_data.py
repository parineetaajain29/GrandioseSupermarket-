"""B2B Client Performance data — Flour Country / Grandiose Bakery's B2B expansion.

============================ DEMO DATA ============================
Everything in this module is illustrative, pending Grandiose-provided
actuals — consistent with the "Figures shown are illustrative benchmarks
pending Grandiose-provided actuals" disclaimer in the app footer. Kept
apart from the page code so swapping in real figures is a change to THIS
FILE ONLY; the page and b2b_calc.py only care about the shapes below.

Client revenues in ACCOUNTS sum exactly to the B2B revenue KPI shown on the
page — that KPI is computed FROM this list, not hardcoded separately, so
the two can never drift apart.
===================================================================
"""

# ---- KPI strip context ----
# Prior month's total B2B revenue, illustrative — used only for the MoM delta.
PRIOR_MONTH_B2B_REVENUE_AED = 132_000.0

# Reused from the existing Performance Tracker benchmark (`baseline
# ["gross_margin_pct"]` in app.py) as the retail comparison point for the
# B2B net-margin KPI. TODO: CONFIRM WITH FINANCE — this is retail *gross*
# margin, not a true retail *net* margin; swap in the real figure once
# finance can provide a directly comparable net number.
RETAIL_MARGIN_PCT_BENCHMARK = 42.1

# Documented Grandiose supplier terms (see Company Profile: "~95% of
# suppliers are on net 90 days, 60 days from statement of account") — shown
# only as context next to average collection days, not used in the calc.
SUPPLIER_TERMS_DAYS = 90

# ---- Revenue vs. service cost, 13 weeks (company-wide B2B) ----
WEEKS = [f"W{i}" for i in range(1, 14)]
WEEKLY_B2B_REVENUE_AED = [
    30200, 31000, 30800, 32100, 33000, 32500, 34200,
    35000, 34800, 36200, 37000, 37500, 38100,
]
WEEKLY_SERVICE_COST_AED = [
    21800, 22300, 22600, 23500, 24400, 24900, 26800,
    27600, 28400, 29800, 31000, 32200, 33400,
]

# ---- Capacity economics ----
# Oven-capacity split. Retail + B2B = ~32.4% utilised, matching the Company
# Profile's stated 30-35% overall utilization range. TODO: CONFIRM WITH GM —
# no per-channel oven-capacity breakdown exists yet; this split is a
# placeholder pending Grandiose's actual figures.
OVEN_CAPACITY_RETAIL_PCT = 24.0
OVEN_CAPACITY_B2B_PCT = 8.4
OVEN_CAPACITY_IDLE_PCT = 67.6

# A representative order, shown twice in the Capacity economics section: once
# assuming it fills idle capacity, once assuming it forces overtime — the
# same formula, different marginal-cost inputs.
MARGINAL_EXAMPLE_ORDER = {
    "order_value": 5000.0,
    "ingredient_cost": 1750.0,
    "packaging_cost": 250.0,
    "delivery_cost": 350.0,
}
MARGINAL_EXAMPLE_IDLE = {"incremental_labour_cost": 0.0, "overtime_premium": 0.0}
MARGINAL_EXAMPLE_OVERTIME = {"incremental_labour_cost": 600.0, "overtime_premium": 450.0}

# Next week's scheduled B2B orders — illustrative, used only to count how
# many land in an overtime slot.
NEXT_WEEK_ORDERS = [
    {"client": "Carrefour Hypermarket", "forces_overtime": False},
    {"client": "Spinneys", "forces_overtime": False},
    {"client": "Choithrams", "forces_overtime": False},
    {"client": "Zoom Supermarket", "forces_overtime": False},
    {"client": "Kibsons", "forces_overtime": False},
    {"client": "Al Maya Supermarket", "forces_overtime": False},
    {"client": "Grand Hyper", "forces_overtime": False},
    {"client": "Carrefour Hypermarket", "forces_overtime": True},
    {"client": "Grand Hyper", "forces_overtime": True},
]

# ---- Accounts ----
# Each account carries both the fully-loaded cost (for Margin, full
# absorption) and the variable-only cost components (for Marginal, per the
# §B2 formula). "Grand Hyper" is deliberately negative on full absorption
# but positive at the margin — it fills otherwise-idle oven capacity, so its
# incremental labour and overtime premium are both ~0 — which is the whole
# point of the Marginal column: dropping this account on the Margin number
# alone would leave that capacity cold while removing real contribution.
ACCOUNTS = [
    {
        "client": "Carrefour Hypermarket", "location": "Dubai",
        "delivery_frequency": "Daily",
        "revenue": 42000.0, "service_cost_full": 27300.0,
        "ingredient_cost": 14700.0, "packaging_cost": 2100.0, "delivery_cost": 2940.0,
        "incremental_labour_cost": 1260.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 29, "total_deliveries": 30,
        "receivable_amount": 8200.0, "days_outstanding": 19,
        "note": None,
    },
    {
        "client": "Spinneys", "location": "Abu Dhabi",
        "delivery_frequency": "Daily",
        "revenue": 38500.0, "service_cost_full": 26000.0,
        "ingredient_cost": 13475.0, "packaging_cost": 1925.0, "delivery_cost": 2695.0,
        "incremental_labour_cost": 1155.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 30, "total_deliveries": 30,
        "receivable_amount": 6100.0, "days_outstanding": 12,
        "note": None,
    },
    {
        "client": "Choithrams", "location": "Sharjah",
        "delivery_frequency": "3x/week",
        "revenue": 21000.0, "service_cost_full": 14700.0,
        "ingredient_cost": 7350.0, "packaging_cost": 1050.0, "delivery_cost": 1470.0,
        "incremental_labour_cost": 630.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 12, "total_deliveries": 13,
        "receivable_amount": 15800.0, "days_outstanding": 48,
        "note": None,
    },
    {
        "client": "Zoom Supermarket", "location": "Dubai",
        "delivery_frequency": "3x/week",
        "revenue": 15200.0, "service_cost_full": 10800.0,
        "ingredient_cost": 5320.0, "packaging_cost": 760.0, "delivery_cost": 1064.0,
        "incremental_labour_cost": 456.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 11, "total_deliveries": 13,
        "receivable_amount": 9300.0, "days_outstanding": 67,
        "note": None,
    },
    {
        "client": "Kibsons", "location": "Dubai",
        "delivery_frequency": "Weekly",
        "revenue": 9800.0, "service_cost_full": 6900.0,
        "ingredient_cost": 3430.0, "packaging_cost": 490.0, "delivery_cost": 686.0,
        "incremental_labour_cost": 294.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 4, "total_deliveries": 4,
        "receivable_amount": 9800.0, "days_outstanding": 35,
        "note": None,
    },
    {
        "client": "Al Maya Supermarket", "location": "Ajman",
        "delivery_frequency": "Weekly",
        "revenue": 7400.0, "service_cost_full": 5300.0,
        "ingredient_cost": 2590.0, "packaging_cost": 370.0, "delivery_cost": 518.0,
        "incremental_labour_cost": 222.0, "overtime_premium": 0.0,
        "uses_idle_capacity": False,
        "on_time_count": 3, "total_deliveries": 4,
        "receivable_amount": 7400.0, "days_outstanding": 28,
        "note": None,
    },
    {
        "client": "Grand Hyper", "location": "Dubai",
        "delivery_frequency": "2x/week",
        "revenue": 12300.0, "service_cost_full": 13000.0,
        "ingredient_cost": 4200.0, "packaging_cost": 620.0, "delivery_cost": 980.0,
        "incremental_labour_cost": 0.0, "overtime_premium": 0.0,
        "uses_idle_capacity": True,
        "on_time_count": 7, "total_deliveries": 8,
        "receivable_amount": 12300.0, "days_outstanding": 95,
        "note": "Negative on full absorption (its allocated overhead share exceeds "
                "its revenue), but positive at the margin — its orders fill oven "
                "time that would otherwise sit idle, so incremental labour and "
                "overtime premium are both ~0. Dropping it on the Margin figure "
                "alone would leave that capacity cold while losing real contribution.",
    },
]

# ---- Recent deliveries feed ----
RECENT_DELIVERIES = [
    {"client": "Carrefour Hypermarket", "location": "Dubai", "time": "Today 08:15",
     "status": "On-time", "value": 1400.0},
    {"client": "Spinneys", "location": "Abu Dhabi", "time": "Today 07:40",
     "status": "On-time", "value": 1250.0},
    {"client": "Zoom Supermarket", "location": "Dubai", "time": "Yesterday 16:20",
     "status": "Late", "value": 620.0},
    {"client": "Choithrams", "location": "Sharjah", "time": "Yesterday 09:10",
     "status": "On-time", "value": 980.0},
    {"client": "Grand Hyper", "location": "Dubai", "time": "2 days ago 14:05",
     "status": "Late", "value": 540.0},
    {"client": "Kibsons", "location": "Dubai", "time": "3 days ago 11:30",
     "status": "On-time", "value": 710.0},
    {"client": "Al Maya Supermarket", "location": "Ajman", "time": "3 days ago 08:50",
     "status": "On-time", "value": 560.0},
    {"client": "Carrefour Hypermarket", "location": "Dubai", "time": "4 days ago 08:05",
     "status": "On-time", "value": 1380.0},
]
