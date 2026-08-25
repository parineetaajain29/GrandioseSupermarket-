"""Product / SKU performance data for the Bakery Product Performance page.

============================ DEMO DATA ============================
Everything in this module is realistic mock data for the prototype.
It is deliberately kept apart from the page code so that swapping in
real figures is a change to THIS FILE ONLY:

  * `load_products()`      -> replace the body with your fetch. Keep the
                              returned shape (see below) and the whole
                              page keeps working unchanged.
  * `_PRIOR_PERIOD`        -> replace with the real previous period's
                              totals. The KPI deltas are computed from
                              it, never written by hand.
  * `SEASONAL_*` constants -> three values swap the seasonal line over.

Shape `load_products()` must return, one dict per SKU:

    sku, product, division, units, sales_aed
    (+ collection, availability on seasonal lines)

`rank` and `contribution_pct` are derived here, so an upstream source
does not need to supply them.
===================================================================
"""

# Division order drives cluster order in the visualisation and accordion
# order in the SKU table, so keep the seasonal line last.
DIVISIONS = [
    "Baklava",
    "French Bakery",
    "Arabic Bread",
    "Viennoiserie",
    "Tahina",
    "Seasonal Collection",
]

SEASONAL_DIVISION = "Seasonal Collection"

# The seasonal line rotates. Only these values change when the collection
# is swapped out (e.g. to a winter range).
SEASONAL_COLLECTION = "Mango Summer Collection"
SEASONAL_AVAILABILITY = "May – Sep 2024"

# Previous period totals, used only to compute the KPI trend lines.
# DEMO VALUES — replace with the real prior month when history is wired up.
_PRIOR_PERIOD = {
    "total_sales": 2_620_000,
    "total_units": 210_400,
    "active_skus": 104,
}

# (sku, product, division, units sold, sales in AED)
# A full catalogue rather than a handful of headline lines, so the
# visualisation reads as a real product range. Values follow a realistic
# long tail: a few hero SKUs, then a broad spread of smaller lines.
_RAW_PRODUCTS = [
    # ------------------------- Baklava (22) -------------------------
    ("BKV-001", "Pistachio Baklava Premium",  "Baklava", 6300, 150000),
    ("BKV-002", "Walnut Baklava Standard",    "Baklava", 4700, 104000),
    ("BKV-003", "Cashew Baklava",             "Baklava", 3900,  78000),
    ("BKV-004", "Mixed Baklava Box",          "Baklava", 2100,  62000),
    ("BKV-005", "Almond Baklava",             "Baklava", 2900,  54000),
    ("BKV-006", "Chocolate Baklava",          "Baklava", 2700,  48000),
    ("BKV-007", "Baklava Fingers",            "Baklava", 2600,  40000),
    ("BKV-008", "Pistachio Burma",            "Baklava", 1700,  34000),
    ("BKV-009", "Rose Baklava",               "Baklava", 1700,  30000),
    ("BKV-010", "Baklava Bites",              "Baklava", 2200,  26000),
    ("BKV-011", "Date Baklava",               "Baklava", 1400,  22000),
    ("BKV-012", "Kunafa Baklava Roll",        "Baklava", 1150,  20000),
    ("BKV-013", "Baklava Gift Tin",           "Baklava",  720,  18000),
    ("BKV-014", "Hazelnut Baklava",           "Baklava",  900,  16000),
    ("BKV-015", "Baklava Assorted 500g",      "Baklava",  620,  14000),
    ("BKV-016", "Semolina Baklava",           "Baklava",  800,  12000),
    ("BKV-017", "Honey Baklava Squares",      "Baklava",  750,  11000),
    ("BKV-018", "Pistachio Baklava Mini",     "Baklava",  900,  10000),
    ("BKV-019", "Walnut Baklava Tray",        "Baklava",  420,   9000),
    ("BKV-020", "Baklava Sampler 6pc",        "Baklava",  560,   8000),
    ("BKV-021", "Saffron Baklava",            "Baklava",  380,   7000),
    ("BKV-022", "Baklava Seasonal Tin",       "Baklava",  300,   7000),

    # ---------------------- French Bakery (26) ----------------------
    ("FRB-001", "Butter Croissant",     "French Bakery", 6200, 82000),
    ("FRB-002", "Classic Baguette",     "French Bakery", 6300, 56000),
    ("FRB-003", "Pain au Chocolat",     "French Bakery", 5400, 54000),
    ("FRB-004", "Sourdough Loaf",       "French Bakery", 4500, 50000),
    ("FRB-005", "Almond Croissant",     "French Bakery", 3500, 46000),
    ("FRB-006", "Brioche Bun",          "French Bakery", 3300, 40000),
    ("FRB-007", "Multigrain Loaf",      "French Bakery", 2100, 36000),
    ("FRB-008", "Ciabatta",             "French Bakery", 2300, 33000),
    ("FRB-009", "Country Sourdough",    "French Bakery", 2200, 31000),
    ("FRB-010", "Focaccia",             "French Bakery", 2500, 28000),
    ("FRB-011", "Rye Loaf",             "French Bakery", 2000, 26000),
    ("FRB-012", "Baguette Tradition",   "French Bakery", 2000, 24000),
    ("FRB-013", "Croissant Mini 6pk",   "French Bakery", 1800, 22000),
    ("FRB-014", "Pain de Campagne",     "French Bakery", 2400, 21000),
    ("FRB-015", "Petit Pain",           "French Bakery", 1900, 19000),
    ("FRB-016", "Olive Fougasse",       "French Bakery", 1400, 17000),
    ("FRB-017", "Seeded Roll",          "French Bakery", 1500, 15000),
    ("FRB-018", "Milk Bread Loaf",      "French Bakery", 1600, 14000),
    ("FRB-019", "Pain aux Raisins",     "French Bakery", 1100, 13000),
    ("FRB-020", "Walnut Loaf",          "French Bakery",  900, 12000),
    ("FRB-021", "Baguette Sesame",      "French Bakery", 1100, 11000),
    ("FRB-022", "Ficelle",              "French Bakery", 1000, 10000),
    ("FRB-023", "Pain Complet",         "French Bakery",  800,  9000),
    ("FRB-024", "Épi Baguette",         "French Bakery",  700,  8000),
    ("FRB-025", "Brioche Tressée",      "French Bakery",  450,  7000),
    ("FRB-026", "Pain Viennois",        "French Bakery",  600,  6000),

    # ----------------------- Arabic Bread (16) ----------------------
    ("ARB-001", "White Arabic Bread Large",   "Arabic Bread", 8200, 62000),
    ("ARB-002", "Brown Arabic Bread",         "Arabic Bread", 6000, 48000),
    ("ARB-003", "Pita Bread 6pk",             "Arabic Bread", 4000, 42000),
    ("ARB-004", "Khubz Rugag",                "Arabic Bread", 2700, 30000),
    ("ARB-005", "Saj Bread",                  "Arabic Bread", 2800, 26000),
    ("ARB-006", "Tannour Bread",              "Arabic Bread", 4800, 24000),
    ("ARB-007", "Whole Wheat Arabic Bread",   "Arabic Bread", 2500, 20000),
    ("ARB-008", "Arabic Bread Family Pack",   "Arabic Bread", 2300, 18000),
    ("ARB-009", "Mini Pita 12pk",             "Arabic Bread", 2100, 14000),
    ("ARB-010", "Zaatar Manakish",            "Arabic Bread", 1600, 13000),
    ("ARB-011", "Cheese Manakish",            "Arabic Bread", 1400, 12000),
    ("ARB-012", "Khubz Arabi Wholemeal",      "Arabic Bread", 1300, 10000),
    ("ARB-013", "Markook Bread",              "Arabic Bread", 1200,  9000),
    ("ARB-014", "Pita Pocket Large",          "Arabic Bread", 1000,  8000),
    ("ARB-015", "Sesame Kaak",                "Arabic Bread",  500,  2000),
    ("ARB-016", "Arabic Flatbread Mini",      "Arabic Bread",  400,  2000),

    # ----------------------- Viennoiserie (20) ----------------------
    ("VNS-001", "Cheese Danish",          "Viennoiserie", 5200, 53000),
    ("VNS-002", "Cinnamon Roll",          "Viennoiserie", 4500, 54000),
    ("VNS-003", "Apple Turnover",         "Viennoiserie", 3700, 44000),
    ("VNS-004", "Chocolate Danish",       "Viennoiserie", 3500, 42000),
    ("VNS-005", "Raisin Swirl",           "Viennoiserie", 3000, 36000),
    ("VNS-006", "Custard Danish",         "Viennoiserie", 2700, 32000),
    ("VNS-007", "Almond Bear Claw",       "Viennoiserie", 2000, 28000),
    ("VNS-008", "Berry Danish",           "Viennoiserie", 2000, 24000),
    ("VNS-009", "Pecan Plait",            "Viennoiserie", 1400, 20000),
    ("VNS-010", "Vanilla Croissant Roll", "Viennoiserie", 1300, 16000),
    ("VNS-011", "Apricot Danish",         "Viennoiserie", 1150, 14000),
    ("VNS-012", "Chocolate Twist",        "Viennoiserie", 1000, 12000),
    ("VNS-013", "Maple Pecan Danish",     "Viennoiserie",  800, 11000),
    ("VNS-014", "Hazelnut Escargot",      "Viennoiserie",  750, 10000),
    ("VNS-015", "Blueberry Danish",       "Viennoiserie",  720,  9000),
    ("VNS-016", "Pistachio Roll",         "Viennoiserie",  500,  7000),
    ("VNS-017", "Cream Cheese Braid",     "Viennoiserie",  420,  6000),
    ("VNS-018", "Lemon Danish",           "Viennoiserie",  400,  5000),
    ("VNS-019", "Cardamom Bun",           "Viennoiserie",  320,  4000),
    ("VNS-020", "Almond Croissant Roll",  "Viennoiserie",  240,  3000),

    # -------------------------- Tahina (12) -------------------------
    ("THN-001", "Premium Tahina 500g",      "Tahina", 5200, 52000),
    ("THN-002", "Classic Tahina 250g",      "Tahina", 4000, 40000),
    ("THN-003", "Organic Tahina 400g",      "Tahina", 2700, 32000),
    ("THN-004", "Tahina Halva Swirl",       "Tahina", 1800, 24000),
    ("THN-005", "Tahina 1kg Catering",      "Tahina", 1100, 20000),
    ("THN-006", "Date Tahina Spread",       "Tahina", 1000, 14000),
    ("THN-007", "Chocolate Tahina Spread",  "Tahina",  850, 12000),
    ("THN-008", "Tahina Sachets 20pk",      "Tahina",  700,  8000),
    ("THN-009", "Tahina Halva Bar",         "Tahina",  600,  6000),
    ("THN-010", "Roasted Tahina 300g",      "Tahina",  420,  5000),
    ("THN-011", "Tahina Gift Pack",         "Tahina",  250,  4000),
    ("THN-012", "Tahina Dressing 250ml",    "Tahina",  300,  3000),

    # ------------- Seasonal Collection — Mango Summer (14) ----------
    ("SEA-M01", "Mango Cheesecake",        SEASONAL_DIVISION, 3600, 72000),
    ("SEA-M02", "Mango Tiramisu",          SEASONAL_DIVISION, 2700, 54000),
    ("SEA-M03", "Mango Éclair",            SEASONAL_DIVISION, 3700, 44000),
    ("SEA-M04", "Mango Croissant",         SEASONAL_DIVISION, 3300, 40000),
    ("SEA-M05", "Mango Cream Tart",        SEASONAL_DIVISION, 2100, 34000),
    ("SEA-M06", "Mango Mousse Cup",        SEASONAL_DIVISION, 2500, 30000),
    ("SEA-M07", "Mango Pistachio Cake",    SEASONAL_DIVISION, 1300, 26000),
    ("SEA-M08", "Mango Danish",            SEASONAL_DIVISION, 2000, 20000),
    ("SEA-M09", "Mango Roll Cake",         SEASONAL_DIVISION, 1100, 16000),
    ("SEA-M10", "Mango Macaron Box",       SEASONAL_DIVISION,  700, 14000),
    ("SEA-M11", "Mango Panna Cotta",       SEASONAL_DIVISION,  900, 12000),
    ("SEA-M12", "Mango Puff Pastry",       SEASONAL_DIVISION,  800,  8000),
    ("SEA-M13", "Mango Sticky Rice Tart",  SEASONAL_DIVISION,  400,  6000),
    ("SEA-M14", "Mango Yoghurt Parfait",   SEASONAL_DIVISION,  350,  4000),
]


def load_products():
    """Return every SKU as a dict, with rank and contribution share added.

    Replace the `_RAW_PRODUCTS` comprehension with a real fetch to go live;
    the derived fields below are recomputed from whatever comes back.
    """
    products = [
        {
            "sku": sku,
            "product": product,
            "division": division,
            "units": units,
            "sales_aed": sales,
        }
        for sku, product, division, units, sales in _RAW_PRODUCTS
    ]

    total_sales = sum(p["sales_aed"] for p in products) or 1

    # Rank is by sales value across the whole catalogue, 1 = best seller.
    for rank, p in enumerate(sorted(products, key=lambda p: -p["sales_aed"]), start=1):
        p["rank"] = rank

    for p in products:
        p["contribution_pct"] = round(p["sales_aed"] / total_sales * 100, 1)
        if p["division"] == SEASONAL_DIVISION:
            p["collection"] = SEASONAL_COLLECTION
            p["availability"] = SEASONAL_AVAILABILITY

    return products


def division_summary(products):
    """Per-division totals, in DIVISIONS order. Divisions with no SKUs are skipped."""
    summary = []
    for division in DIVISIONS:
        items = [p for p in products if p["division"] == division]
        if not items:
            continue
        summary.append({
            "division": division,
            "sku_count": len(items),
            "units": sum(p["units"] for p in items),
            "sales_aed": sum(p["sales_aed"] for p in items),
        })
    return summary


def kpis(products):
    """The five headline figures, all derived — nothing hard-coded."""
    if not products:
        return {"total_sales": 0, "total_units": 0, "active_skus": 0,
                "top_division": "—", "top_product": "—"}

    by_division = division_summary(products)
    top_division = max(by_division, key=lambda d: d["sales_aed"])
    top_product = max(products, key=lambda p: p["sales_aed"])

    return {
        "total_sales": sum(p["sales_aed"] for p in products),
        "total_units": sum(p["units"] for p in products),
        "active_skus": len(products),
        "top_division": top_division["division"],
        "top_product": top_product["product"],
        "top_division_sales": top_division["sales_aed"],
        "top_product_sales": top_product["sales_aed"],
    }


def _pct_change(current, previous):
    if not previous:
        return None
    return (current - previous) / previous * 100


def kpi_cards(products):
    """The KPI strip as data: label, value, sub-line and trend direction.

    Trend lines for the three numeric KPIs are computed against
    `_PRIOR_PERIOD`; the two categorical KPIs get a derived context line
    instead of a fabricated trend, since "last month's top product" is not
    a percentage change. `trend` is one of "up" / "down" / None and is left
    for the page to colour.
    """
    k = kpis(products)
    total = k["total_sales"] or 1

    sales_delta = _pct_change(k["total_sales"], _PRIOR_PERIOD["total_sales"])
    units_delta = _pct_change(k["total_units"], _PRIOR_PERIOD["total_units"])
    sku_delta = k["active_skus"] - _PRIOR_PERIOD["active_skus"]

    def signed(value):
        return f"{value:+.1f}% vs last month"

    return [
        {
            "label": "Total sales",
            "value": format_aed(k["total_sales"]),
            "sub": signed(sales_delta) if sales_delta is not None else "",
            "trend": "up" if (sales_delta or 0) >= 0 else "down",
        },
        {
            "label": "Units sold",
            "value": format_units(k["total_units"]),
            "sub": signed(units_delta) if units_delta is not None else "",
            "trend": "up" if (units_delta or 0) >= 0 else "down",
        },
        {
            "label": "Active SKUs",
            "value": f"{k['active_skus']}",
            "sub": (f"{sku_delta:+d} vs last month" if sku_delta
                    else "unchanged vs last month"),
            "trend": "up" if sku_delta > 0 else ("down" if sku_delta < 0 else None),
        },
        {
            "label": "Top division",
            "value": k["top_division"],
            "sub": (f"{format_aed(k['top_division_sales'])} · "
                    f"{k['top_division_sales'] / total * 100:.0f}% of sales"),
            "trend": None,
            "accent": True,
        },
        {
            "label": "Top product",
            "value": k["top_product"],
            "sub": (f"Rank #1 · {k['top_product_sales'] / total * 100:.1f}% of sales"),
            "trend": None,
        },
    ]


def format_aed(value):
    """AED figures shortened for KPI display: 2_840_000 -> 'AED 2.84M'."""
    if value >= 1_000_000:
        return f"AED {value / 1_000_000:.2f}M"
    if value >= 1_000:
        return f"AED {value / 1_000:.0f}k"
    return f"AED {value:,.0f}"


def format_units(value):
    """Unit counts shortened for KPI display: 156_400 -> '156.4K'."""
    if value >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,}"
