"""Product / SKU performance data for the Bakery Product Performance page.

Everything that page renders is derived from `load_products()`. Ranks,
contribution shares, division roll-ups and the KPI strip are all computed
here rather than stored, so swapping the mock table below for a database
or API call is the only change needed to go live — no page code moves.

To connect real data, replace the body of `load_products()` with the
fetch and keep returning the same list-of-dicts shape. The keys the page
relies on are:

    sku, product, division, units, sales_aed
    (+ collection, availability on seasonal lines)

`rank` and `contribution_pct` are added by this module; do not expect the
upstream source to supply them.
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

# The seasonal line rotates. Only these three values change when the
# collection is swapped out (e.g. to a winter range).
SEASONAL_COLLECTION = "Mango Summer Collection"
SEASONAL_AVAILABILITY = "May – Sep 2024"

# (sku, product, division, units sold, sales in AED)
_RAW_PRODUCTS = [
    # --- Baklava (12 SKUs) ---
    ("BKV-001", "Pistachio Baklava Premium", "Baklava", 8000, 190000),
    ("BKV-002", "Walnut Baklava Standard",   "Baklava", 6000, 132000),
    ("BKV-003", "Cashew Baklava",            "Baklava", 4800,  96000),
    ("BKV-004", "Mixed Baklava Box",         "Baklava", 2600,  78000),
    ("BKV-005", "Almond Baklava",            "Baklava", 3400,  62000),
    ("BKV-006", "Chocolate Baklava",         "Baklava", 3200,  58000),
    ("BKV-007", "Baklava Fingers",           "Baklava", 2900,  44000),
    ("BKV-008", "Rose Baklava",              "Baklava", 1800,  32000),
    ("BKV-009", "Pistachio Burma",           "Baklava", 1400,  28000),
    ("BKV-010", "Baklava Bites",             "Baklava", 2000,  24000),
    ("BKV-011", "Date Baklava",              "Baklava", 1300,  20000),
    ("BKV-012", "Baklava Gift Tin",          "Baklava",  640,  16000),

    # --- French Bakery (18 SKUs) ---
    ("FRB-001", "Butter Croissant",     "French Bakery", 7000, 86000),
    ("FRB-002", "Classic Baguette",     "French Bakery", 7000, 62000),
    ("FRB-003", "Pain au Chocolat",     "French Bakery", 6200, 62000),
    ("FRB-004", "Sourdough Loaf",       "French Bakery", 5200, 58000),
    ("FRB-005", "Almond Croissant",     "French Bakery", 4400, 58000),
    ("FRB-006", "Brioche Bun",          "French Bakery", 3800, 46000),
    ("FRB-007", "Multigrain Loaf",      "French Bakery", 2400, 42000),
    ("FRB-008", "Ciabatta",             "French Bakery", 2600, 38000),
    ("FRB-009", "Country Sourdough",    "French Bakery", 2600, 36000),
    ("FRB-010", "Focaccia",             "French Bakery", 2900, 32000),
    ("FRB-011", "Rye Loaf",             "French Bakery", 2200, 28000),
    ("FRB-012", "Baguette Tradition",   "French Bakery", 2200, 26000),
    ("FRB-013", "Croissant Mini 6pk",   "French Bakery", 2000, 24000),
    ("FRB-014", "Pain de Campagne",     "French Bakery", 2800, 24000),
    ("FRB-015", "Petit Pain",           "French Bakery", 2200, 22000),
    ("FRB-016", "Olive Fougasse",       "French Bakery", 1500, 18000),
    ("FRB-017", "Seeded Roll",          "French Bakery", 1600, 16000),
    ("FRB-018", "Milk Bread Loaf",      "French Bakery", 1400, 12000),

    # --- Arabic Bread (9 SKUs) ---
    ("ARB-001", "White Arabic Bread Large",  "Arabic Bread", 10000, 78000),
    ("ARB-002", "Brown Arabic Bread",        "Arabic Bread",  7500, 62000),
    ("ARB-003", "Pita Bread 6pk",            "Arabic Bread",  5000, 54000),
    ("ARB-004", "Khubz Rugag",               "Arabic Bread",  3000, 34000),
    ("ARB-005", "Saj Bread",                 "Arabic Bread",  3200, 30000),
    ("ARB-006", "Tannour Bread",             "Arabic Bread",  5600, 28000),
    ("ARB-007", "Whole Wheat Arabic Bread",  "Arabic Bread",  2800, 22000),
    ("ARB-008", "Arabic Bread Family Pack",  "Arabic Bread",  2600, 20000),
    ("ARB-009", "Mini Pita 12pk",            "Arabic Bread",  1800, 12000),

    # --- Viennoiserie (10 SKUs) ---
    ("VNS-001", "Cheese Danish",          "Viennoiserie", 6500, 78000),
    ("VNS-002", "Cinnamon Roll",          "Viennoiserie", 5700, 68000),
    ("VNS-003", "Apple Turnover",         "Viennoiserie", 4500, 54000),
    ("VNS-004", "Chocolate Danish",       "Viennoiserie", 4300, 52000),
    ("VNS-005", "Raisin Swirl",           "Viennoiserie", 3700, 44000),
    ("VNS-006", "Custard Danish",         "Viennoiserie", 3300, 40000),
    ("VNS-007", "Almond Bear Claw",       "Viennoiserie", 2400, 34000),
    ("VNS-008", "Berry Danish",           "Viennoiserie", 2200, 26000),
    ("VNS-009", "Pecan Plait",            "Viennoiserie", 1400, 20000),
    ("VNS-010", "Vanilla Croissant Roll", "Viennoiserie", 1200, 14000),

    # --- Tahina (7 SKUs) ---
    ("THN-001", "Premium Tahina 500g",   "Tahina", 6200, 62000),
    ("THN-002", "Classic Tahina 250g",   "Tahina", 4800, 48000),
    ("THN-003", "Organic Tahina 400g",   "Tahina", 3200, 38000),
    ("THN-004", "Tahina Halva Swirl",    "Tahina", 2100, 28000),
    ("THN-005", "Tahina 1kg Catering",   "Tahina", 1300, 24000),
    ("THN-006", "Date Tahina Spread",    "Tahina",  900, 12000),
    ("THN-007", "Tahina Sachets 20pk",   "Tahina",  700,  8000),

    # --- Seasonal Collection — Mango Summer (8 SKUs) ---
    ("SEA-M01", "Mango Cheesecake",      SEASONAL_DIVISION, 4600, 92000),
    ("SEA-M02", "Mango Tiramisu",        SEASONAL_DIVISION, 3400, 68000),
    ("SEA-M03", "Mango Éclair",          SEASONAL_DIVISION, 4500, 54000),
    ("SEA-M04", "Mango Croissant",       SEASONAL_DIVISION, 4000, 48000),
    ("SEA-M05", "Mango Cream Tart",      SEASONAL_DIVISION, 2500, 40000),
    ("SEA-M06", "Mango Mousse Cup",      SEASONAL_DIVISION, 2800, 34000),
    ("SEA-M07", "Mango Pistachio Cake",  SEASONAL_DIVISION, 1400, 28000),
    ("SEA-M08", "Mango Danish",          SEASONAL_DIVISION, 1600, 16000),
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
    }


def format_aed(value):
    """AED figures shortened for KPI display: 2_840_000 -> 'AED 2.84M'."""
    if value >= 1_000_000:
        return f"AED {value / 1_000_000:.2f}M"
    if value >= 1_000:
        return f"AED {value / 1_000:.0f}k"
    return f"AED {value:,.0f}"


def format_units(value):
    """Unit counts shortened for KPI display: 217_740 -> '217.7K'."""
    if value >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,}"
