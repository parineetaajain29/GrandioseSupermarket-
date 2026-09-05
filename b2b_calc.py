"""
Pure calculations for the B2B Client Performance page (§B of the brief).

No Streamlit calls, no I/O — every function takes plain numbers or plain
lists/tuples and returns plain numbers (or `None` on an undefined ratio),
so the whole module is unit-testable without spinning up the app and
without any coupling to b2b_data.py's dict shape. Full float precision is
carried through every intermediate; round only at the display layer.
"""


def safe_divide(numerator, denominator, multiplier=1.0):
    """Guarded ratio: `None` (never inf/nan) when the denominator is zero or
    either input is missing."""
    if numerator is None or denominator is None or denominator == 0:
        return None
    return (numerator / denominator) * multiplier


def pct_change(current, previous):
    """Period-over-period % change (e.g. MoM). `None` if the prior period
    is missing or zero."""
    if current is None or previous is None or previous == 0:
        return None
    return (current - previous) / previous * 100.0


def pp_delta(current, benchmark):
    """Percentage-point delta vs. a benchmark rate (e.g. B2B margin vs.
    retail margin). `None` if either side is missing."""
    if current is None or benchmark is None:
        return None
    return current - benchmark


def net_margin_pct(revenue, cost):
    """(revenue - cost) / revenue * 100 — used for both the company-wide
    net-margin KPI and any per-account margin %. `None` if revenue or cost
    is missing, or revenue is zero."""
    if revenue is None or cost is None:
        return None
    return safe_divide(revenue - cost, revenue, 100.0)


def otif_rate_pct(on_time_count, total_count):
    return safe_divide(on_time_count, total_count, 100.0)


def late_count(total_count, on_time_count):
    if total_count is None or on_time_count is None:
        return None
    return total_count - on_time_count


def weighted_avg_days(amounts_and_days):
    """Weighted mean of days-outstanding, weighted by amount. `amounts_and_days`
    is an iterable of (amount, days) pairs. `None` if the total amount is 0
    or the iterable is empty."""
    amounts_and_days = list(amounts_and_days)
    total_amount = sum(a for a, _ in amounts_and_days)
    if not amounts_and_days or total_amount == 0:
        return None
    return sum(a * d for a, d in amounts_and_days) / total_amount


def marginal_contribution(order_value, ingredient_cost, packaging_cost, delivery_cost,
                            incremental_labour_cost, overtime_premium):
    """§B2's marginal-contribution formula, verbatim. `None` if the order
    value itself is missing."""
    if order_value is None:
        return None
    return (order_value - ingredient_cost - packaging_cost - delivery_cost
            - incremental_labour_cost - overtime_premium)


def account_margin(revenue, service_cost_full):
    """Full-absorption margin (AED) — revenue less the fully-loaded cost to
    serve (ingredients, packaging, delivery, labour, allocated overhead)."""
    if revenue is None or service_cost_full is None:
        return None
    return revenue - service_cost_full


def total_revenue(revenues):
    return sum(r for r in revenues if r is not None)


def total_of(values):
    """Generic sum, skipping missing values — used for totals other than
    revenue (service cost, on-time counts, delivery counts)."""
    return sum(v for v in values if v is not None)


def utilisation_pct(retail_pct, b2b_pct):
    """Combined oven utilisation — retail + B2B share of total capacity."""
    if retail_pct is None or b2b_pct is None:
        return None
    return retail_pct + b2b_pct


def top_n_share_pct(revenues, n=2):
    """Concentration: the top-n values' share of the total. `None` if the
    list is empty or the total is 0."""
    revenues = list(revenues)
    total = sum(revenues)
    if not revenues or total == 0:
        return None
    top_n_total = sum(sorted(revenues, reverse=True)[:n])
    return top_n_total / total * 100.0


def normalize_shares_pct(values):
    """Renormalize a set of raw quantities to percentages summing to
    exactly 100 — guards against rounding drift in source data. Returns a
    tuple of zeros if the total is 0."""
    values = list(values)
    total = sum(values)
    if total == 0:
        return tuple(0.0 for _ in values)
    return tuple(v / total * 100.0 for v in values)


def count_overtime_orders(forces_overtime_flags):
    return sum(1 for flag in forces_overtime_flags if flag)


_AGING_BUCKET_LABELS = ("0-30", "31-60", "61-90", "90+")


def aging_buckets(amounts_and_days):
    """Buckets (amount, days_outstanding) pairs into 0-30/31-60/61-90/90+.
    Returns an ordered dict-like mapping of bucket label -> total amount."""
    buckets = {label: 0.0 for label in _AGING_BUCKET_LABELS}
    for amount, days in amounts_and_days:
        if days <= 30:
            buckets["0-30"] += amount
        elif days <= 60:
            buckets["31-60"] += amount
        elif days <= 90:
            buckets["61-90"] += amount
        else:
            buckets["90+"] += amount
    return buckets


def total_outstanding(amounts_and_days):
    return sum(a for a, _ in amounts_and_days)


def past_60_days_total(amounts_and_days):
    """Amount in the 61-90 and 90+ buckets combined."""
    return sum(a for a, d in amounts_and_days if d > 60)
