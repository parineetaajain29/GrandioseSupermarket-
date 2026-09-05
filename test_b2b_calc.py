"""
Tests for b2b_calc.py — pytest.

Covers each pure function's guard behaviour (missing/zero -> None, never
inf/nan) plus a couple of hand-verified worked examples, and two
integration checks against the actual b2b_data.py dataset: that client
revenues sum to a specific known total (§B3 — "client revenues must sum
exactly to the summary revenue figure"), and that at least one account is
negative on full absorption but positive at the margin (§B2's Marginal
column requirement).
"""

import pytest

import b2b_data
from b2b_calc import (
    safe_divide,
    pct_change,
    pp_delta,
    net_margin_pct,
    otif_rate_pct,
    late_count,
    total_of,
    utilisation_pct,
    weighted_avg_days,
    marginal_contribution,
    account_margin,
    total_revenue,
    top_n_share_pct,
    normalize_shares_pct,
    count_overtime_orders,
    aging_buckets,
    total_outstanding,
    past_60_days_total,
)


# ---------------------------------------------------------------------
# Guards: zero/missing denominators and missing inputs return None
# ---------------------------------------------------------------------

def test_guards_return_none_not_inf_or_nan():
    assert safe_divide(100.0, 0.0) is None
    assert pct_change(100.0, 0.0) is None
    assert pct_change(100.0, None) is None
    assert pct_change(None, 100.0) is None
    assert pp_delta(None, 42.1) is None
    assert pp_delta(18.0, None) is None
    assert net_margin_pct(0.0, 100.0) is None
    assert net_margin_pct(None, 100.0) is None
    assert net_margin_pct(100.0, None) is None
    assert otif_rate_pct(5, 0) is None
    assert otif_rate_pct(None, 10) is None
    assert late_count(None, 5) is None
    assert weighted_avg_days([]) is None
    assert weighted_avg_days([(0.0, 30)]) is None
    assert marginal_contribution(None, 100, 10, 10, 0, 0) is None
    assert account_margin(None, 100.0) is None
    assert account_margin(100.0, None) is None
    assert top_n_share_pct([]) is None
    assert top_n_share_pct([0, 0, 0]) is None


# ---------------------------------------------------------------------
# Hand-verified worked examples
# ---------------------------------------------------------------------

def test_pct_change_and_pp_delta():
    assert pct_change(146200.0, 132000.0) == pytest.approx(10.757575757575758)
    assert pp_delta(18.2, 42.1) == pytest.approx(-23.9)


def test_net_margin_pct():
    assert net_margin_pct(42000.0, 27300.0) == pytest.approx(35.0)
    assert net_margin_pct(12300.0, 13000.0) == pytest.approx(-5.691056910569106)


def test_otif_rate_and_late_count():
    assert otif_rate_pct(96, 102) == pytest.approx(94.11764705882352)
    assert late_count(102, 96) == 6


def test_weighted_avg_days():
    # AED 8,200 at 19 days and AED 6,100 at 12 days.
    result = weighted_avg_days([(8200.0, 19), (6100.0, 12)])
    expected = (8200 * 19 + 6100 * 12) / (8200 + 6100)
    assert result == pytest.approx(expected)


def test_marginal_contribution_idle_vs_overtime():
    order = b2b_data.MARGINAL_EXAMPLE_ORDER
    idle = marginal_contribution(
        order["order_value"], order["ingredient_cost"], order["packaging_cost"],
        order["delivery_cost"], **b2b_data.MARGINAL_EXAMPLE_IDLE)
    overtime = marginal_contribution(
        order["order_value"], order["ingredient_cost"], order["packaging_cost"],
        order["delivery_cost"], **b2b_data.MARGINAL_EXAMPLE_OVERTIME)
    assert idle == pytest.approx(2650.0)
    assert overtime == pytest.approx(1600.0)
    assert idle > overtime  # idle capacity is strictly more profitable at the margin


def test_top_n_share_pct():
    assert top_n_share_pct([42000.0, 38500.0, 21000.0, 15200.0], n=2) == pytest.approx(
        (42000.0 + 38500.0) / (42000.0 + 38500.0 + 21000.0 + 15200.0) * 100.0)


def test_normalize_shares_pct_sums_to_100():
    shares = normalize_shares_pct([24.0, 8.4, 67.6])
    assert sum(shares) == pytest.approx(100.0)
    assert shares[0] == pytest.approx(24.0)


def test_total_of_skips_none():
    assert total_of([10.0, None, 5.0]) == pytest.approx(15.0)
    assert total_of([]) == 0


def test_utilisation_pct():
    assert utilisation_pct(24.0, 8.4) == pytest.approx(32.4)
    assert utilisation_pct(None, 8.4) is None


def test_count_overtime_orders():
    assert count_overtime_orders([False, False, True, False, True]) == 2
    assert count_overtime_orders([]) == 0


def test_aging_buckets_and_totals():
    rows = [(8200.0, 19), (15800.0, 48), (9300.0, 67), (12300.0, 95)]
    buckets = aging_buckets(rows)
    assert buckets["0-30"] == pytest.approx(8200.0)
    assert buckets["31-60"] == pytest.approx(15800.0)
    assert buckets["61-90"] == pytest.approx(9300.0)
    assert buckets["90+"] == pytest.approx(12300.0)
    assert total_outstanding(rows) == pytest.approx(45600.0)
    assert past_60_days_total(rows) == pytest.approx(9300.0 + 12300.0)


# ---------------------------------------------------------------------
# Integration checks against the actual b2b_data.py dataset
# ---------------------------------------------------------------------

def test_account_revenues_sum_to_known_total():
    # §B3 — client revenues must sum exactly to the summary revenue figure;
    # this pins the dataset's total so a silent edit gets caught.
    assert total_revenue(a["revenue"] for a in b2b_data.ACCOUNTS) == pytest.approx(146200.0)


def test_at_least_one_account_negative_margin_positive_marginal():
    # §B2 — the whole point of the Marginal column: at least one account
    # must be margin-negative on full absorption but marginal-positive
    # because it fills idle capacity.
    found = False
    for a in b2b_data.ACCOUNTS:
        margin = account_margin(a["revenue"], a["service_cost_full"])
        marginal = marginal_contribution(
            a["revenue"], a["ingredient_cost"], a["packaging_cost"], a["delivery_cost"],
            a["incremental_labour_cost"], a["overtime_premium"])
        if margin < 0 and marginal > 0:
            found = True
            assert a["uses_idle_capacity"] is True
            assert a["note"]
    assert found, "expected at least one account negative on margin but positive at the margin"


def test_avg_collection_days_vs_customer_terms():
    # The KPI strip compares collection days against what customers are
    # invoiced on (receivables), not what Grandiose owes its own suppliers
    # (payables) — pins the exact badge text a live-review fix required.
    rows = [(a["receivable_amount"], a["days_outstanding"]) for a in b2b_data.ACCOUNTS]
    avg_days = weighted_avg_days(rows)
    delta = pp_delta(avg_days, b2b_data.CUSTOMER_PAYMENT_TERMS_DAYS)
    assert avg_days == pytest.approx(48.31930333817126)
    assert delta == pytest.approx(18.31930333817126)
