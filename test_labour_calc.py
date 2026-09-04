"""
Tests for labour_calc.py — pytest.

Covers the six cases the brief (§A11) requires: a hand-verified worked
example, breaks_are_paid True vs False diverging correctly, department
rollup differing from a naive mean of percentages, zero productive hours
returning None, missing revenue returning None, and exact minutes-to-hours
conversion at boundary values.
"""

import pytest

from labour_calc import (
    minutes_to_hours,
    safe_divide,
    compute_shift_hours,
    utilisation_pct,
    true_efficiency_pct,
    performance_while_working_pct,
    cost_per_productive_hour,
    revenue_per_labour_dirham,
    revenue_per_day,
    compute_revenue_attributed,
    rollup_hours,
)


# ---------------------------------------------------------------------
# Minutes -> hours: exact at boundary values
# ---------------------------------------------------------------------

def test_minutes_to_hours_boundary_values():
    assert minutes_to_hours(0) == 0.0
    assert minutes_to_hours(60) == 1.0
    assert minutes_to_hours(30) == 0.5
    assert minutes_to_hours(35) == pytest.approx(0.5833333333333334)


# ---------------------------------------------------------------------
# A hand-verified worked example
# ---------------------------------------------------------------------

def test_worked_example_hand_verified():
    # 9h shift, 35-minute break (paid), 15 min changeover, 10 min downtime.
    sh = compute_shift_hours(
        paid_hours=9.0, break_minutes=35, changeover_minutes=15, downtime_minutes=10,
        breaks_are_paid=True,
    )
    assert sh["break_hours"] == pytest.approx(35 / 60)
    assert sh["changeover_hours"] == pytest.approx(15 / 60)
    assert sh["downtime_hours"] == pytest.approx(10 / 60)
    # available = 9.0 - 0 (paid break) - 0.25 - 0.16666... = 8.58333...
    assert sh["available_hours"] == pytest.approx(8.583333333333334)
    # no idle_waiting captured in Part A -> productive == available
    assert sh["productive_hours"] == pytest.approx(8.583333333333334)
    # bar_productive always subtracts break time, unlike productive_hours above
    assert sh["bar_productive_hours"] == pytest.approx(9.0 - 35 / 60 - 15 / 60 - 10 / 60)
    # deduction-bar segments always sum to paid_hours
    assert (sh["bar_productive_hours"] + sh["changeover_hours"] + sh["downtime_hours"]
            + sh["break_hours"] + sh["idle_waiting_hours"]) == pytest.approx(sh["paid_hours"])

    assert utilisation_pct(sh["available_hours"], sh["paid_hours"]) == pytest.approx(95.37037037037037)
    assert true_efficiency_pct(sh["productive_hours"], sh["paid_hours"]) == pytest.approx(95.37037037037037)
    assert performance_while_working_pct(sh["productive_hours"], sh["available_hours"]) == pytest.approx(100.0)

    # Labour economics on a clean 9-paid-hour, no-deduction day for round numbers.
    sh_clean = compute_shift_hours(
        paid_hours=9.0, break_minutes=0, changeover_minutes=0, downtime_minutes=0,
        breaks_are_paid=True,
    )
    assert sh_clean["productive_hours"] == pytest.approx(9.0)
    assert cost_per_productive_hour(total_salary_cost=180.0, productive_hours=sh_clean["productive_hours"]) == pytest.approx(20.0)
    revenue_attributed = compute_revenue_attributed(division_revenue=900.0, allocation_weight=1.0)
    assert revenue_attributed == pytest.approx(900.0)
    assert revenue_per_labour_dirham(revenue_attributed, total_salary_cost=180.0) == pytest.approx(5.0)
    assert revenue_per_day(revenue_attributed, days_worked=1) == pytest.approx(900.0)


# ---------------------------------------------------------------------
# breaks_are_paid True vs False produce correctly different results
# ---------------------------------------------------------------------

def test_breaks_are_paid_flips_efficiency_correctly():
    common = dict(paid_hours=9.0, break_minutes=60, changeover_minutes=0, downtime_minutes=0)

    paid = compute_shift_hours(**common, breaks_are_paid=True)
    unpaid = compute_shift_hours(**common, breaks_are_paid=False)

    assert paid["available_hours"] == pytest.approx(9.0)
    assert unpaid["available_hours"] == pytest.approx(8.0)

    assert true_efficiency_pct(paid["productive_hours"], paid["paid_hours"]) == pytest.approx(100.0)
    assert true_efficiency_pct(unpaid["productive_hours"], unpaid["paid_hours"]) == pytest.approx(88.88888888888889)


# ---------------------------------------------------------------------
# Department rollup differs from the naive mean of percentages
# ---------------------------------------------------------------------

def test_rollup_differs_from_naive_mean_of_percentages():
    # Employee A: 1 day, no deductions -> 100% true efficiency.
    a = compute_shift_hours(paid_hours=9.0, break_minutes=0, changeover_minutes=0, downtime_minutes=0, breaks_are_paid=True)
    # Employee B: 2 days (paid_hours pre-summed to 18), 2h total deductions -> 16/18 = 88.888...%
    b = compute_shift_hours(paid_hours=18.0, break_minutes=0, changeover_minutes=120, downtime_minutes=0, breaks_are_paid=True)

    a_pct = true_efficiency_pct(a["productive_hours"], a["paid_hours"])
    b_pct = true_efficiency_pct(b["productive_hours"], b["paid_hours"])
    naive_mean = (a_pct + b_pct) / 2

    rolled = rollup_hours([a, b])
    correct_pct = rolled["true_efficiency_pct"]

    assert naive_mean == pytest.approx(94.44444444444444)
    assert correct_pct == pytest.approx(92.5925925925926)
    assert naive_mean != pytest.approx(correct_pct)


# ---------------------------------------------------------------------
# Zero productive hours returns None, not inf/nan
# ---------------------------------------------------------------------

def test_zero_denominator_returns_none_not_inf_or_nan():
    assert safe_divide(100.0, 0.0) is None
    assert cost_per_productive_hour(total_salary_cost=500.0, productive_hours=0.0) is None
    assert true_efficiency_pct(productive_hours=0.0, paid_hours=0.0) is None
    assert utilisation_pct(available_hours=5.0, paid_hours=0.0) is None


# ---------------------------------------------------------------------
# Missing revenue returns None, not 0
# ---------------------------------------------------------------------

def test_missing_revenue_returns_none_not_zero():
    assert revenue_per_labour_dirham(None, total_salary_cost=180.0) is None
    assert revenue_per_day(None, days_worked=1) is None
    assert compute_revenue_attributed(division_revenue=None, allocation_weight=0.3) is None
    assert compute_revenue_attributed(division_revenue=1000.0, allocation_weight=None) is None
