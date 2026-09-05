"""
Pure labour-efficiency calculations for the Employee Portal.

No Streamlit calls, no I/O — every function takes plain numbers and returns
plain numbers (or `None` on an undefined ratio), which is what makes the
whole chain unit-testable without spinning up the app. Ports the same
formula chain the brief specifies in §A4-§A8; see
streamlit-employee-portal-and-b2b-brief.md for the source spec and
labour_config.py for the assumptions it reads.

Full float precision is carried through every intermediate. Round only at
the display layer (in app.py) — round-then-feed-forward compounds error
across a rollup of many shifts.

These functions are pure — they take assumptions (e.g. `breaks_are_paid`)
as plain arguments rather than reading LABOUR_CONFIG directly, so callers
(app.py) own reading the config and testing can pass either flag value
without monkeypatching anything.
"""


def minutes_to_hours(minutes):
    """Exact minutes -> hours conversion."""
    return minutes / 60.0


def safe_divide(numerator, denominator, multiplier=1.0):
    """Guarded ratio: `None` (never inf/nan) when the denominator is zero or
    either input is missing (e.g. unattributed revenue)."""
    if numerator is None or denominator is None or denominator == 0:
        return None
    return (numerator / denominator) * multiplier


def compute_shift_hours(paid_hours, break_minutes, changeover_minutes, downtime_minutes,
                          breaks_are_paid, idle_waiting_hours=0.0):
    """
    The §A4 hours chain for one shift (or any group of shifts whose paid
    hours have already been summed).

    `idle_waiting_hours` is a separate bucket from `downtime_minutes`
    (machine down / oven not ready / waiting on materials), reserved for the
    formula's general chain. Part A's Daily Log only captures break/
    changeover/downtime, so this defaults to 0 and the "idle" segment of the
    §A10 deduction bar renders zero-width until a field for it exists.

    Returns a dict: paid_hours, break_hours, changeover_hours, downtime_hours,
    idle_waiting_hours, available_hours, productive_hours.
    """
    break_hours = minutes_to_hours(break_minutes)
    changeover_hours = minutes_to_hours(changeover_minutes)
    downtime_hours = minutes_to_hours(downtime_minutes)

    available_hours = (
        paid_hours
        - (0.0 if breaks_are_paid else break_hours)
        - changeover_hours
        - downtime_hours
    )
    productive_hours = available_hours - idle_waiting_hours

    # Unconditional physical decomposition of the paid shift, for the §A10.2
    # deduction bar only — break time is real elapsed time regardless of pay
    # treatment, so (unlike `productive_hours` above) this always subtracts
    # it. This is what keeps productive + changeover + downtime + breaks +
    # idle summing to paid_hours even when breaks_are_paid is True.
    bar_productive_hours = paid_hours - break_hours - changeover_hours - downtime_hours - idle_waiting_hours

    return {
        "paid_hours": paid_hours,
        "break_hours": break_hours,
        "changeover_hours": changeover_hours,
        "downtime_hours": downtime_hours,
        "idle_waiting_hours": idle_waiting_hours,
        "available_hours": available_hours,
        "productive_hours": productive_hours,
        "bar_productive_hours": bar_productive_hours,
    }


def utilisation_pct(available_hours, paid_hours):
    return safe_divide(available_hours, paid_hours, 100.0)


def true_efficiency_pct(productive_hours, paid_hours):
    """Management's number — the cost view. Downtime reduces it, correctly,
    because Grandiose paid for that time."""
    return safe_divide(productive_hours, paid_hours, 100.0)


def performance_while_working_pct(productive_hours, available_hours):
    """The employee's number. Unaffected by equipment failure or changeover,
    so honestly reporting downtime never lowers it."""
    return safe_divide(productive_hours, available_hours, 100.0)


def quality_yield_pct(units_produced, units_wasted, units_failed_qc):
    """
    Share of ATTEMPTED output that was both kept (not wasted) and passed QC
    — the "Quality" factor in an OEE-style Availability x Performance x
    Quality decomposition. `units_produced` is already "good units" (not
    wasted) per the Daily Log's own field label, and `units_failed_qc` is a
    subset of it, so:

        attempted  = units_produced + units_wasted
        good       = units_produced - units_failed_qc
        yield      = good / attempted

    `None` if there was no attempted output at all.
    """
    total_attempted = units_produced + units_wasted
    good_units = max(units_produced - units_failed_qc, 0)
    return safe_divide(good_units, total_attempted, 100.0)


def quality_adjusted_performance_pct(performance_while_working_pct, quality_yield_pct):
    """
    Performance x Quality — folds wastage/QC failures into the employee-
    facing headline, so a shift with clean hours but heavy wastage no
    longer reads as 100%. Deliberately NOT applied to `true_efficiency_pct`
    (management's number): that one is a pure cost-of-time view by design
    (§A5 — downtime reduces it because Grandiose paid for that time), and
    quality/wastage is already tracked as its own KPI for managers.

    `None` if either factor is undefined.
    """
    if performance_while_working_pct is None or quality_yield_pct is None:
        return None
    return performance_while_working_pct * quality_yield_pct / 100.0


def cost_per_productive_hour(total_salary_cost, productive_hours):
    return safe_divide(total_salary_cost, productive_hours)


def revenue_per_labour_dirham(revenue_attributed, total_salary_cost):
    return safe_divide(revenue_attributed, total_salary_cost)


def revenue_per_day(revenue_attributed, days_worked):
    return safe_divide(revenue_attributed, days_worked)


def compute_revenue_attributed(division_revenue, allocation_weight):
    """§A7 — non-production departments (QC, packing & dispatch, admin,
    maintenance) receive an allocated share of division revenue rather than
    direct attribution. `None` if either input is missing."""
    if division_revenue is None or allocation_weight is None:
        return None
    return division_revenue * allocation_weight


_ROLLUP_KEYS = (
    "paid_hours", "break_hours", "changeover_hours", "downtime_hours",
    "idle_waiting_hours", "available_hours", "productive_hours", "bar_productive_hours",
)


def rollup_hours(shift_hours_list):
    """
    §A8 — aggregate hours first, then compute percentages from the
    aggregated hours. Averaging individual shifts'/employees' percentages is
    wrong whenever they cover different numbers of hours, and it is a silent
    error (the React build proved a 91.319% vs 90.960% divergence on a
    two-employee example).

    `shift_hours_list` is an iterable of dicts shaped like
    `compute_shift_hours`'s return value — one per shift, employee, or
    however the group is composed. Works the same whether the group is a
    whole department or one employee's own history across days.

    Returns the summed hours plus the three efficiency ratios computed from
    those sums.
    """
    totals = {k: 0.0 for k in _ROLLUP_KEYS}
    for sh in shift_hours_list:
        for k in _ROLLUP_KEYS:
            totals[k] += sh.get(k, 0.0)

    totals["utilisation_pct"] = utilisation_pct(totals["available_hours"], totals["paid_hours"])
    totals["true_efficiency_pct"] = true_efficiency_pct(totals["productive_hours"], totals["paid_hours"])
    totals["performance_while_working_pct"] = performance_while_working_pct(
        totals["productive_hours"], totals["available_hours"])
    return totals
