# Build Brief: Streamlit Employee Portal Calculations + B2B Page

**For:** Claude Code CLI
**Repo:** `parineetaajain29/GrandioseSupermarket-` — the **Streamlit** app (`grandiosesupermarket3.streamlit.app`)
**Author:** Parineeta Jain, with Rajveer Singh & Tarang Gupta — GIP III, SP Jain

This is the primary dashboard. Two jobs:

**A.** Fix the Employee Portal's labour calculations — the client asked specifically that these be accurate "to the decimal point" and account for changeover, breaks, and machine downtime.
**B.** Add a new B2B Client Performance page.

> A parallel React app (`grandiose-bakery-command-center`) already implements this calculation logic in TypeScript with 26 passing tests. **Port the logic, not the components.** If that repo is accessible, read `src/lib/labourCalc.ts` and `src/lib/labourCalc.test.ts` and translate them to Python — the formulas, the config-flag pattern, the guards, and the test cases all transfer directly. If it isn't accessible, build from §A2 below, which specifies the same chain.

---

## §0. Styling — read the existing CSS first, do not guess

Before writing any UI, find and read the app's existing CSS injection block (the `st.markdown("<style>…")` or equivalent) and any theme config. Extract the **actual** values in use:

- Background and card/panel colours
- The heading font (currently **Anton**, uppercase, wide letterforms) and the body font
- The accent colour (an amber/gold — used on the "Ramp-up focus" badge and focused input rings)
- The secondary badge colour (muted olive — used for the line name pill)
- Border treatment, corner radius, and the container padding rhythm

Every new element must reuse those exact values. Do not introduce new colours, fonts, radii, or spacing units. New sections must be visually indistinguishable from existing ones — someone should not be able to tell which parts are new.

Match existing component conventions too: the same `st.expander` styling, the same number-input steppers with `?` help tooltips, the same tab pattern (`My performance` / `Daily log` / `Goals & feedback`), the same section-heading treatment.

---

# PART A — Employee Portal calculations

## A1. The core problem

The Daily Log currently captures: units produced, units wasted, hours worked, units failed QC, batches completed, batches on standard time, revenue (optional), notes.

**Hours worked is a single whole-shift number.** There is no capture of breaks, changeover, or machine downtime — so true efficiency cannot be computed from existing data. Any efficiency figure shown today implicitly treats non-productive time as productive.

Fix the capture first, then the maths, then the display. In that order.

## A2. Add three fields to the Daily Log form

Place them in a clearly labelled group after "Hours worked". Whole minutes, integer steppers, styled identically to the existing number inputs, each with a `?` help tooltip.

| Label | Key | Default | Help text |
|---|---|---|---|
| Break time taken (minutes) | `break_minutes` | 35 | Total break minutes this shift. |
| Waiting for equipment (minutes) | `downtime_minutes` | 0 | Machine down, oven not ready, or waiting on materials. |
| Changeover / setup (minutes) | `changeover_minutes` | 0 | Switching between products or batch setup. |

Minutes not decimal hours — an employee knows "about 20 minutes", not "0.33 hours". Convert in code.

Add a short line above the group, in the existing muted body style:

> These are used to work out true efficiency. Reporting downtime honestly does not lower your personal performance score — it flags an equipment or scheduling issue.

That line matters. It is what makes the data trustworthy (see §A5).

## A3. Fix the revenue field

`Revenue/value generated (AED, optional)` currently defaults to `0.00`. Left blank it silently produces zero revenue per labour dirham and breaks the AED 1,000/day benchmark comparison.

Change to: keep it optional, but default to **empty/None rather than 0.00**, and distinguish "not entered" from "genuinely zero" in the data model. Where revenue is missing, downstream metrics must render `—`, never `0.00` or `0%`. Add a small muted note under the field: `Leave blank if not tracked for your line — it will be excluded rather than counted as zero.`

## A4. The calculation chain

New module `lib/labour_calc.py`. **Pure functions only** — no Streamlit calls, no I/O. This is what makes it testable.

```
paid_hours       = shift_length_hours * days_worked        # shift_length_hours = 9.0

break_hours      = break_minutes / 60
changeover_hours = changeover_minutes / 60
downtime_hours   = downtime_minutes / 60

available_hours  = paid_hours
                   - (0 if breaks_are_paid else break_hours)
                   - changeover_hours
                   - downtime_hours

productive_hours = available_hours - idle_waiting_hours

utilisation_pct       = available_hours  / paid_hours      * 100
true_efficiency_pct   = productive_hours / paid_hours      * 100
performance_while_working_pct = productive_hours / available_hours * 100

cost_per_productive_hour  = total_salary_cost / productive_hours
revenue_per_labour_dirham = revenue_attributed / total_salary_cost
revenue_per_day           = revenue_attributed / days_worked
```

Carry **full float precision** through every intermediate. Round only at display. Never round then feed forward.

**Guards:** any denominator of zero returns `None`, never `inf` or `nan`. Missing revenue returns `None`, not `0`.

## A5. Two efficiency figures, two owners — important

Because employees self-report their own deductions, a single efficiency number creates a bad incentive: reporting downtime would lower your score, so people stop reporting it and the data degrades.

Show both, clearly labelled:

- **True efficiency** (`productive / paid`) — the cost view. Management's number. Downtime reduces it, correctly, because Grandiose paid for that time.
- **Performance while working** (`productive / available`) — the employee's number. Unaffected by equipment failure or changeover.

In the employee's own "My performance" tab, lead with **performance while working**, and show true efficiency below it as division context. In the Manager / HR view, lead with **true efficiency**. Same underlying maths, different framing per audience.

## A6. Config

New `config/labour_config.py`:

```python
LABOUR_CONFIG = {
    "shift_length_hours": 9.0,        # documented Grandiose standard — not 8
    "breaks_are_paid": True,          # TODO: CONFIRM WITH GM — flips every efficiency figure
    "default_break_minutes": 35,      # TODO: CONFIRM WITH GM
    "revenue_attribution_basis": "allocated",
    "gm_daily_benchmark_aed": 1000.0, # TODO: reconfirm still current
}
```

Every calculation reads from here. Flipping `breaks_are_paid` must correctly cascade with no other edits — verify this by running the app both ways.

Where a displayed figure depends on an unconfirmed value, show a muted footnote in the existing caption style: `Breaks treated as paid — pending confirmation.`

## A7. Revenue attribution

Confirmed: non-production departments (QC, packing & dispatch, admin, maintenance) receive an **allocated share of division revenue**, not direct attribution.

```
revenue_attributed = division_revenue * allocation_weight
```

Store `allocation_weight` per department in the data layer. Document the basis in a comment — headcount share is the current placeholder and needs real figures from finance. Where a support department is shown, add the muted note: `Revenue allocated, not directly attributed.`

## A8. Department rollup

Aggregate **hours first, then compute percentages from the aggregated hours.** Do not average individual employees' percentages — that is wrong whenever employees worked different numbers of days, and it is a silent error. The React build proved a 91.319% vs 90.960% divergence on a two-employee example.

## A9. Display precision

| Value | Format |
|---|---|
| Hours | 2 dp (`165.40`) |
| Percentages | 1 dp (`71.6%`) |
| AED | 2 dp, thousands separator (`AED 4,200.00`) |
| Ratios | 2 dp + `×` (`6.23×`) |
| Missing / undefined | `—` |

Format at the display layer only. Never let a raw float reach the page — visible float artefacts will undermine the client's confidence in every other number on the dashboard.

## A10. Rework "My performance" tab

Show the chain, not just the answer. The client needs to see which numbers produced the result.

1. **Headline** — performance while working (employee view) or true efficiency (manager view), large, in the existing heading treatment
2. **Stacked horizontal bar** — productive / changeover / downtime / breaks / idle, proportional, with hours in the legend. **All segments must sum to paid hours.**
3. **Deduction table** — monospace, right-aligned: paid hours, each deduction as a subtraction, productive hours on a bordered total row
4. **Labour economics** — salary cost, cost per productive hour, revenue attributed, revenue per labour dirham
5. **Benchmark footer** — actual AED/day vs the AED 1,000/day benchmark with variance

## A11. Tests

`tests/test_labour_calc.py`, using pytest:

- A worked example with hand-verified expected outputs
- `breaks_are_paid` True vs False produce correctly different results
- Department rollup differs from the naive mean of percentages
- Zero productive hours returns `None`, not `inf`/`nan`
- Missing revenue returns `None`, not `0`
- Minutes-to-hours conversion is exact at boundary values

---

# PART B — B2B Client Performance page

New page in the existing multi-page structure, matching current navigation conventions.

## B1. Purpose

Flour Country / Grandiose Bakery is expanding into B2B. This page tracks that expansion at account and location level. The insight it exists to surface: **revenue-positive accounts can be margin-negative once service cost is included, and different again once idle capacity is accounted for.**

## B2. Sections, in order

**KPI strip (4)** — B2B revenue (MoM delta) · net margin after service cost (delta vs retail) · OTIF rate (late drop count) · average collection days (with 90-day supplier terms as context). Reuse the existing KPI/metric styling.

**Revenue vs service cost, 13 weeks** — two-line chart. Caption: `The gap between the lines is your margin. Watch for convergence.` Use whatever charting library the app already uses; do not add a new dependency.

**Capacity economics** — horizontal stacked bar of retail / B2B / idle oven capacity (currently ~32.4% utilised). Below it, marginal margin when an order fills idle capacity vs. when it forces overtime, and a count of next week's orders landing in overtime slots.

```
marginal_contribution = order_value
                        - ingredient_cost - packaging_cost - delivery_cost
                        - incremental_labour_cost      # ~0 when idle capacity absorbs it
                        - overtime_premium             # 0 unless the slot forces overtime
```

Caveat line, in the muted caption style: `Marginal view assumes fixed costs are absorbed by retail volume. Holds while utilisation stays low.`

**Concentration risk** — top-2 accounts as a share of B2B revenue, a proportional segmented bar of all accounts, and a plain-English consequence line. This mirrors the existing supplier concentration analysis, pointed at customers.

**Account profitability table** — Client (with location and delivery frequency), Revenue, Margin, **Marginal**, OTIF. Sortable and searchable by client name or location.

The **Marginal column beside Margin is the point of the table.** Include at least one account that is negative on full absorption but positive at the margin because it fills idle capacity, with a footnote explaining it. Without that column, management's obvious move is to drop such an account — which would leave capacity cold while removing real contribution.

**Receivables** — total outstanding, amount past 60 days, aging buckets (0–30 / 31–60 / 61–90 / 90+), CSV export.

**Recent deliveries feed** — client, location, time, on-time/late status, value. Late deliveries flagged in the existing warning colour.

## B3. Data

Follow whatever data pattern the app already uses (session state, Supabase table, or local file — match the existing convention). Mark placeholder data clearly, consistent with the existing `Figures shown are illustrative benchmarks pending Grandiose-provided actuals` disclaimer in the footer.

Client revenues must sum exactly to the summary revenue figure. The client will add them up.

---

## Constraints

- Reuse the existing colour palette, fonts, and spacing exactly — pull values from the existing CSS, never invent
- AED throughout, never `$`
- No catering metrics — explicitly out of project scope
- All arithmetic in pure functions in `lib/`; page files format and render only
- Where a value rests on an unconfirmed assumption, say so in a muted footnote rather than presenting it as settled
- Do not break the existing Performance Tracker, Scenario module, Company Profile, or Data Processor pages

## On finishing

Report: what changed, pytest results, any deviations and why, and an explicit list of every value still pending GM confirmation.
