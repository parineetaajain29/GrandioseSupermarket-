"""
Labour calculation configuration for the Employee Portal.

Every efficiency/cost figure the portal shows reads its assumptions from
here — nowhere else. That is what makes flipping `breaks_are_paid` (or any
other value below) cascade correctly through every efficiency figure with
no other edits: change it once, here, and every caller of labour_calc.py
picks it up on the next render.

Values marked "TODO: CONFIRM WITH GM" are working assumptions, not
confirmed figures — see the Employee Portal's pending-confirmation
footnotes, and the final build report, for the full list.
"""

LABOUR_CONFIG = {
    "shift_length_hours": 9.0,        # documented Grandiose standard — not 8
    "breaks_are_paid": True,          # TODO: CONFIRM WITH GM — flips every efficiency figure
    "default_break_minutes": 35,      # TODO: CONFIRM WITH GM
    "revenue_attribution_basis": "allocated",
    "gm_daily_benchmark_aed": 1000.0, # TODO: reconfirm still current
}
