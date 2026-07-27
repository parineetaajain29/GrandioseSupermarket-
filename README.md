# Grandiose Bakery Dashboard (GIP III)

Two-tab Streamlit dashboard for the Grandiose bakery financial performance and
cost optimization project.

- **Tab A — Performance tracker**: headline KPIs, food-cost trend, and four
  category panels (Production, Procurement, Warehouse & Inventory, Cost
  Control & Margin).
- **Tab B — Scenario & resilience**: four interactive modules —
  inflation-adjusted cost sensitivity, natural calamity / supply disruption
  risk, pandemic preparedness, and supplier concentration analysis.

Scope: bakery operations only. Catering is excluded, per GIP III scope.
All figures are illustrative benchmarks pending Grandiose's actual bakery
data (flag for the company mentor).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## Deploy on Streamlit Community Cloud (free)

1. Push this folder (`app.py`, `requirements.txt`) to a GitHub repository.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **New app**, select the repo/branch, and set the main file path to
   `app.py`.
4. Click **Deploy** — Streamlit Cloud installs `requirements.txt`
   automatically and gives you a public `*.streamlit.app` URL you can share
   for the mid-review.

## Updating with real data

Replace the placeholder values in the `baseline`, `category_panels`, and the
default supplier table in `app.py` with Grandiose's actual bakery figures
once received. No structural changes are needed — the sliders, charts, and
thresholds will recalculate automatically against whatever base numbers are
in those variables.

## Color palette

Colors live in the `COLORS` dict near the top of `app.py` — muted sage green,
warm terracotta, and cream tones intended to echo Grandiose's fresh
neighbourhood-grocery identity without using bold/saturated primary colors.
Adjust hex values there if you have exact brand guideline colors to match.
