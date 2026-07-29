# Grandiose Bakery Dashboard (GIP III)

Streamlit dashboard for the Grandiose bakery financial performance and cost
optimization project. Navigate between sections using the dropdown at the
top of the page:

- **📊 Performance tracker** — headline KPIs (with live sparklines), a
  glowing food-cost trend chart, and four category panels (Production,
  Procurement, Warehouse & Inventory, Cost Control & Margin).
- **🧭 Scenario & resilience** — four tabs, one per topic: inflation-adjusted
  cost sensitivity (with a live donut chart that redraws as you move the
  inflation sliders), natural calamity / supply disruption risk, pandemic
  preparedness, and supplier concentration analysis.
- **👤 Employee portal** — self-service performance tracking for bakery staff
  (see below).

Scope: bakery operations only. Catering is excluded, per GIP III scope. All
figures are illustrative benchmarks pending Grandiose's actual bakery data
(flag for the company mentor).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).
Without any secrets configured, the app still runs fully — email sending and
persistent employee-portal storage just show a "not configured yet" message
until you add the secrets below.

## Deploy on Streamlit Community Cloud (free)

1. Push this folder (`app.py`, `requirements.txt`, `.streamlit/config.toml`)
   to a GitHub repository. **Do not push a real `secrets.toml`** — only
   `secrets.toml.example` should ever be in the repo.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **New app**, select the repo/branch, and set the main file path to
   `app.py`.
4. Click **Deploy**.
5. Add real secrets: **App → Settings → Secrets**, paste in values based on
   `secrets.toml.example` (see the two setup sections below for where each
   value comes from). Save — the app restarts automatically.

## Email setup (send report as email)

The sidebar's "📧 Email this report" panel generates a one-page PDF snapshot
of the current KPIs and sends it as an attachment via SMTP. It needs a
sending email account:

1. Use a Gmail account (or create a free one dedicated to this project).
2. Turn on 2-Step Verification: Google Account → Security → 2-Step
   Verification.
3. Create an App Password: Google Account → Security → 2-Step Verification →
   App passwords. Choose "Mail" as the app, generate, and copy the 16-character
   password.
4. In Streamlit Cloud's Secrets panel, add:
   ```toml
   EMAIL_ADDRESS = "your-address@gmail.com"
   EMAIL_APP_PASSWORD = "the-16-character-app-password"
   ```
5. That's it — no code changes needed. Any other SMTP provider (Outlook,
   SendGrid's SMTP relay, etc.) works too; just override `SMTP_SERVER` /
   `SMTP_PORT` in secrets if it isn't Gmail.

**Never use your real Google account password** — only an app password, and
only in Streamlit's Secrets panel, never committed to GitHub.

## Employee portal setup (persistent storage via Supabase)

The employee portal works immediately with no setup — entries are kept
in-session so the whole interface is demoable right away. They just don't
survive an app restart until Supabase is connected. To make it persistent:

1. Create a free project at https://supabase.com.
2. In the project's **SQL Editor**, run:
   ```sql
   create table employee_performance (
     id bigint generated always as identity primary key,
     employee_name text not null,
     employee_id text not null,
     department text,
     log_date date not null default current_date,
     shift text,
     units_produced int,
     wastage_pct numeric,
     hours_worked numeric,
     batch_time_adherence_pct numeric,
     quality_pass_pct numeric,
     revenue_generated numeric,
     notes text,
     created_at timestamptz default now()
   );

   create table employee_goals (
     id bigint generated always as identity primary key,
     employee_name text not null,
     employee_id text not null,
     department text,
     goal_text text,
     self_assessment text,
     manager_feedback text,
     created_at timestamptz default now()
   );

   -- This prototype uses the shared "anon" key with no per-employee login,
   -- so Row Level Security must be disabled (or given a permissive policy)
   -- for the anon key to read/write. Fine for a class project; revisit
   -- before using this with real employee data in production.
   alter table employee_performance disable row level security;
   alter table employee_goals disable row level security;
   ```
3. Go to **Project Settings → API** and copy the **Project URL** and the
   **anon public** key.
4. In Streamlit Cloud's Secrets panel, add:
   ```toml
   SUPABASE_URL = "https://your-project-ref.supabase.co"
   SUPABASE_KEY = "your-anon-public-key"
   ```
5. Redeploy/restart the app. The portal will show "💾 Storage: Supabase
   (persistent)" instead of "In-session only" once connected — no code
   changes needed, the same forms now write to the real database.

**Who's who:** `DEPARTMENT_HEADCOUNT` (near the top of `app.py`) holds the real
approximate headcount per line — Baklava, French Bakery, Arabic Bread,
Viennoiserie, Tahina, QC, Packing & Dispatch, Store, Maintenance, GM, Office,
Drivers, Third-Party Cleaners. `EMPLOYEES` holds a small illustrative sample
of named staff per department (2-3 each) so the portal and department
comparisons are demoable today — replace with Grandiose's full roster once
available; no other code changes are needed since every form and query reads
from these two lists. Since there's no real login, anyone with the link can
pick any department/name from the dropdown; this is fine for a class
prototype but should not be used for real HR data without adding actual
authentication.

## Benchmarks & targets from the 27 Jul MoM

A few figures baked into the portal come directly from the personal meeting
with the GM, and are marked in `app.py` for easy updating:

- `VALUE_PER_EMPLOYEE_DAY_TARGET = 1000` — GM's rule of thumb that each
  employee should generate ~AED 1,000/day of value. Populated from the
  optional "Revenue/value generated" field in the Daily log.
- `WASTAGE_TARGET_PCT = 1.0` — company target is 1%; current company-wide
  average is 2-3%. Used to flag wastage in both the employee and manager views.
- `DEPARTMENT_PRIORITY_TAG` — Arabic Bread is called out as a top
  revenue/profit driver; Tahina and Baklava are called out as the
  least-utilized lines. Shown as pills next to the department name.
- `DEPARTMENT_CAPACITY_TARGET` — **illustrative placeholder** daily-output
  targets for the five production lines. The MoM gives an overall utilization
  figure (30-35%, target 100% within a year) but not a per-line breakdown, so
  these numbers are directional only — replace with Grandiose's actual
  max-output figures per line as soon as they're available.
- `COMPANY_HEADCOUNT_CURRENT` / `COMPANY_HEADCOUNT_TARGET` — ~110 today,
  target 200 within a year, shown in the "Staff breakdown by line" panel.
- The Arabic Bread department shows a Morning/Night shift selector in the
  Daily log (and shift-comparison tables), since the MoM specifically notes
  Arabic Bread runs a separate night shift while other lines don't.

## Updating with real data

Replace the placeholder values in the `baseline`, `category_panels`, and the
default supplier table in `app.py` with Grandiose's actual bakery figures
once received. No structural changes are needed — the sliders, charts, and
thresholds will recalculate automatically against whatever base numbers are
in those variables.

## Color palette

Colors live in the `COLORS` dict near the top of `app.py` — a bold dark
theme (near-black canvas) with vibrant amber/gold as the primary accent and
coral/violet as secondary tones, echoing bakery warmth rather than a generic
blue/purple SaaS look. Adjust hex values there, and in
`.streamlit/config.toml`, to restyle.
