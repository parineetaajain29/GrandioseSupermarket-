# Grandiose Bakery — Financial Performance & Cost Optimization Dashboard

GIP III project for Grandiose Bakery's UAE bakery division. A single-page
Streamlit app covering performance tracking, scenario/resilience modelling,
an employee portal, a company profile overview, and an AI-assisted data
processor for messy uploaded files.

[![CI](https://github.com/parineetaajain29/grandiosesupermarket-/actions/workflows/ci.yml/badge.svg)](https://github.com/parineetaajain29/grandiosesupermarket-/actions/workflows/ci.yml)

## Features

- **Performance tracker** — KPI cards with sparklines, wastage gauge, cost
  structure donut, food-cost trend chart with a projected view.
- **Scenario & resilience** — inflation sensitivity, supply disruption,
  pandemic preparedness, and supplier concentration modules.
- **Employee portal** — shift logging and self-assessment goals for
  individual employees, plus a manager/HR roll-up view. Persists to
  Supabase when configured, otherwise falls back to in-session storage
  automatically — no setup required to try it out.
- **Company profile** — production capacity, shift structure, sales mix,
  cost structure, and inventory/wastage overview.
- **Data processor** — drag-and-drop Excel/CSV/PDF/Word upload that uses
  Claude to interpret messy files into structured, exportable sheets.
- **Excel export & email** — every section can be exported as a formatted
  `.xlsx` workbook, or emailed directly from the app.

All AI, database, and email integrations degrade gracefully when their
secrets aren't configured — the app is fully usable locally with zero
setup beyond `pip install`.

## Tech stack

Streamlit · Pandas / NumPy · Plotly · Anthropic (Claude) · Supabase ·
openpyxl · pypdf · python-docx

## Project structure

```
.
├── app.py                        # the entire app (single Streamlit script)
├── requirements.txt
├── runtime.txt                   # pins the Python version for Streamlit Cloud
├── .streamlit/
│   ├── config.toml                # theme (dark, gold accent)
│   └── secrets.toml.example       # template — copy to secrets.toml locally
└── .github/workflows/ci.yml       # compiles + boots the app on every push
```

## Local development

```bash
git clone https://github.com/parineetaajain29/grandiosesupermarket-.git
cd grandiosesupermarket-
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit .streamlit/secrets.toml with real values (all optional, see below)

streamlit run app.py
```

The app opens at `http://localhost:8501`. Every integration below is
optional — the app runs and looks the same without any of them, just
with AI processing / persistence / email switched off and a clear message
shown wherever a feature needs a key it doesn't have.

## Configuration (secrets)

All secrets live in `.streamlit/secrets.toml` locally, or in the
Streamlit Community Cloud "Secrets" panel in production. **Never commit
real credentials** — `.streamlit/secrets.toml` is gitignored; only the
`.example` template is tracked.

| Key | Used for | Required? |
|---|---|---|
| `ANTHROPIC_API_KEY` | Data processor's AI file interpretation | Optional |
| `SUPABASE_URL`, `SUPABASE_KEY` | Employee portal persistence | Optional |
| `EMAIL_ADDRESS`, `EMAIL_APP_PASSWORD` | "Email this report" | Optional |
| `SMTP_SERVER`, `SMTP_PORT` | Override email defaults (Gmail) | Optional |

### Data processor setup

1. Get an API key from the [Anthropic Console](https://console.anthropic.com/).
2. Set `ANTHROPIC_API_KEY` in your secrets.
3. Without it, the Data Processor still accepts uploads and shows a clear
   inline message explaining that AI interpretation isn't configured yet.

### Employee portal setup

The portal uses two Supabase tables: `employee_performance` (shift logs)
and `employee_goals` (goals / self-assessments / manager feedback).

1. Create a free project at [supabase.com](https://supabase.com/).
2. In the SQL editor, run:

   ```sql
   create table employee_performance (
     id bigint generated always as identity primary key,
     employee_name text not null,
     employee_id text not null,
     department text not null,
     log_date date not null,
     shift text not null default 'Morning',
     units_produced integer not null default 0,
     units_wasted integer not null default 0,
     batches_completed integer not null default 0,
     batches_on_time integer not null default 0,
     units_failed_qc integer not null default 0,
     wastage_pct numeric not null default 0,
     hours_worked numeric not null default 0,
     batch_time_adherence_pct integer not null default 0,
     quality_pass_pct integer not null default 0,
     revenue_generated numeric not null default 0,
     notes text,
     created_at timestamptz not null default now()
   );

   create table employee_goals (
     id bigint generated always as identity primary key,
     employee_name text not null,
     employee_id text not null,
     department text not null,
     goal_text text,
     self_assessment text,
     manager_feedback text,
     created_at timestamptz not null default now()
   );
   ```

3. Copy the **Project URL** and **anon public** API key from
   Project Settings → API into `SUPABASE_URL` / `SUPABASE_KEY`.
4. This app has no login system, so leave Row Level Security **off** on
   both tables (the default for a freshly created table) so the anon key
   can read and write. If you later add auth, enable RLS and add policies
   scoped to `employee_id` instead of leaving the tables open.
5. Without Supabase configured, shift logs and goals are still saved —
   just in the browser session only, and lost on refresh/restart. The
   portal shows which mode it's in at all times.

### Email setup

Uses Gmail SMTP with an [app password](https://myaccount.google.com/apppasswords)
(not your regular Gmail password — requires 2-Step Verification enabled
on the Google account). Set `EMAIL_ADDRESS` and `EMAIL_APP_PASSWORD`; to
use a non-Gmail SMTP provider, also set `SMTP_SERVER` / `SMTP_PORT`.

## Deployment (Streamlit Community Cloud)

1. Push your changes to GitHub (see below) — the app deploys straight
   from the repo, no build step or Dockerfile needed.
2. Go to [share.streamlit.io](https://share.streamlit.io/) and sign in
   with GitHub.
3. **Create app** → pick this repository, the branch to deploy
   (`main` for production), and set the main file path to `app.py`.
4. Before or after the first deploy, open **App settings → Secrets** and
   paste in the contents of your `.streamlit/secrets.toml` (only the
   keys you actually use — all are optional, see above).
5. Deploy. `runtime.txt` pins the Python version so the build is
   reproducible; `requirements.txt` is installed automatically.
6. From then on, every push to the deployed branch triggers an automatic
   redeploy — no manual step needed.

## Continuous integration

`.github/workflows/ci.yml` runs on every push and PR: it installs
dependencies, byte-compiles `app.py`, then boots it headless with
`streamlit run` and checks it actually serves a `200` — catching import
errors, syntax errors, and startup crashes before they reach a deploy.

## Updating dependencies

Dependency floors live in `requirements.txt` (e.g. `streamlit>=1.38`).
Bump a floor when you rely on a newer feature; otherwise leave the range
open so deploys pick up patch releases automatically. Run the CI smoke
test locally (`streamlit run app.py`) after any dependency change.
