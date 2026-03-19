# MSP Quote Generator

Internal web tool for MSP Construction Ltd. Upload an Excel workbook and instantly download a formatted landscape PDF quote.

## How it works

1. Upload an `.xlsx` workbook with the required sheets (see below)
2. Click **Generate PDF Quote**
3. The PDF downloads automatically

### Required Excel sheets

| Sheet name | Contents |
|---|---|
| `Client Summary (2)` | Client name, address, date, scope, and summary table |
| `MSP LMS` | Full labour & materials estimate with section breakdowns |

---

## Running locally

**Prerequisites:** Python 3.10+

```bash
# 1. Clone and enter the repo
git clone <repo-url>
cd msp-quote

# 2. Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. (Optional) Set up environment variables
cp .env.example .env
# Edit .env — leave APP_PASSWORD blank to skip auth on localhost

# 5. Start the dev server
python app.py
```

Then open [http://localhost:3000](http://localhost:3000).

> If `APP_PASSWORD` is not set, the app runs without authentication — safe for local use.

---

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `SECRET_KEY` | `msp-dev-key-change-in-production` | Flask session secret — change in production |
| `APP_USER` | `msp` | HTTP Basic Auth username |
| `APP_PASSWORD` | *(empty)* | HTTP Basic Auth password — leave blank to disable auth |
| `DEPLOYED_AT` | *(process start time)* | ISO timestamp shown in the footer, set by your deploy script |

---

## Deploying

The app includes a `Procfile` for Heroku / Railway / Render:

```
web: gunicorn app:app
```

Set the environment variables above in your platform's dashboard, and stamp the deploy time in your deploy script:

```bash
export DEPLOYED_AT="$(date -u +%Y-%m-%dT%H:%M:%S)"
```

---

## Project structure

```
msp-quote/
├── app.py                  # Flask app & routes
├── quote_engine/
│   ├── __init__.py
│   └── generator.py        # PDF generation logic
├── templates/
│   └── index.html          # Upload UI
├── requirements.txt
├── Procfile
└── .env.example
```
