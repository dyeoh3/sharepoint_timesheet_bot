# SharePoint Timesheet Bot

Automates filling out weekly timesheets on SharePoint PWA using Python + Playwright.

## Setup

### Option 1: Virtual Environment (Recommended)

```bash
# 1. Create & activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Install Playwright browsers
playwright install chromium

# 4. Copy .env.example → .env and fill in your SharePoint URLs
cp .env.example .env
```

When using a virtual environment, run scripts with `python` directly:
```bash
python scripts/test_open_site.py
python main.py login
```

### Option 2: System Python

```bash
# 1. Install dependencies globally
pip3 install -r requirements.txt

# 2. Install Playwright browsers
playwright install chromium

# 3. Copy .env.example → .env and fill in your SharePoint URLs
cp .env.example .env
```

When using system Python without a virtual environment, use `python3`:
```bash
python3 scripts/test_open_site.py
python3 main.py login
```

## Configuration

- **`.env`** — SharePoint URLs and runtime flags (see `.env.example`).
- **`config.yaml`** — Projects, default hour allocations, and browser settings.

## Usage

**Note:** Use `python` if you activated a virtual environment, or `python3` for system Python.

```bash
# Smoke test — opens the site, handles login, checks page elements
python scripts/test_open_site.py

# First run — log in and save session
python main.py login

# Inspect the page (useful for debugging selectors)
python main.py inspect

# Fill out the current week (dry run — no save)
python main.py fill --dry-run

# Fill and save
python main.py fill

# Fill, save, and submit for approval
python main.py fill --submit
```

## How It Works

1. Launches Chrome via Playwright using a **persistent browser profile**
2. Navigates to the SharePoint Timesheet page
3. Handles Microsoft SSO login (manual first time; session is reused automatically on subsequent runs)
4. Reads your project/hours config from `config.yaml`
5. Fills in the timesheet grid
6. Optionally saves and submits

## Authentication & Session Persistence

The bot uses a **persistent Chromium profile** (`browser_state/profile/`) rather than
cookie-only storage. This preserves the full browser state — cookies, localStorage,
IndexedDB, and service workers — so Microsoft SSO sessions survive between runs
without needing to log in again.

- **First run**: You'll be prompted to log in manually in the browser window (including MFA).
- **Subsequent runs**: The saved profile auto-redirects past the login page.
- **Session expired?**: The bot detects this and falls back to the manual login prompt.
- **Clear session**: Delete the `browser_state/profile/` directory.

## Important Notes

- **Selectors need tuning**: SharePoint PWA HTML structure varies. Run `python scripts/test_open_site.py` first, open DevTools (Cmd+Opt+I), and update the selectors in `bot/timesheet.py` to match your actual page.
- **MFA**: Run `python scripts/test_open_site.py` or `python main.py login` first. Complete MFA manually once, and the persistent profile will be reused.
- **Security**: Never commit your `.env` file or `browser_state/` directory. Both are in `.gitignore`.

## Project Structure

```
sharepoint_timesheet_bot/
├── bot/
│   ├── __init__.py       # Package init
│   ├── browser.py        # Playwright browser lifecycle & auth
│   ├── config.py         # Config/env loader
│   ├── runner.py         # Main orchestrator
│   └── timesheet.py      # SharePoint page objects & interactions
├── scripts/
│   └── test_open_site.py # Smoke test — login, page element checks
├── browser_state/
│   └── profile/          # Persistent Chromium profile (gitignored)
├── .env                  # Your SharePoint URLs (gitignored)
├── .env.example          # Template for .env
├── .gitignore
├── config.yaml           # Project/hours configuration
├── main.py               # CLI entry point
├── README.md
└── requirements.txt
```
