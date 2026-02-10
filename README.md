# SharePoint Timesheet Bot

Automates filling out weekly timesheets on SharePoint PWA using Python + Playwright.

## Setup

```bash
# 1. Create & activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Install Playwright browsers (uses your installed Chrome)
playwright install chromium

# 4. Copy .env.example → .env and add your credentials
cp .env.example .env
# Edit .env with your MS_EMAIL and MS_PASSWORD
```

## Configuration

Edit **config.yaml** to match your projects and default hour allocations.

## Usage

```bash
# First run — log in and save auth state
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

1. Launches Chrome via Playwright
2. Navigates to the SharePoint Timesheet page
3. Handles Microsoft SSO login (saves session for reuse)
4. Reads your project/hours config from `config.yaml`
5. Fills in the timesheet grid
6. Optionally saves and submits

## Important Notes

- **Selectors need tuning**: SharePoint PWA HTML structure varies. Run `python main.py inspect` first, open DevTools, and update the selectors in `bot/timesheet.py` to match your actual page.
- **MFA**: If your org uses MFA, run `python main.py login` first. Complete MFA manually once, and the saved auth state will be reused.
- **Security**: Never commit your `.env` file. It's in `.gitignore` by default.

## Project Structure

```
sharepoint_timesheet_bot/
├── bot/
│   ├── __init__.py       # Package init
│   ├── browser.py        # Playwright browser lifecycle & auth
│   ├── config.py         # Config/env loader
│   ├── runner.py         # Main orchestrator
│   └── timesheet.py      # SharePoint page objects & interactions
├── browser_state/        # Saved auth cookies (gitignored)
├── .env                  # Your credentials (gitignored)
├── .env.example          # Template for .env
├── .gitignore
├── config.yaml           # Project/hours configuration
├── main.py               # CLI entry point
├── README.md
└── requirements.txt
```
