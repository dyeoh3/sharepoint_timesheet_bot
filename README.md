# SharePoint Timesheet Bot

Automates filling, saving, submitting, and recalling weekly timesheets on SharePoint PWA using Python + Playwright.

## Features

- **Batch fill** — fill hours for one or many weeks in a single run
- **Public holiday awareness** — automatically skips Australian public holidays (configurable state)
- **Planned → Actual** — optionally copy server-side Planned hours as Actual
- **Clear Planned** — zero-out Planned hours after filling
- **Non-config task cleanup** — clears hours from tasks not listed in your config
- **Save & verify** — saves and checks the summary page to confirm hours persisted
- **Submit** — submits timesheets for approval after filling
- **Recall** — recalls submitted/approved timesheets so they can be re-edited
- **Persistent SSO** — Microsoft login is saved across runs (manual login only needed once)

## Setup

### 1. Create a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate   # macOS / Linux
# .venv\Scripts\activate    # Windows
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 3. Configure environment

```bash
cp .env.example .env
```

Edit `.env` with your SharePoint URLs:

```dotenv
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_TIMESHEET_URL=https://yourcompany.sharepoint.com/sites/YourPWA/_layouts/15/pwa/Timesheet/MyTSSummary.aspx
HEADLESS=false
```

### 4. First-time login

```bash
.venv/bin/python main.py login
```

A browser window opens. Log in with your Microsoft account (including MFA). The session is saved to `browser_state/profile/` and reused automatically on subsequent runs.

## Configuration

Edit `config.yaml` to define your projects and defaults:

```yaml
defaults:
  total_hours_per_day: 8
  region: "NSW"              # Australian state for public holidays
  work_days:
    - Monday
    - Tuesday
    - Wednesday
    - Thursday
    - Friday

projects:
  - name: "ST-333 Marketplace 2.0"
    default_hours_per_day: 8
    clear_planned: true

  # Copy Planned hours as Actual instead of a fixed value:
  # - name: "ST-456 Other Project"
  #   use_planned: true

browser:
  slow_mo: 100               # ms delay between actions
  timeout: 60000             # ms timeout for page loads
  user_data_dir: "browser_state/profile"
```

### Project options

| Key | Type | Description |
|---|---|---|
| `name` | `str` | Task name as it appears in the SharePoint grid (substring match) |
| `default_hours_per_day` | `float` | Fixed hours to fill per work day |
| `use_planned` | `bool` | Copy server Planned hours as Actual instead of a fixed value |
| `clear_planned` | `bool` | Zero-out Planned hours for this task after filling |

Tasks **not** listed in `projects` will have their Actual and Planned hours cleared automatically.

## Usage

> All commands below use `.venv/bin/python`. If your venv is activated, you can use `python` directly.

### Fill the current week (dry run)

```bash
.venv/bin/python scripts/test_fill_timesheet.py --dry-run
```

### Fill the current week and save

```bash
.venv/bin/python scripts/test_fill_timesheet.py
```

### Fill the current week, save, and submit

```bash
.venv/bin/python scripts/test_fill_timesheet.py --submit
```

### Fill a date range

Use `--from DDMMYYYY` and `--to DDMMYYYY` to fill multiple weeks:

```bash
# Fill and submit all weeks from 5 Jan to 8 Feb 2026
.venv/bin/python scripts/test_fill_timesheet.py \
  --from 05012026 --to 08022026 --submit
```

Omitting `--to` defaults to the current week.

### Recall submitted/approved timesheets

```bash
# Recall a single week
.venv/bin/python scripts/test_fill_timesheet.py \
  --recall --from 02022026 --to 02022026

# Recall multiple weeks
.venv/bin/python scripts/test_fill_timesheet.py \
  --recall --from 06072026 --to 13072026
```

After recalling, re-run with `--from` / `--to` and `--submit` to re-fill and resubmit.

### Keep the browser open after completion

Add `--no-close` to any command to keep the browser window open for inspection:

```bash
.venv/bin/python scripts/test_fill_timesheet.py --submit --no-close
```

### CLI entry point (main.py)

`main.py` provides a Click-based CLI for single-week operations:

```bash
.venv/bin/python main.py login      # Save Microsoft SSO session
.venv/bin/python main.py inspect    # Open browser for manual inspection
.venv/bin/python main.py fill       # Fill current week and save
.venv/bin/python main.py fill --dry-run
.venv/bin/python main.py fill --submit
```

### Utility scripts

| Script | Purpose |
|---|---|
| `scripts/test_open_site.py` | Smoke test — opens SharePoint, checks page elements |
| `scripts/test_open_timesheet.py` | Opens the current week's timesheet for inspection |
| `scripts/test_fill_timesheet.py` | Main batch fill / submit / recall script |

## How It Works

1. Launches Chromium via Playwright using a **persistent browser profile**
2. Navigates to the SharePoint PWA Timesheet summary page
3. Handles Microsoft SSO (manual first time; auto-redirect on subsequent runs)
4. For each target week:
   - Opens the timesheet (creates it if status is "Not Yet Created")
   - Clears hours from tasks not in the config
   - Fills Actual hours via the JSGrid `UpdateProperties` API
   - Optionally clears Planned hours
   - Saves and verifies the total on the summary page
   - Optionally submits for approval

### Technical details

The SharePoint PWA grid (JSGrid) stores work-duration values in units of **1/1000th of a minute** — i.e. `hours × 60,000`. For example, 8h = `480,000`.

All writes use `grid.UpdateProperties()` with `SP.JsGrid.CreateValidatedPropertyUpdate()` to go through the grid's change-tracking pipeline. This ensures `IsDirty()` returns `true` and the Save button persists the changes server-side.

## Authentication & Session Persistence

The bot uses a **persistent Chromium profile** (`browser_state/profile/`) rather than cookie-only storage. This preserves the full browser state — cookies, localStorage, IndexedDB, and service workers — so Microsoft SSO sessions survive between runs.

| Scenario | What happens |
|---|---|
| First run | Browser opens; you log in manually (including MFA) |
| Subsequent runs | Saved profile auto-redirects past the login page |
| Session expired | Bot detects this and falls back to manual login prompt |
| Clear session | Delete the `browser_state/profile/` directory |

## Project Structure

```
sharepoint_timesheet_bot/
├── bot/
│   ├── __init__.py          # Package init
│   ├── browser.py           # Playwright browser lifecycle & Microsoft SSO auth
│   ├── config.py            # YAML config & .env loader
│   ├── holidays.py          # Australian public holiday detection
│   ├── runner.py            # High-level orchestrator (used by main.py)
│   └── timesheet.py         # SharePoint page objects & JSGrid interactions
├── scripts/
│   ├── test_open_site.py    # Smoke test — login & page element checks
│   ├── test_open_timesheet.py  # Open current week's timesheet
│   └── test_fill_timesheet.py  # Batch fill / submit / recall
├── browser_state/
│   └── profile/             # Persistent Chromium profile (gitignored)
├── .env                     # SharePoint URLs & runtime flags (gitignored)
├── .env.example             # Template for .env
├── .gitignore
├── config.yaml              # Project & hours configuration
├── main.py                  # Click CLI entry point
├── pyproject.toml           # Project metadata & commitizen config
├── requirements.txt         # Python dependencies
├── CHANGELOG.md
└── README.md
```

## Security

- **Never commit** your `.env` file or `browser_state/` directory — both are in `.gitignore`.
- The persistent browser profile contains your full authenticated session. Treat it like a password.
