## 0.4.0 (2026-02-10)

### Feat

- **fill**: batch fill multiple weeks with `--from DDMMYYYY` / `--to DDMMYYYY`
- **fill**: fill Actual hours via JSGrid `UpdateProperties` API (reliable, bypasses click-and-type)
- **fill**: auto-skip Australian public holidays per configured region
- **fill**: `use_planned` option — copy server Planned hours as Actual
- **fill**: `clear_planned` option — zero-out Planned hours after filling
- **fill**: auto-clear Actual & Planned hours on tasks not listed in config
- **submit**: submit timesheets for approval via ribbon Send → Turn in Final Timesheet
- **recall**: recall submitted/approved timesheets with `--recall` flag
- **save**: force-save with retry logic; verify totals on summary page after save

### Fix

- fix data-value unit: grid uses 1/1000th of a minute (`hours × 60,000`), not milliseconds
- fix JS syntax error from escaped newline in join separator
- fix recall dialog handling — use Playwright `page.on('dialog')` for native `window.confirm()`
- fix navigation context destruction with retry logic in `_get_controller_name()`

### Refactor

- switch from `WriteLocalizedValueByKey` (silently fails) to `UpdateProperties` for all grid writes
- persistent browser profile replaces cookie-only `state.json` storage
- remove legacy `state_file` config option

## 0.3.0 (2026-02-10)

### Feat

- added open website functionality and manual login prompt if creds not saved

## 0.2.0 (2026-02-10)

### Feat

- **ci**: add pyproject.toml, commitizen, and GitHub Actions release workflow
- scaffold SharePoint timesheet bot project
