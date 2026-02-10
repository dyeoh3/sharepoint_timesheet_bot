"""
Test script — opens timesheet(s) and fills hours for the projects
defined in config.yaml.

This script will:
1. Open the timesheet summary page
2. For each target week:
   a. Open/create the week's timesheet
   b. Fill the configured hours for each work day
   c. Save the timesheet
   d. Optionally submit it
3. Verify each week after saving

Run with: python scripts/test_fill_timesheet.py
Options:
  --dry-run          Skip saving and submitting
  --submit           Submit each timesheet after saving
  --from DDMMYYYY    Start filling from this Monday (e.g. 05012026)
  --to   DDMMYYYY    Fill up to and including this Monday (default: current week)
"""

import sys
from datetime import date, timedelta
from pathlib import Path

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import load_config
from bot.timesheet import (
    TimesheetEditPage,
    TimesheetSummaryPage,
    get_current_week_range,
)


def _parse_date_arg(value: str) -> date:
    """Parse a DDMMYYYY string into a date."""
    return date(int(value[4:8]), int(value[2:4]), int(value[0:2]))


def _mondays_in_range(start: date, end: date) -> list[date]:
    """Return all Mondays from *start* to *end* inclusive."""
    # Snap start to Monday
    start = start - timedelta(days=start.weekday())
    end = end - timedelta(days=end.weekday())
    weeks = []
    cur = start
    while cur <= end:
        weeks.append(cur)
        cur += timedelta(days=7)
    return weeks


def _fill_single_week(
    page,
    summary: TimesheetSummaryPage,
    target_monday: date,
    projects: list[dict],
    work_days: list[str],
    region: str,
    dry_run: bool,
    submit: bool,
):
    """Fill, save, and optionally submit a single week's timesheet."""

    # 1. Open the week's timesheet
    try:
        status = summary.open_timesheet(target_monday)
        print(f"   📋 Timesheet opened (was: {status})")
    except RuntimeError as e:
        print(f"   ❌ {e}")
        return False

    page.wait_for_timeout(2000)

    # 2. Fill hours
    editor = TimesheetEditPage(page)
    editor.fill_week_from_config(
        projects, work_days, region=region, period_start=target_monday
    )

    print(f"\n   ✅ All hours filled for week {target_monday.strftime('%d/%m/%Y')}")

    if dry_run:
        print("   🏃 Dry run — skipping save & submit")
        # Navigate back to summary for the next week
        summary.navigate()
        page.wait_for_timeout(1000)
        return True

    # 3. Save (force=True because JS API writes may not set dirty flag)
    print("\n   💾 Saving timesheet...")
    editor.save(force=True)

    # 4. Navigate back to summary and verify
    summary.navigate()
    page.wait_for_timeout(2000)

    save_verified = False
    result = summary.find_row_for_week(target_monday)
    if result:
        row, row_status = result
        texts = summary._row_cell_texts(row, max_cells=6)
        total = next(
            (t for t in texts if "h" in t.lower() and any(c.isdigit() for c in t)),
            None,
        )
        print(f"   📊 Summary: {' | '.join(texts)}")
        if total and total != "0h":
            print(f"   ✅ Saved total: {total}")
            save_verified = True
        else:
            print(f"   ❌ Total shows: {total or 'N/A'} — hours did NOT persist!")
    else:
        print("   ❌ Could not find week row in summary")

    if not save_verified:
        print(f"   ⛔ Skipping submit for {target_monday.strftime('%d/%m/%Y')} — save did not persist")
        return False

    # 5. Submit if requested
    if submit:
        print("\n   📤 Submitting timesheet...")
        try:
            summary.open_timesheet(target_monday)
            page.wait_for_timeout(2000)
            editor = TimesheetEditPage(page)
            editor.submit()

            # Verify
            page.wait_for_timeout(2000)
            summary.navigate()
            page.wait_for_timeout(2000)
            result = summary.find_row_for_week(target_monday)
            if result:
                _, row_status = result
                status_lower = row_status.lower() if row_status else ""
                if "submitted" in status_lower or "approved" in status_lower:
                    print(f"   ✅ Submitted (status: {row_status})")
                else:
                    print(f"   ⚠️  Status after submit: {row_status}")
            else:
                print("   ⚠️  Could not find row after submit")
        except Exception as e:
            print(f"   ❌ Submit failed: {e}")

    return True


def _recall_weeks(
    page,
    summary: TimesheetSummaryPage,
    weeks: list[date],
):
    """Recall one or more submitted/approved timesheets."""
    for i, monday in enumerate(weeks, 1):
        end = monday + timedelta(days=6)
        print("\n" + "=" * 60)
        print(
            f"📥 Recall {i}/{len(weeks)}: "
            f"{monday.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"
        )
        print("=" * 60)

        try:
            summary.recall(monday)
        except RuntimeError as e:
            print(f"   ❌ {e}")
            continue

    print(f"\n🏁 Finished recalling {len(weeks)} week(s)")


def test_fill_timesheet(
    dry_run: bool = False,
    submit: bool = False,
    recall: bool = False,
    from_date: date | None = None,
    to_date: date | None = None,
):
    """Fill hours for one or more weeks from config."""
    config = load_config()
    projects = config.get("projects", [])
    defaults = config.get("defaults", {})
    work_days = defaults.get("work_days", [])
    region = defaults.get("region", "NSW")

    if not recall and not projects:
        print("❌ No projects configured in config.yaml")
        return

    # Determine the range of Mondays to process
    current_monday = get_current_week_range()[0]
    if to_date is None:
        to_date = current_monday
    if from_date is None:
        from_date = current_monday

    weeks = _mondays_in_range(from_date, to_date)

    if recall:
        print("📥 RECALL MODE")
        print(f"📆 Weeks to recall: {len(weeks)}")
        for w in weeks:
            end = w + timedelta(days=6)
            print(f"   • {w.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}")
        print()

        with BrowserManager(config) as bm:
            page = bm.page
            print("🌐 Opening timesheet summary page...\n")
            summary = TimesheetSummaryPage(page)
            summary.navigate()

            if bm.is_on_login_page(page):
                bm.wait_for_manual_login(page)
                summary.navigate()

            _recall_weeks(page, summary, weeks)

            print("\n" + "=" * 60)
            print(f"🏁 Finished recall for {len(weeks)} week(s)!")
            print("=" * 60)
            if "--no-close" in sys.argv:
                input("\nPress Enter to close the browser...")

        print("👋 Browser closed.")
        return

    print("📋 Projects to fill:")
    for p in projects:
        if p.get("use_planned"):
            print(f"   • {p['name']}: use Planned hours")
        else:
            print(f"   • {p['name']}: {p.get('default_hours_per_day', 0)}h/day")
    print(f"📅 Work days: {', '.join(work_days)}")
    print(f"🌏 Region: {region}")
    print(f"📆 Weeks to fill: {len(weeks)}")
    for w in weeks:
        end = w + timedelta(days=6)
        print(f"   • {w.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}")
    print()

    with BrowserManager(config) as bm:
        page = bm.page

        # Navigate to timesheet summary
        print("🌐 Opening timesheet summary page...\n")
        summary = TimesheetSummaryPage(page)
        summary.navigate()

        # Handle login if needed
        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            summary.navigate()

        # Process each week
        for i, monday in enumerate(weeks, 1):
            end = monday + timedelta(days=6)
            print("\n" + "=" * 60)
            print(
                f"📅 Week {i}/{len(weeks)}: "
                f"{monday.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"
            )
            print("=" * 60)

            success = _fill_single_week(
                page, summary, monday,
                projects, work_days, region,
                dry_run, submit,
            )
            if not success:
                print(f"   ⚠️  Skipping week {monday.strftime('%d/%m/%Y')}")

        print("\n" + "=" * 60)
        print(f"🏁 Finished processing {len(weeks)} week(s)!")
        print("=" * 60)
        if "--no-close" in sys.argv:
            input("\nPress Enter to close the browser...")

    print("👋 Browser closed.")


if __name__ == "__main__":
    dry = "--dry-run" in sys.argv
    do_submit = "--submit" in sys.argv
    do_recall = "--recall" in sys.argv

    from_date = None
    to_date = None
    if "--from" in sys.argv:
        idx = sys.argv.index("--from")
        from_date = _parse_date_arg(sys.argv[idx + 1])
    if "--to" in sys.argv:
        idx = sys.argv.index("--to")
        to_date = _parse_date_arg(sys.argv[idx + 1])

    test_fill_timesheet(
        dry_run=dry, submit=do_submit, recall=do_recall,
        from_date=from_date, to_date=to_date,
    )
