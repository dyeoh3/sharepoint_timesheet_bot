"""
Test script — opens the current week's timesheet and fills hours
for the projects defined in config.yaml.

This script will:
1. Open the timesheet summary page
2. Open/create the current week's timesheet
3. For each project in config.yaml:
   a. Find the task row in the grid (or add it via existing assignments)
   b. Fill the configured hours for each work day
4. Save the timesheet

Run with: python scripts/test_fill_timesheet.py
Add --dry-run to skip saving.
"""

import sys
from pathlib import Path

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import load_config
from bot.timesheet import TimesheetEditPage, TimesheetSummaryPage


def test_fill_timesheet(dry_run: bool = False):
    """Open the current week's timesheet and fill hours from config."""
    config = load_config()
    projects = config.get("projects", [])
    work_days = config.get("defaults", {}).get("work_days", [])

    if not projects:
        print("❌ No projects configured in config.yaml")
        return

    print("📋 Projects to fill:")
    for p in projects:
        print(f"   • {p['name']}: {p.get('default_hours_per_day', 0)}h/day")
    print(f"📅 Work days: {', '.join(work_days)}\n")

    with BrowserManager(config) as bm:
        page = bm.page

        # 1. Navigate to timesheet summary
        print("🌐 Opening timesheet summary page...\n")
        summary = TimesheetSummaryPage(page)
        summary.navigate()

        # Handle login if needed
        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            summary.navigate()

        # 2. Open current week's timesheet
        try:
            status = summary.open_timesheet()
            print(f"📋 Timesheet opened (was: {status})\n")
        except RuntimeError as e:
            print(f"❌ {e}")
            input("\nPress Enter to close the browser...")
            return

        # Give the edit page time to fully render
        page.wait_for_timeout(2000)

        print(f"📄 Edit page loaded: {page.title()}")
        print(f"   URL: {page.url}\n")

        # 3. Dump visible grid content for debugging
        print("=" * 60)
        print("🔍 Current grid content:")
        print("=" * 60)
        editor = TimesheetEditPage(page)
        task_rows = editor.get_task_rows()
        if task_rows:
            for item in task_rows:
                print(f"   Row {item['left_index']:2d}: {item['name']}")
        else:
            print("   (no task rows detected)")
        print("=" * 60)
        print()

        # 4. Fill hours
        editor.fill_week_from_config(projects, work_days)

        print("\n" + "=" * 60)
        print("✅ All hours filled!")
        print("=" * 60)

        if dry_run:
            print("\n🏃 Dry run — NOT saving.")
            print("🔗 Browser is still open — inspect with DevTools (Cmd+Opt+I)")
            input("   Press Enter to close the browser...")
            return

        # 5. Save
        print("\n💾 Saving timesheet...")
        editor.save()
        print("✅ Timesheet saved!\n")

        input("Press Enter to close the browser...")

    print("👋 Browser closed.")


if __name__ == "__main__":
    dry = "--dry-run" in sys.argv
    test_fill_timesheet(dry_run=dry)
