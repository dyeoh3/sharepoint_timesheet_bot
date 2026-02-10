"""
Main orchestrator — ties together browser, auth, and timesheet page objects.
"""

from bot.browser import BrowserManager
from bot.config import load_config
from bot.timesheet import TimesheetEditPage, TimesheetSummaryPage, get_current_week_range


def run_timesheet_bot(
    config_path: str | None = None,
    dry_run: bool = False,
    submit: bool = False,
):
    """
    Main entry point: open SharePoint, log in, and fill out the timesheet.

    Args:
        config_path: Optional path to a YAML config file.
        dry_run: If True, fill hours but don't save or submit.
        submit: If True, submit the timesheet after saving.
    """
    config = load_config(config_path)
    monday, friday = get_current_week_range()
    print(f"📅 Filling timesheet for week: {monday} → {friday}")

    with BrowserManager(config) as bm:
        page = bm.page

        # 1. Navigate to timesheet summary
        summary = TimesheetSummaryPage(page)
        summary.navigate()

        # 2. Handle login — let user authenticate manually if needed
        if bm.is_on_login_page(page):
            if not bm.has_valid_session():
                print("⚠️  No saved session found. Run 'python main.py login' first,")
                print("   or log in now in the browser window.")
            bm.wait_for_manual_login(page)
            summary.navigate()

        # 4. Create timesheet for current period if needed
        summary.create_timesheet_if_needed()

        # 5. Open the current period's timesheet
        status = summary.get_current_period_status()
        print(f"📋 Current period status: {status}")
        summary.click_current_period()

        # 6. Fill in hours from config
        editor = TimesheetEditPage(page)
        projects = config.get("projects", [])
        work_days = config.get("defaults", {}).get("work_days", [])
        editor.fill_week_from_config(projects, work_days)
        print("✅ Hours filled")

        if dry_run:
            print("🏃 Dry run — not saving. Browser will stay open for inspection.")
            input("Press Enter to close the browser...")
            return

        # 7. Save
        editor.save()
        print("💾 Timesheet saved")

        # 8. Optionally submit
        if submit:
            editor.submit()
            print("🚀 Timesheet submitted for approval")

    print("✅ Done!")
