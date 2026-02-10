"""
CLI entry point for the SharePoint Timesheet Bot.
"""

import click

from bot.runner import run_timesheet_bot


@click.group()
@click.version_option(version="0.1.0")
def cli():
    """SharePoint Timesheet Bot — automate your weekly timesheet."""
    pass


@cli.command()
@click.option("--config", "-c", default=None, help="Path to config.yaml")
@click.option("--dry-run", is_flag=True, help="Fill hours but don't save (browser stays open)")
@click.option("--submit", is_flag=True, help="Submit timesheet after saving")
def fill(config, dry_run, submit):
    """Fill out the current week's timesheet from config defaults."""
    run_timesheet_bot(config_path=config, dry_run=dry_run, submit=submit)


@cli.command()
def login():
    """Open browser, let you log in manually, and save the session for reuse."""
    from bot.browser import BrowserManager
    from bot.config import load_config
    from bot.timesheet import TimesheetSummaryPage

    config = load_config()
    with BrowserManager(config) as bm:
        page = bm.page
        summary = TimesheetSummaryPage(page)
        summary.navigate()
        bm.wait_for_manual_login(page)
        print("💾 Auth state saved to:", bm.state_file)
        print("   Future runs will skip login automatically.")
        input("Press Enter to close the browser...")


@cli.command()
def inspect():
    """Open the timesheet page in a visible browser for manual inspection."""
    from bot.browser import BrowserManager
    from bot.config import load_config
    from bot.timesheet import TimesheetSummaryPage

    config = load_config()
    with BrowserManager(config) as bm:
        page = bm.page
        summary = TimesheetSummaryPage(page)
        summary.navigate()

        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            summary.navigate()

        print("🔍 Browser open for inspection. Check selectors with DevTools.")
        print(f"   Current URL: {page.url}")
        input("Press Enter to close the browser...")


if __name__ == "__main__":
    cli()
