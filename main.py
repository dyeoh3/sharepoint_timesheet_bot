"""
CLI entry point for the SharePoint Timesheet Bot.
"""

import os
import subprocess
from pathlib import Path

import click

from bot.runner import run_timesheet_bot

PROJECT_DIR = Path(__file__).resolve().parent
PLIST_NAME = "com.darren.timesheet-bot.plist"
PLIST_SRC = PROJECT_DIR / PLIST_NAME
PLIST_DST = Path.home() / "Library" / "LaunchAgents" / PLIST_NAME
LOG_DIR = PROJECT_DIR / "logs"


@click.group()
@click.version_option(version="0.3.2")
def cli():
    """SharePoint Timesheet Bot — automate your weekly timesheet."""
    pass


# ---------------------------------------------------------------------------
# Timesheet commands
# ---------------------------------------------------------------------------


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
        print("💾 Auth state saved to:", bm.user_data_dir)
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


# ---------------------------------------------------------------------------
# Schedule commands — manage the macOS LaunchAgent
# ---------------------------------------------------------------------------


@cli.group()
def schedule():
    """Manage the macOS LaunchAgent for scheduled runs."""
    pass


@schedule.command("install")
def schedule_install():
    """Install the LaunchAgent to run every Friday at 9 AM."""
    if not PLIST_SRC.exists():
        raise click.ClickException(f"Plist not found: {PLIST_SRC}")

    LOG_DIR.mkdir(parents=True, exist_ok=True)

    # Make wrapper script executable
    wrapper = PROJECT_DIR / "scripts" / "run_timesheet.sh"
    if wrapper.exists():
        wrapper.chmod(0o755)

    # Unload existing agent if present
    if PLIST_DST.exists():
        subprocess.run(["launchctl", "unload", str(PLIST_DST)],
                        capture_output=True)

    # Create symlink
    PLIST_DST.unlink(missing_ok=True)
    PLIST_DST.symlink_to(PLIST_SRC)
    click.echo(f"🔗 Symlinked: {PLIST_DST} → {PLIST_SRC}")

    # Load the agent
    result = subprocess.run(["launchctl", "load", str(PLIST_DST)],
                            capture_output=True, text=True)
    if result.returncode != 0:
        raise click.ClickException(f"Failed to load agent: {result.stderr}")

    click.echo("✅ LaunchAgent installed — runs every Friday at 09:00")
    click.echo(f"📂 Logs: {LOG_DIR}/")


@schedule.command("uninstall")
def schedule_uninstall():
    """Uninstall the LaunchAgent and remove the symlink."""
    if PLIST_DST.exists():
        subprocess.run(["launchctl", "unload", str(PLIST_DST)],
                        capture_output=True)
        PLIST_DST.unlink(missing_ok=True)
        click.echo("✅ LaunchAgent uninstalled")
    else:
        click.echo("ℹ️  LaunchAgent was not installed")


@schedule.command("status")
def schedule_status():
    """Show whether the LaunchAgent is loaded and its schedule."""
    result = subprocess.run(
        ["launchctl", "list"],
        capture_output=True, text=True,
    )
    is_loaded = "timesheet-bot" in result.stdout

    click.echo("📋 LaunchAgent Status")
    click.echo(f"   Installed : {'Yes' if PLIST_DST.exists() else 'No'}")
    click.echo(f"   Loaded    : {'Yes' if is_loaded else 'No'}")
    click.echo(f"   Schedule  : Every Friday at 09:00")
    click.echo(f"   Plist     : {PLIST_DST}")
    click.echo(f"   Logs      : {LOG_DIR}/")

    # Show most recent run
    if LOG_DIR.exists():
        logs = sorted(LOG_DIR.glob("timesheet_*.log"), reverse=True)
        if logs:
            click.echo(f"   Last run  : {logs[0].name}")
        else:
            click.echo("   Last run  : (no runs yet)")


@schedule.command("logs")
@click.option("-n", "--lines", default=50, help="Number of lines to show")
@click.option("-f", "--follow", is_flag=True, help="Follow the log file (like tail -f)")
def schedule_logs(lines, follow):
    """Show recent log output from the last scheduled run."""
    if not LOG_DIR.exists():
        raise click.ClickException(f"Log directory not found: {LOG_DIR}")

    logs = sorted(LOG_DIR.glob("timesheet_*.log"), reverse=True)
    if not logs:
        raise click.ClickException("No log files found yet")

    latest = logs[0]
    click.echo(f"📄 {latest.name}")
    click.echo("=" * 60)

    if follow:
        os.execvp("tail", ["tail", "-f", str(latest)])
    else:
        result = subprocess.run(["tail", f"-{lines}", str(latest)],
                                capture_output=True, text=True)
        click.echo(result.stdout)


@schedule.command("run")
def schedule_run():
    """Manually trigger the scheduled wrapper script now."""
    wrapper = PROJECT_DIR / "scripts" / "run_timesheet.sh"
    if not wrapper.exists():
        raise click.ClickException(f"Wrapper script not found: {wrapper}")

    click.echo("🚀 Running timesheet bot (same as scheduled run)...")
    click.echo("")
    os.execvp("bash", ["bash", str(wrapper)])


if __name__ == "__main__":
    cli()
