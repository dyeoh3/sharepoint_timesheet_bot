"""
SharePoint Timesheet Page Object
Encapsulates all interactions with the SharePoint PWA Timesheet pages.

NOTE: The selectors below are best-guess starting points for SharePoint PWA.
You will almost certainly need to inspect the actual page and adjust selectors
once you run the bot for the first time with HEADLESS=false.
"""

from __future__ import annotations

from datetime import date, timedelta

from playwright.sync_api import Page


class TimesheetSummaryPage:
    """Interactions with MyTSSummary.aspx — the timesheet list page."""

    URL = "https://lionco.sharepoint.com/sites/ITProjects/_layouts/15/pwa/Timesheet/MyTSSummary.aspx"

    def __init__(self, page: Page):
        self.page = page

    def navigate(self):
        """Go to the timesheet summary page."""
        self.page.goto(self.URL, wait_until="networkidle")

    def get_current_period_status(self) -> str | None:
        """Read the status of the current timesheet period (e.g., 'Not Created', 'In Progress')."""
        # TODO: inspect the actual page and pin down the correct selector
        status_cell = self.page.locator("table.ms-listviewtable tr.ms-itmHover td:nth-child(3)")
        if status_cell.count() > 0:
            return status_cell.first.inner_text()
        return None

    def click_current_period(self):
        """Click into the current (most recent) timesheet period to edit it."""
        # TODO: adjust selector — typically the first row link in the timesheet list
        self.page.locator("table.ms-listviewtable tr.ms-itmHover td a").first.click()
        self.page.wait_for_load_state("networkidle")

    def create_timesheet_if_needed(self):
        """Click 'Create' if no timesheet exists for the current period."""
        create_btn = self.page.locator("text=Click here to create")
        if create_btn.is_visible(timeout=3000):
            create_btn.click()
            self.page.wait_for_load_state("networkidle")


class TimesheetEditPage:
    """Interactions with the timesheet data-entry grid."""

    def __init__(self, page: Page):
        self.page = page

    def get_project_rows(self) -> list[str]:
        """Return the list of project/task names visible in the timesheet grid."""
        # TODO: adjust selector after inspecting actual grid
        rows = self.page.locator("table.pwa-grid tr td.pwa-taskname")
        return [rows.nth(i).inner_text() for i in range(rows.count())]

    def fill_hours(self, row_index: int, day_index: int, hours: float):
        """
        Fill in hours for a specific row (project) and day column.

        Args:
            row_index: 0-based index of the project row.
            day_index: 0-based index of the day column (0=Mon, 4=Fri).
            hours: Number of hours to enter.
        """
        # TODO: the actual grid input selectors will vary —
        #       inspect the page with HEADLESS=false to get the right ones
        cell_selector = f"table.pwa-grid tr:nth-child({row_index + 2}) td.pwa-day:nth-child({day_index + 2}) input"
        cell = self.page.locator(cell_selector)
        cell.fill(str(hours))

    def fill_week_from_config(self, projects: list[dict], work_days: list[str]):
        """
        Fill an entire week of hours using the config defaults.

        Args:
            projects: List of project dicts from config (name, default_hours_per_day).
            work_days: List of day names to fill (e.g., ["Monday", ..., "Friday"]).
        """
        day_map = {"Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3, "Friday": 4}
        for row_idx, project in enumerate(projects):
            hours = project.get("default_hours_per_day", 0)
            if hours <= 0:
                continue
            for day_name in work_days:
                day_idx = day_map.get(day_name)
                if day_idx is not None:
                    self.fill_hours(row_idx, day_idx, hours)

    def save(self):
        """Click the Save button."""
        self.page.locator("text=Save").click()
        self.page.wait_for_load_state("networkidle")

    def submit(self):
        """Click the Submit button to send the timesheet for approval."""
        self.page.locator("text=Submit").click()
        # Handle confirmation dialog if present
        try:
            self.page.locator("text=OK").click(timeout=3000)
        except Exception:
            pass
        self.page.wait_for_load_state("networkidle")


def get_current_week_range() -> tuple[date, date]:
    """Return (Monday, Friday) of the current week."""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    return monday, friday
