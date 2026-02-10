"""
SharePoint Timesheet Page Object
Encapsulates all interactions with the SharePoint PWA Timesheet pages.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

from playwright.sync_api import Locator, Page

from bot.config import get_sharepoint_urls


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_current_week_range() -> tuple[date, date]:
    """Return (Monday, Sunday) of the current week."""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday


def _parse_period_dates(text: str) -> tuple[date, date] | None:
    """
    Extract start/end dates from a period string like
    '29 (20/07/2026 - 26/07/2026)' or 'WC - 1005 (10/05/2026 - 16/05/2026)'.
    """
    match = re.search(r"\((\d{1,2}/\d{2}/\d{4})\s*-\s*(\d{1,2}/\d{2}/\d{4})\)", text)
    if not match:
        return None
    try:
        parts_s = match.group(1).split("/")
        parts_e = match.group(2).split("/")
        start = date(int(parts_s[2]), int(parts_s[1]), int(parts_s[0]))
        end = date(int(parts_e[2]), int(parts_e[1]), int(parts_e[0]))
        return start, end
    except (ValueError, IndexError):
        return None


# Status values a timesheet can have on the summary page.
EDITABLE_STATUSES = {"not yet created", "in progress"}


# ---------------------------------------------------------------------------
# Page Objects
# ---------------------------------------------------------------------------

class TimesheetSummaryPage:
    """Interactions with MyTSSummary.aspx — the timesheet list page."""

    def __init__(self, page: Page):
        self.page = page
        _, self.url = get_sharepoint_urls()

    def navigate(self):
        """Go to the timesheet summary page."""
        self.page.goto(self.url, wait_until="domcontentloaded")
        self.page.wait_for_load_state("load")

    # ----- Row parsing ---------------------------------------------------

    def _get_data_rows(self) -> list[Locator]:
        """Return all <tr> elements that contain at least 3 <td> cells."""
        rows = self.page.locator("table tr")
        result = []
        for i in range(rows.count()):
            row = rows.nth(i)
            if row.locator("td").count() >= 3:
                result.append(row)
        return result

    @staticmethod
    def _row_cell_texts(row: Locator, max_cells: int = 6) -> list[str]:
        """Extract the text of up to *max_cells* <td> cells in a row."""
        cells = row.locator("td")
        count = min(cells.count(), max_cells)
        return [cells.nth(j).inner_text(timeout=2000).strip() for j in range(count)]

    @staticmethod
    def _extract_status(cell_texts: list[str]) -> str:
        """Find the status value among the cell texts."""
        known = {
            "not yet created", "in progress", "approved",
            "rejected", "submitted", "period closed",
        }
        for ct in cell_texts:
            if ct.strip().lower() in known:
                return ct.strip()
        return "Unknown"

    @staticmethod
    def _extract_period(cell_texts: list[str]) -> tuple[date, date] | None:
        """Find the first cell that contains a date range and parse it."""
        for ct in cell_texts:
            dates = _parse_period_dates(ct)
            if dates:
                return dates
        return None

    # ----- Public API ----------------------------------------------------

    def find_row_for_week(self, target_monday: date | None = None) -> tuple[Locator, str] | None:
        """
        Locate the table row whose period contains *target_monday*.

        Args:
            target_monday: The Monday of the desired week.
                           Defaults to the current week's Monday.

        Returns:
            (row_locator, status_string) or None if not found.
        """
        if target_monday is None:
            target_monday = get_current_week_range()[0]

        for row in self._get_data_rows():
            texts = self._row_cell_texts(row)
            dates = self._extract_period(texts)
            if dates is None:
                continue
            start, end = dates
            if start <= target_monday <= end:
                return row, self._extract_status(texts)
        return None

    def open_timesheet(self, target_monday: date | None = None) -> str:
        """
        Open (or create) the timesheet for the week containing *target_monday*.

        Only timesheets with status **Not Yet Created** or **In Progress**
        can be opened.  Other statuses will raise a ``RuntimeError``.

        Args:
            target_monday: The Monday of the desired week (defaults to current).

        Returns:
            The status of the row that was opened.
        """
        if target_monday is None:
            target_monday = get_current_week_range()[0]

        result = self.find_row_for_week(target_monday)
        if result is None:
            raise RuntimeError(
                f"Could not find a timesheet period for week starting "
                f"{target_monday.strftime('%d/%m/%Y')}. "
                f"It may not be visible in the current view."
            )

        row, status = result
        status_lower = status.lower()

        if status_lower not in EDITABLE_STATUSES:
            raise RuntimeError(
                f"Timesheet for week {target_monday.strftime('%d/%m/%Y')} "
                f"has status '{status}' — only 'Not Yet Created' or "
                f"'In Progress' timesheets can be opened."
            )

        # Click the first link in the row — this is either
        # "Click to Create" or "My Timesheet".
        link = row.locator("a").first
        link_text = link.inner_text(timeout=2000)
        print(f"🖱️  Clicking: \"{link_text}\" (status: {status})")
        link.click()
        self.page.wait_for_load_state("load")

        return status

    def get_all_periods(self) -> list[dict]:
        """
        Return a summary list of every visible timesheet period.

        Each dict has keys: name, period, start, end, status.
        """
        periods = []
        for row in self._get_data_rows():
            texts = self._row_cell_texts(row)
            dates = self._extract_period(texts)
            if dates is None:
                continue
            start, end = dates
            periods.append({
                "name": texts[0] if texts else "",
                "period": next((t for t in texts if re.search(r"\d{1,2}/\d{2}/\d{4}", t)), ""),
                "start": start,
                "end": end,
                "status": self._extract_status(texts),
            })
        return periods


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
        self.page.wait_for_load_state("load")

    def submit(self):
        """Click the Submit button to send the timesheet for approval."""
        self.page.locator("text=Submit").click()
        # Handle confirmation dialog if present
        try:
            self.page.locator("text=OK").click(timeout=3000)
        except Exception:
            pass
        self.page.wait_for_load_state("load")


def get_current_week_range() -> tuple[date, date]:
    """Return (Monday, Sunday) of the current week."""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday
