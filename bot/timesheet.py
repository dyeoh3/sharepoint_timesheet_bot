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
    """
    Interactions with the timesheet data-entry grid (Timesheet.aspx).

    The grid is a SharePoint PWA JSGrid with **two panes**:
    - **Left pane** — task metadata (Task Name, Project, Comment, Billing Category).
      One row per task, plus a header row (row 0).
    - **Right pane** — day columns (Mon .. Fri) with **two sub-rows per task**:
      *Actual* (odd indices) and *Planned* (even indices), plus a header row (row 0).

    Mapping:  left row *i* (i ≥ 1)  →  right row ``(i-1)*2 + 1`` (Actual),
                                        right row ``(i-1)*2 + 2`` (Planned).

    Clicking a day cell spawns a **floating** ``<input class="jsgrid-control-editbox">``
    that is *not* a child of the cell — it's positioned absolutely on top.
    """

    # Known table IDs (stable across sessions, may differ per tenant)
    _LEFT_TABLE_SUFFIX = "TimesheetPartJSGridControl_leftpane_mainTable"
    _RIGHT_TABLE_SUFFIX = "TimesheetPartJSGridControl_rightpane_mainTable"

    # Ribbon element IDs (discovered via debug inspection)
    _TIMESHEET_TAB_SEL = "a[title*='Timesheet']"
    _ADD_ROW_BTN_ID = "Ribbon.ContextualTabs.TiedMode.Home.Tasks.AddLine-Large"
    _SAVE_BTN_ID = "Ribbon.ContextualTabs.TiedMode.Home.Sheet.Save-Large"

    def __init__(self, page: Page):
        self.page = page

    # ----- Locators for the two-pane tables ------------------------------

    def _left_table(self) -> Locator:
        """Locate the left-pane table (task names)."""
        return self.page.locator(f"table[id$='{self._LEFT_TABLE_SUFFIX}']")

    def _right_table(self) -> Locator:
        """Locate the right-pane table (hour cells)."""
        return self.page.locator(f"table[id$='{self._RIGHT_TABLE_SUFFIX}']")

    # ----- Row discovery -------------------------------------------------

    def get_task_rows(self) -> list[dict]:
        """
        Return every task visible in the left pane.

        Each dict: ``{"left_index": int, "name": str}``
        where *left_index* is the 0-based row index in the left-pane table
        (row 0 is always the header).
        """
        left = self._left_table()
        rows = left.locator("tr")
        count = rows.count()
        tasks: list[dict] = []

        for i in range(1, count):  # skip header row 0
            row = rows.nth(i)
            cells = row.locator("td")
            if cells.count() < 3:
                continue
            try:
                # Task Name is in cell index 2 (after indent + checkbox cells)
                name = cells.nth(2).inner_text(timeout=1000).strip()
            except Exception:
                continue
            if not name or name.lower() in ("task name/description",):
                continue
            tasks.append({"left_index": i, "name": name})

        return tasks

    def find_task_row_index(self, task_name: str) -> int | None:
        """
        Return the left-pane row index for a task whose name contains
        *task_name* (case-insensitive substring match), or None.
        """
        target = task_name.lower().strip()
        for item in self.get_task_rows():
            if target in item["name"].lower():
                print(f"   ✅ Found task in grid: \"{item['name']}\" (left row {item['left_index']})")
                return item["left_index"]
        return None

    # ----- Ensure the TIMESHEET ribbon tab is active ---------------------

    def _activate_timesheet_tab(self):
        """Click the TIMESHEET ribbon tab (next to BROWSE) if not already active."""
        # The tab <li> has a known ID; click the <a> inside it.
        tab_li = self.page.locator(
            "li#Ribbon\\.ContextualTabs\\.TiedMode\\.Home-title"
        )
        if tab_li.count() > 0:
            anchor = tab_li.locator("a").first
            # Only click if the tab is not already selected
            selected = tab_li.get_attribute("aria-selected") or ""
            if selected.lower() != "true":
                anchor.click()
                self.page.wait_for_timeout(500)
                print("   🖱️  Activated TIMESHEET ribbon tab")
        else:
            # Fallback: try clicking by title
            tab = self.page.locator(self._TIMESHEET_TAB_SEL).first
            tab.click()
            self.page.wait_for_timeout(500)
            print("   🖱️  Activated TIMESHEET ribbon tab (fallback)")

    # ----- Add row from existing assignments -----------------------------

    def add_row_from_existing_assignments(self, task_name: str):
        """
        Add a new task row via:  TIMESHEET tab → Add Row → From Existing
        Assignments → select task → OK.
        """
        print(f"\n➕ Adding row for \"{task_name}\" via existing assignments...")

        # 1. Click the TIMESHEET tab in the ribbon
        self._activate_timesheet_tab()

        # 2. Click "Add Row" button (id confirmed via debug)
        add_row = self.page.locator(f"a[id='{self._ADD_ROW_BTN_ID}']")
        add_row.click()
        self.page.wait_for_timeout(800)
        print("   🖱️  Clicked Add Row")

        # 3. A dropdown menu appears — pick "From Existing Assignments"
        #    SharePoint renders the menu items as <a> inside a popup <ul>.
        menu_item = self.page.locator(
            "a:has-text('From Existing Assignments'), "
            "span.ms-cui-ctl-mediumlabel:has-text('From Existing Assignments'), "
            "span:text('From Existing Assignments')"
        ).first
        menu_item.click()
        print("   🖱️  Clicked 'From Existing Assignments'")

        # 4. Wait for the popup dialog / iframe to appear
        self.page.wait_for_timeout(2000)

        # 5. Interact with the assignment dialog
        self._select_task_in_assignment_dialog(task_name)

    def _select_task_in_assignment_dialog(self, task_name: str):
        """
        Inside the "Add From Existing Assignments" dialog, expand tree
        nodes, select the matching task, and click OK.
        """
        # The dialog may be an iframe — try to find it
        dialog_frame = None
        for frame in self.page.frames:
            url = frame.url
            if any(kw in url for kw in ("AddAssignment", "AddTask", "Existing")):
                dialog_frame = frame
                print(f"   📦 Found dialog iframe: {url[:80]}")
                break

        ctx = dialog_frame if dialog_frame else self.page

        # Expand all collapsed tree nodes so the task becomes visible
        self._expand_all_tree_nodes(ctx)

        # Find the task element by text
        task_el = None

        # Strategy 1: regex text match
        text_matches = ctx.locator(f"text=/{re.escape(task_name)}/i")
        if text_matches.count() > 0:
            task_el = text_matches.first
            print(f"   ✅ Found task in dialog: \"{task_name}\"")

        # Strategy 2: scan labels/spans
        if task_el is None:
            target_lower = task_name.lower().strip()
            for tag in ("label", "span", "td", "a", "div"):
                elements = ctx.locator(tag)
                for j in range(elements.count()):
                    try:
                        txt = elements.nth(j).inner_text(timeout=300).strip()
                        if target_lower in txt.lower():
                            task_el = elements.nth(j)
                            print(f"   ✅ Found task in dialog (scan): \"{txt}\"")
                            break
                    except Exception:
                        continue
                if task_el is not None:
                    break

        if task_el is None:
            raise RuntimeError(
                f"Could not find \"{task_name}\" in the existing assignments dialog."
            )

        # Check the checkbox next to it, or click the element directly
        parent = task_el.locator("xpath=..")
        cb = parent.locator("input[type='checkbox']")
        if cb.count() > 0:
            if not cb.is_checked():
                cb.check()
                print("   ☑️  Checked task checkbox")
        else:
            task_el.click()
            print("   🖱️  Clicked task element")

        self.page.wait_for_timeout(500)

        # Click OK / Add
        ok = ctx.locator(
            "input[value='OK'], button:has-text('OK'), "
            "input[value='Add'], button:has-text('Add')"
        ).first
        ok.click()
        print("   🖱️  Clicked OK")

        self.page.wait_for_timeout(2000)
        self.page.wait_for_load_state("load")
        print("   ✅ Task added to timesheet grid")

    def _expand_all_tree_nodes(self, ctx):
        """Click all collapsed tree nodes in the dialog."""
        for _ in range(5):
            expandable = ctx.locator(
                "[aria-expanded='false'], "
                "img[alt*='expand' i], "
                "a[title*='expand' i]"
            )
            count = expandable.count()
            if count == 0:
                break
            for i in range(count):
                try:
                    expandable.nth(i).click(timeout=1000)
                    self.page.wait_for_timeout(300)
                except Exception:
                    continue

    # ----- Fill hours ----------------------------------------------------

    def _actual_row_index(self, left_index: int) -> int:
        """Convert a left-pane row index to the right-pane *Actual* row index."""
        return (left_index - 1) * 2 + 1

    def fill_hours_for_task(
        self,
        left_index: int,
        hours_per_day: float,
        work_days: list[str],
    ):
        """
        Fill daily hours into the *Actual* cells of the right-pane row
        that corresponds to left-pane row *left_index*.

        The right-pane Actual row has 5 cells:
           cell 0 = "Actual" label
           cells 1‥4 = day columns (Mon .. Thu/Fri depending on period)

        Clicking a day cell spawns a floating ``<input class="jsgrid-control-editbox">``
        on the page.  We click the cell, wait for the input to appear, type
        the hours value, then press Tab to commit.
        """
        hours_str = (
            f"{int(hours_per_day)}h"
            if hours_per_day == int(hours_per_day)
            else f"{hours_per_day}h"
        )

        actual_idx = self._actual_row_index(left_index)
        right_table = self._right_table()
        right_rows = right_table.locator("tr")

        if actual_idx >= right_rows.count():
            print(f"   ❌ Right-pane row {actual_idx} out of range ({right_rows.count()} rows)")
            return

        actual_row = right_rows.nth(actual_idx)
        day_cells = actual_row.locator("td[role='gridcell']")
        total_cells = day_cells.count()

        # cell 0 = "Actual" label, cells 1+ = day columns
        day_count = total_cells - 1  # number of day columns
        if day_count <= 0:
            print("   ❌ No day columns found in Actual row")
            return

        # Map work_days to 0-based day indices
        day_index_map = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2,
            "Thursday": 3, "Friday": 4,
        }
        target_indices = sorted(day_index_map[d] for d in work_days if d in day_index_map)

        filled = 0
        for day_idx in target_indices:
            cell_offset = day_idx + 1  # +1 because cell 0 is the "Actual" label
            if cell_offset >= total_cells:
                # This day column doesn't exist in the grid (e.g. short week)
                continue

            cell = day_cells.nth(cell_offset)

            # Check if already filled
            try:
                existing = cell.inner_text(timeout=500).strip()
                if existing and existing not in ("", "0h", "0"):
                    print(f"   ⏭️  Day {day_idx} already has \"{existing}\" — skipping")
                    continue
            except Exception:
                pass

            # Click the cell to activate the floating editbox
            cell.click()
            self.page.wait_for_timeout(300)

            # The floating input should now be the active element
            # It has class "jsgrid-control-editbox"
            editbox = self.page.locator("input.jsgrid-control-editbox:visible")
            if editbox.count() == 0:
                print(f"   ⚠️  Editbox not found after clicking day {day_idx} — retrying")
                cell.dblclick()
                self.page.wait_for_timeout(300)
                editbox = self.page.locator("input.jsgrid-control-editbox:visible")

            if editbox.count() > 0:
                editbox.first.fill("")
                editbox.first.fill(hours_str)
                editbox.first.press("Tab")
                self.page.wait_for_timeout(200)
                filled += 1
            else:
                print(f"   ❌ Could not activate editbox for day {day_idx}")

        print(f"   ✅ Filled {hours_str} in {filled}/{len(target_indices)} day cells")

    # ----- High-level fill from config -----------------------------------

    def fill_week_from_config(self, projects: list[dict], work_days: list[str]):
        """
        Fill an entire week of hours using the config defaults.

        For each project in *projects*:
        1. Check if its row already exists in the grid.
        2. If not found, add it via "From Existing Assignments".
        3. Fill the daily Actual hours.
        """
        for project in projects:
            name = project["name"]
            hours = project.get("default_hours_per_day", 0)
            if hours <= 0:
                print(f"   ⏭️  Skipping {name} (0 hours)")
                continue

            print(f"\n📝 Processing: {name} ({hours}h/day)")

            # 1. Check if task already exists in the grid
            left_index = self.find_task_row_index(name)

            # 2. If not found, add via existing assignments
            if left_index is None:
                print(f"   ℹ️  Task \"{name}\" not found in grid — adding...")
                self.add_row_from_existing_assignments(name)
                # Re-scan after adding
                left_index = self.find_task_row_index(name)
                if left_index is None:
                    print(f"   ❌ Could not find \"{name}\" even after adding — skipping")
                    continue

            # 3. Fill hours in the Actual row
            self.fill_hours_for_task(left_index, hours, work_days)

    # ----- Save / Submit -------------------------------------------------

    def save(self):
        """Click the Save button in the TIMESHEET ribbon."""
        save_btn = self.page.locator(
            f"a[id='{self._SAVE_BTN_ID}']"
        )
        # If the Save button isn't visible, activate the tab first
        if not save_btn.is_visible():
            self._activate_timesheet_tab()
        save_btn.click()
        self.page.wait_for_load_state("load")
        print("💾 Timesheet saved")

    def submit(self):
        """Click the Send / Submit button in the TIMESHEET ribbon."""
        submit_btn = self.page.locator("a[id*='Submit'], a[id*='Send']").first
        if not submit_btn.is_visible():
            self._activate_timesheet_tab()
        submit_btn.click()
        try:
            self.page.locator(
                "input[value='OK'], button:has-text('OK')"
            ).first.click(timeout=3000)
        except Exception:
            pass
        self.page.wait_for_load_state("load")
        print("🚀 Timesheet submitted")
