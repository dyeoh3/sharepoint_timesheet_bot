"""
SharePoint PWA Timesheet Page Objects.

Provides two page objects for automating the SharePoint Project Web App
timesheet workflow:

- ``TimesheetSummaryPage`` — interacts with MyTSSummary.aspx (the list of
  all timesheet periods).  Supports opening, selecting, and recalling
  timesheets.

- ``TimesheetEditPage`` — interacts with Timesheet.aspx (the JSGrid data-
  entry grid).  Fills Actual hours, clears Planned hours, saves, and
  submits via the SharePoint ribbon UI.

Key implementation detail:
    The JSGrid stores work-duration values internally in units of
    **1/1000th of a minute** — i.e. ``hours × 60,000``.  For example
    8 hours = ``480,000``.  All writes go through
    ``grid.UpdateProperties()`` with ``SP.JsGrid.CreateValidatedPropertyUpdate``
    to reliably track changes via the grid's diff pipeline.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

from playwright.sync_api import Locator, Page

from bot.config import get_sharepoint_urls
from bot.holidays import get_holidays_in_range


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
RECALLABLE_STATUSES = {"submitted", "approved"}


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
        # SharePoint may redirect after creating a new timesheet;
        # wait for the network to settle so the JSGrid is ready.
        try:
            self.page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass  # best-effort — fall through and let retry logic handle it

        return status

    def select_timesheet_row(self, target_monday: date | None = None) -> tuple[Locator, str]:
        """
        Select (click) a timesheet row on the summary page so that
        ribbon buttons like *Recall* become enabled.

        Returns:
            (row_locator, status_string)

        Raises:
            RuntimeError: If the row cannot be found.
        """
        if target_monday is None:
            target_monday = get_current_week_range()[0]

        result = self.find_row_for_week(target_monday)
        if result is None:
            raise RuntimeError(
                f"Could not find a timesheet period for week starting "
                f"{target_monday.strftime('%d/%m/%Y')}."
            )

        row, status = result

        # Click the row to select it (click the first <td> that isn't a link
        # to avoid accidentally opening the timesheet).
        cells = row.locator("td")
        clicked = False
        for i in range(cells.count()):
            cell = cells.nth(i)
            # Skip cells that contain links (clicking those would navigate)
            if cell.locator("a").count() > 0:
                continue
            try:
                cell.click(timeout=2000)
                clicked = True
                break
            except Exception:
                continue

        if not clicked:
            # Fallback — click the row itself
            row.click()

        self.page.wait_for_timeout(1000)
        print(f"   🖱️  Selected timesheet row (status: {status})")
        return row, status

    def recall(self, target_monday: date | None = None):
        """
        Recall a submitted/approved timesheet so it can be edited again.

        The flow on the My Timesheets summary page is:
        1. Click the timesheet row to select it.
        2. Click the **Recall** button in the ribbon.
        3. Confirm in the dialog if one appears.

        After recall the timesheet status changes back to **In Progress**.

        Args:
            target_monday: The Monday of the week to recall.

        Raises:
            RuntimeError: If the timesheet is not in a recallable state.
        """
        if target_monday is None:
            target_monday = get_current_week_range()[0]

        # Ensure we're on the summary page
        if "MyTSSummary" not in self.page.url:
            self.navigate()
            self.page.wait_for_timeout(2000)

        # 1. Find and validate the row
        result = self.find_row_for_week(target_monday)
        if result is None:
            raise RuntimeError(
                f"Could not find a timesheet period for week starting "
                f"{target_monday.strftime('%d/%m/%Y')}."
            )

        _, status = result
        status_lower = status.lower()

        if status_lower not in RECALLABLE_STATUSES:
            raise RuntimeError(
                f"Timesheet for week {target_monday.strftime('%d/%m/%Y')} "
                f"has status '{status}' — only Submitted or Approved "
                f"timesheets can be recalled."
            )

        # 2. Select the row (click it so ribbon buttons are enabled)
        self.select_timesheet_row(target_monday)

        # 3. Click the Recall button in the ribbon
        #    SharePoint PWA uses a TIMESHEET contextual tab on the summary
        #    page with a Recall button.
        print(f"   📥 Recalling timesheet for {target_monday.strftime('%d/%m/%Y')}...")

        # Try known ribbon ID patterns for the Recall button
        recall_btn = None
        recall_selectors = [
            "a[id*='Recall']",
            "a[id*='recall']",
            "span:has-text('Recall')",
            "a:has(span:has-text('Recall'))",
            "a:has(img[alt*='Recall'])",
        ]

        for sel in recall_selectors:
            loc = self.page.locator(sel).first
            try:
                if loc.is_visible(timeout=1000):
                    recall_btn = loc
                    break
            except Exception:
                continue

        if recall_btn is None:
            # The ribbon tab may not be active — try activating TIMESHEET tab
            tab_selectors = [
                "li#Ribbon\\.ContextualTabs\\.TiedMode\\.Home-title a",
                "a[title*='Timesheet']",
                "a[title*='TIMESHEET']",
            ]
            for sel in tab_selectors:
                try:
                    tab = self.page.locator(sel).first
                    if tab.is_visible(timeout=1000):
                        tab.click()
                        self.page.wait_for_timeout(1000)
                        break
                except Exception:
                    continue

            # Re-select the row (tab click may deselect)
            self.select_timesheet_row(target_monday)
            self.page.wait_for_timeout(500)

            # Retry finding the Recall button
            for sel in recall_selectors:
                loc = self.page.locator(sel).first
                try:
                    if loc.is_visible(timeout=1000):
                        recall_btn = loc
                        break
                except Exception:
                    continue

        if recall_btn is None:
            # Last resort — dump all ribbon button IDs for debugging
            buttons = self.page.evaluate("""() => {
                let btns = document.querySelectorAll('a[id*="Ribbon"]');
                return Array.from(btns).map(b => ({
                    id: b.id,
                    text: b.innerText?.substring(0, 50),
                    visible: b.offsetParent !== null
                }));
            }""")
            visible = [b for b in buttons if b.get("visible")]
            print("   🔍 Visible ribbon buttons:")
            for b in visible:
                print(f"      - {b['id']}: {b.get('text', '')}")
            raise RuntimeError("Could not find the Recall button in the ribbon")

        # 4. Set up a dialog handler BEFORE clicking — Recall triggers
        #    a native window.confirm() ("Are you sure you want to recall…")
        dialog_handled = False

        def _handle_dialog(dialog):
            nonlocal dialog_handled
            print(f"   🗨️  Dialog: {dialog.message}")
            dialog.accept()
            dialog_handled = True

        self.page.on("dialog", _handle_dialog)

        recall_btn.click()
        # Give the dialog a moment to fire and be accepted
        self.page.wait_for_timeout(3000)
        self.page.remove_listener("dialog", _handle_dialog)
        print("   🖱️  Clicked Recall button")

        if dialog_handled:
            print("   ✅ Confirmed recall in dialog")

        # Wait for page to settle
        self.page.wait_for_load_state("load")
        try:
            self.page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        self.page.wait_for_timeout(2000)

        # 5. Verify the recall succeeded
        self.navigate()
        self.page.wait_for_timeout(2000)
        result = self.find_row_for_week(target_monday)
        if result:
            _, new_status = result
            new_lower = new_status.lower()
            if new_lower in EDITABLE_STATUSES:
                print(f"   ✅ Timesheet recalled (status: {new_status})")
            else:
                print(f"   ⚠️  Status after recall: {new_status}")
        else:
            print("   ⚠️  Could not verify recall — row not found")

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

    # ----- JSGrid Controller helpers --------------------------------------

    def _get_controller_name(self) -> str:
        """Return the global JS variable name for the JSGrid controller.

        Retries a few times to handle cases where the execution context
        is destroyed by a late navigation (e.g. SharePoint redirect after
        creating a new timesheet).
        """
        from playwright._impl._errors import Error as PlaywrightError

        for attempt in range(5):
            try:
                self.page.wait_for_load_state("load")
                name = self.page.evaluate("""() => {
                    for (let key in window) {
                        if (key.includes('JSGridController')) return key;
                    }
                    return null;
                }""")
                if name:
                    return name
            except PlaywrightError:
                pass  # context destroyed — page is still navigating
            self.page.wait_for_timeout(2000)

        raise RuntimeError("Could not find JSGridController on the page")

    def _find_record_key(self, task_name: str) -> str | None:
        """
        Find the JSGrid record key for a task by matching against
        the cached assignment name (``TS_LINE_CACHED_ASSIGN_NAME``).
        """
        ctrl = self._get_controller_name()
        target = task_name.lower().strip()
        return self.page.evaluate(f"""() => {{
            let ctrl = window['{ctrl}'];
            let grid = ctrl._jsGridControl;
            let count = grid.GetViewRecordCount();
            let target = {repr(target)};

            for (let i = 0; i < count; i++) {{
                let key = grid.GetRecordKeyByViewIndex(i);
                let rec = grid.GetRecord(key);
                if (!rec) continue;

                // Match on cached assignment name (most reliable)
                let assignProp = rec.properties?.['TS_LINE_CACHED_ASSIGN_NAME'];
                if (assignProp) {{
                    for (let k in assignProp) {{
                        let v = assignProp[k];
                        if (typeof v === 'string' && v.toLowerCase().includes(target)) {{
                            return key;
                        }}
                    }}
                }}
            }}
            return null;
        }}""")

    # ----- Fill hours (via JSGrid JS API) ---------------------------------

    def _actual_row_index(self, left_index: int) -> int:
        """Convert a left-pane row index to the right-pane *Actual* row index."""
        return (left_index - 1) * 2 + 1

    def fill_hours_for_task(
        self,
        left_index: int,
        hours_per_day: float,
        work_days: list[str],
        task_name: str = "",
        period_start: date | None = None,
        region: str = "NSW",
    ):
        """
        Fill daily Actual hours for a single task row.

        Uses the JSGrid ``UpdateProperties`` API to batch-write values
        directly to the grid's data model — bypassing the unreliable
        click-and-type approach.

        Data values are stored in **1/1000th of a minute** units
        (``hours × 60,000``).  For example 8h = 480,000.

        Field keys: ``TPD_col{N}a`` where N = 0..6 (Mon..Sun), suffix
        ``a`` = Actual.

        Public holidays are automatically skipped and any pre-existing
        value on those days is cleared to ``0h``.

        Args:
            left_index:    1-based row index in the left pane.
            hours_per_day: Hours to fill per work day.
            work_days:     Day names to fill.
            task_name:     Task name for record-key lookup.
            period_start:  Monday of the target week.
            region:        Australian state code for holiday detection.
        """
        hours_str = (
            f"{int(hours_per_day)}h"
            if hours_per_day == int(hours_per_day)
            else f"{hours_per_day}h"
        )

        # Determine the Monday of the target week
        if period_start is None:
            period_start = get_current_week_range()[0]

        # Build the set of holiday dates for the week
        period_end = period_start + timedelta(days=6)
        holiday_dates = get_holidays_in_range(period_start, period_end, state=region)

        # Find the record key via task_name or left_index correlation
        record_key = None
        if task_name:
            record_key = self._find_record_key(task_name)

        if record_key is None:
            # Fallback: correlate left_index → viewIndex
            ctrl = self._get_controller_name()
            record_key = self.page.evaluate(f"""() => {{
                let grid = window['{ctrl}']._jsGridControl;
                let viewIdx = {left_index - 1};
                if (viewIdx >= 0 && viewIdx < grid.GetViewRecordCount()) {{
                    return grid.GetRecordKeyByViewIndex(viewIdx);
                }}
                return null;
            }}""")

        if record_key is None:
            print(f"   ❌ Could not resolve record key for left row {left_index}")
            return

        # Map work_days to TPD column indices (0-based: Mon=0 .. Sun=6)
        day_index_map = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2,
            "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6,
        }
        target_indices = sorted(day_index_map[d] for d in work_days if d in day_index_map)

        ctrl = self._get_controller_name()
        filled = 0
        cleared = 0
        skipped_holiday = 0

        # Collect all property updates to batch via UpdateProperties
        updates_js_parts: list[str] = []

        for day_idx in target_indices:
            field_key = f"TPD_col{day_idx}a"
            day_date = period_start + timedelta(days=day_idx)

            # --- Holiday check ---
            if day_date in holiday_dates:
                hol_name = holiday_dates[day_date]
                print(f"   🏖️  {day_date.strftime('%A %d/%m')} is \"{hol_name}\" — skipping")
                skipped_holiday += 1

                # Clear any existing value on this holiday
                current = self.page.evaluate(f"""() => {{
                    let grid = window['{ctrl}']._jsGridControl;
                    let rec = grid.GetRecord('{record_key}');
                    if (!rec) return null;
                    try {{ return rec.GetLocalizedValue('{field_key}'); }}
                    catch(e) {{ return null; }}
                }}""")
                if current and current not in ("", "0h", "0", None):
                    updates_js_parts.append(
                        f"SP.JsGrid.CreateValidatedPropertyUpdate('{record_key}', '{field_key}', 0, '0h')"
                    )
                    cleared += 1
                    print(f"       ↳ Clearing existing \"{current}\" → 0h")
                continue

            # --- Read current value ---
            current = self.page.evaluate(f"""() => {{
                let grid = window['{ctrl}']._jsGridControl;
                let rec = grid.GetRecord('{record_key}');
                if (!rec) return null;
                try {{ return rec.GetLocalizedValue('{field_key}'); }}
                catch(e) {{ return null; }}
            }}""")

            if current == hours_str:
                # Already correct — no need to rewrite
                filled += 1
                continue

            if current and current not in ("", "0h", "0", None):
                print(f"   🔄 {day_date.strftime('%A')}: overwriting \"{current}\" → \"{hours_str}\"")

            # Queue update — data value in 1/1000th of a minute (hours × 60,000)
            ms_value = int(hours_per_day * 60_000)
            updates_js_parts.append(
                f"SP.JsGrid.CreateValidatedPropertyUpdate('{record_key}', '{field_key}', {ms_value}, '{hours_str}')"
            )
            filled += 1

        # Apply all updates in a single UpdateProperties call
        if updates_js_parts:
            updates_array = ",\n                ".join(updates_js_parts)
            self.page.evaluate(f"""() => {{
                let ctrl = window['{ctrl}'];
                let grid = ctrl._jsGridControl;
                let updates = [
                    {updates_array}
                ];
                let changeKey = grid.UpdateProperties(updates, null, null);
                if (changeKey) {{
                    ctrl.notifyWritePending(changeKey);
                }}
                grid.RefreshAllRows();
            }}""")
            self.page.wait_for_timeout(500)

        parts = [f"✅ Filled {hours_str} in {filled}/{len(target_indices) - skipped_holiday} day cells"]
        if skipped_holiday:
            parts.append(f"{skipped_holiday} holiday(s) skipped")
        if cleared:
            parts.append(f"{cleared} cleared")
        print(f"   {', '.join(parts)}")

        # Verify values were written
        verify_indices = [i for i in target_indices if (period_start + timedelta(days=i)) not in holiday_dates]
        self._verify_fill(record_key, verify_indices, hours_str)

    def _verify_fill(self, record_key: str, day_indices: list[int], expected: str):
        """Read back the TPD_col values to confirm they were written."""
        ctrl = self._get_controller_name()
        mismatches = []
        for day_idx in day_indices:
            field_key = f"TPD_col{day_idx}a"
            actual = self.page.evaluate(f"""() => {{
                let grid = window['{ctrl}']._jsGridControl;
                let rec = grid.GetRecord('{record_key}');
                if (!rec) return null;
                try {{ return rec.GetLocalizedValue('{field_key}'); }}
                catch(e) {{ return null; }}
            }}""")
            if actual != expected:
                mismatches.append((day_idx, actual))

        if mismatches:
            for day_idx, actual in mismatches:
                print(f"   ⚠️  Verify: day {day_idx} has \"{actual}\" (expected \"{expected}\")")
        else:
            print(f"   ✅ Verified: all {len(day_indices)} day cells read back \"{expected}\"")

    def _fill_from_planned(
        self,
        record_key: str,
        planned: dict[int, str],
        work_days: list[str],
        task_name: str = "",
        period_start: date | None = None,
        region: str = "NSW",
    ):
        """
        Copy Planned values as Actual for each work day.

        *planned* is a dict mapping day index → localized value string
        (from ``_read_planned_values``).  Days not in *planned* or that
        fall on a public holiday are cleared to 0.
        """
        if period_start is None:
            period_start = get_current_week_range()[0]

        period_end = period_start + timedelta(days=6)
        holiday_dates = get_holidays_in_range(period_start, period_end, state=region)

        day_index_map = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2,
            "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6,
        }
        target_indices = sorted(day_index_map[d] for d in work_days if d in day_index_map)

        ctrl = self._get_controller_name()
        filled = 0
        cleared = 0
        skipped_holiday = 0

        # Collect updates for batch UpdateProperties call
        updates_js_parts: list[str] = []

        for day_idx in target_indices:
            field_key = f"TPD_col{day_idx}a"
            day_date = period_start + timedelta(days=day_idx)

            # Holiday check
            if day_date in holiday_dates:
                hol_name = holiday_dates[day_date]
                print(f"   🏖️  {day_date.strftime('%A %d/%m')} is \"{hol_name}\" — skipping")
                skipped_holiday += 1
                # Clear any existing Actual on this holiday
                current = self.page.evaluate(f"""() => {{
                    let grid = window['{ctrl}']._jsGridControl;
                    let rec = grid.GetRecord('{record_key}');
                    if (!rec) return null;
                    try {{ return rec.GetLocalizedValue('{field_key}'); }}
                    catch(e) {{ return null; }}
                }}""")
                if current and current not in ("", "0h", "0", None):
                    updates_js_parts.append(
                        f"SP.JsGrid.CreateValidatedPropertyUpdate('{record_key}', '{field_key}', 0, '0h')"
                    )
                    cleared += 1
                    print(f"       ↳ Clearing existing \"{current}\" → 0h")
                continue

            # Get the planned value for this day
            planned_val = planned.get(day_idx)
            if not planned_val:
                continue

            # Read current Actual
            current = self.page.evaluate(f"""() => {{
                let grid = window['{ctrl}']._jsGridControl;
                let rec = grid.GetRecord('{record_key}');
                if (!rec) return null;
                try {{ return rec.GetLocalizedValue('{field_key}'); }}
                catch(e) {{ return null; }}
            }}""")

            if current == planned_val:
                filled += 1
                continue

            if current and current not in ("", "0h", "0", None):
                print(f"   🔄 {day_date.strftime('%A')}: overwriting \"{current}\" → \"{planned_val}\"")
            else:
                print(f"   📋 {day_date.strftime('%A')}: Planned → Actual \"{planned_val}\"")

            # Parse hours from planned_val (e.g. "8h", "8.24h") → 1/1000th min
            try:
                hours_num = float(planned_val.replace("h", ""))
                ms_value = int(hours_num * 60_000)
            except (ValueError, AttributeError):
                ms_value = 0
            updates_js_parts.append(
                f"SP.JsGrid.CreateValidatedPropertyUpdate('{record_key}', '{field_key}', {ms_value}, '{planned_val}')"
            )
            filled += 1

        # Apply all updates
        if updates_js_parts:
            updates_array = ",\n                ".join(updates_js_parts)
            self.page.evaluate(f"""() => {{
                let ctrl = window['{ctrl}'];
                let grid = ctrl._jsGridControl;
                let updates = [
                    {updates_array}
                ];
                let changeKey = grid.UpdateProperties(updates, null, null);
                if (changeKey) {{
                    ctrl.notifyWritePending(changeKey);
                }}
                grid.RefreshAllRows();
            }}""")
            self.page.wait_for_timeout(500)

        total_workdays = len(target_indices) - skipped_holiday
        parts = [f"✅ Filled Planned→Actual in {filled}/{total_workdays} day cells"]
        if skipped_holiday:
            parts.append(f"{skipped_holiday} holiday(s) skipped")
        if cleared:
            parts.append(f"{cleared} cleared")
        print(f"   {', '.join(parts)}")

    # ----- Read Planned values -----------------------------------------

    def _read_planned_values(self, record_key: str) -> dict[int, str]:
        """
        Read the server-side Planned values (``TPD_col{N}p``) for a record.

        Returns a dict mapping day index (0–6) to the localized value
        string (e.g. ``'8.24h'``).  Days with no Planned value are
        omitted from the dict.
        """
        ctrl = self._get_controller_name()
        raw = self.page.evaluate(f"""() => {{
            let grid = window['{ctrl}']._jsGridControl;
            let rec = grid.GetRecord('{record_key}');
            if (!rec) return {{}};
            let result = {{}};
            for (let c = 0; c < 7; c++) {{
                try {{
                    let v = rec.GetLocalizedValue('TPD_col' + c + 'p');
                    if (v && v !== '0h' && v !== '0') result[c] = v;
                }} catch(e) {{}}
            }}
            return result;
        }}""")
        return {int(k): v for k, v in raw.items()} if raw else {}

    # ----- Clear Planned values -----------------------------------------

    def _clear_planned_hours(
        self,
        record_key: str,
        work_days: list[str],
        task_name: str = "",
        period_start: date | None = None,
    ):
        """
        Zero-out Planned values (``TPD_col{N}p``) for a task row.

        Uses ``grid.UpdateProperties`` with
        ``SP.JsGrid.CreateValidatedPropertyUpdate`` to reliably clear
        planned cells via the grid's change-tracking pipeline.
        """
        if period_start is None:
            period_start = get_current_week_range()[0]

        day_index_map = {
            "Monday": 0, "Tuesday": 1, "Wednesday": 2,
            "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6,
        }
        target_indices = sorted(day_index_map[d] for d in work_days if d in day_index_map)

        ctrl = self._get_controller_name()
        updates_js_parts: list[str] = []
        cleared = 0

        for day_idx in target_indices:
            field_key = f"TPD_col{day_idx}p"

            # Read current planned value
            current = self.page.evaluate(f"""() => {{
                let grid = window['{ctrl}']._jsGridControl;
                let rec = grid.GetRecord('{record_key}');
                if (!rec) return null;
                try {{ return rec.GetLocalizedValue('{field_key}'); }}
                catch(e) {{ return null; }}
            }}""")

            if not current or current in ("", "0h", "0", None):
                continue

            updates_js_parts.append(
                f"SP.JsGrid.CreateValidatedPropertyUpdate('{record_key}', '{field_key}', 0, '0h')"
            )

            day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
            print(f"   🗑️  Planned {day_names[day_idx]}: \"{current}\" → 0h")
            cleared += 1

        if updates_js_parts:
            updates_array = ",\n                ".join(updates_js_parts)
            self.page.evaluate(f"""() => {{
                let ctrl = window['{ctrl}'];
                let grid = ctrl._jsGridControl;
                let updates = [
                    {updates_array}
                ];
                let changeKey = grid.UpdateProperties(updates, null, null);
                if (changeKey) {{
                    ctrl.notifyWritePending(changeKey);
                }}
                grid.RefreshAllRows();
            }}""")
            self.page.wait_for_timeout(500)
            print(f"   ✅ Cleared Planned values in {cleared} cells")

            # Verify the clear
            remaining = self._read_planned_values(record_key)
            work_remaining = {k: v for k, v in remaining.items() if k in target_indices}
            if work_remaining:
                summary = ", ".join(
                    f"{['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][k]}={v}"
                    for k, v in sorted(work_remaining.items())
                )
                print(f"   ⚠️  Some Planned values may not have cleared: {summary}")
            else:
                print(f"   ✅ Verified: all Planned values are now 0")
        else:
            print(f"   ✅ Planned hours already clear for {task_name}")

    # ----- Clear tasks not in config ------------------------------------

    def _clear_non_config_tasks(self, config_task_names: list[str]):
        """
        Zero-out **Actual** and **Planned** hours for any grid task whose
        name is NOT in *config_task_names*.  The config is the single
        source of truth — only listed tasks should have filled values.
        """
        ctrl = self._get_controller_name()
        # Collect all tasks with their Actual and Planned values
        tasks = self.page.evaluate(f"""() => {{
            let ctrl = window['{ctrl}'];
            let grid = ctrl._jsGridControl;
            let count = grid.GetViewRecordCount();
            let result = [];
            for (let i = 0; i < count; i++) {{
                let key = grid.GetRecordKeyByViewIndex(i);
                let rec = grid.GetRecord(key);
                if (!rec) continue;
                let name = '';
                let assignProp = rec.properties?.['TS_LINE_CACHED_ASSIGN_NAME'];
                if (assignProp) {{
                    for (let k in assignProp) {{
                        if (typeof assignProp[k] === 'string') {{ name = assignProp[k]; break; }}
                    }}
                }}
                if (!name) continue;  // skip summary/total rows
                let actual = [];
                let planned = [];
                for (let c = 0; c < 7; c++) {{
                    try {{ actual.push(rec.GetLocalizedValue('TPD_col' + c + 'a') || ''); }}
                    catch(e) {{ actual.push(''); }}
                    try {{ planned.push(rec.GetLocalizedValue('TPD_col' + c + 'p') || ''); }}
                    catch(e) {{ planned.push(''); }}
                }}
                result.push({{ key: key, name: name, actual: actual, planned: planned }});
            }}
            return result;
        }}""")

        config_names_lower = [n.lower().strip() for n in config_task_names]
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

        # Collect ALL property updates across all non-config tasks
        all_updates_js_parts: list[str] = []

        for task in tasks:
            task_name = task["name"]
            task_lower = task_name.lower()
            is_configured = any(
                cfg in task_lower or task_lower in cfg
                for cfg in config_names_lower
            )
            if is_configured:
                continue

            # Find non-zero Actual values
            non_zero_actual = [
                (i, v) for i, v in enumerate(task["actual"])
                if v and v not in ("", "0h", "0")
            ]

            # Find non-zero Planned values
            non_zero_planned = [
                (i, v) for i, v in enumerate(task["planned"])
                if v and v not in ("", "0h", "0")
            ]

            if not non_zero_actual and not non_zero_planned:
                continue

            print(f"\n🧹 Clearing non-config task: {task_name}")

            for col_idx, old_val in non_zero_actual:
                all_updates_js_parts.append(
                    f"SP.JsGrid.CreateValidatedPropertyUpdate('{task['key']}', 'TPD_col{col_idx}a', 0, '0h')"
                )
                print(f"   🗑️  Actual  {day_names[col_idx]}: \"{old_val}\" → 0h")

            for col_idx, old_val in non_zero_planned:
                all_updates_js_parts.append(
                    f"SP.JsGrid.CreateValidatedPropertyUpdate('{task['key']}', 'TPD_col{col_idx}p', 0, '0h')"
                )
                print(f"   🗑️  Planned {day_names[col_idx]}: \"{old_val}\" → 0h")

        if all_updates_js_parts:
            updates_array = ",\n                ".join(all_updates_js_parts)
            self.page.evaluate(f"""() => {{
                let ctrl = window['{ctrl}'];
                let grid = ctrl._jsGridControl;
                let updates = [
                    {updates_array}
                ];
                let changeKey = grid.UpdateProperties(updates, null, null);
                if (changeKey) {{
                    ctrl.notifyWritePending(changeKey);
                }}
                grid.RefreshAllRows();
            }}""")
            self.page.wait_for_timeout(500)
        else:
            print("\n✅ No non-config tasks to clear")

    # ----- High-level fill from config -----------------------------------

    def fill_week_from_config(
        self,
        projects: list[dict],
        work_days: list[str],
        region: str = "NSW",
        period_start: date | None = None,
    ):
        """
        Fill an entire week of hours using the project config.

        For each project in *projects*:

        1. Clear Actual & Planned hours from any grid task **not** in the
           config (single source of truth).
        2. Locate the task row in the grid (or add it via "From Existing
           Assignments").
        3. Fill daily Actual hours — either a fixed ``default_hours_per_day``
           or by copying server Planned values (``use_planned: true``).
        4. Optionally zero-out Planned values (``clear_planned: true``).

        Public holidays (based on *region*) are automatically skipped.

        Args:
            projects:     List of project dicts from ``config.yaml``.
            work_days:    Day names to fill, e.g. ``["Monday", …, "Friday"]``.
            region:       Australian state code for holiday detection.
            period_start: Monday of the target week (defaults to current).
        """
        if period_start is None:
            period_start = get_current_week_range()[0]

        # Show holiday info for the week
        period_end = period_start + timedelta(days=6)
        from bot.holidays import get_holidays_in_range as _get_hols
        hols = _get_hols(period_start, period_end, state=region)
        if hols:
            print(f"\n📅 Public holidays this week ({region}):")
            for d, name in sorted(hols.items()):
                print(f"   🏖️  {d.strftime('%A %d/%m/%Y')}: {name}")

        # Clear hours from any tasks NOT in the config
        config_task_names = [p["name"] for p in projects]
        self._clear_non_config_tasks(config_task_names)

        for project in projects:
            name = project["name"]
            use_planned = project.get("use_planned", False)
            clear_planned = project.get("clear_planned", False)
            hours = project.get("default_hours_per_day", 0)

            if not use_planned and hours <= 0 and not clear_planned:
                print(f"   ⏭️  Skipping {name} (0 hours, use_planned=false)")
                continue

            mode = "Planned → Actual" if use_planned else f"{hours}h/day"
            if clear_planned:
                mode += " + clear Planned"
            print(f"\n📝 Processing: {name} ({mode})")

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
            if use_planned:
                record_key = self._find_record_key(name)
                if record_key is None:
                    print(f"   ❌ Could not resolve record key for \"{name}\"")
                    continue
                planned = self._read_planned_values(record_key)
                if not planned:
                    print(f"   ⚠️  No Planned values found — nothing to copy")
                    continue
                planned_summary = ", ".join(
                    f"{['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][k]}={v}"
                    for k, v in sorted(planned.items())
                )
                print(f"   📋 Planned: {planned_summary}")
                self._fill_from_planned(
                    record_key, planned, work_days,
                    task_name=name,
                    period_start=period_start,
                    region=region,
                )
            else:
                self.fill_hours_for_task(
                    left_index, hours, work_days,
                    task_name=name,
                    period_start=period_start,
                    region=region,
                )

            # Clear Planned hours if configured
            if clear_planned:
                rec_key = self._find_record_key(name) if not use_planned else record_key
                if rec_key:
                    print(f"\n🧹 Clearing Planned hours for: {name}")
                    self._clear_planned_hours(
                        rec_key, work_days,
                        task_name=name,
                        period_start=period_start,
                    )
                else:
                    print(f"   ⚠️  Could not find record key for {name} — skipping planned clear")

    # ----- Save / Submit -------------------------------------------------

    def save(self, force: bool = False):
        """Click the Save button in the TIMESHEET ribbon.

        Args:
            force: Always click Save even if IsDirty() returns False.
                   Useful after JS API writes which may not set the
                   dirty flag.
        """
        if not force:
            try:
                ctrl = self._get_controller_name()
                is_dirty = self.page.evaluate(f"() => window['{ctrl}'].IsDirty()")
                if not is_dirty:
                    print("💾 No unsaved changes — skipping save")
                    return
            except Exception:
                pass  # proceed with save anyway

        save_btn = self.page.locator(
            f"a[id='{self._SAVE_BTN_ID}']"
        )
        # If the Save button isn't visible, activate the tab first
        if not save_btn.is_visible():
            self._activate_timesheet_tab()
        save_btn.click()
        self.page.wait_for_load_state("load")
        try:
            self.page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        self.page.wait_for_timeout(3000)

        # Verify save completed — retry if still dirty
        ctrl = self._get_controller_name()
        for attempt in range(3):
            try:
                is_dirty = self.page.evaluate(f"() => window['{ctrl}'].IsDirty()")
                if not is_dirty:
                    print("💾 Timesheet saved successfully")
                    return
            except Exception:
                pass
            if attempt < 2:
                self.page.wait_for_timeout(2000)

        # Still dirty after retries — try clicking save again
        print("   ⏳ Still dirty, retrying save...")
        try:
            save_btn.click()
            self.page.wait_for_load_state("load")
            self.page.wait_for_timeout(4000)
            is_dirty = self.page.evaluate(f"() => window['{ctrl}'].IsDirty()")
            if not is_dirty:
                print("💾 Timesheet saved successfully (on retry)")
            else:
                print("⚠️  Timesheet may not have saved (still dirty after retry)")
        except Exception:
            print("💾 Timesheet saved (could not verify)")

    def submit(self):
        """
        Submit the timesheet via:
        TIMESHEET tab → Send → Turn in Final Timesheet → OK.
        """
        # 1. Activate the TIMESHEET ribbon tab
        self._activate_timesheet_tab()
        self.page.wait_for_timeout(500)

        # 2. Click the "Send" menu button (known ribbon ID)
        send_btn_id = "Ribbon.ContextualTabs.TiedMode.Home.Sheet.SubmitMenu-Large"
        send_btn = self.page.locator(f"a[id='{send_btn_id}']")
        if not send_btn.is_visible():
            send_btn = self.page.locator("a[id*='SubmitMenu']").first
        print("   📤 Clicking Send...")
        send_btn.click()
        self.page.wait_for_timeout(1500)

        # 3. Click "Turn in Final Timesheet" from the dropdown menu
        print("   📤 Clicking 'Turn in Final Timesheet'...")
        # The menu item span has class ms-cui-ctl-mediumlabel; the clickable
        # element is the nearest ancestor <a> with class ms-cui-ctl.
        turn_in = self.page.locator(
            "a.ms-cui-ctl:has(span:has-text('Turn in Final Timesheet'))"
        ).first
        try:
            turn_in.wait_for(state="visible", timeout=3000)
            turn_in.click()
        except Exception:
            # Fallback: click the span directly
            try:
                span = self.page.locator(
                    "span:has-text('Turn in Final Timesheet')"
                ).first
                span.click(timeout=5000)
            except Exception:
                # Last resort: try re-clicking Send and then the menu item
                print("   ⚠️  Menu item not found, retrying Send...")
                send_btn.click()
                self.page.wait_for_timeout(2000)
                self.page.locator(
                    "a.ms-cui-ctl:has(span:has-text('Turn in Final Timesheet')), "
                    "span:has-text('Turn in Final Timesheet')"
                ).first.click(timeout=5000)

        # Wait for the dialog iframe to appear
        self.page.wait_for_timeout(3000)

        # 4. Click OK in the confirmation dialog (SharePoint modal iframe)
        print("   📤 Confirming submission...")
        ok_clicked = False

        # Wait and retry — the dialog iframe may take a moment to load
        for attempt in range(6):
            frames = self.page.frames
            if attempt == 0:
                print(f"   🔍 Found {len(frames)} frame(s):")
                for f in frames:
                    print(f"      - name=\"{f.name}\" url=\"{f.url[:80]}\"")

            # Try matching the SubmitTSDlg iframe specifically
            for frame in frames:
                if "SubmitTSDlg" in frame.url or "DlgFrame" in frame.name:
                    try:
                        ok_btn = frame.locator("input[value='OK']").first
                        ok_btn.wait_for(state="visible", timeout=3000)
                        ok_btn.click()
                        ok_clicked = True
                        print(f"   ✅ Clicked OK in frame: {frame.name}")
                        break
                    except Exception:
                        continue
            if ok_clicked:
                break

            # Fallback: try any frame with a visible OK input
            for frame in frames:
                try:
                    ok_btn = frame.locator("input[value='OK']").first
                    if ok_btn.is_visible():
                        ok_btn.click()
                        ok_clicked = True
                        print(f"   ✅ Clicked OK in frame (fallback): {frame.name}")
                        break
                except Exception:
                    continue
            if ok_clicked:
                break

            # Also try the ms-dlgContent overlay approach — click OK via JS
            try:
                ok_found = self.page.evaluate("""() => {
                    // Search all iframes for OK button
                    let iframes = document.querySelectorAll('iframe');
                    for (let iframe of iframes) {
                        try {
                            let doc = iframe.contentDocument || iframe.contentWindow.document;
                            let btn = doc.querySelector("input[value='OK']");
                            if (btn) { btn.click(); return true; }
                        } catch(e) {}
                    }
                    return false;
                }""")
                if ok_found:
                    ok_clicked = True
                    print("   ✅ Clicked OK via JS iframe traversal")
                    break
            except Exception:
                pass

            self.page.wait_for_timeout(1000)

        if not ok_clicked:
            print("   ⚠️  Could not find OK button in dialog")

        self.page.wait_for_load_state("load")
        self.page.wait_for_timeout(3000)
        print("🚀 Timesheet submitted")
