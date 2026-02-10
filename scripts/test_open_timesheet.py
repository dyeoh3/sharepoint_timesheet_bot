"""
Test script — opens the timesheet summary page, finds the current week's
timesheet row, and opens it (creates it if needed).

Run with: python scripts/test_open_timesheet.py
"""

import re
import sys
from datetime import date, timedelta
from pathlib import Path

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import get_sharepoint_urls, load_config


def get_current_week_monday() -> date:
    """Return the Monday of the current week."""
    today = date.today()
    return today - timedelta(days=today.weekday())


def parse_period_dates(period_text: str) -> tuple[date, date] | None:
    """
    Parse a period string like '29 (20/07/2026 - 26/07/2026)' or
    'WC - 1005 (10/05/2026 - 16/05/2026)' and return (start_date, end_date).
    """
    match = re.search(r"\((\d{1,2}/\d{2}/\d{4})\s*-\s*(\d{1,2}/\d{2}/\d{4})\)", period_text)
    if not match:
        return None
    try:
        start = date(
            int(match.group(1).split("/")[2]),
            int(match.group(1).split("/")[1]),
            int(match.group(1).split("/")[0]),
        )
        end = date(
            int(match.group(2).split("/")[2]),
            int(match.group(2).split("/")[1]),
            int(match.group(2).split("/")[0]),
        )
        return start, end
    except (ValueError, IndexError):
        return None


def open_timesheet():
    """Navigate to the summary page, find this week's row, and open it."""
    _, timesheet_url = get_sharepoint_urls()
    config = load_config()

    print(f"🌐 Opening: {timesheet_url}\n")

    with BrowserManager(config) as bm:
        page = bm.page
        page.goto(timesheet_url, wait_until="domcontentloaded")

        # Handle login if needed
        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            page.goto(timesheet_url, wait_until="domcontentloaded")

        page.wait_for_load_state("load")

        print(f"📄 Page loaded: {page.title()}")
        print(f"   URL: {page.url}\n")

        # --- Dump table structure for debugging ---
        print("🔍 Inspecting timesheet table rows...\n")

        # Get all table rows (skip header)
        rows = page.locator("table tr")
        row_count = rows.count()
        print(f"   Found {row_count} table rows total\n")

        # Parse each row to find columns
        current_monday = get_current_week_monday()
        print(f"📅 Looking for week starting: {current_monday.strftime('%d/%m/%Y')}\n")

        target_row = None
        target_status = None
        target_name = None

        for i in range(row_count):
            row = rows.nth(i)
            cells = row.locator("td")
            cell_count = cells.count()

            if cell_count < 3:
                continue  # skip header or empty rows

            # Extract cell text
            cell_texts = []
            for j in range(min(cell_count, 6)):
                try:
                    text = cells.nth(j).inner_text(timeout=1000).strip()
                    cell_texts.append(text)
                except Exception:
                    cell_texts.append("")

            # Try to find the period column with date range
            row_text = " | ".join(cell_texts)
            period_text = ""
            for ct in cell_texts:
                if re.search(r"\d{1,2}/\d{2}/\d{4}", ct):
                    period_text = ct
                    break

            if not period_text:
                continue

            dates = parse_period_dates(period_text)
            if not dates:
                continue

            start_date, end_date = dates

            # Determine status from row text
            status = "Unknown"
            for ct in cell_texts:
                ct_lower = ct.strip().lower()
                if ct_lower in ("not yet created", "in progress", "approved", "rejected", "submitted"):
                    status = ct.strip()
                    break

            # Determine name
            name = cell_texts[0] if cell_texts else ""

            # Check if this is the current week
            is_current = start_date <= current_monday <= end_date

            marker = " 👈 CURRENT WEEK" if is_current else ""
            print(f"   Row {i:2d}: {name:<20s} | {period_text:<40s} | {status}{marker}")

            if is_current:
                target_row = row
                target_status = status
                target_name = name

        print()

        if target_row is None:
            print("❌ Could not find the current week's timesheet period!")
            print("   The current week may not be visible in the current view.")
            input("\n   Press Enter to close the browser...")
            return

        print(f"✅ Found current week: {target_name} | Status: {target_status}\n")

        if target_status.lower() == "approved":
            print("⚠️  This timesheet is already approved — cannot edit.")
            input("\n   Press Enter to close the browser...")
            return

        if target_status.lower() in ("rejected", "submitted"):
            print(f"⚠️  This timesheet has status '{target_status}' — cannot edit.")
            input("\n   Press Enter to close the browser...")
            return

        # Click to open or create
        link = target_row.locator("a").first
        link_text = link.inner_text(timeout=2000)
        print(f"🖱️  Clicking: \"{link_text}\"")
        link.click()

        page.wait_for_load_state("load")

        print(f"\n📄 Timesheet page loaded!")
        print(f"   Title: {page.title()}")
        print(f"   URL:   {page.url}\n")

        # Dump the timesheet edit page content
        print("=" * 60)
        print("📝 Timesheet page content (first 3000 chars):")
        print("=" * 60)
        try:
            body_text = page.locator("body").inner_text(timeout=5000)
            print(body_text[:3000])
        except Exception as e:
            print(f"   Could not get page text: {e}")
        print("=" * 60)

        print("\n🔗 Browser is still open — inspect with DevTools (Cmd+Opt+I)")
        input("   Press Enter to close the browser...")

    print("👋 Browser closed.")


if __name__ == "__main__":
    open_timesheet()
