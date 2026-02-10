"""
Smoke test — opens Chrome, navigates to the SharePoint timesheet URL,
handles manual login if needed, and checks that key page elements exist.

Run with: python scripts/test_open_site.py
"""

import sys
from pathlib import Path

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import get_sharepoint_urls, load_config


def check_element(page, description: str, selector: str, timeout: int = 5000) -> bool:
    """Check if an element exists on the page and report the result."""
    try:
        loc = page.locator(selector)
        loc.first.wait_for(state="visible", timeout=timeout)
        count = loc.count()
        text = loc.first.inner_text(timeout=2000)[:80] if count > 0 else ""
        print(f"   ✅ {description}: found ({count} match{'es' if count > 1 else ''}) → \"{text}\"")
        return True
    except Exception:
        print(f"   ❌ {description}: NOT found  (selector: {selector})")
        return False


def open_website():
    """Open Chrome, navigate to the timesheet URL, verify elements exist."""
    _, timesheet_url = get_sharepoint_urls()
    config = load_config()

    print(f"🌐 Opening: {timesheet_url}\n")

    with BrowserManager(config) as bm:
        page = bm.page
        page.goto(timesheet_url, wait_until="domcontentloaded")

        # Handle login — if we have a saved session, this will wait
        # for the auto-redirect; otherwise it prompts for manual login.
        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            # Navigate again after login
            page.goto(timesheet_url, wait_until="domcontentloaded")

        page.wait_for_load_state("networkidle")

        print(f"📄 Page loaded!")
        print(f"   Title: {page.title()}")
        print(f"   URL:   {page.url}\n")

        # ----- Element discovery -----------------------------------------
        print("🔍 Checking for page elements...\n")

        # Dump the full page HTML tag structure for debugging
        # Common SharePoint PWA timesheet elements to look for
        selectors = [
            ("Page heading / title",            "h1, h2, .ms-core-pageTitle, [role='heading']"),
            ("Timesheet grid / table",          "table, [role='grid'], .ms-listviewtable, .pwa-grid"),
            ("Any input fields",                "input:visible"),
            ("Save button",                     "input[value='Save'], button:has-text('Save'), a:has-text('Save')"),
            ("Submit button",                   "input[value='Submit'], button:has-text('Submit'), a:has-text('Submit')"),
            ("Create timesheet link",           "a:has-text('Create'), a:has-text('Click here')"),
            ("Timesheet period rows",           "table tr, [role='row']"),
            ("Navigation / breadcrumbs",        "nav, .ms-breadcrumb, .breadcrumb, [role='navigation']"),
            ("Any links on page",               "a[href]:visible"),
            ("Any buttons on page",             "button:visible, input[type='button']:visible, input[type='submit']:visible"),
        ]

        found = 0
        total = len(selectors)
        for desc, sel in selectors:
            if check_element(page, desc, sel):
                found += 1

        print(f"\n📊 Results: {found}/{total} element checks passed")

        # Dump all visible text for debugging selectors
        print("\n" + "=" * 60)
        print("📝 Page snapshot (all visible text, first 2000 chars):")
        print("=" * 60)
        try:
            body_text = page.locator("body").inner_text(timeout=5000)
            print(body_text[:2000])
        except Exception as e:
            print(f"   Could not get page text: {e}")
        print("=" * 60)

        print("\n🔗 The browser is still open — inspect elements with DevTools (Cmd+Opt+I)")
        input("   Press Enter to close the browser...")

    print("👋 Browser closed.")


if __name__ == "__main__":
    open_website()
