"""
Browser manager — handles Playwright browser lifecycle and Microsoft SSO auth.
"""

from pathlib import Path

from playwright.sync_api import BrowserContext, Page, sync_playwright

from bot.config import get_sharepoint_urls, is_headless, load_config


class BrowserManager:
    """Manages the Playwright browser instance and authentication state."""

    def __init__(self, config: dict | None = None):
        self.config = config or load_config()
        self._playwright = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

        browser_cfg = self.config.get("browser", {})
        self.slow_mo = browser_cfg.get("slow_mo", 100)
        self.timeout = browser_cfg.get("timeout", 60000)
        self.user_data_dir = Path(browser_cfg.get("user_data_dir", "browser_state/profile"))

    # -- Lifecycle --------------------------------------------------------

    def start(self) -> Page:
        """Launch the browser with a persistent profile and return the main page."""
        self._playwright = sync_playwright().start()
        self.user_data_dir.mkdir(parents=True, exist_ok=True)

        # Persistent context keeps the full browser profile (cookies,
        # localStorage, IndexedDB, service workers, etc.) across runs —
        # this is what makes Microsoft SSO sessions survive restarts.
        self._context = self._playwright.chromium.launch_persistent_context(
            user_data_dir=str(self.user_data_dir),
            headless=is_headless(),
            slow_mo=self.slow_mo,
            channel="chrome",
        )
        self._context.set_default_timeout(self.timeout)
        # Persistent contexts come with one page already; use it.
        self._page = self._context.pages[0] if self._context.pages else self._context.new_page()
        return self._page

    def stop(self):
        """Close the browser (profile is auto-persisted)."""
        if self._context:
            self._context.close()
        if self._playwright:
            self._playwright.stop()

    # -- Authentication ---------------------------------------------------

    def has_valid_session(self) -> bool:
        """Check if a persistent browser profile exists with prior data."""
        return self.user_data_dir.exists() and any(self.user_data_dir.iterdir())

    def is_on_login_page(self, page: Page) -> bool:
        """Check if the current page is a Microsoft SSO / auth page."""
        auth_hosts = (
            "login.microsoftonline.com",
            "login.live.com",
            "login.windows.net",
            "mysignins.microsoft.com",
            "aadcdn.msauth.net",
            "device.login.microsoftonline.com",
        )
        return any(host in page.url for host in auth_hosts)

    def wait_for_manual_login(self, page: Page, timeout: int = 300_000):
        """
        Wait for the user to complete Microsoft SSO login manually.
        Handles MFA, Authenticator prompts, conditional access — anything.

        Args:
            page: The Playwright page currently showing the login form.
            timeout: Max time to wait in ms (default 5 minutes).
        """
        if not self.is_on_login_page(page):
            return  # already authenticated

        base_url, _ = get_sharepoint_urls()
        # Extract the SharePoint hostname so we detect arrival on *any*
        # SharePoint page, not just the exact timesheet URL.
        from urllib.parse import urlparse

        sp_host = urlparse(base_url).hostname  # e.g. "lionco.sharepoint.com"

        # If we have a saved profile, the login page may auto-redirect.
        # Give it time before asking the user to log in manually.
        if self.has_valid_session():
            print("\n🔄 Saved session found — waiting for auto-redirect...")
            try:
                page.wait_for_url(lambda url: sp_host in url, timeout=15_000)
                page.wait_for_load_state("networkidle")
                print("✅ Auto-logged in via saved session!\n")
                return
            except Exception:
                print("⚠️  Saved session expired — manual login required.")

        print("\n🔐 Microsoft login detected!")
        print("   Please log in manually in the browser window.")
        print("   (Handle MFA / Authenticator as needed)")
        print("   Waiting for you to reach SharePoint...\n")

        # Wait until the URL actually lands on the SharePoint domain.
        # Simply checking "not login.microsoftonline.com" fires too early
        # on intermediate auth redirects (aadcdn, mysignins, etc.).
        page.wait_for_url(
            lambda url: sp_host in url,
            timeout=timeout,
        )

        # Give the page a moment to fully load after redirect
        page.wait_for_load_state("networkidle")
        print(f"✅ Logged in! Now at: {page.url}")
        print("   Session will be saved for future runs.\n")

    # -- Context manager --------------------------------------------------

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *_):
        self.stop()

    @property
    def page(self) -> Page:
        if self._page is None:
            raise RuntimeError("Browser not started. Call start() first.")
        return self._page
