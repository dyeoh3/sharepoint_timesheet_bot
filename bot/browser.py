"""
Browser manager — handles Playwright browser lifecycle and Microsoft SSO auth.
"""

from pathlib import Path

from playwright.sync_api import Browser, BrowserContext, Page, sync_playwright

from bot.config import get_credentials, is_headless, load_config


class BrowserManager:
    """Manages the Playwright browser instance and authentication state."""

    def __init__(self, config: dict | None = None):
        self.config = config or load_config()
        self._playwright = None
        self._browser: Browser | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

        browser_cfg = self.config.get("browser", {})
        self.slow_mo = browser_cfg.get("slow_mo", 100)
        self.timeout = browser_cfg.get("timeout", 60000)
        self.state_file = Path(browser_cfg.get("state_file", "browser_state/state.json"))

    # -- Lifecycle --------------------------------------------------------

    def start(self) -> Page:
        """Launch the browser and return the main page."""
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(
            headless=is_headless(),
            slow_mo=self.slow_mo,
            channel="chrome",  # use installed Chrome
        )
        self._context = self._load_or_create_context()
        self._context.set_default_timeout(self.timeout)
        self._page = self._context.new_page()
        return self._page

    def stop(self):
        """Save auth state and close everything."""
        if self._context:
            self._save_state()
            self._context.close()
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()

    # -- Auth state persistence -------------------------------------------

    def _load_or_create_context(self) -> BrowserContext:
        """Load saved auth cookies/state if available, else create fresh context."""
        if self.state_file.exists():
            return self._browser.new_context(storage_state=str(self.state_file))
        return self._browser.new_context()

    def _save_state(self):
        """Persist browser auth state for reuse."""
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        self._context.storage_state(path=str(self.state_file))

    # -- Microsoft SSO login ----------------------------------------------

    def login_if_needed(self, page: Page):
        """
        Detect the Microsoft login page and authenticate.
        If already logged in (via saved state), this is a no-op.
        """
        # Check if we've been redirected to Microsoft login
        if "login.microsoftonline.com" not in page.url:
            return  # already authenticated

        email, password = get_credentials()

        # Enter email
        page.locator('input[type="email"]').fill(email)
        page.locator('input[type="submit"]').click()

        # Wait for password page
        page.wait_for_selector('input[type="password"]', state="visible")
        page.locator('input[type="password"]').fill(password)
        page.locator('input[type="submit"]').click()

        # Handle "Stay signed in?" prompt
        try:
            page.locator("text=Yes").click(timeout=5000)
        except Exception:
            pass  # prompt may not appear

        # Wait for SharePoint to load
        page.wait_for_url("**/sharepoint.com/**", timeout=30000)

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
