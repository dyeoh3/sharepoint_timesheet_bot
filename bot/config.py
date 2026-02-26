"""
Configuration loader — reads config.yaml and .env
"""

import os
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

# Load .env from project root
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(_PROJECT_ROOT / ".env")


def load_config(config_path: str | None = None) -> dict[str, Any]:
    """Load and return the YAML configuration dictionary."""
    path = Path(config_path) if config_path else _PROJECT_ROOT / "config.yaml"
    with open(path, "r") as f:
        return yaml.safe_load(f)


def get_sharepoint_urls() -> tuple[str, str]:
    """Return (base_url, timesheet_url) from environment variables."""
    base_url = os.getenv("SHAREPOINT_BASE_URL")
    timesheet_url = os.getenv("SHAREPOINT_TIMESHEET_URL")
    if not base_url or not timesheet_url:
        raise EnvironmentError(
            "SHAREPOINT_BASE_URL and SHAREPOINT_TIMESHEET_URL must be set in .env file. "
            "Copy .env.example to .env and fill in your URLs."
        )
    return base_url, timesheet_url


def is_headless() -> bool:
    """Check if the browser should run headless."""
    return os.getenv("HEADLESS", "false").lower() == "true"
