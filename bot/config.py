"""
Configuration loader — reads config.yaml and .env
"""

import os
from pathlib import Path

import yaml
from dotenv import load_dotenv

# Load .env from project root
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(_PROJECT_ROOT / ".env")


def load_config(config_path: str | None = None) -> dict:
    """Load and return the YAML configuration dictionary."""
    path = Path(config_path) if config_path else _PROJECT_ROOT / "config.yaml"
    with open(path, "r") as f:
        return yaml.safe_load(f)


def get_credentials() -> tuple[str, str]:
    """Return (email, password) from environment variables."""
    email = os.getenv("MS_EMAIL")
    password = os.getenv("MS_PASSWORD")
    if not email or not password:
        raise EnvironmentError(
            "MS_EMAIL and MS_PASSWORD must be set in .env file. "
            "Copy .env.example to .env and fill in your credentials."
        )
    return email, password


def is_headless() -> bool:
    """Check if the browser should run headless."""
    return os.getenv("HEADLESS", "false").lower() == "true"
