"""Configuration loading: WCL API key + base URL from environment / .env file."""

from __future__ import annotations

import os
from dataclasses import dataclass

from dotenv import load_dotenv

DEFAULT_BASE_URL = "https://fresh.warcraftlogs.com/v1"


class ConfigError(Exception):
    """Raised when required configuration (e.g. API key) is missing."""


@dataclass(frozen=True)
class Config:
    api_key: str
    base_url: str = DEFAULT_BASE_URL


def get_config(api_key_override: str | None = None) -> Config:
    """Load configuration from .env / environment variables.

    `api_key_override` lets the CLI's `--api-key` flag take precedence over
    whatever is in the environment.
    """
    load_dotenv()

    api_key = api_key_override or os.environ.get("WCL_API_KEY")
    if not api_key:
        raise ConfigError(
            "WCL_API_KEY is not set. Copy .env.example to .env and fill in "
            "your Warcraft Logs v1 API key, or pass --api-key."
        )

    base_url = os.environ.get("WCL_BASE_URL", DEFAULT_BASE_URL)
    return Config(api_key=api_key, base_url=base_url)
