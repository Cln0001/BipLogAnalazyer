"""Thin client for the Warcraft Logs v1 REST API (fresh.warcraftlogs.com/v1/docs/).

Centralizes api_key injection, timeouts, 429 backoff/retry, and error
surfacing. Pure HTTP/parsing concerns only — aggregation logic lives in
analyze.py.
"""

from __future__ import annotations

import re
import time
from collections.abc import Iterator
from typing import Any

import requests

REPORT_URL_RE = re.compile(r"/reports/([A-Za-z0-9]+)")
BARE_CODE_RE = re.compile(r"^[A-Za-z0-9]{8,20}$")

MAX_RETRIES = 5
DEFAULT_BACKOFF_SECONDS = 2.0
REQUEST_TIMEOUT_SECONDS = 30
MAX_EVENT_PAGES = 1000  # defensive cap against a malformed pagination loop


class WCLAPIError(Exception):
    """Raised on a non-2xx response from the WCL API, or retries exhausted."""

    def __init__(self, message: str, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code


class InvalidReportURLError(Exception):
    """Raised when a report URL/code can't be parsed."""


def extract_report_code(url_or_code: str) -> str:
    """Extracts the report code from a full WCL report URL, or passes through
    a bare code if it looks like one.
    """
    match = REPORT_URL_RE.search(url_or_code)
    if match:
        return match.group(1)
    if BARE_CODE_RE.match(url_or_code):
        return url_or_code
    raise InvalidReportURLError(
        f"Could not extract a report code from: {url_or_code!r}"
    )


class WCLClient:
    def __init__(self, api_key: str, base_url: str):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()

    def _get(self, path: str, **params: Any) -> Any:
        url = f"{self.base_url}{path}"
        query = {"api_key": self.api_key, **params}

        attempt = 0
        while True:
            attempt += 1
            response = self.session.get(url, params=query, timeout=REQUEST_TIMEOUT_SECONDS)

            if response.status_code == 429:
                if attempt >= MAX_RETRIES:
                    raise WCLAPIError(
                        "Rate limited (429) and retries exhausted. Try again later.",
                        status_code=429,
                    )
                retry_after = response.headers.get("Retry-After")
                delay = float(retry_after) if retry_after else DEFAULT_BACKOFF_SECONDS * attempt
                print(f"  rate limited (429) on {path}, retrying in {delay:.0f}s (attempt {attempt}/{MAX_RETRIES})...")
                time.sleep(delay)
                continue

            if not response.ok:
                snippet = response.text[:200]
                raise WCLAPIError(
                    f"WCL API request to {path} failed with status "
                    f"{response.status_code}: {snippet}",
                    status_code=response.status_code,
                )

            return response.json()

    def get_report_fights(self, code: str) -> dict:
        """GET /report/fights/{code} -> fight list + friendlies (players/pets).

        Note: the top-level "start"/"end" in this response are absolute
        Unix epoch ms (real wall-clock time of the report) — NOT the same
        scale as fights[].start_time/end_time, which are ms offsets
        *relative to report start*. The tables/events endpoints below expect
        the relative scale, so always derive the analysis window from
        fights[].start_time/end_time, never from the top-level start/end.
        """
        return self._get(f"/report/fights/{code}")

    def get_guild_reports(
        self,
        guild_name: str,
        server_name: str,
        server_region: str,
        start: int | None = None,
        end: int | None = None,
    ) -> list[dict]:
        """GET /reports/guild/{guild}/{server}/{region}?start=&end= -> list of
        report summaries ({id, title, owner, start, end, zone}) for the
        guild. `start`/`end` here are absolute Unix epoch ms (same scale as
        get_report_fights' top-level "start"/"end", NOT the relative scale
        used by /report/tables and /report/events) and scope the results to
        a date range — useful for limiting to one phase.

        Path confirmed live (2026-06-21): note the plural "reports" — unlike
        every other endpoint here ("report/fights", "report/tables", ...),
        this one doesn't follow that singular pattern.
        """
        params: dict[str, Any] = {}
        if start is not None:
            params["start"] = start
        if end is not None:
            params["end"] = end
        return self._get(f"/reports/guild/{guild_name}/{server_name}/{server_region}", **params)

    def get_table(self, view: str, code: str, start: int | None = None, end: int | None = None, **params: Any) -> dict:
        """GET /report/tables/{view}/{code}?start=&end=... -> aggregated per-player table.

        `view` e.g. "damage-done", "healing", "casts", "deaths". `start`/`end`
        are relative ms offsets (see get_report_fights docstring); omitting
        them returns full-report aggregates for some views, but per-player
        breakdowns need them explicitly.
        """
        if start is not None:
            params["start"] = start
        if end is not None:
            params["end"] = end
        return self._get(f"/report/tables/{view}/{code}", **params)

    def get_buff_uptime(self, code: str, start: int, end: int, ability_id: int) -> list[dict]:
        """GET /report/tables/buffs/{code}?start=&end=&abilityid=<id> -> per-
        player uptime ("auras" list) for one specific buff ability.

        Confirmed live: passing a *single* ability id switches the response
        from a guild-wide aggregate (one entry per buff seen, "id"/"name"
        meaning the ability) to a per-player breakdown (one entry per
        player who had it, "id"/"name" now meaning the player, "totalUptime"
        in ms, same scale as fights[].start_time/end_time). A comma-joined
        list of ids does NOT filter — it silently falls back to the
        unfiltered aggregate — so callers needing several ids (e.g. several
        ranks of the same scroll) must call this once per id and sum.
        """
        return self._get(f"/report/tables/buffs/{code}", start=start, end=end, abilityid=ability_id).get("auras", [])

    def get_events(self, code: str, start: int, end: int, **params: Any) -> Iterator[dict]:
        """GET /report/events/{code}?start=&end=&filter=..., paginated.

        Yields each raw event dict across all pages, following
        `nextPageTimestamp` until it's absent/null or the defensive page cap
        is hit. `start`/`end` are relative ms offsets, matching
        fights[].start_time/end_time.
        """
        page_start = start
        pages_fetched = 0

        while True:
            pages_fetched += 1
            if pages_fetched > MAX_EVENT_PAGES:
                raise WCLAPIError(
                    f"Aborting after {MAX_EVENT_PAGES} event pages — possible "
                    "pagination loop."
                )

            data = self._get(f"/report/events/{code}", start=page_start, end=end, **params)
            for event in data.get("events", []):
                yield event

            next_page = data.get("nextPageTimestamp")
            if not next_page or next_page <= page_start:
                break
            page_start = next_page
