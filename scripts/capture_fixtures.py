"""One-off dev helper: hits the real WCL v1 API for a given report and saves
raw responses into fixtures/ so field-name assumptions in analyze.py can be
verified/corrected, and so tests have real (not synthetic) fixture data.

Usage:
    python scripts/capture_fixtures.py QgPvwcBdHX78zaxJ

Requires WCL_API_KEY in .env (see .env.example).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from log_analyzer.config import get_config
from log_analyzer.wcl_client import WCLClient

FIXTURES_DIR = Path(__file__).resolve().parent.parent / "fixtures"


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python scripts/capture_fixtures.py <report_code>", file=sys.stderr)
        sys.exit(1)

    code = sys.argv[1]
    config = get_config()
    client = WCLClient(api_key=config.api_key, base_url=config.base_url)

    print(f"Fetching fights for {code}...")
    fights = client.get_report_fights(code)
    (FIXTURES_DIR / f"fights_{code}.json").write_text(
        json.dumps(fights, indent=2), encoding="utf-8"
    )
    print(json.dumps(fights, indent=2)[:2000])

    fights_list = fights.get("fights", [])
    start = fights.get("start", min((f["start_time"] for f in fights_list), default=0))
    end = fights.get("end", max((f["end_time"] for f in fights_list), default=0))

    print(f"\nFetching tables/summary for {code} ({start}-{end})...")
    table = client.get_table("summary", start, end, code)
    (FIXTURES_DIR / f"table_summary_{code}.json").write_text(
        json.dumps(table, indent=2), encoding="utf-8"
    )
    print(json.dumps(table, indent=2)[:2000])

    print(f"\nFetching first page of events for {code}...")
    events_iter = client.get_events(start, end, code)
    first_events = []
    for i, event in enumerate(events_iter):
        first_events.append(event)
        if i >= 99:
            break
    (FIXTURES_DIR / f"events_casts_{code}_page1.json").write_text(
        json.dumps({"events": first_events}, indent=2), encoding="utf-8"
    )
    print(json.dumps(first_events[:10], indent=2))

    print(f"\nSaved fixtures to {FIXTURES_DIR}")


if __name__ == "__main__":
    main()
