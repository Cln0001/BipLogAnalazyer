"""CLI entrypoint: fetch a WCL report (or every report in a guild's phase),
summarize per-player consumable usage, write to .xlsx. Prints only
status/error lines — no console table (results live in the spreadsheet).
"""

from __future__ import annotations

import argparse
import json
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path

from log_analyzer.analyze import (
    ROLE_UNKNOWN,
    PlayerSummary,
    build_player_index,
    classify_roles,
    count_consumable_casts,
    merge_player_summary,
    merge_role_dicts,
    phase_fight_windows,
)
from log_analyzer.config import ConfigError, get_config
from log_analyzer.consumables import CONSUMABLES, KNOWN_CONSUMABLE_IDS
from log_analyzer.excel_export import sheet_already_exists, write_report
from log_analyzer.json_export import write_report as write_json_report
from log_analyzer.wcl_client import InvalidReportURLError, WCLAPIError, WCLClient, extract_report_code

# Defaults so `--guild`/dates don't have to be retyped every run — override
# with --guild NAME SERVER REGION if ever needed for a different guild.
DEFAULT_GUILD = ("Bananas in Pyjamas", "Spineshatter", "EU")

PHASES_PATH = Path(__file__).resolve().parent.parent / "data" / "phases.json"
LOGS_DIR = Path(__file__).resolve().parent.parent / "Logs"


class _Tee:
    """Writes everything to several streams at once — used to mirror
    stdout/stderr into a per-run log file without changing any of the
    existing print() call sites.
    """

    def __init__(self, *streams):
        self.streams = streams

    def write(self, data: str) -> None:
        for stream in self.streams:
            stream.write(data)

    def flush(self) -> None:
        for stream in self.streams:
            stream.flush()


def _phase_zones(phase: str) -> list[str]:
    """Looks up `phase` in data/phases.json (hand-edited zone-keyword list
    per raid phase, e.g. "2" -> SSC/TK) — same idea as data/consumables.json,
    a plain file so a new phase doesn't need a code change.
    """
    phases = json.loads(PHASES_PATH.read_text(encoding="utf-8"))
    zones = phases.get(str(phase))
    if zones is None:
        raise ConfigError(f"Unknown --phase {phase!r}. Known phases in {PHASES_PATH}: {sorted(phases)}")
    return zones


def _report_time_window(fights_response: dict) -> tuple[int, int]:
    """Report-wide start/end, as ms offsets relative to report start.

    Always derived from fights[].start_time/end_time — the top-level
    "start"/"end" fields on this response are absolute Unix epoch ms
    (real wall-clock time) and are on a different scale; using them here
    would send the wrong window to /report/tables and /report/events.
    """
    fights = fights_response.get("fights", []) or []
    if not fights:
        raise WCLAPIError("Report has no fights — nothing to analyze.")
    start = min(f["start_time"] for f in fights)
    end = max(f["end_time"] for f in fights)
    return start, end


def _consumable_cast_filter(known_ids: set[int]) -> str:
    ids = ", ".join(str(i) for i in sorted(known_ids))
    return f'type = "cast" and ability.id in ({ids})'


def _log_date(fights_response: dict) -> str:
    """UTC calendar date of the report, from the top-level "start" (absolute
    epoch ms) — used to bucket reports into raid nights for sheet naming
    and same-night merging.
    """
    return datetime.fromtimestamp(fights_response["start"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")


def _log_sheet_name(log_date: str, code: str) -> str:
    """"<YYYY-MM-DD> (<code-prefix>)" so two different logs on the same
    calendar day get distinct sheets instead of colliding.
    """
    return f"{log_date} ({code[:12]})"


def _process_report(
    client: WCLClient, code: str, fights_response: dict, zone_keywords: list[str] | None = None
) -> list[PlayerSummary]:
    """Runs the role + consumable pipeline for one already-fetched report.

    If `zone_keywords` is given, only the fights in matching zones are
    queried (see `phase_fight_windows` — a report can span several raid
    zones back to back, so the report's overall start/end would otherwise
    pull in other zones' data too), and players who never show up in any
    matching window's role buckets or consumable casts are dropped from
    the result entirely (they just weren't part of this phase's content).
    """
    player_index = build_player_index(fights_response)
    if not player_index:
        raise WCLAPIError(f"No players found in report {code}.")

    if zone_keywords:
        windows = phase_fight_windows(fights_response, zone_keywords)
        if not windows:
            raise WCLAPIError(f"No fights in report {code} match zones {zone_keywords}.")
    else:
        windows = [_report_time_window(fights_response)]

    cast_filter = _consumable_cast_filter(KNOWN_CONSUMABLE_IDS)
    role_dicts = []
    consumable_counts: dict[int, dict[int, int]] = {}
    for start, end in windows:
        summary_table = client.get_table("summary", code, start=start, end=end)
        role_dicts.append(classify_roles(summary_table, player_index))

        events = client.get_events(code, start, end, filter=cast_filter)
        for player_id, per_ability in count_consumable_casts(events, KNOWN_CONSUMABLE_IDS, player_index).items():
            totals = consumable_counts.setdefault(player_id, {})
            for ability_id, count in per_ability.items():
                totals[ability_id] = totals.get(ability_id, 0) + count

    roles = merge_role_dicts(role_dicts)

    if zone_keywords:
        relevant_ids = {pid for pid, role in roles.items() if role != ROLE_UNKNOWN} | set(consumable_counts)
        player_index = {pid: info for pid, info in player_index.items() if pid in relevant_ids}

    return merge_player_summary(player_index, consumable_counts, CONSUMABLES, roles)


def run(
    report_url: str,
    api_key_override: str | None,
    output_path: str | None,
    zone_keywords: list[str] | None = None,
    json_output_path: str | None = None,
    phase_label: str | None = None,
) -> int:
    try:
        config = get_config(api_key_override)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 3

    try:
        code = extract_report_code(report_url)
    except InvalidReportURLError as exc:
        print(f"Invalid report URL: {exc}", file=sys.stderr)
        return 2

    client = WCLClient(api_key=config.api_key, base_url=config.base_url)

    try:
        print(f"Fetching report {code}...")
        fights_response = client.get_report_fights(code)
        print("Fetching role composition and consumable cast events...")
        summaries = _process_report(client, code, fights_response, zone_keywords)
    except WCLAPIError as exc:
        print(f"WCL API error: {exc}", file=sys.stderr)
        return 3

    path = output_path or "report.xlsx"
    json_path = json_output_path or "docs/data.json"
    log_date = _log_date(fights_response)
    sheet_name = _log_sheet_name(log_date, code)
    write_report(summaries, path, sheet_name)
    print(f"Wrote {path}")
    write_json_report(
        summaries, json_path, sheet_name, log_date=log_date, report_code=code, phase=phase_label or "unphased"
    )
    print(f"Wrote {json_path}")
    return 0


def _date_to_epoch_ms(date_str: str, *, end_of_day: bool = False) -> int:
    dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    if end_of_day:
        dt = dt.replace(hour=23, minute=59, second=59, microsecond=999000)
    return int(dt.timestamp() * 1000)


def _relevant_duration_ms(fights_response: dict, zone_keywords: list[str] | None) -> int:
    """How much of this report is actually phase-relevant content — used to
    pick the most complete log when several reports land on the same
    calendar date (different guild members logging the same raid
    independently; only one of them counts, see `run_guild`)."""
    if not zone_keywords:
        return fights_response["end"] - fights_response["start"]
    return sum(end - start for start, end in phase_fight_windows(fights_response, zone_keywords))


def run_guild(
    guild_name: str,
    server_name: str,
    server_region: str,
    zones: list[str],
    date_from: str | None,
    date_to: str | None,
    api_key_override: str | None,
    output_path: str | None,
    json_output_path: str | None = None,
    phase_label: str | None = None,
) -> int:
    try:
        config = get_config(api_key_override)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 3

    client = WCLClient(api_key=config.api_key, base_url=config.base_url)
    start_ms = _date_to_epoch_ms(date_from) if date_from else None
    end_ms = _date_to_epoch_ms(date_to, end_of_day=True) if date_to else None

    try:
        print(f"Fetching guild report list for {guild_name} ({server_name}-{server_region})...")
        reports = client.get_guild_reports(guild_name, server_name, server_region, start=start_ms, end=end_ms)
    except WCLAPIError as exc:
        print(f"WCL API error: {exc}", file=sys.stderr)
        return 3

    reports = sorted(reports, key=lambda r: r.get("start", 0))
    if not reports:
        print("No reports found for this guild/date range.", file=sys.stderr)
        return 1

    # First pass: just fetch fights (cheap) for every candidate, bucket by
    # raid-night date, no summary/events calls yet — those only happen below
    # for whichever report(s) we actually decide to keep.
    fights_by_date: dict[str, list[tuple[str, dict]]] = {}

    for report in reports:
        code = report["id"]
        try:
            print(f"Fetching report {code} ({report.get('title', '')})...")
            fights_response = client.get_report_fights(code)
            log_date = _log_date(fights_response)

            # Guild batch mode only: stray off-roster/pug runs (e.g. a "Bip
            # Pug" team) show up in the same guild's report list with no way
            # to tell them apart via the API — restricting to the main raid
            # day (Sunday) filters those out. Doesn't apply to a directly
            # given report_url, which is processed regardless of weekday.
            if datetime.strptime(log_date, "%Y-%m-%d").weekday() != 6:
                print(f"  skipping {code} — {log_date} isn't a Sunday")
                continue

            if zones and not phase_fight_windows(fights_response, zones):
                print(f"  skipping {code} — no fights match zones {zones}")
                continue
        except WCLAPIError as exc:
            print(f"  skipping {code} — {exc}", file=sys.stderr)
            continue

        fights_by_date.setdefault(log_date, []).append((code, fights_response))

    if not fights_by_date:
        print("No reports matched the day/zone filter.", file=sys.stderr)
        return 1

    path = output_path or "report.xlsx"
    json_path = json_output_path or "docs/data.json"
    for log_date, candidates in sorted(fights_by_date.items()):
        # Only one report counts per calendar date — several reports on the
        # same date means several guild members independently logged the
        # same raid, not a real split, so pick the most complete one and
        # drop the rest rather than summing (would multiply every
        # consumable count by however many people logged it).
        best_code, best_fights = max(candidates, key=lambda c: _relevant_duration_ms(c[1], zones))
        dropped = [code for code, _ in candidates if code != best_code]
        if dropped:
            print(f"  {log_date}: kept {best_code}, dropped duplicate log(s) {', '.join(dropped)}")

        sheet_name = _log_sheet_name(log_date, best_code)
        if sheet_already_exists(path, sheet_name):
            print(f"  {log_date}: sheet for {best_code} already exists, skipping re-parse.")
            continue

        try:
            print(f"Fetching role composition and consumable cast events for {best_code}...")
            summaries = _process_report(client, best_code, best_fights, zones)
        except WCLAPIError as exc:
            print(f"  skipping {best_code} — {exc}", file=sys.stderr)
            continue

        write_report(summaries, path, sheet_name)
        write_json_report(
            summaries,
            json_path,
            sheet_name,
            log_date=log_date,
            report_code=best_code,
            phase=phase_label or "unphased",
        )

    print(f"Wrote {path}")
    print(f"Wrote {json_path}")
    return 0


def main(argv: list[str] | None = None) -> None:
    LOGS_DIR.mkdir(exist_ok=True)
    log_path = LOGS_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    with open(log_path, "w", encoding="utf-8") as log_file:
        real_stdout, real_stderr = sys.stdout, sys.stderr
        sys.stdout = _Tee(real_stdout, log_file)
        sys.stderr = _Tee(real_stderr, log_file)
        try:
            exit_code = _main(argv)
        except SystemExit:
            raise
        except Exception:
            traceback.print_exc()  # goes through sys.stderr -> also lands in the log file
            exit_code = 1
        finally:
            sys.stdout, sys.stderr = real_stdout, real_stderr

    sys.exit(exit_code)


def _main(argv: list[str] | None) -> int:
    # Logged before argparse can fail/exit — when launched via run.bat, the
    # batch file's own "Phase number / WCL Report URL" prompts and whatever
    # was typed there only ever show in the console window, never in the
    # Logs/ file (the stdout/stderr Tee only wraps this Python process, not
    # cmd.exe's own set /p echo) — without this line a bad-input error in
    # the log has no record of what was actually passed in.
    print(f"Args: {argv if argv is not None else sys.argv[1:]}")

    parser = argparse.ArgumentParser(
        description="Summarize Warcraft Logs (fresh.warcraftlogs.com) TBC reports "
        "into per-player consumable usage, exported to Excel."
    )
    parser.add_argument("report_url", nargs="?", default=None, help="Full WCL report URL or bare report code")
    parser.add_argument(
        "--guild",
        nargs=3,
        metavar=("NAME", "SERVER", "REGION"),
        default=None,
        help=f"Instead of a single report, process every guild report (default: {DEFAULT_GUILD})",
    )
    parser.add_argument(
        "--phase",
        default=None,
        help="Raid phase number, e.g. 2 — looks up its zones in data/phases.json (alternative to --zones; "
        "required with no report_url, optional with one — restricts a multi-zone report to just this phase's fights)",
    )
    parser.add_argument(
        "--zones",
        default=None,
        help="Comma-separated zone-name substrings to keep, e.g. 'serpentshrine,tempest keep,the eye' "
        "(alternative to --phase, same scoping rules)",
    )
    parser.add_argument("--from", dest="date_from", default=None, help="YYYY-MM-DD, optionally further narrows the guild report list by date")
    parser.add_argument("--to", dest="date_to", default=None, help="YYYY-MM-DD, optionally further narrows the guild report list by date")
    parser.add_argument("--api-key", dest="api_key", default=None, help="Override WCL_API_KEY")
    parser.add_argument(
        "--output-path",
        dest="output_path",
        default=None,
        help="Output .xlsx path (default: report.xlsx — re-running adds/updates a sheet in the same file)",
    )
    parser.add_argument(
        "--json-output-path",
        dest="json_output_path",
        default=None,
        help="Output JSON path for the web UI (default: docs/data.json — re-running adds/updates an entry in the same file)",
    )
    args = parser.parse_args(argv)

    if not args.zones and not args.phase and not args.report_url:
        parser.error("need --phase or --zones when there's no report_url (which phase's logs to keep?)")

    if args.phase and "://" in args.phase:
        parser.error(
            f"--phase got a URL ({args.phase!r}) instead of a phase number — looks like a report URL was "
            "pasted into the wrong prompt. Pass the URL as the plain report_url argument instead (no --phase)."
        )

    zones = None
    if args.zones:
        zones = [z.strip().lower() for z in args.zones.split(",") if z.strip()]
    elif args.phase:
        try:
            zones = _phase_zones(args.phase)
        except ConfigError as exc:
            parser.error(str(exc))

    # Identifies which "Phase" bucket this run's JSON output belongs to (see
    # json_export.py) — the phase number when --phase was given, else a
    # zones-based fallback label so --zones runs still land somewhere
    # findable in the web UI instead of crashing on a missing key.
    phase_label = args.phase or (args.zones if args.zones else None)

    if args.report_url:
        return run(args.report_url, args.api_key, args.output_path, zones, args.json_output_path, phase_label)
    else:
        guild_name, server_name, server_region = args.guild or DEFAULT_GUILD
        return run_guild(
            guild_name,
            server_name,
            server_region,
            zones,
            args.date_from,
            args.date_to,
            args.api_key,
            args.output_path,
            args.json_output_path,
            phase_label,
        )


if __name__ == "__main__":
    main()
