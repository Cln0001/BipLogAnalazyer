# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

WoW TBC Classic log analyzer. Pulls a Warcraft Logs ("fresh" realms) report via the **v1 REST API** and writes a per-player consumable usage summary (role, class, consumable counts) into a `.xlsx` — one sheet per log, accumulated across a whole raid season in the same file, plus an aggregated "Summary" overview sheet.

## Commands

```
# setup
.venv\Scripts\python.exe -m pip install -r requirements.txt
cp .env.example .env   # then edit: paste WCL v1 API key into WCL_API_KEY

# run — always use the venv interpreter, plain `python` lacks deps
.venv\Scripts\python.exe -m log_analyzer.cli "https://fresh.warcraftlogs.com/reports/<code>"

# or process an entire guild's phase in one go (every report, deduped per raid night)
.venv\Scripts\python.exe -m log_analyzer.cli --phase 2

# or via the Windows wrapper (prompts for the URL interactively)
run.bat
```

No test suite — do not add one (project policy: don't write tests for this codebase).

`--api-key` overrides `WCL_API_KEY`; `--output-path` overrides the default `report.xlsx` — re-running against the same path (the default) keeps adding/correcting sheets in that one file rather than starting a fresh one per log. Omitting `report_url` switches to guild-batch mode: defaults to `DEFAULT_GUILD` in `cli.py` ("Bananas in Pyjamas" / Spineshatter / EU), overridable with `--guild NAME SERVER REGION`; requires either `--phase N` (looks up zone keywords in `data/phases.json`) or `--zones` (comma-separated, case-insensitive substrings matched against each report's zone names) to know which raid phase's logs to keep — the guild reports list otherwise spans the whole guild's history. `--from`/`--to` (`YYYY-MM-DD`, UTC) are optional, purely to cut down how many reports get listed/fetched in the first place. Guild-batch mode also only keeps reports logged on a Sunday (the guild's main raid day) — the v1 API exposes no "guild team" field on a report, so this is the only way to filter out stray off-roster/pug runs (e.g. a "Bip Pug" team) that show up in the same guild's report list. Doesn't apply when a `report_url` is given directly.

## Architecture

Pipeline, wired together in `log_analyzer/cli.py:run` (single report) / `run_guild` (whole-guild batch):

1. **`wcl_client.py`** — thin HTTP client for the v1 API (`requests`, 429 backoff/retry, and a console print on every 429 retry so a long guild-batch run doesn't look hung). `extract_report_code` parses a report code out of a full URL or accepts a bare code. Endpoints used: `/report/fights`, `/report/tables/summary` (role composition), `/report/events` (paginated, for consumable casts), `/reports/guild/{name}/{server}/{region}` (guild's report list, only used by `--guild`; note the plural "reports" — confirmed live, the one endpoint that breaks the otherwise-singular `/report/...` pattern).
2. **`analyze.py`** — pure aggregation, no I/O. Builds the player index from `friendlies[]` (filtered to a TBC class allowlist to drop pets/NPCs — no Death Knight in TBC), classifies roles (Tank > Healer > DPS priority, DPS further split into Caster/Physical by spec name), and tallies consumable casts. `phase_fight_windows` finds the contiguous (start_time, end_time) windows of just the fights matching `--zones`/`--phase` — a single report can log a whole raid night across *several* zones back to back (e.g. SSC, then TK, then Gruul, then AQ40, all one report code), so filtering has to happen at fight granularity, not "does this report contain a matching zone anywhere" (that would pull in every other zone's data too). `merge_role_dicts` combines the role classification from each window of the *same* report (player ids are stable within one report) by Tank > Healer > DPS priority. `merge_summaries_across_reports` combines *different* reports' `PlayerSummary` lists into one — for a raid night genuinely split across multiple report codes (disconnect/server reset, confirmed non-overlapping wall-clock time) — summing consumables by name and resolving role the same way. **Don't** call this for reports that merely share a date; see the duplicate-log note below.
3. **`consumables.py`** — loads `data/consumables.json` (hand-edited spell_id → name/category map) into a module-level cache (`CONSUMABLES`, `KNOWN_CONSUMABLE_IDS`) used by both `analyze.py` and the event filter built in `cli.py`.
4. **`excel_export.py`** — `write_report(summaries, output_path, sheet_name)` loads the existing workbook at `output_path` (or creates one) and writes/corrects one sheet per log, named `"<YYYY-MM-DD> (<code-prefix>)"`. Each sheet is a transposed table: players grouped into column blocks by role (Tank/Healer/Caster DPS/Physical DPS/Unknown), each block headed by a merged role label and separated by a blank column, sorted by class then name within a block. Rows are Class, Role, then one row per consumable. A `"Summary"` sheet is recomputed from every per-date sheet on each run (`_recompute_total_sheet`): per player, the *average* consumable count across the logs they attended (not a raw sum, via `_read_sheet` parsing each per-date sheet back) plus a `"Logs Attended"` row — averaging only over attended logs means missing a raid doesn't drag a player's numbers down. Re-running on a log whose sheet already exists compares old vs new consumable counts and either leaves it (`"unchanged"`) or replaces it (`"corrected"`) — it never duplicates a sheet for the same log.

### Non-obvious API quirks (load-bearing, don't "fix")

- `/report/fights` top-level `start`/`end` are absolute Unix epoch ms — a different scale than `fights[].start_time`/`end_time` (ms relative to report start). The analysis window is always derived from `min(start_time)`/`max(end_time)` across fights (or, when `--zones`/`--phase` is given, from `phase_fight_windows`' per-zone windows), never the top-level fields. The top-level `start` *is* used for one thing: deriving the log's calendar date (UTC) for its sheet name, in `cli.py:_log_sheet_name`.
- A report can contain fights from multiple raid zones in one continuous log (confirmed live: a single report code with SSC, TK, Gruul's Lair, and AQ40 fights all back to back). `_process_report` in `cli.py` queries `/report/tables/summary` and `/report/events` once per `phase_fight_windows` window rather than once for the whole report, and drops players who never appear in any matching window (no role bucket hit, no consumable casts) — otherwise off-phase zones' role/consumable data leaks into the sheet.
- Consumable counts use a server-side filter (`filter=type = "cast" and ability.id in (...)`) on `/report/events` rather than downloading and filtering every event client-side.
- `_read_sheet` in `excel_export.py` is the exact inverse of `_write_sheet`'s layout (row 2 = names, row 3 = Class, row 4 = Role, row 5+ = consumables) — if that layout changes, `_read_sheet` and the Total-sheet recompute must change with it, since nothing else describes the sheet's shape.
- Only one report counts per calendar date — confirmed live that multiple report codes on the same date are several different guild members independently logging the exact same raid (near-fully-overlapping wall-clock time), not a genuine split raid. `run_guild` picks the single most phase-relevant one per date (`_relevant_duration_ms`) and drops the rest. Summing them instead (an earlier version of this code did, via `merge_summaries_across_reports`) was an actual bug: it inflated a player's "Destro" count from a real ~20 to a reported 60 (4 people's logs of one raid, summed). `merge_summaries_across_reports` is still in `analyze.py` but currently unused by `cli.py` — don't wire it back into `run_guild` without re-confirming a real non-overlapping split case exists.

### Removed scope (deliberate)

Performance (DPS/HPS via damage/healing tables), parse % (`/report/rankings`), a "Gear Issues" sheet, and a "Buff Consumables" (flask/elixir/food/weapon-enchant uptime) sheet were tried and removed — didn't work as wanted on fresh-realm reports / added complexity without payoff. The tool is scoped back down to: role/class grouping + consumable cast counts, one sheet. Don't re-add without discussing first.

## Logging

Every `main()` invocation mirrors everything printed to stdout/stderr (including an unexpected exception's traceback, caught at the top level) into a timestamped file under `Logs/` (gitignored) — see `cli.py`'s `_Tee` class. This is in addition to, not instead of, the console output.

## Data files

- `data/consumables.json` — plain hand-editable list, not a hardcoded Python dict, so spell ID fixes don't need code changes. Entries with `"verified": false` were compiled from general TBC reference knowledge, not confirmed against a live report; cross-check via the WCL web UI (`?type=casts&ability=<spell_id>`) before trusting the count.
- `data/phases.json` — phase number (string key) -> list of zone-name keyword substrings for `--phase`. Add a new entry here when a new raid phase opens; don't hardcode it into `cli.py`. Only `"2"` (SSC/TK) is populated — its date range was deliberately *not* recorded anywhere, since the zone-name filter alone already scopes a report to the right phase regardless of when it was logged.
