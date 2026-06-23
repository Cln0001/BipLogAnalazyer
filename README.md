# WoW TBC Classic Log Analyzer

Pulls a [Warcraft Logs](https://fresh.warcraftlogs.com/) (TBC "Fresh" realms) report via the **v1 REST API** and writes a per-player summary — overall performance (damage/healing, DPS/HPS) and consumable usage counts — to an Excel file.

## Setup

```
pip install -r requirements.txt
cp .env.example .env   # then edit .env and paste your WCL v1 API key
```

Get a v1 API key at https://fresh.warcraftlogs.com/v1/docs/ (account → API access).

## Usage

```
python -m log_analyzer.cli "https://fresh.warcraftlogs.com/reports/QgPvwcBdHX78zaxJ"
```

Writes `report_QgPvwcBdHX78zaxJ.xlsx` in the current directory (one row per player: name, class, total damage, DPS, total healing, HPS, then one column per consumable used in the report). Override the output path with `--output-path`, or the API key with `--api-key`.

Console output is status/errors only — the actual results are in the spreadsheet.

## Consumable list

`data/consumables.json` maps spell id → consumable name/category. It's a plain hand-editable JSON list — add missing items or fix wrong ones directly. Entries with `"verified": false` were compiled from general TBC reference knowledge, not individually confirmed against a live report; cross-check a suspicious count against the WCL web UI (`?type=casts&ability=<spell_id>` on a report URL) and correct the JSON if needed.

## Tests

```
pip install -r requirements-dev.txt
pytest
```

Unit tests run against fixtures captured in `fixtures/` — no live API calls / API key needed.

## How it talks to the API (confirmed against a live report)

- `/report/fights/{code}` has no usable top-level `start`/`end` for windowing — those are absolute Unix epoch ms (wall-clock), on a different scale than `fights[].start_time`/`end_time` (ms relative to report start). The tool always derives its analysis window from `min(start_time)`/`max(end_time)` across fights.
- Per-player damage/healing comes from two separate calls, `/report/tables/damage-done/{code}` and `/report/tables/healing/{code}` (both with `start`/`end` query params) — there's no single "summary" table with both. Neither view has a precomputed `dps`/`hps` field; both are derived here as `total / (activeTime / 1000)`.
- Consumable counts use a server-side filter on `/report/events/{code}`: `filter=type = "cast" and ability.id in (id1, id2, ...)` built from the known consumable id list — far cheaper than downloading every event and filtering client-side (98 matching events vs. tens of thousands of raw casts on the example report).
- Pets/NPCs are excluded from `friendlies` via a TBC-class allowlist; real data confirms hunter pets carry `"type": "Pet"` and boss adds/totem effects carry `"type": "NPC"`, neither of which matches a class name.

## Known limitations / next steps

- First cut analyzes the **whole report** time window (no per-boss/per-fight filtering yet).
- No JSON/CSV export yet, just `.xlsx`.
- No multi-log comparison yet.
