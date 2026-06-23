"""Writes per-log + aggregated "Summary" data into one JSON file for the
static web UI (docs/), mirroring excel_export.py's accumulation model
(upsert by sheet_name, recompute the Summary block from everything in the
file) but without any sheet/round-trip parsing — the JSON itself is the
durable per-log store, so accumulation is just dict upsert + recompute.

Logs/summary are bucketed under a "phase" key (e.g. "2") so the UI can
switch between raid phases — each phase has its own independent log list
and Summary, recomputed only from that phase's logs.
"""

from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

from log_analyzer.analyze import (
    ROLE_CASTER_DPS,
    ROLE_HEALER,
    ROLE_PHYSICAL_DPS,
    ROLE_PRIORITY,
    ROLE_TANK,
    ROLE_UNKNOWN,
    PlayerSummary,
)
from log_analyzer.excel_export import _CATEGORY_FILL, _CATEGORY_ORDER, _category_for

ROLE_ORDER = [ROLE_TANK, ROLE_HEALER, ROLE_CASTER_DPS, ROLE_PHYSICAL_DPS, ROLE_UNKNOWN]


def _skeleton() -> dict:
    return {
        "version": 2,
        "generated_at": None,
        "role_order": ROLE_ORDER,
        "category_order": _CATEGORY_ORDER,
        "category_colors": _CATEGORY_FILL,
        "consumable_categories": {},
        "phase_order": [],
        "phases": {},
    }


def _empty_phase() -> dict:
    return {"log_order": [], "logs": {}, "summary": {}}


def _categories_for_names(names: set[str]) -> dict[str, str]:
    """Consumable display name -> section category (Potions/Others/
    Explosives/Buffs-Items/Other), same mapping excel_export.py uses to
    group consumable rows into sheet sections — computed once here so the
    frontend doesn't need to duplicate consumables.json's category rules.
    """
    return {name: _category_for(name) for name in names}


def _load(output_path: str) -> dict:
    if not Path(output_path).exists():
        return _skeleton()
    return json.loads(Path(output_path).read_text(encoding="utf-8"))


def _phase_sort_key(phase: str):
    # Numeric phases ("2", "3") sort numerically; anything non-numeric
    # (e.g. a zones-based fallback label) sorts after, alphabetically.
    return (0, int(phase)) if phase.isdigit() else (1, phase)


def _summary_from_logs(logs: dict[str, dict]) -> dict:
    """Same algorithm as excel_export._recompute_total_sheet: per player,
    average consumable count across the logs they attended (not a raw sum)
    plus a "Logs Attended" count, computed here in Python so the frontend
    does zero aggregation math. A player's role on the Summary is whichever
    role they played most often across their attended logs (ties broken by
    ROLE_PRIORITY) — a single night spent off-role (e.g. a healer dps'ing
    for one log) shouldn't flip their season-long Summary role; per-log
    roles already reflect that night's actual spec correctly (see
    analyze.py's classify_roles), this only changes how they're aggregated.
    """
    logs_attended: dict[str, int] = {}
    consumable_sums: dict[str, dict[str, int]] = {}
    class_by_name: dict[str, str] = {}
    role_counts: dict[str, dict[str, int]] = {}

    for sheet_name in sorted(logs, key=lambda k: logs[k]["log_date"]):
        for player in logs[sheet_name]["players"]:
            name = player["name"]
            logs_attended[name] = logs_attended.get(name, 0) + 1
            totals = consumable_sums.setdefault(name, {})
            for consumable, count in player["consumables"].items():
                totals[consumable] = totals.get(consumable, 0) + count
            class_by_name[name] = player["class_name"]
            counts = role_counts.setdefault(name, {})
            counts[player["role"]] = counts.get(player["role"], 0) + 1

    def majority_role(name: str) -> str:
        counts = role_counts[name]
        best_count = max(counts.values())
        tied = [role for role, count in counts.items() if count == best_count]
        return next((role for role in ROLE_PRIORITY if role in tied), tied[0])

    players = [
        {
            "name": name,
            "class_name": class_by_name[name],
            "role": majority_role(name),
            "logs_attended": logs_attended[name],
            "consumables_avg": {
                consumable: round(total / logs_attended[name], 1)
                for consumable, total in consumable_sums.get(name, {}).items()
            },
        }
        for name in class_by_name
    ]
    players.sort(key=lambda p: p["name"])

    return {"logs_attended": logs_attended, "players": players}


def _players_signature(players: list[dict]) -> set[tuple]:
    """Drop zero-count entries before comparing, mirroring
    excel_export._summary_signature — only the nonzero consumables matter
    for deciding "unchanged" vs "corrected".
    """
    return {
        (p["name"], p["class_name"], p["role"], tuple(sorted((k, v) for k, v in p["consumables"].items() if v)))
        for p in players
    }


def sheet_already_exists(output_path: str, sheet_name: str, phase: str) -> bool:
    if not Path(output_path).exists():
        return False
    data = _load(output_path)
    return sheet_name in data.get("phases", {}).get(phase, {}).get("logs", {})


def write_report(
    summaries: list[PlayerSummary],
    output_path: str,
    sheet_name: str,
    log_date: str,
    report_code: str,
    phase: str,
) -> None:
    data = _load(output_path)
    phase_data = data["phases"].setdefault(phase, _empty_phase())

    new_players = [asdict(s) for s in summaries]
    existing_entry = phase_data["logs"].get(sheet_name)
    if existing_entry is not None:
        changed = _players_signature(existing_entry["players"]) != _players_signature(new_players)
        print(f"JSON log {sheet_name} (phase {phase}) {'corrected (consumable counts differed)' if changed else 'unchanged'}.")
    else:
        print(f"JSON log {sheet_name} (phase {phase}) added.")

    phase_data["logs"][sheet_name] = {"log_date": log_date, "report_code": report_code, "players": new_players}
    phase_data["log_order"] = sorted(phase_data["logs"], key=lambda k: phase_data["logs"][k]["log_date"])
    phase_data["summary"] = _summary_from_logs(phase_data["logs"])

    data["phase_order"] = sorted(data["phases"], key=_phase_sort_key)
    data["role_order"] = ROLE_ORDER
    data["category_order"] = _CATEGORY_ORDER
    data["category_colors"] = _CATEGORY_FILL

    all_names: set[str] = set()
    for phase_entry in data["phases"].values():
        for log in phase_entry["logs"].values():
            for player in log["players"]:
                all_names.update(player["consumables"].keys())
    data["consumable_categories"] = _categories_for_names(all_names)

    data["version"] = 2
    data["generated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    tmp = output.with_suffix(output.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(output)
