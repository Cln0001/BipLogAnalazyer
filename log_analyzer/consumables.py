"""Loads and indexes the TBC consumable reference table (data/consumables.json).

The JSON file is a hand-editable list, not a hardcoded Python dict, so adding
or correcting a spell id doesn't require touching code. Names/categories were
taken from the raid's own "General" consumables breakdown; spell ids were then
resolved against a live report (matching ability.name via the WCL API) rather
than guessed. Entries with empty `spell_ids` are known-by-name but their id
hasn't been observed yet — they're harmless no-ops until someone fills them
in. Entries marked "verified": false should be cross-checked against the WCL
UI (``?type=casts&ability=<spell_id>``) before fully trusting their count.

Most entries are tallied by *cast count* (`/report/events`, see cli.py). An
entry can instead set `"metric": "uptime"` to be tallied by *buff uptime %*
(`/report/tables/buffs`, see `UptimeMetric`/`UPTIME_METRICS` below and
cli.py's `_uptime_metrics`) — used for buffs that are typically held rather
than repeatedly cast (scrolls, flasks, elixirs, raid food). An uptime entry
can optionally restrict which roles it's computed for via `"roles"` (a list
of analyze.py role name strings); omitting it means every role.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parent.parent / "data" / "consumables.json"


@dataclass(frozen=True)
class ConsumableInfo:
    name: str
    category: str
    verified: bool = True


@dataclass(frozen=True)
class UptimeMetric:
    name: str
    spell_ids: tuple[int, ...]
    roles: frozenset[str] | None  # None means every role


def load_consumables(path: Path = DATA_PATH) -> dict[int, ConsumableInfo]:
    """Returns spell_id -> ConsumableInfo for every *cast-count* entry (i.e.
    not `"metric": "uptime"`), loaded from the JSON data file.

    Multiple spell ids can map to the same ConsumableInfo (e.g. "Super Mana
    Potion equivalents" covers several distinct items with one display name).
    Uptime entries are deliberately excluded — they're tallied via
    `/report/tables/buffs` instead (see `UPTIME_METRICS`), so their spell ids
    must not also feed the `/report/events` cast-count filter.
    """
    raw = json.loads(path.read_text(encoding="utf-8"))
    result: dict[int, ConsumableInfo] = {}
    for entry in raw:
        if entry.get("metric") == "uptime":
            continue
        info = ConsumableInfo(
            name=entry["name"],
            category=entry["category"],
            verified=entry.get("verified", True),
        )
        for spell_id in entry.get("spell_ids", []):
            result[spell_id] = info
    return result


def load_uptime_metrics(path: Path = DATA_PATH) -> list[UptimeMetric]:
    """Returns every `"metric": "uptime"` entry, grouped by display name (a
    name can span several entries — not currently used, but mirrors
    load_consumables' multi-entries-per-name support). Entries with empty
    spell_ids (unconfirmed placeholders) are skipped.
    """
    raw = json.loads(path.read_text(encoding="utf-8"))
    by_name: dict[str, tuple[list[int], frozenset[str] | None]] = {}
    for entry in raw:
        if entry.get("metric") != "uptime" or not entry.get("spell_ids"):
            continue
        name = entry["name"]
        roles = frozenset(entry["roles"]) if entry.get("roles") else None
        ids, _ = by_name.setdefault(name, ([], roles))
        ids.extend(entry["spell_ids"])
    return [UptimeMetric(name=name, spell_ids=tuple(ids), roles=roles) for name, (ids, roles) in by_name.items()]


def _load_name_categories(path: Path = DATA_PATH) -> dict[str, str]:
    """display name -> category, including placeholder entries with empty
    spell_ids — unlike CONSUMABLES (keyed by spell_id), this covers every
    name in the data file regardless of whether it's been seen live yet.
    Used by json_export.py to group consumable rows into UI sections.
    """
    raw = json.loads(path.read_text(encoding="utf-8"))
    return {entry["name"]: entry["category"] for entry in raw}


# Consumable rows are colored by section (not by individual name) in the web
# UI. data/consumables.json's raw "category" field is finer-grained than
# what's useful there, so several raw categories fold into one section here.
DISPLAY_CATEGORY = {
    "potion": "Potions",
    "mana_item": "Potions",
    "energy_item": "Potions",
    "rune": "Others",
    "explosive": "Explosives",
    "scroll": "Buffs/Items",
    "flask": "Buffs/Items",
    "elixir": "Buffs/Items",
    "food": "Buffs/Items",
    "bandage": "Buffs/Items",
}
# Per-name override where the raw consumables.json category (and thus the
# default display-category mapping above) doesn't match what should show in
# the UI — "Flame Cap" is a "potion" by raw category, but belongs in
# "Others" alongside Demonic Rune/Dark Rune here, same color and everything.
CATEGORY_OVERRIDE = {"Flame Cap": "Others"}
CATEGORY_ORDER = ["Potions", "Others", "Explosives", "Buffs/Items", "Other"]

# Section color, applied to every row in that category — used by the web UI.
CATEGORY_FILL = {
    "Potions": "DAEEF3",
    "Others": "F7CDEE",
    "Explosives": "FFB3B3",
}

# "Explosives/Drums" is a synthetic total row (Sapper/Nades + Drums, added
# at export time, not a real consumable.name from consumables.json).
EXPLOSIVE_TOTAL_LABEL = "Explosives/Drums"
ROW_TOTAL_PREFIX = "§TOTAL§"

# Every "metric": "uptime" entry's display row is named "<json name> Uptime
# %" (see cli.py's _uptime_metrics), not the bare json "name" — strip this
# suffix before looking up NAME_CATEGORIES, which is keyed by the bare name.
UPTIME_SUFFIX = " Uptime %"


def category_for(name: str) -> str:
    if name == EXPLOSIVE_TOTAL_LABEL:
        return "Explosives"
    if name.startswith(ROW_TOTAL_PREFIX):
        return name[len(ROW_TOTAL_PREFIX):]
    if name in CATEGORY_OVERRIDE:
        return CATEGORY_OVERRIDE[name]
    base_name = name[: -len(UPTIME_SUFFIX)] if name.endswith(UPTIME_SUFFIX) else name
    return DISPLAY_CATEGORY.get(NAME_CATEGORIES.get(base_name, ""), "Other")


# Module-level cache: loaded once, reused by analyze.py / cli.py.
CONSUMABLES: dict[int, ConsumableInfo] = load_consumables()
KNOWN_CONSUMABLE_IDS: set[int] = set(CONSUMABLES.keys())
NAME_CATEGORIES: dict[str, str] = _load_name_categories()
UPTIME_METRICS: list[UptimeMetric] = load_uptime_metrics()
