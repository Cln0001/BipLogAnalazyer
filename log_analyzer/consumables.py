"""Loads and indexes the TBC consumable reference table (data/consumables.json).

The JSON file is a hand-editable list, not a hardcoded Python dict, so adding
or correcting a spell id doesn't require touching code. Names/categories were
taken from the raid's own "General" consumables breakdown; spell ids were then
resolved against a live report (matching ability.name via the WCL API) rather
than guessed. Entries with empty `spell_ids` are known-by-name but their id
hasn't been observed yet — they're harmless no-ops until someone fills them
in. Entries marked "verified": false should be cross-checked against the WCL
UI (``?type=casts&ability=<spell_id>``) before fully trusting their count.
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


def load_consumables(path: Path = DATA_PATH) -> dict[int, ConsumableInfo]:
    """Returns spell_id -> ConsumableInfo, loaded from the JSON data file.

    Multiple spell ids can map to the same ConsumableInfo (e.g. "Super Mana
    Potion equivalents" covers several distinct items with one display name).
    """
    raw = json.loads(path.read_text(encoding="utf-8"))
    result: dict[int, ConsumableInfo] = {}
    for entry in raw:
        info = ConsumableInfo(
            name=entry["name"],
            category=entry["category"],
            verified=entry.get("verified", True),
        )
        for spell_id in entry.get("spell_ids", []):
            result[spell_id] = info
    return result


def _load_name_categories(path: Path = DATA_PATH) -> dict[str, str]:
    """display name -> category, including placeholder entries with empty
    spell_ids — unlike CONSUMABLES (keyed by spell_id), this covers every
    name in the data file regardless of whether it's been seen live yet.
    Used by excel_export.py to group consumable rows into sheet sections.
    """
    raw = json.loads(path.read_text(encoding="utf-8"))
    return {entry["name"]: entry["category"] for entry in raw}


# Module-level cache: loaded once, reused by analyze.py / cli.py.
CONSUMABLES: dict[int, ConsumableInfo] = load_consumables()
KNOWN_CONSUMABLE_IDS: set[int] = set(CONSUMABLES.keys())
NAME_CATEGORIES: dict[str, str] = _load_name_categories()
