"""Pure aggregation logic — no network I/O, so this is the easiest module to
unit-test against captured fixtures.

Several field-name assumptions below are based on the legacy WCL v1 API
shape and are marked for verification against a real response (see
fixtures/ and the project plan). Code is written defensively (``.get`` with
fallbacks) so a slightly different real shape degrades gracefully rather
than crashing.
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field

from log_analyzer.consumables import ConsumableInfo

# TBC has no Death Knight — only these 9 playable classes. Anything in
# `friendlies` that isn't one of these (pets, NPCs allied to the raid, etc.)
# is excluded from the player index.
TBC_PLAYABLE_CLASSES = {
    "Warrior",
    "Paladin",
    "Hunter",
    "Rogue",
    "Priest",
    "Shaman",
    "Mage",
    "Warlock",
    "Druid",
}

ROLE_TANK = "Tank"
ROLE_HEALER = "Healer"
ROLE_CASTER_DPS = "Caster DPS"
ROLE_PHYSICAL_DPS = "Physical DPS"
ROLE_UNKNOWN = "Unknown"

# Priority used when the same player shows up under different roles across
# reports merged into one raid night (e.g. off-tanked one report, dps'd the
# other) — same Tank > Healer > DPS priority as classify_roles, just applied
# across reports instead of across fights within one report.
ROLE_PRIORITY = [ROLE_TANK, ROLE_HEALER, ROLE_CASTER_DPS, ROLE_PHYSICAL_DPS, ROLE_UNKNOWN]

# DPS specs that deal damage primarily via spells. Anything else seen in the
# "dps" playerDetails bucket (Enhancement, Feral, Retribution, Fury/Arms,
# BeastMastery/Survival/Marksmanship, Combat/Assassination/Subtlety, ...) is
# treated as Physical DPS by default.
CASTER_DPS_SPECS = {
    "Elemental",
    "Destruction",
    "Affliction",
    "Demonology",
    "Balance",
    "Arcane",
    "Fire",
    "Frost",
    "Shadow",
}


@dataclass(frozen=True)
class PlayerInfo:
    player_id: int
    name: str
    class_name: str


@dataclass
class PlayerSummary:
    name: str
    class_name: str
    role: str
    consumables: dict[str, int] = field(default_factory=dict)


def build_player_index(fights_response: dict) -> dict[int, PlayerInfo]:
    """Maps player id -> PlayerInfo, filtered down to real players (no
    pets/NPCs) using a TBC-class allowlist on `friendlies[].type`.
    """
    players: dict[int, PlayerInfo] = {}
    for friendly in fights_response.get("friendlies", []) or []:
        class_name = friendly.get("type")
        if class_name not in TBC_PLAYABLE_CLASSES:
            continue
        player_id = friendly.get("id")
        if player_id is None:
            continue
        players[player_id] = PlayerInfo(
            player_id=player_id,
            name=friendly.get("name", f"Unknown-{player_id}"),
            class_name=class_name,
        )
    return players


def phase_fight_windows(fights_response: dict, zone_keywords: list[str]) -> list[tuple[int, int]]:
    """Contiguous (start_time, end_time) windows — ms offsets, same scale as
    fights[].start_time/end_time — covering only the fights whose zone name
    matches one of `zone_keywords` (case-insensitive substring).

    Adjacent matching fights are merged into one window. This matters
    because a single report can log a whole raid night across *several*
    raid zones back to back (e.g. SSC, then TK, then Gruul, then AQ40, all
    in one report code) — using the report's overall min/max window would
    pull in every other zone's role/consumable data too. Returns one window
    per contiguous run, so callers query/aggregate only the relevant slices.
    """
    zone_names_by_id = {zone.get("id"): zone.get("name", "") for zone in fights_response.get("zones", []) or []}

    windows: list[tuple[int, int]] = []
    current: tuple[int, int] | None = None
    for fight in fights_response.get("fights", []) or []:
        zone_name = (fight.get("zoneName") or zone_names_by_id.get(fight.get("zoneID"), "")).lower()
        if any(keyword in zone_name for keyword in zone_keywords):
            current = (current[0], fight["end_time"]) if current else (fight["start_time"], fight["end_time"])
        elif current is not None:
            windows.append(current)
            current = None
    if current is not None:
        windows.append(current)
    return windows


def merge_role_dicts(role_dicts: list[dict[int, str]]) -> dict[int, str]:
    """Combines several role classifications of the *same report* (one per
    `phase_fight_windows` window) into one, keeping each player's
    highest-priority role across windows (Tank > Healer > DPS, same as
    `classify_roles`' own within-report priority) rather than letting a
    later window's "Unknown" overwrite an earlier window's real role.
    """
    merged: dict[int, str] = {}
    for roles in role_dicts:
        for player_id, role in roles.items():
            if player_id not in merged or ROLE_PRIORITY.index(role) < ROLE_PRIORITY.index(merged[player_id]):
                merged[player_id] = role
    return merged


def matching_fights(fights_response: dict, zone_keywords: list[str] | None) -> list[dict]:
    """Individual fight entries (not merged into windows) whose zone matches
    `zone_keywords` — or every fight in the report when `zone_keywords` is
    None (whole-report mode). Used to classify role *per encounter* rather
    than per `phase_fight_windows` window, since a window can span several
    fights and a single off-role pull within it (e.g. an emergency off-tank
    on one trash pull) shouldn't dominate role classification for the whole
    window the way `merge_role_dicts`' priority merge would.
    """
    fights = fights_response.get("fights", []) or []
    if not zone_keywords:
        return fights

    zone_names_by_id = {zone.get("id"): zone.get("name", "") for zone in fights_response.get("zones", []) or []}
    matched = []
    for fight in fights:
        zone_name = (fight.get("zoneName") or zone_names_by_id.get(fight.get("zoneID"), "")).lower()
        if any(keyword in zone_name for keyword in zone_keywords):
            matched.append(fight)
    return matched


def majority_role_dicts(role_dicts: list[dict[int, str]]) -> dict[int, str]:
    """Combines several per-fight role classifications (one per
    `matching_fights` entry) into one per player: whichever role they were
    classified as most often across fights, ties broken by ROLE_PRIORITY.
    Unlike `merge_role_dicts`, a single off-role fight (e.g. one emergency
    off-tank pull) no longer locks in that role for the whole log — it's
    just one vote among many.

    ROLE_UNKNOWN votes are excluded from the count — many short/ambiguous
    trash pulls can fail to bucket a player at all (table summary is
    unreliable for very short fights), and those shouldn't be able to
    outvote the fights that *did* classify the player into a real role.
    Unknown is only the result when a player has no real-role votes at all.
    """
    counts: dict[int, dict[str, int]] = {}
    for roles in role_dicts:
        for player_id, role in roles.items():
            player_counts = counts.setdefault(player_id, {})
            player_counts[role] = player_counts.get(role, 0) + 1

    result: dict[int, str] = {}
    for player_id, role_counts in counts.items():
        real_counts = {role: count for role, count in role_counts.items() if role != ROLE_UNKNOWN}
        if not real_counts:
            result[player_id] = ROLE_UNKNOWN
            continue
        best_count = max(real_counts.values())
        tied = [role for role, count in real_counts.items() if count == best_count]
        result[player_id] = next((role for role in ROLE_PRIORITY if role in tied), tied[0])
    return result


def _entry_player_id(entry: dict, player_index: dict[int, PlayerInfo]) -> int | None:
    """Resolves a table/event entry to a known player id, trying the most
    likely field names ("id" / "sourceID") and falling back to a name match.
    """
    for key in ("id", "sourceID"):
        candidate = entry.get(key)
        if candidate in player_index:
            return candidate
    name = entry.get("name")
    if name:
        for pid, info in player_index.items():
            if info.name == name:
                return pid
    return None


def classify_roles(
    summary_table: dict, player_index: dict[int, PlayerInfo]
) -> dict[int, str]:
    """Maps player id -> Tank / Healer / Caster DPS / Physical DPS / Unknown,
    using /report/tables/summary's `playerDetails` (which buckets each
    player into "tanks"/"healers"/"dps" based on the spec they actually
    played). A player who swapped roles across fights (e.g. off-tank,
    or healer who occasionally dps'd) can appear in more than one bucket;
    priority is Tank > Healer > DPS, since that reflects the role they
    were needed for that raid night.

    Within "dps", spec name decides Caster vs Physical (see
    CASTER_DPS_SPECS).
    """
    player_details = summary_table.get("playerDetails", {}) or {}

    def ids_in_bucket(bucket: str) -> set[int]:
        return {
            _entry_player_id(entry, player_index)
            for entry in player_details.get(bucket, []) or []
            if _entry_player_id(entry, player_index) is not None
        }

    tank_ids = ids_in_bucket("tanks")
    healer_ids = ids_in_bucket("healers")

    roles: dict[int, str] = {}
    for player_id in player_index:
        if player_id in tank_ids:
            roles[player_id] = ROLE_TANK
        elif player_id in healer_ids:
            roles[player_id] = ROLE_HEALER
        else:
            roles[player_id] = ROLE_UNKNOWN  # filled in below if found in "dps"

    for entry in player_details.get("dps", []) or []:
        player_id = _entry_player_id(entry, player_index)
        if player_id is None or player_id in tank_ids or player_id in healer_ids:
            continue
        specs = entry.get("specs") or []
        spec = specs[0] if specs else None
        roles[player_id] = ROLE_CASTER_DPS if spec in CASTER_DPS_SPECS else ROLE_PHYSICAL_DPS

    return roles


def _event_ability_id(event: dict) -> int | None:
    ability_id = event.get("abilityGameID")
    if ability_id is not None:
        return ability_id
    ability = event.get("ability") or {}
    return ability.get("guid")


def count_consumable_casts(
    events: Iterable[dict],
    known_ids: set[int],
    player_index: dict[int, PlayerInfo],
) -> dict[int, dict[int, int]]:
    """Tallies cast events for known consumable spell ids, per player.

    Returns player_id -> {ability_id: count}.
    """
    counts: dict[int, dict[int, int]] = {}

    for event in events:
        if event.get("type") not in (None, "cast", "casts"):
            continue
        ability_id = _event_ability_id(event)
        if ability_id not in known_ids:
            continue
        source_id = event.get("sourceID")
        if source_id not in player_index:
            continue
        per_ability = counts.setdefault(source_id, {})
        per_ability[ability_id] = per_ability.get(ability_id, 0) + 1

    return counts


def is_encounter_fight(fight: dict) -> bool:
    """True for a real boss pull (kill or wipe), false for trash — `boss` is
    a nonzero encounter id on a boss fight, 0/absent on trash. Used to scope
    buff-uptime metrics to actual encounter time, excluding the (often
    substantial — roughly half a raid night) trash/travel/repair time
    between pulls that `phase_fight_windows`' merged window otherwise
    includes.
    """
    return bool(fight.get("boss"))


def _band_overlap_ms(band: dict, intervals: list[tuple[int, int]]) -> int:
    return sum(
        max(0, min(band["endTime"], end) - max(band["startTime"], start))
        for start, end in intervals
    )


def aggregate_buff_uptime_ms(
    auras_lists: list[list[dict]],
    player_index: dict[int, PlayerInfo],
    intervals: list[tuple[int, int]] | None = None,
) -> dict[int, int]:
    """Sums buff uptime (ms) per player across one or more single-ability
    `/report/tables/buffs` responses (see `WCLClient.get_buff_uptime`) — one
    list per (window, spell id) pair so several ranks of the same named
    consumable (e.g. every Scroll of Strength rank) land in one total per
    player.

    Without `intervals`, sums each entry's `totalUptime` as reported for the
    queried window. With `intervals` (e.g. a report's encounter-fight
    start/end pairs from `is_encounter_fight`), sums only the portion of each
    `bands` entry that overlaps those intervals instead — lets the uptime
    query itself span a whole merged window (cheap, one call per spell id)
    while still scoping the result to just encounter time, not trash/travel
    time in between.

    Doesn't union overlapping time ranges across different ids/intervals, so
    a player somehow holding two ranks of the same scroll simultaneously
    would be double-counted — disregarded as a practically-never case.
    """
    totals: dict[int, int] = {}
    for auras in auras_lists:
        for entry in auras:
            player_id = _entry_player_id(entry, player_index)
            if player_id is None:
                continue
            if intervals is None:
                ms = entry.get("totalUptime", 0)
            else:
                ms = sum(_band_overlap_ms(band, intervals) for band in entry.get("bands", []))
            totals[player_id] = totals.get(player_id, 0) + ms
    return totals


def merge_player_summary(
    player_index: dict[int, PlayerInfo],
    consumable_counts: dict[int, dict[int, int]],
    consumable_lookup: dict[int, ConsumableInfo],
    role_lookup: dict[int, str],
    extra_metrics: dict[int, dict[str, float]] | None = None,
) -> list[PlayerSummary]:
    """Combines player identity, role, and consumable tallies into the
    final per-player summary list (sorted by name).

    `extra_metrics` (player_id -> {row_label: value}) are merged into each
    player's `consumables` dict alongside the spell-id-derived counts — used
    for derived metrics with no spell id of their own (e.g. "Scrolls Uptime
    %", from `aggregate_buff_uptime_ms`), so excel_export.py/json_export.py
    don't need any special-casing: they already treat every key in
    `consumables` generically.
    """
    extra_metrics = extra_metrics or {}
    summaries: list[PlayerSummary] = []

    for player_id, info in player_index.items():
        raw_counts = consumable_counts.get(player_id, {})

        # Several spell ids can share one display name (e.g. "Super Mana
        # Potion equivalents"), so sum into the name rather than overwrite.
        consumables: dict[str, int] = {}
        for ability_id, count in raw_counts.items():
            consumable_info = consumable_lookup.get(ability_id)
            if consumable_info is None:
                continue
            consumables[consumable_info.name] = consumables.get(consumable_info.name, 0) + count

        consumables.update(extra_metrics.get(player_id, {}))

        summaries.append(
            PlayerSummary(
                name=info.name,
                class_name=info.class_name,
                role=role_lookup.get(player_id, ROLE_UNKNOWN),
                consumables=consumables,
            )
        )

    summaries.sort(key=lambda s: s.name)
    return summaries


def merge_summaries_across_reports(summary_lists: list[list[PlayerSummary]]) -> list[PlayerSummary]:
    """Combines several reports' PlayerSummary lists into one — for a raid
    night that got split into multiple report codes (e.g. a disconnect or
    server reset mid-raid), so it ends up as one sheet instead of double-
    counting consumables across separate per-report sheets.

    Consumable counts are summed by name across reports; role is resolved
    by the same Tank > Healer > DPS priority as a single report's role
    swaps (see ROLE_PRIORITY); class/name are assumed stable across reports
    of the same player.
    """
    class_by_name: dict[str, str] = {}
    roles_seen: dict[str, set[str]] = {}
    consumables_by_name: dict[str, dict[str, int]] = {}

    for summaries in summary_lists:
        for summary in summaries:
            class_by_name[summary.name] = summary.class_name
            roles_seen.setdefault(summary.name, set()).add(summary.role)
            totals = consumables_by_name.setdefault(summary.name, {})
            for consumable, count in summary.consumables.items():
                totals[consumable] = totals.get(consumable, 0) + count

    merged = [
        PlayerSummary(
            name=name,
            class_name=class_by_name[name],
            role=next((r for r in ROLE_PRIORITY if r in roles_seen[name]), ROLE_UNKNOWN),
            consumables=consumables_by_name.get(name, {}),
        )
        for name in class_by_name
    ]
    merged.sort(key=lambda s: s.name)
    return merged
