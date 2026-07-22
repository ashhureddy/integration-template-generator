"""
Maps each checklist item to which columns are real VALUE slots to fill vs. fixed structural
labels (like 'From:'/'To:') that must stay untouched — derived directly from the exact column
layout inspected in the real Legacy_MCA template, not guessed generically.
"""

# key -> list of value-column letters (1-indexed column numbers) IN ORDER of the values a
# detect() result / manual entry would supply, for a SINGLE row slot.
VALUE_COLUMNS = {
    "integration": [3, 4],                  # C=band labels, D=Node ID
    "controller_integration": [3],          # C=Controller ID
    "port_conversion": [3],                 # C=Node ID
    "moved_sectors": [3, 5, 7],             # C=Sectors, E=From Node ID, G=To Node ID (D/F are literal 'From:'/'To:')
    "deleted_node": [3],
    "deleted_sector": [3],
    "retune": [3, 5, 7],                    # C=Sector, E=From Freq/BW, G=To Freq/BW
    "fdd_renaming": [3, 5, 8],              # C=Sector, E=Old FDD Name, H=New FDD Name
    "radio_swap": [3, 5, 7],                # C=Sector, E=From Radio, G=To Radio
    "dss_activation": [3],                  # C=Sectors
    "ngs_activation": [3, 4],               # C=Sectors, D=Node ID
    "gps_installation": [3, 5],             # C=Node ID, E=GPS Version (D stays 'Version: ')
    "lkf_installation": [3],                # C=Node ID | Controller ID (combined)
    "psap_moved_lte": [3, 5],               # C=band list, E=PSAP Schedule ID (D stays 'PSAP Schedule ID:')
    "speedtest_new_lte": [3],
    "speedtest_5g": [3],
    "calltest_fnet": [3],
    "transport_sfp": [3, 4],                # C=Node ID, D=SFP models combined
    "sfp_installation": [3, 4],             # C=SFP Type, D=Sector Details
    "ret_configuration": [],                # no value columns, just a checkbox
    "alarm_scripting": [3],
    "alarm_testing": [3],
    "sau_connections": [3],
    "sup_connections": [3],
    "xmu_installation": [3],
    "idl_connections": [],                  # no value columns (Build Type is its own separate row 19)
    "script_6673": [3],
    "installation_generic": [],
    "area_test": [3],                       # C = Node ID | 6610 Controller ID (Pending-only row 146)
}


def _result_to_column_values(item_key, detect_result, manual_extra=None):
    """Turns a detect() result dict into an ordered list of plain values matching
    VALUE_COLUMNS[item_key]'s column order. manual_extra appends any manual-entry values
    (e.g. GPS Version, PSAP Schedule ID) after whatever was auto-filled."""
    values = []
    if detect_result:
        fill = detect_result.get("fill", {})
        if "nodes" in fill:
            values.append("|".join(fill["nodes"]))
        if "controller" in fill and "nodes" not in fill:
            values.append(fill["controller"])
        elif "controller" in fill:
            values[-1] = f"{values[-1]}|{fill['controller']}" if values else fill["controller"]
        if "bands" in fill:
            values.append(fill.get("bands"))

        # "lines"-based items (Integration, 6610 Controller Integration, NGS activation,
        # Moved Sectors, Retune, etc.) — parse the FIRST detected scope line's own
        # tab-separated parts into column values, dropping the leading label token itself
        # (e.g. "Integration:") and any "From:"/"To:" literal tokens that are already fixed
        # in the template (those columns are intentionally skipped in VALUE_COLUMNS).
        lines = detect_result.get("lines")
        if lines:
            parts = [p for p in lines[0].split("\t")]
            parts = parts[1:]  # drop the leading label token
            parts = [p for p in parts if p.strip().rstrip(":").lower() not in ("from", "to")]
            values.extend(parts)

        fdd = detect_result.get("fdd")
        if fdd:
            node, old_name, new_name = fdd[0]
            values.extend([node, old_name, new_name])
    if manual_extra:
        values.extend(manual_extra)
    return values


def build_xlsm_row_writes(checklist_results, choices, row_map):
    """checklist_results: output of evaluate_checklist(). choices: dict of item key ->
    {"section": "completed"|"pending", "checked": bool, "manual_extra": [...]} — the engineer's
    actual choices from the UI (defaults to checked_by_default / item['section'] if not present).
    Returns the row_writes list ready for fill_legacy_mca()."""
    row_writes = []
    for item in checklist_results:
        key = item["key"]
        if key not in row_map or key not in VALUE_COLUMNS:
            continue
        choice = choices.get(key, {})
        checked = choice.get("checked", item["checked_by_default"])
        section = choice.get("section", item["section"])
        manual_extra = choice.get("manual_extra", [])

        mapping = row_map[key]
        target_rows = mapping.get(section) if isinstance(mapping, dict) else None
        if not target_rows:
            continue

        values = _result_to_column_values(key, item.get("result"), manual_extra)
        cols = VALUE_COLUMNS[key]
        col_value_pairs = list(zip(cols, values))

        # multi-slot items: only the FIRST row slot gets checked+filled per single detected
        # instance; if detect() found multiple concrete instances (e.g. several Moved Sectors
        # events), each instance uses the next available slot in order.
        instances = item.get("result", {}).get("lines") if item.get("result") else None
        if instances and len(target_rows) > 1:
            for i, row_num in enumerate(target_rows):
                if i < len(instances):
                    row_writes.append((row_num, checked, col_value_pairs))
                else:
                    row_writes.append((row_num, False, []))
        else:
            row_writes.append((target_rows[0], checked, col_value_pairs))
            for extra_row in target_rows[1:]:
                row_writes.append((extra_row, False, []))

    return row_writes
