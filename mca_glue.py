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
    "post_configuration_pending": [],
    "siad_provisioning": [],
    "edp_publish": [],
    "rilinks_scripting": [],
    "script_6673_config": [],
    "port_config_enm": [],
    "link_failure": [],
    "sfp_not_present": [],
    "mo_inconsistent_alarm": [],
    "fiberloss": [],
    "high_rssi": [],
    "low_rssi": [],
    "high_vswr": [],
    "low_vswr": [],
    "vswr_overthreshold": [],
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
        # Confirmed fix: deleted_node is a genuine exception — each deleted node produces
        # its OWN separate scope line (not multiple parts of one line), but has only a
        # single dedicated .xlsm row/column, so multiple deleted nodes must be combined
        # with "|" into that one value rather than silently dropping every node after the
        # first (which is what taking only lines[0] did previously).
        lines = detect_result.get("lines")
        if lines:
            if item_key == "deleted_node" and len(lines) > 1:
                nodes = [ln.split("\t")[-1].strip() for ln in lines if ln.split("\t")[-1].strip()]
                values.append("|".join(nodes))
            else:
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


def _line_to_column_values(line):
    """Extract one 'lines'-style scope line's own tab-separated column values — same
    parsing _result_to_column_values already does for lines[0], factored out so every
    instance (not just the first) can be extracted correctly."""
    parts = [p for p in line.split("\t")]
    parts = parts[1:]  # drop the leading label token
    return [p for p in parts if p.strip().rstrip(":").lower() not in ("from", "to")]


def _combine_overflow_lines(lines):
    """Combine multiple 'lines'-style instances into ONE set of column values, joining
    each column position across instances with '|' (position-paired) — used for the LAST
    available .xlsm row when there are more instances than reserved rows (e.g. Integration
    with 4+ newly-added nodes but only 3 template rows: rows/cols beyond the buffer get
    folded into the final row rather than dropped, matching the same overflow convention
    already used elsewhere (GPS/Transport SFP/deleted_node). The separate plain-text report
    is unaffected — it already lists every instance individually with no row limit."""
    per_instance = [_line_to_column_values(ln) for ln in lines]
    per_instance = [v for v in per_instance if v]
    if not per_instance:
        return []
    n_cols = max(len(v) for v in per_instance)
    combined = []
    for col_i in range(n_cols):
        col_vals = [v[col_i] for v in per_instance if col_i < len(v) and v[col_i]]
        combined.append("|".join(col_vals))
    return combined


def _apply_stakeholder(col_value_pairs, stakeholder):
    """Appends ' (stakeholder)' to the FIRST value column — matching the exact same tag
    build_mca_report_text already applies to the text report line for this item, so the
    .xlsm write stops silently diverging from the report it's meant to mirror. Confirmed
    real gap: the raw scope_line text this function reads was never tagged with a
    stakeholder at the source — only build_mca_report_text's separate rendering step
    added it, so every generic Pending item with a declared stakeholder (retune, FDD
    renaming, controller integration, port conversion, NGS activation, alarm scripting/
    testing, 6673 script load, SAU connections, etc.) silently wrote untagged text into
    the .xlsm."""
    if not stakeholder or not col_value_pairs:
        return col_value_pairs
    col, val = col_value_pairs[0]
    if val:
        col_value_pairs = [(col, f"{val} ({stakeholder})")] + list(col_value_pairs[1:])
    return col_value_pairs


def build_xlsm_row_writes(checklist_results, choices, row_map, stakeholders=None):
    """checklist_results: output of evaluate_checklist(). choices: dict of item key ->
    {"section": "completed"|"pending", "checked": bool, "manual_extra": [...]} — the engineer's
    actual choices from the UI (defaults to checked_by_default / item['section'] if not present).
    stakeholders: item key -> resolved stakeholder string (same dict already built for
    build_mca_report_text's stakeholder_by_key) — applied to Pending items' first value column.
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
        stakeholder = (stakeholders or {}).get(key) if section == "pending" else None

        mapping = row_map[key]
        target_rows = mapping.get(section) if isinstance(mapping, dict) else None
        if not target_rows:
            continue

        per_node_values = choice.get("per_node_manual")
        cols = VALUE_COLUMNS[key]
        if per_node_values:
            nodes = list(per_node_values.keys())
            if len(target_rows) > 1:
                # Transport SFP: 3 real slots, one node per row. Its remaining value columns
                # are fewer than the per-node manual fields (Node ID + ONE combined SFP-models
                # column, but 2 manual fields: BBU End + SIAD End) — join extra values into the
                # last available column rather than silently dropping them.
                for i, row_num in enumerate(target_rows):
                    if i < len(nodes):
                        node = nodes[i]
                        vals = [v for v in per_node_values[node] if v]
                        rest_cols = cols[1:]
                        if len(vals) > len(rest_cols) and rest_cols:
                            vals = vals[:len(rest_cols) - 1] + [" & ".join(vals[len(rest_cols) - 1:])]
                        pairs = _apply_stakeholder([(cols[0], node)] + list(zip(rest_cols, vals)), stakeholder)
                        row_writes.append((row_num, checked, pairs))
                    else:
                        row_writes.append((row_num, False, []))
            else:
                # GPS Installation: only 1 real row in the template — fill the first node's
                # values, note any additional nodes since there's nowhere else to put them
                # (confirmed acceptable: the plain-text report shows all of them individually).
                first_node = nodes[0]
                first_vals = [v for v in per_node_values[first_node] if v]
                node_display = first_node if len(nodes) == 1 else f"{first_node} (+{len(nodes)-1} more \u2014 see report)"
                pairs = _apply_stakeholder([(cols[0], node_display)] + list(zip(cols[1:], first_vals)), stakeholder)
                row_writes.append((target_rows[0], checked, pairs))
            continue

        values = _result_to_column_values(key, item.get("result"), manual_extra)
        col_value_pairs = list(zip(cols, values))

        # multi-slot items: only the FIRST row slot gets checked+filled per single detected
        # instance; if detect() found multiple concrete instances (e.g. several Moved Sectors
        # events), each instance uses the next available slot in order.
        instances = item.get("result", {}).get("lines") if item.get("result") else None
        if instances and len(target_rows) > 1:
            # Confirmed real bug found: every reserved row was writing the SAME first
            # instance's values (col_value_pairs computed once, outside this loop) —
            # rows beyond the first never got their own instance's data at all. Each row
            # now gets its own instance; when there are more instances than reserved rows,
            # the overflow instances combine into the LAST row via _combine_overflow_lines
            # (per-column "|" join) instead of being silently dropped or duplicating row 1.
            n_rows = len(target_rows)
            for i, row_num in enumerate(target_rows):
                if i >= len(instances):
                    row_writes.append((row_num, False, []))
                elif i == n_rows - 1 and len(instances) > n_rows:
                    overflow_values = _combine_overflow_lines(instances[i:])
                    row_writes.append((row_num, checked, _apply_stakeholder(list(zip(cols, overflow_values)), stakeholder)))
                else:
                    row_values = _line_to_column_values(instances[i])
                    row_writes.append((row_num, checked, _apply_stakeholder(list(zip(cols, row_values)), stakeholder)))
        else:
            row_writes.append((target_rows[0], checked, _apply_stakeholder(col_value_pairs, stakeholder)))
            for extra_row in target_rows[1:]:
                row_writes.append((extra_row, False, []))

    return row_writes
