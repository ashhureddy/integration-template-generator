"""
MCA Integration Report — full checklist mechanism.

Every item is a dict describing one checkable line. `section` is where it defaults to
("completed" or "pending"). `default_checked` follows the same pre-check-if-detected pattern
established throughout this build. `stakeholder` is the fixed tag shown when an item lives in
Pending (per the real macro's tags: MIC, MIC PM, AT&T, Tower Crew). Items with `toggle=True`
(currently only DSS Activation) let the engineer choose Completed vs Pending, prompting for a
stakeholder (AT&T or MIC) only if placed in Pending.

Each item's `detect(ctx)` function returns either:
  - None                          -> condition not met, item stays unchecked / not shown
  - {"fill": {...}}                -> condition met, with these auto-filled values
  - {}                             -> condition met, nothing to auto-fill (plain trigger)
`ctx` is a dict of everything pre-computed once per report: scope_lines, new_nodes,
board_swaps, fdd_renames, controller_id, controller_in_edp, mm_objs, etc.
"""

import re


def _scope_lines_matching(ctx, *prefixes):
    return [l for l in ctx["scope_lines"] if l.startswith(prefixes)]


def _reword_tab_line(line, sep=" "):
    """'Moved Sectors:\\tLTE_700\\tFrom:\\tA\\tTo:\\tB' -> 'Moved Sectors: LTE_700 from A to B.'
    Generic reformatter matching the confirmed real-sample prose style."""
    parts = [p.strip() for p in line.split("\t") if p.strip()]
    if not parts:
        return ""
    head = parts[0]
    rest = " ".join(p.lower() if p.rstrip(":").lower() in ("from", "to") else p for p in parts[1:])
    return f"{head} {rest}.".replace("  ", " ")


CHECKLIST = [
    # ---------------- Completed: fully auto-detected ----------------
    {"key": "integration", "label": "Integration:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Integration:")} if _scope_lines_matching(ctx, "Integration:") else None},

    {"key": "controller_integration", "label": "6610 Controller Integration:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "6610 Controller Integration:")} if _scope_lines_matching(ctx, "6610 Controller Integration:") else None},

    {"key": "port_conversion", "label": "Port speed 1G to 10G conversion with MPST:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Port speed 1G to 10G conversion")} if _scope_lines_matching(ctx, "Port speed 1G to 10G conversion") else None},

    {"key": "moved_sectors", "label": "Moved Sectors:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Moved Sectors:")} if _scope_lines_matching(ctx, "Moved Sectors:") else None},

    {"key": "deleted_node", "label": "Deleted Node from ENM:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Deleted Node from ENM:")} if _scope_lines_matching(ctx, "Deleted Node from ENM:") else None},

    {"key": "deleted_sector", "label": "Deleted Sector :", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Deleted Sector")} if _scope_lines_matching(ctx, "Deleted Sector") else None},

    {"key": "retune", "label": "Retune on:", "section": "completed", "toggle": True, "stakeholder": "MIC",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Retune on:")} if _scope_lines_matching(ctx, "Retune on:") else None},

    {"key": "fdd_renaming", "label": "FDD Renaming on:", "section": "completed", "toggle": True, "stakeholder": "MIC",
     "detect": lambda ctx: {"fdd": ctx["fdd_renames"]} if ctx["fdd_renames"] else None},

    {"key": "ngs_activation", "label": "NGS activation:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "NGS Activation on")} if _scope_lines_matching(ctx, "NGS Activation on") else None},

    {"key": "idl_connections", "label": "IDL connections", "section": "completed", "toggle": True, "stakeholder": "MIC PM",
     "detect": lambda ctx: {"build_type": ctx.get("idl_build_type")} if ctx.get("idl_build_type") else None},

    # ---------------- Completed: universal toggle (Completed/Pending + stakeholder) ----------------
    {"key": "dss_activation", "label": "DSS Activation:", "section": "completed", "toggle": True,
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "DSS Activation")} if _scope_lines_matching(ctx, "DSS Activation") else None},

    # ---------------- Completed: conditional trigger + auto-fill + manual completion ----------------
    {"key": "gps_installation", "label": "GPS Installation:", "section": "completed", "toggle": True, "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"]}} if ctx["new_nodes"] else None,
     "per_node_manual": ["GPS Version"]},

    {"key": "lkf_installation", "label": "LKF Installation:", "section": "completed", "toggle": True, "stakeholder": "MIC",
     # New node, board swap, and 6610 present are 3 INDEPENDENT triggers (confirmed) — not an
     # AND requiring controller alongside new-node/board-swap. A prior version wrongly required
     # the controller, which meant a genuinely new node with no controller never triggered LKF.
     "detect": lambda ctx: {"fill": {"nodes": (ctx["new_nodes"] + [n for n, _, _ in ctx["board_swaps"]]), "controller": ctx.get("controller_id")}}
                if (ctx["new_nodes"] or ctx["board_swaps"] or ctx.get("controller_id")) else None},

    {"key": "transport_sfp", "label": "Transport SFP Installation on", "pending_label": "Compatible Transport SFP Installation on", "section": "completed", "toggle": True, "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"]}} if ctx["new_nodes"] else None,
     "per_node_manual": ["SFP Model (BBU End)", "SFP Model (SIAD End)"]},

    {"key": "sau_connections", "label": "SAU Connections:", "section": "completed", "toggle": True, "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if ctx.get("controller_id") else {}},

    {"key": "alarm_scripting", "label": "External alarm Scripting on", "section": "completed", "toggle": True, "stakeholder": "MIC",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if (ctx.get("controller_id") and ctx.get("controller_in_edp")) else None},

    {"key": "alarm_testing", "label": "External alarm testing:", "section": "completed", "toggle": True, "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if (ctx.get("controller_id") and ctx.get("controller_in_edp")) else None},

    # ---------------- Completed: Call Test — moved to a dedicated section in
    # mca_report_ui.py (Completed/Pending/Partially Completed), same pattern as NSB.
    # Removed from here to avoid double-rendering.

    # ---------------- Completed: universal, new-node-triggered ----------------
    # Confirmed new trigger: when External alarm testing is Pending (every scripted
    # port locked, NEA pending), Area test also fires with the controller ID
    # included. Controller ID is appended directly into the SAME "nodes" list
    # (rather than the separate "controller" fill key, which renders with " | "
    # spacing) so the pipe-join produces the confirmed exact format with no extra
    # spaces: "Area test: {node}|{controller}." when a node is present, or just
    # "Area test: {controller}." when testing is Pending with no new node at all.
    # Second trigger, same fill shape: SAU disabled on the 6610 itself (per
    # Controller-checks) — independent of whether the External alarms table shows
    # testing as Pending, since SAU and the alarm-scripting table are separate
    # FRUs/checks on the controller.
    {"key": "area_test", "label": "Area test", "section": "pending", "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"] + ([ctx["controller_id"]] if ((ctx.get("testing_section") == "Pending" or ctx.get("sau_placement") == "Pending") and ctx.get("controller_id")) else [])}}
               if (ctx["new_nodes"] or ((ctx.get("testing_section") == "Pending" or ctx.get("sau_placement") == "Pending") and ctx.get("controller_id"))) else None},

    # ---------------- Pending: default-pending items ----------------
    {"key": "radio_swap", "label": "Radio Swap on:", "section": "pending", "stakeholder": "Tower Crew",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Radio Swap on")} if _scope_lines_matching(ctx, "Radio Swap on") else None},

    # ---------------- Fully manual (no auto-detection at all) ----------------
    {"key": "sup_connections", "label": "SUP Connections:", "section": "completed", "toggle": True, "stakeholder": "MIC PM", "manual": True},
    {"key": "xmu_installation", "label": "XMU Installation:", "section": "completed", "toggle": True, "stakeholder": "MIC PM", "manual": True},
    {"key": "sfp_installation", "label": "SFP Installation on", "section": "completed", "toggle": True, "stakeholder": "MIC PM", "manual": True},
    {"key": "ret_configuration", "label": "RET configuration", "section": "completed", "toggle": True, "stakeholder": "Tower Crew", "manual": True},
    {"key": "script_6673", "label": "6673 Script load", "section": "completed", "toggle": True, "stakeholder": "MIC PM", "manual": True},
    {"key": "installation_generic", "label": "Installation", "section": "completed", "manual": True},

    # ---------------- Pending-only items (no Completed counterpart, fixed stakeholder tags) ----------------
    {"key": "post_configuration_pending", "label": "Post Configuration", "section": "pending", "stakeholder": "MIC PM", "manual": True},
    {"key": "siad_provisioning", "label": "SIAD provisioning", "section": "pending", "stakeholder": "AT&T", "manual": True},
    {"key": "edp_publish", "label": "EDP Publish", "section": "pending", "stakeholder": "AT&T", "manual": True},
    {"key": "rilinks_scripting", "label": "Rilinks Scripting", "section": "pending", "stakeholder": "MIC PM", "manual": True},
    {"key": "script_6673_config", "label": "6673 Configuration", "section": "pending", "stakeholder": "AT&T", "manual": True},
    {"key": "port_config_enm", "label": "6673 Port Configuration in ENM", "section": "pending", "stakeholder": "AT&T", "manual": True},
    {"key": "link_failure", "label": "Link failure", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "sfp_not_present", "label": "SFP Not Present", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "mo_inconsistent_alarm", "label": "Mo Inconsistent configuration alarm", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "fiberloss", "label": "Fiberloss (Data Link_1/2)", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "high_rssi", "label": "High RSSI", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "low_rssi", "label": "Low RSSI", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "high_vswr", "label": "High VSWR", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "low_vswr", "label": "Low VSWR", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
    {"key": "vswr_overthreshold", "label": "VSWR overthreshold", "section": "pending", "stakeholder": "Tower Crew", "manual": True},
]


def _line_from_result(item, choice=None):
    """Builds the prose Completed/Pending line(s) for one checked item, reusing whatever
    detect() already found. Multi-instance items (Radio Swap, Moved Sectors, Retune, etc.)
    render EVERY detected instance, not just the first — confirmed bug: a site with 5 separate
    Radio Swap events was only showing 1. Items with per_node_manual (GPS Installation,
    Transport SFP) render ONE LINE PER NODE, since different nodes can have different GPS/SFP
    models — confirmed real gap: these were previously combined into a single shared line."""
    choice = choice or {}
    choice_manual_extra = choice.get("manual_extra")
    result = item.get("result") or {}
    label = item["label"]

    per_node_values = choice.get("per_node_manual")
    if per_node_values:
        out_lines = []
        for node, vals in per_node_values.items():
            filled_vals = [v for v in vals if v]
            suffix = f" | {' | '.join(filled_vals)}" if filled_vals else ""
            out_lines.append(f"{label} {node}{suffix}.")
        return "\n".join(out_lines) if out_lines else f"{label}."

    if result.get("lines"):
        out_lines = []
        for raw_line in result["lines"]:
            parts = [p.strip() for p in raw_line.split("\t") if p.strip()]
            rest = " ".join(parts[1:])
            out_lines.append(f"{parts[0]} {rest}.".replace("  ", " "))
        return "\n".join(out_lines)
    if result.get("fdd"):
        return "\n".join(f"FDD Renaming on: {node} From: {old_name} To: {new_name}." for node, old_name, new_name in result["fdd"])
    fill = result.get("fill", {})
    bits = []
    if fill.get("nodes"):
        bits.append("|".join(fill["nodes"]))
    if fill.get("controller"):
        bits.append(fill["controller"])
    if fill.get("bands"):
        bits.append(fill["bands"])
    if choice_manual_extra:
        bits.extend([b for b in choice_manual_extra if b])
    return f"{label} {' | '.join(bits)}.".strip()


def build_mca_report_text(mm_objs, checklist_results, choices, header_fields, stakeholder_by_key=None):
    """header_fields: {mic, market, status, site_name, fa_code, site_ids, sow, iwm_details,
    pre_configuration, current_configuration, post_configuration, wll_node, controller_id,
    software_version, gs_version, idl_build_type, idle, idly, switch, slot_port}.
    stakeholder_by_key: item key -> chosen stakeholder tag, for items placed in Pending."""
    lines = []
    lines.append("Subject")
    lines.append(f"{header_fields.get('mic','MIC')} | {header_fields.get('market','')} | "
                  f"{header_fields.get('status','')} | {header_fields.get('site_name','')} | "
                  f"{header_fields.get('fa_code','')} | {header_fields.get('site_ids','')} | "
                  f"{header_fields.get('sow','')}")
    lines.append("")
    lines.append("IWM Details:")
    lines.append(header_fields.get("iwm_details", ""))
    lines.append("")
    lines.append("Configuration")
    lines.append(f"Pre Configuration : {header_fields.get('pre_configuration','')}")
    if (header_fields.get("current_configuration") or "").strip():
        lines.append(f"Current Configuration : {header_fields['current_configuration']}")
    lines.append(f"Post Configuration : {header_fields.get('post_configuration','')}")
    if (header_fields.get("wll_node") or "").strip():
        lines.append(f"WLL  node : {header_fields['wll_node']}")
    lines.append(f"6610 Controller : {header_fields.get('controller_id','')}")
    lines.append(f"Software version: {header_fields.get('software_version','')}")
    lines.append(f"GS Version: {header_fields.get('gs_version','')}")
    lines.append("")
    lines.append("IDL Connections")
    if header_fields.get("idl_build_type"):
        lines.append(f"Build Type : {header_fields['idl_build_type']}")
    if (header_fields.get("idle") or "").strip():
        lines.append(header_fields["idle"])
    if (header_fields.get("idly") or "").strip():
        lines.append(header_fields["idly"])
    if (header_fields.get("switch") or "").strip():
        lines.append(header_fields["switch"])
    if (header_fields.get("slot_port") or "").strip():
        lines.append(header_fields["slot_port"])
    lines.append("")

    completed_lines, pending_lines = [], []
    for item in checklist_results:
        key = item["key"]
        choice = choices.get(key, {})
        checked = choice.get("checked", item["checked_by_default"])
        if not checked:
            continue
        section = choice.get("section", item["section"])
        effective_item = item
        if section == "pending" and item.get("pending_label"):
            effective_item = {**item, "label": item["pending_label"]}
        text = _line_from_result(effective_item, choice)
        if section == "pending":
            stakeholder = (stakeholder_by_key or {}).get(key, item.get("stakeholder", ""))
            if stakeholder:
                tagged = "\n".join(f"{ln} ({stakeholder})" for ln in text.split("\n"))
                pending_lines.append(tagged)
            else:
                pending_lines.append(text)
        else:
            completed_lines.append(text)

    lines.append("Completed:")
    lines.extend(completed_lines)
    if (choices.get("additional_completed", {}).get("text") or "").strip():
        lines.append(choices["additional_completed"]["text"])
    lines.append("")
    lines.append("Pending:")
    lines.extend(pending_lines)
    if (choices.get("additional_pending", {}).get("text") or "").strip():
        lines.append(choices["additional_pending"]["text"])
    lines.append("")
    lines.append("Pre-Existing Issues:")
    lines.append(choices.get("pre_existing_issues_text", ""))
    lines.append("")
    lines.append("Notes:")
    for note_key in ("notes_final_port_config", "notes_nr_verified", "notes_cpri_sfp",
                     "notes_no_external_alarms", "notes_mme_config", "notes_monitored", "notes_not_monitored",
                     "notes_ctrl_monitored", "notes_sau_enabled", "notes_testing"):
        if choices.get(note_key, {}).get("checked"):
            text = choices[note_key].get("text", "")
            lines.append(text)
    if (choices.get("notes_generic_text") or "").strip():
        lines.append(choices["notes_generic_text"])

    return "\n".join(lines)


def evaluate_checklist(ctx):
    """Runs every item's detect() against ctx. Returns a list of
    {"key","label","section","stakeholder","checked_by_default","toggle","prompt","manual","result"}."""
    out = []
    for item in CHECKLIST:
        result = None if item.get("manual") or item.get("prompt") else item["detect"](ctx)
        checked_by_default = bool(result) and not item.get("prompt")
        out.append({
            **item,
            "checked_by_default": checked_by_default,
            "result": result,
        })
    return out
