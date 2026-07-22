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

    {"key": "retune", "label": "Retune on:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Retune on:")} if _scope_lines_matching(ctx, "Retune on:") else None},

    {"key": "fdd_renaming", "label": "FDD Renaming on:", "section": "completed",
     "detect": lambda ctx: {"fdd": ctx["fdd_renames"]} if ctx["fdd_renames"] else None},

    {"key": "ngs_activation", "label": "NGS activation:", "section": "completed",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "NGS Activation on")} if _scope_lines_matching(ctx, "NGS Activation on") else None},

    {"key": "idl_connections", "label": "IDL connections", "section": "completed",
     "detect": lambda ctx: {"build_type": ctx.get("idl_build_type")} if ctx.get("idl_build_type") else None},

    # ---------------- Completed: universal toggle (Completed/Pending + stakeholder) ----------------
    {"key": "dss_activation", "label": "DSS Activation:", "section": "completed", "toggle": True,
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "DSS Activation")} if _scope_lines_matching(ctx, "DSS Activation") else None},

    # ---------------- Completed: conditional trigger + auto-fill + manual completion ----------------
    {"key": "gps_installation", "label": "GPS Installation:", "section": "completed",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"]}} if ctx["new_nodes"] else None,
     "manual_fields": ["GPS Version"]},

    {"key": "lkf_installation", "label": "LKF Installation:", "section": "completed",
     "detect": lambda ctx: {"fill": {"nodes": (ctx["new_nodes"] + [n for n, _, _ in ctx["board_swaps"]]), "controller": ctx.get("controller_id")}}
                if (ctx["new_nodes"] or ctx["board_swaps"]) and ctx.get("controller_id") else None},

    {"key": "transport_sfp", "label": "Transport SFP Installation on", "section": "completed",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"]}} if ctx["new_nodes"] else None,
     "manual_fields": ["SFP Model (BBU End)", "SFP Model (SIAD End)"]},

    {"key": "sau_connections", "label": "SAU Connections:", "section": "completed",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if ctx.get("controller_id") else {}},

    {"key": "alarm_scripting", "label": "External alarm Scripting on", "section": "completed",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if (ctx.get("controller_id") and ctx.get("controller_in_edp")) else None},

    {"key": "alarm_testing", "label": "External alarm testing:", "section": "completed",
     "detect": lambda ctx: {"fill": {"controller": ctx.get("controller_id")}} if (ctx.get("controller_id") and ctx.get("controller_in_edp")) else None},

    # ---------------- Completed: field-test sub-section (automatic pair) ----------------
    {"key": "psap_moved_lte", "label": "PSAP test/Speedtest/VoLTE voice calltest:", "section": "completed",
     "detect": lambda ctx: {"fill": {"bands": ctx.get("moved_lte_bands")}} if ctx.get("moved_lte_bands") else None},

    {"key": "calltest_fnet", "label": "Calltest with F-NET SIM:", "section": "completed",
     "detect": lambda ctx: {} if ctx.get("fnet_moved_or_new") else None},

    # ---------------- Completed: conditional PROMPT (only shown if the underlying condition exists) ----------------
    {"key": "psap_new_lte", "label": "PSAP test/Speedtest/VoLTE voice calltest:", "section": "completed", "prompt": True,
     "detect": lambda ctx: {} if ctx.get("new_lte_bands") else None},
    {"key": "speedtest_new_lte", "label": "Speedtest/VoLTE voice calltest:", "section": "completed", "prompt": True,
     "detect": lambda ctx: {} if ctx.get("new_lte_bands") else None},
    {"key": "speedtest_moving_5g", "label": "Speed test: MOVING 5G Sectors", "section": "completed", "prompt": True,
     "detect": lambda ctx: {} if ctx.get("moving_5g_bands_incl_cband") else None},
    {"key": "speedtest_new_5g", "label": "Speed test: Newly adding 5G Sectors", "section": "completed", "prompt": True,
     "detect": lambda ctx: {} if ctx.get("new_5g_bands_excl_cband") else None},
    {"key": "speedtest_new_cband", "label": "Speed test: Newly adding CBAND|DOD", "section": "completed", "prompt": True,
     "detect": lambda ctx: {} if ctx.get("new_cband_dod") else None},

    # ---------------- Completed: universal, new-node-triggered ----------------
    {"key": "area_test", "label": "Area test", "section": "pending", "stakeholder": "MIC PM",
     "detect": lambda ctx: {"fill": {"nodes": ctx["new_nodes"]}} if ctx["new_nodes"] else None},

    # ---------------- Pending: default-pending items ----------------
    {"key": "radio_swap", "label": "Radio Swap on:", "section": "pending", "stakeholder": "Tower Crew",
     "detect": lambda ctx: {"lines": _scope_lines_matching(ctx, "Radio Swap on")} if _scope_lines_matching(ctx, "Radio Swap on") else None},

    # ---------------- Fully manual (no auto-detection at all) ----------------
    {"key": "sup_connections", "label": "SUP Connections:", "section": "completed", "manual": True},
    {"key": "xmu_installation", "label": "XMU Installation:", "section": "completed", "manual": True},
    {"key": "sfp_installation", "label": "SFP Installation on", "section": "completed", "manual": True},
    {"key": "ret_configuration", "label": "RET configuration", "section": "completed", "manual": True},
    {"key": "script_6673", "label": "6673 Script load", "section": "completed", "manual": True},
    {"key": "installation_generic", "label": "Installation", "section": "completed", "manual": True},
]


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
