"""
MCA Integration Report — interactive 'Generate Report' UI section.
Fixes applied: correct 6610 controller ID (Controller Info has TWO columns — 'Controller' is
just the literal string '6610' as a type marker, 'Controller ID' is the real instance name),
live preview of detected values (no longer hidden until Generate), multi-instance items show
ALL detected lines, manual-entry space + stakeholder selector for every item, high-contrast
display of auto-fetched read-only values, and a more organized bordered-card layout.
"""
import streamlit as st
import re
import io
import zipfile

import report_detect
import mca_checklist
import mca_glue
import mca_report_text
import mca_completed_logic as mcl
from mca_row_map import ROW_MAP
from mca_xlsm_surgical import fill_legacy_mca_surgical

from pathlib import Path

TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "Legacy_MCA_Macro_Template_v6_1.xlsm"
STAKEHOLDER_OPTIONS = ["MIC", "MIC PM", "AT&T", "Tower Crew"]


def _get_controller_id(controller_objs):
    """Controller Info has a 'Controller' column that's just the literal string '6610' (a type
    marker) and a SEPARATE 'Controller ID' column with the real instance name (e.g.
    LSPC273360_C001) — confirmed bug: using 'Controller' directly showed the literal '6610'."""
    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    return ctrl_rows[0].get("Controller ID") if ctrl_rows else ""


def _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp, testing_section=None, sau_placement=None):
    new_nodes, board_swaps = report_detect.detect_node_board_changes(app, ciq_wb, mm_objs, precheck_text)
    fdd_renames = report_detect.detect_fdd_renaming(app, ciq_wb)
    return {
        "scope_lines": scope_lines, "new_nodes": new_nodes, "board_swaps": board_swaps,
        "fdd_renames": fdd_renames, "controller_id": controller_id, "controller_in_edp": controller_in_edp,
        "idl_build_type": idl_build_type, "testing_section": testing_section, "sau_placement": sau_placement,
        "moved_lte_bands": None, "fnet_moved_or_new": False, "new_lte_bands": None,
        "moving_5g_bands_incl_cband": None, "new_5g_bands_excl_cband": None, "new_cband_dod": None,
    }


def _detected_preview(item):
    """Plain-language preview of what QUICKIX already found for this item — shown immediately,
    not hidden until Generate."""
    result = item.get("result")
    if not result:
        return None
    if result.get("lines"):
        return " / ".join(l.replace("\t", " ") for l in result["lines"])
    if result.get("fdd"):
        return " / ".join(f"{n}: {o} \u2192 {new}" for n, o, new in result["fdd"])
    fill = result.get("fill", {})
    bits = []
    if fill.get("nodes"):
        bits.append(", ".join(fill["nodes"]))
    if fill.get("controller"):
        bits.append(fill["controller"])
    if fill.get("bands"):
        bits.append(fill["bands"])
    return ", ".join(bits) if bits else None


def _humanize_scope_line(raw_line):
    """Radio Swap's lines use the same raw tab-separated convention as every other
    scope_line in this codebase (meant to flow through mca_report_text.py's automatic
    tab->sentence conversion) — but displaying/injecting them directly here (via
    _checked_group, not the normal per-item mechanism) would show literal tab characters
    instead of a clean sentence. Same conversion algorithm as mca_report_text.py, applied
    explicitly since this bypasses that normal path."""
    parts = [p.strip() for p in raw_line.split("\t") if p.strip()]
    if len(parts) <= 1:
        return raw_line
    rest = " ".join(parts[1:])
    return f"{parts[0]} {rest}.".replace("  ", " ")


def _checked_group(label, lines, key):
    """Renders one checkbox (checked by default) covering a whole group of auto-computed
    lines, with each line shown as a caption underneath while checked. Confirmed fix for a
    real Streamlit bug: a text_area's `value=` default is only honored on that widget's
    FIRST render — once its `key` exists in session_state, later reruns (e.g. triggered by
    typing into an unrelated field) silently ignore any newly recomputed default, so content
    that only becomes known after the user interacts with something else (like Transport
    SFP's manual model fields) could never actually appear. A checkbox+caption recomputes
    fresh on every rerun, so nothing goes stale. Returns (checked, lines_or_empty)."""
    if not lines:
        return False, []
    checked = st.checkbox(label, value=True, key=key)
    if checked:
        for l in lines:
            st.caption(l)
    return checked, (lines if checked else [])


def _simple_item_row(item):
    """Simplified per your instruction: just a checkbox (checked if QUICKIX detected it,
    unchecked otherwise) with the live detected preview underneath. No dropdowns, no
    stakeholder pickers, no per-item manual fields — manual entry is one shared box under
    Completed/Pending instead. Section placement uses whatever the item's own detect()
    already computed; unchecking simply excludes the item from the report entirely."""
    key = item["key"]
    checked = st.checkbox(item["label"], value=item["checked_by_default"], key=f"chk_{key}")
    preview = _detected_preview(item)
    if preview:
        st.caption(f"Detected: {preview}")
    stakeholder = item.get("stakeholder", "").split("|")[0] if item.get("stakeholder") else "MIC PM"
    return {"checked": checked, "section": item["section"], "manual_extra": [], "per_node_manual": {}}, stakeholder


def render(app, ciq_wb, mm_objs, controller_objs, precheck_text, pre_line, post_line, scope_lines,
           postcheck_text="", controller_checks_text="", edp_index=None):
    # Confirmed hard block: at least one CIQ node ID must appear across every
    # uploaded document together — otherwise this is treated as a wrong/mismatched
    # site upload, and the report must not be generated at all.
    is_mismatch, _mismatch_labels = mcl.detect_site_mismatch(
        mm_objs, controller_objs, precheck_text=precheck_text, postcheck_text=postcheck_text, controller_checks_text=controller_checks_text)
    if is_mismatch:
        _doc_names = {"precheck_text": "Pre-checks", "postcheck_text": "Post-checks", "controller_checks_text": "Controller-checks"}
        _bad_docs = [_doc_names.get(l, l) for l in _mismatch_labels]
        _detail = f" The following document(s) don't contain any of this CIQ's node IDs: {', '.join(_bad_docs)}." \
            if _bad_docs else " The uploaded documents don't share a common node ID with each other."
        st.error("Wrong input given: none of this CIQ's node IDs were found together across the uploaded documents." + _detail + " Please confirm you've uploaded the correct files for this site.")
        st.stop()
    st.subheader("Generate Report")

    # Confirmed feedback: manual-entry fields were too visually plain to notice. Targets
    # Streamlit's stable data-testid attributes (not internal CSS class names, which change
    # between versions) — every text_input/text_area in this page is a manual-entry field,
    # so this is a comprehensive fix, not per-field patching.
    st.markdown("""
        <style>
        div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
            border: 3px solid #5B9BD5 !important;
            background-color: #CCE5FF !important;
            color: #0D1B2A !important;
            font-weight: 600 !important;
        }
        div[data-testid="stTextInput"] input:focus, div[data-testid="stTextArea"] textarea:focus {
            border: 3px solid #0D47A1 !important;
            background-color: #ffffff !important;
        }
        div[data-testid="stSelectbox"] > div > div {
            border: 3px solid #5B9BD5 !important; background-color: #CCE5FF !important;
        }
        div[data-testid="stSelectbox"] * {
            color: #0D1B2A !important; font-weight: 600 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    mcl.set_app_module(app)  # bind the already-loaded app module — see mca_completed_logic
                              # docstring for why this can never be a fresh "import app"

    idl_build_type = app.derive_idl_build_type_label(ciq_wb, mm_objs)
    controller_id = _get_controller_id(controller_objs)
    controller_in_edp = bool(controller_id)

    # ---- Real classification, computed once, feeding both the checklist and every
    # verification function below (confirmed this session — reuses qx.classify_carriers
    # directly rather than re-deriving anything). ----
    classification = app.classify_carriers(ciq_wb, mm_objs, precheck_text)
    new_nodes, board_swaps = report_detect.detect_node_board_changes(app, ciq_wb, mm_objs, precheck_text)

    # ---- Retune fix: replace the old sector-dropping "Retune on:" lines with the
    # corrected, sector-tracked version before the checklist ever sees them — this way
    # mca_checklist.py's existing "_scope_lines_matching(ctx, 'Retune on:')" mechanism
    # picks up the FIXED lines automatically, no change needed to the checklist item itself. ----
    retune_events = mcl.classify_retunes_with_sectors(ciq_wb)
    corrected_retune_lines = mcl.format_retunes(retune_events)
    scope_lines = [l for l in scope_lines if not l.startswith("Retune on:")] + corrected_retune_lines

    # ---- Call Test: market-table-driven, confirmed and built this session, replacing the
    # previously dead stub items entirely. Requires categorizing added/moved bands by
    # tech (lte / 5g / cband_dod) via band_label() — reused directly, nothing new derived. ----
    added_bands_by_tech = {"lte": set(), "5g": set(), "cband_dod": set()}
    moved_bands_by_tech = {"lte": set(), "5g": set(), "cband_dod": set()}
    for node, cells in classification.get("added", {}).items():
        for cell in cells:
            label, _sector = app.band_label(cell)
            if label in ("CBAND", "DOD", "DOD_BWE"):
                added_bands_by_tech["cband_dod"].add(label)
            elif label and label.startswith("5G_"):
                added_bands_by_tech["5g"].add(label)
            elif label:
                added_bands_by_tech["lte"].add(label)
    for mv in classification.get("moved", []):
        label, _sector = app.band_label(mv["cell"])
        if label in ("CBAND", "DOD", "DOD_BWE"):
            moved_bands_by_tech["cband_dod"].add(label)
        elif label and label.startswith("5G_"):
            moved_bands_by_tech["5g"].add(label)
        elif label:
            moved_bands_by_tech["lte"].add(label)

    # Confirmed fix: PSAP was reporting the whole band even when only some sectors
    # move (e.g. Beta/Gamma/Epsilon), since moved_bands_by_tech only tracked band
    # labels. Reusing the same per-band sector tracking + whole-band detection
    # already proven in app.py's "Moved Sectors:" line generation, rather than
    # touching call_test_lines() itself (shared with N2E/NSB).
    _WHOLE_BAND_SET = {"Alpha", "Beta", "Gamma"}
    _moved_lte_per_label = {}
    for mv in classification.get("moved", []):
        label, sector = app.band_label(mv["cell"])
        if label and label not in ("CBAND", "DOD", "DOD_BWE") and not label.startswith("5G_") and sector:
            _moved_lte_per_label.setdefault(label, set()).add(sector)
    moved_bands_by_tech["lte"] = set()
    # Confirmed grouping fix: bands sharing the exact same moved-sector set get
    # combined into one entry "[band1/band2/...] sector" rather than each band
    # producing its own separate "{band} {sector}" string.
    _sectorset_to_labels = {}
    for label, sset in _moved_lte_per_label.items():
        is_whole = _WHOLE_BAND_SET <= sset
        if is_whole:
            moved_bands_by_tech["lte"].add(label)
        else:
            _sectorset_to_labels.setdefault(frozenset(sset), []).append(label)
    for sset, labels in _sectorset_to_labels.items():
        sector_names = sorted(sset, key=lambda s: app.SECTOR_ORDER.index(s) if s in app.SECTOR_ORDER else 99)
        labels_fmt = "/".join(sorted(labels))
        moved_bands_by_tech["lte"].add(f"[{labels_fmt}] {', '.join(sector_names)}")

    calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
    market = None
    if calltest_path.exists() and mm_objs:
        prefix_to_market, calltest_rules = mcl.load_calltest_table(calltest_path)
        market = mcl.determine_market(mm_objs[0].get("Node to be built as"), prefix_to_market)
        if market:
            scope_lines = scope_lines + mcl.call_test_lines(
                classification, market, calltest_rules,
                moved_bands_by_tech["lte"], added_bands_by_tech, moved_bands_by_tech)

    # ---- Florida-only: newly added CBAND/DOD/DOD_BWE cells, one per row (93-104), overflow
    # appended to the last row with '|' (confirmed this session — previously untouched).
    # Rendered as a proper Completed checklist checkbox below, not a standalone block. ----
    florida_cells = mcl.florida_newly_added_cells(market, classification)
    florida_rows = mcl.florida_cells_to_rows(florida_cells)

    # ---- Radio Swap: remove the old single toggle-based line entirely (it can no longer
    # represent "some swaps done, some not" correctly) and compute the real split instead.
    # Placement is DETERMINED by Post-checks, confirmed — not a manual toggle anymore. ----
    scope_lines = [l for l in scope_lines if not l.startswith("Radio Swap on:")]
    radio_swap_completed_lines, radio_swap_pending_lines = [], []
    if postcheck_text:
        rs_completed, rs_pending = mcl.classify_radio_swap_placement(precheck_text, postcheck_text, ciq_wb)
        radio_swap_completed_lines = mcl.format_radio_swaps(rs_completed)
        radio_swap_pending_lines = mcl.format_radio_swaps(rs_pending, label_prefix="Radio Swap on:")

    # ---- Port Conversion via board swap: NEW completion path, confirmed against real
    # ECL02586 data — a board swap can itself complete the 1G->10G conversion, which the
    # existing same-board-only logic never recognized. Confirmed: merge ALL completed nodes
    # (same-board AND via-swap) onto ONE line through Port Conversion's existing dedicated
    # checklist row — no separate buffer injection, regardless of how many nodes. ----
    port_conv_swap_completed = []
    if postcheck_text:
        port_conv_swap_completed = mcl.check_port_conversion_via_board_swap(
            ciq_wb, mm_objs, precheck_text, postcheck_text)
        same_board_nodes = {
            l.split("MPST:")[-1].strip().rstrip(".")
            for l in scope_lines if l.startswith("Port speed 1G to 10G conversion with MPST:")
        }
        swap_nodes = {r["node"] for r in port_conv_swap_completed}
        all_completed_nodes = sorted(same_board_nodes | swap_nodes)
        scope_lines = [l for l in scope_lines if not l.startswith("Port speed 1G to 10G conversion with MPST:")]
        if all_completed_nodes:
            scope_lines = scope_lines + [f"Port speed 1G to 10G conversion with MPST: {'|'.join(all_completed_nodes)}."]

    # ---- 6610 cascade: if a 6610 is present/EDP-published but the controller-checks file
    # doesn't confirm alarm scripting, 4 items move to Pending together, no warning. ----
    controller_checks_data = mcl.extract_controller_checks(controller_checks_text) if controller_checks_text else {}
    cascade_fires = mcl.controller_integration_cascade(
        bool(controller_in_edp), controller_checks_data, controller_id)
    sau_placement = mcl.sau_connections_placement(controller_checks_data, controller_id) if controller_id else None
    testing_section, testing_note, _ = mcl.external_alarm_testing_placement(controller_checks_data) \
        if controller_checks_data else (None, None, None)

    # ---- GPS: Installation (new nodes, grouped by Post-checks type), Upgrade (existing
    # nodes, type changed), and the two site-health checks — all fully automatic now that
    # GPS Version comes from Post-checks, no manual entry needed. Computed here (not later)
    # since none of it needs interactive widgets, and Status needs to see it. ----
    gps_extra_completed, gps_extra_pending = [], []
    if postcheck_text:
        post_gps = mcl.extract_gps_status(postcheck_text)
        pre_gps = mcl.extract_gps_status(precheck_text) if precheck_text else {}
        existing_nodes = [row.get("Node to be built as") for row in mm_objs
                           if row.get("Node to be built as") not in new_nodes]
        gps_dedicated, gps_overflow = mcl.gps_installation_lines(new_nodes, post_gps)
        if gps_dedicated:
            gps_extra_completed.append(gps_dedicated)
        gps_extra_completed += gps_overflow
        gps_extra_completed += mcl.gps_upgrade_lines(existing_nodes, pre_gps, post_gps)

        gps_unconfirmed = mcl.gps_unconfirmed_type_check(mm_objs, post_gps)
        if gps_unconfirmed:
            gps_extra_pending.append(gps_unconfirmed)
        post_sync = mcl.extract_sync_status_2(postcheck_text)
        gps_sync_line = mcl.gps_sync_disabled_check(mm_objs, post_sync)
        if gps_sync_line:
            gps_extra_pending.append(f"{gps_sync_line} ({mcl.gps_pending_stakeholder(market)})")

    # ---- EDP Publish fallback: 6610 present in CIQ but NOT published in EDP -> this
    # replaces the old generic message entirely; 6610 Controller Integration does NOT
    # appear in Completed at all in this case. ----
    edp_publish_text = ""
    if controller_id and not controller_in_edp:
        edp_publish_text = mcl.edp_publish_line(
            mm_objs[0].get("Node to be built as") if mm_objs else "", controller_id, "")

    # ---- Current Configuration: only populated when Post-checks actually differs from
    # the CIQ target (equipment still missing) — confirmed rule, built this pass. ----
    current_config_auto = mcl.current_configuration_line(
        ciq_wb, mm_objs, postcheck_text, derive_identity_from_checks=True) if postcheck_text else ""

    # ---- FDD Renaming, corrected: band-label grouping instead of raw ungrouped cell
    # tuples — confirmed gap, built this pass. ----
    fdd_lines_fixed = mcl.fdd_renaming_lines(ciq_wb)

    ctx = _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp, testing_section, sau_placement)
    results = mca_checklist.evaluate_checklist(ctx)

    if cascade_fires:
        # Force these 5 items to Pending, drop them from wherever the normal detection put
        # them, no warning per confirmed decision. (Confirmed gap: sau_connections was
        # missing from this set — only 4 of the 5 real cascade items were being forced.)
        cascade_keys = {"controller_integration", "alarm_scripting", "lkf_installation",
                         "alarm_testing", "sau_connections"}
        for item in results:
            if item["key"] in cascade_keys:
                item["section"] = "pending"
                item["checked_by_default"] = True
    if sau_placement and not cascade_fires:
        for item in results:
            if item["key"] == "sau_connections":
                item["section"] = "completed" if sau_placement == "Completed" else "pending"
                item["checked_by_default"] = True
    sau_disabled_on_6610 = (sau_placement == "Pending") and not cascade_fires
    if testing_section and not cascade_fires:
        for item in results:
            if item["key"] == "alarm_testing":
                item["section"] = "completed" if testing_section == "Completed" else "pending"
                item["checked_by_default"] = True
    # SAU disabled on the 6610 specifically (per Controller-checks) — a narrower case than
    # the full cascade above (alarm scripting itself may still be confirmed, and the
    # External alarms table above may independently say Completed). Per confirmed rule:
    # a disabled SAU alone still pushes External alarm testing and Area test to Pending,
    # since it means testing/area verification can't have genuinely happened via the
    # controller — SAU may still be enabled on the node instead, tracked separately in
    # Notes ("SAU enabled on : {Node ID}"). Applied AFTER the testing_section block above
    # so it wins regardless of what that block set.
    if sau_disabled_on_6610:
        for item in results:
            if item["key"] in {"alarm_testing", "area_test"}:
                item["section"] = "pending"
                item["checked_by_default"] = True
    if edp_publish_text:
        # Suppress the checklist's own 6610 Controller Integration item entirely when the
        # EDP Publish fallback applies — confirmed: it goes to Pending via EDP Publish
        # instead, not shown in Completed at all.
        for item in results:
            if item["key"] == "controller_integration":
                item["checked_by_default"] = False
    if fdd_lines_fixed:
        # Suppress the checklist's own FDD Renaming item — it uses report_detect's
        # ungrouped per-cell tuples; the corrected, band-label-grouped version is
        # injected separately via extra_completed_text instead.
        for item in results:
            if item["key"] == "fdd_renaming":
                item["checked_by_default"] = False
    # Suppress the checklist's own NGS activation item unconditionally — confirmed change:
    # was auto-Completed with no toggle; now needs a per-pair 3-way choice
    # (Completed/Pending/Pre-Existing), replaced by the custom section below.
    for item in results:
        if item["key"] == "ngs_activation":
            item["checked_by_default"] = False
    # Suppress the checklist's own DSS Activation item unconditionally — confirmed real
    # gap: it originally had toggle=True (Completed/Pending + AT&T/MIC stakeholder if
    # Pending), but that got silently lost when the UI was simplified to a plain checkbox.
    # Replaced by the custom section below, matching the original spec exactly.
    for item in results:
        if item["key"] == "dss_activation":
            item["checked_by_default"] = False
    # Suppress the checklist's own LKF Installation item unconditionally — confirmed bug:
    # it was showing alongside the new per-node Completed/Pending dropdown section as a
    # confusing duplicate. The custom section (built this pass) fully replaces it, since it
    # covers the same 3 original triggers plus the 4th (single-tech node gaining a second
    # tech) the old item never had.
    for item in results:
        if item["key"] == "lkf_installation":
            item["checked_by_default"] = False
    # Same confirmed bug found for Transport SFP — never suppressed, only triggered on
    # new_nodes in the old item, so it could show as a duplicate alongside the new
    # multi-trigger (new node/Port Conversion/board swap) section on any site where
    # new_nodes happened to be non-empty.
    for item in results:
        if item["key"] == "transport_sfp":
            item["checked_by_default"] = False
    # Same confirmed mechanism as NSB/N2E, ported here — MCA's own checklist items for
    # these two ("sup_connections"/"xmu_installation") were plain manual toggles with no
    # auto-detection at all. Suppressed in favor of the auto-detected group built below.
    for item in results:
        if item["key"] in {"sup_connections", "xmu_installation"}:
            item["checked_by_default"] = False
    # Same fix, same reason: these two never had a working way to actually be selected.
    for item in results:
        if item["key"] in {"ret_configuration", "idl_connections"}:
            item["checked_by_default"] = False

    # ---- SUP / XMU auto-detection — confirmed same mechanism as NSB/N2E, reusing the
    # same shared helpers (nodes_expecting_sup/xmu, _hardware_component_state) rather than
    # re-deriving anything. SUP Connections is expected per-node (5216 or XMU present in
    # that node's own CIQ target hardware, not site-wide). XMU Installation is Completed
    # only when BOTH the CIQ target confirms XMU AND Post-checks' Hardware Status shows it
    # ENABLED; a node expected to have XMU but missing from Post-checks entirely (not just
    # DISABLED) still reports Pending rather than being silently dropped. ----
    xmu_present_in_ciq = mcl.xmu_in_ciq(post_line)
    all_site_nodes = {row.get("Node to be built as") for row in mm_objs if row.get("Node to be built as")}
    sup_completed_lines, sup_pending_lines = [], []
    xmu_completed_lines, xmu_pending_lines = [], []
    if postcheck_text:
        sup_expecting_nodes = mcl.nodes_expecting_sup(mm_objs, ciq_wb) & all_site_nodes
        if sup_expecting_nodes:
            sup_found = mcl._hardware_component_state(postcheck_text, "SUP")
            sup_state = {n: s for n, s in sup_found.items() if n in sup_expecting_nodes}
            sup_missing = sup_expecting_nodes - set(sup_found.keys())
            for node, state in sup_state.items():
                (sup_completed_lines if state == "ENABLED" else sup_pending_lines).append(
                    f"SUP Connections: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
            for node in sorted(sup_missing):
                sup_pending_lines.append(f"SUP Connections: {node} (MIC PM)")

        if xmu_present_in_ciq:
            xmu_state = mcl._hardware_component_state(postcheck_text, "XMU")
            for node, state in xmu_state.items():
                (xmu_completed_lines if state == "ENABLED" else xmu_pending_lines).append(
                    f"XMU Installation: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
            xmu_expected_nodes = mcl.nodes_expecting_xmu(mm_objs, ciq_wb) & all_site_nodes
            for node in sorted(xmu_expected_nodes - set(xmu_state.keys())):
                xmu_pending_lines.append(f"XMU Installation: {node} (MIC PM)")

    # ---- Warnings tab collection: every verification function feeds here. Confirmed
    # design: warning-only for most items (never touches report placement), except
    # Radio Swap (which changes Completed/Pending placement itself) and the
    # board-swap-triggered Transport SFP case (which is both Pending AND a warning). ----
    warnings = []
    if postcheck_text:
        warnings += mcl.verify_integration_against_postcheck(classification, postcheck_text)
        warnings += mcl.verify_moved_sectors_against_postcheck(classification, postcheck_text, ciq_wb)
        warnings += mcl.verify_deleted_sectors_against_postcheck(classification, postcheck_text)
        warnings += mcl.verify_retune_against_checks(ciq_wb, retune_events, postcheck_text)
        if edp_index:
            warnings += mcl.verify_port_conversion_against_postcheck(
                ciq_wb, mm_objs, precheck_text, postcheck_text, edp_index)
        if gps_unconfirmed:
            warnings.append({"type": "gps_unconfirmed_type", "text": gps_unconfirmed})

    site_ids = "/".join(r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as"))
    fa_code = ""
    if "5G Info" in ciq_wb.sheetnames:
        for row in app.sheet_objs(ciq_wb["5G Info"]):
            if app.is_populated(row.get("FA Code")):
                fa_code = row.get("FA Code")
                break

    # ---- Status: ATP only if there are truly no Pending items ANYWHERE — confirmed bug,
    # this used to only look at the base checklist `results`, missing every new Pending
    # source added this session (GPS, Radio Swap, EDP Publish, the locked-alarm-ports Notes
    # case). LKF/Transport SFP choices are made via widgets further down and can't be known
    # yet at this point in the render — same limitation Streamlit's rerun model imposes on
    # any interactive choice made after this line; those will be reflected once the engineer
    # interacts with those widgets and the script reruns. ----
    has_pending = (
        any(r["section"] == "pending" and r["checked_by_default"] for r in results)
        or bool(gps_extra_pending) or bool(radio_swap_pending_lines)
        or bool(edp_publish_text) or bool(testing_note)
    )
    default_status = "IX-STF" if has_pending else "IX-ATP"

    with st.container(border=True):
        st.markdown("**Subject**")
        c = st.columns(7)
        with c[0]: st.markdown(f"MIC\n\n**MIC**")
        with c[1]: market_subject_input = st.selectbox("\U0001F4DD Market", ["\u2014 Select \u2014", "MNS", "AT&T"], key="rpt_market")
        with c[2]:
            status = st.selectbox("\U0001F4DD Status", ["\u2014 Select \u2014", "IX-STF", "IX-ATP"],
                                   index=["\u2014 Select \u2014", "IX-STF", "IX-ATP"].index(default_status), key="rpt_status")
        status = "" if status == "\u2014 Select \u2014" else status
        market_subject_input = "" if market_subject_input == "\u2014 Select \u2014" else market_subject_input
        with c[3]: site_name = st.text_input("\U0001F4DD Site Name", key="rpt_site_name")
        with c[4]: st.markdown(f"FA CODE\n\n**{fa_code or '(not found)'}**")
        with c[5]: st.markdown(f"Site ID's\n\n**{site_ids}**")
        with c[6]: sow = st.text_input("\U0001F4DD SOW", key="rpt_sow")

    with st.container(border=True):
        st.markdown("**IWM Details**")
        iwm_details = st.text_input("IWM Details", key="rpt_iwm", label_visibility="collapsed")

    with st.container(border=True):
        st.markdown("**Configuration**")
        st.markdown(f"Pre Configuration : **{pre_line}**")
        st.markdown(f"Post Configuration : **{post_line}**")
        st.markdown(f"6610 Controller : **{controller_id or '(none detected)'}**")
        if current_config_auto:
            current_config = st.text_input(
                "\U0001F4DD Current Configuration \u2014 Post-checks differs from CIQ target, review/edit:",
                value=current_config_auto, key="rpt_current_config")
        else:
            current_config = ""
        # Confirmed: auto-fill WLL node from the CIQ when detectable (same convention
        # as N2E — node name ending in "L"), but stays a fully editable manual field
        # otherwise — engineer can always override/clear it regardless of detection.
        _mca_wll_detected = [row.get("Node to be built as") for row in mm_objs
                              if row.get("Node to be built as") and str(row.get("Node to be built as")).strip().upper().endswith("L")]
        c1, c2 = st.columns(2)
        with c1:
            wll_node = st.text_input("\U0001F4DD WLL node (if applicable)", value="|".join(_mca_wll_detected), key="rpt_wll")
        with c2:
            software_version = st.text_input("\U0001F4DD Software version", key="rpt_sw")
            gs_version = st.text_input("\U0001F4DD GS Version", key="rpt_gs")

    idl_completed_line, idl_pending_line = None, None
    idle = idly = switch = slot_port = ""
    if len(mm_objs) > 1:
        with st.container(border=True):
            st.markdown(f"**IDL Connections** \u2014 Build Type: **{idl_build_type or '(not detected)'}**")
            # Confirmed same as NSB/N2E: IDL connections status is a real user choice
            # (Completed/Pending), not auto-detected — the generic checklist item's
            # detect() only ever knew the build type, never actually asked this.
            idl_choice = st.selectbox("IDL connections status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                        key="mca_idlconn")
            if idl_choice == "Completed":
                idl_completed_line = "IDL connections"
            elif idl_choice == "Pending":
                idl_pending_line = "IDL connections (MIC PM)"
            c1, c2 = st.columns(2)
            with c1:
                idle = st.text_area("\U0001F4DD IDLe cable details (manual)", key="rpt_idle", height=60)
            with c2:
                idly = st.text_area("\U0001F4DD IDLy cable details (manual)", key="rpt_idly", height=60)

            sidehaul_rows = mcl.sidehaul_display_rows(ciq_wb)
            if sidehaul_rows:
                st.caption("Switch / Slot-Port — auto-filled from Sidehaul Info, Cable part number is manual:")
                cable_pns = {}
                for i, srow in enumerate(sidehaul_rows):
                    sc1, sc2, sc3, sc4, sc5 = st.columns([1, 1, 1, 1, 1])
                    with sc1: st.caption(f"**{srow['switch_type']}**")
                    with sc2: st.caption(srow["switch_id"])
                    with sc3: st.caption(srow["slot_port"])
                    with sc4: cable_pns[i] = st.text_input("Cable P/N", key=f"cable_pn_{i}", label_visibility="collapsed", placeholder="Cable part number")
                    with sc5: st.caption(srow["node_id"])
                switch = "\n".join(mcl.format_sidehaul_lines(sidehaul_rows, cable_pns))
                slot_port = ""  # folded into the switch lines above — same source, one combined display
            else:
                switch = st.text_area("\U0001F4DD Switch details (manual)", key="rpt_switch", height=60)
                slot_port = st.text_area("\U0001F4DD Slot/Port/Cable/Node ID (manual)", key="rpt_slotport", height=60)

    # ---- RET configuration: Completed/Pending choice, confirmed same as NSB — the
    # generic checklist item was "manual": True with no way to actually select it at
    # all (filtered out of both Completed/Pending lists entirely, since manual items
    # never get checked_by_default=True). ----
    ret_completed_line, ret_pending_line = None, None
    ret_choice = st.selectbox("RET configuration", ["\u2014 Select \u2014", "Completed", "Pending"], key="mca_ret")
    if ret_choice == "Completed":
        ret_completed_line = "RET configuration"
    elif ret_choice == "Pending":
        ret_pending_line = "RET configuration (Tower Crew)"

    # ---- DSS Activation: Completed/Pending choice, stakeholder (AT&T/MIC) prompt only if
    # Pending — confirmed real gap, restoring the original spec exactly. No default —
    # same "require an active pick" pattern used elsewhere. Confirmed new rule: DSS bands
    # that are ALSO scripted/locked (configuredMaxTxPower=5) go DIRECTLY to Pending with
    # stakeholder AT&T, bypassing the choice — except in NTX market, where those bands
    # aren't reported for DSS at all. Remaining bands still go through the normal choice.
    dss_raw_lines = [l for l in scope_lines if l.startswith("DSS Activation")]
    dss_completed_line, dss_pending_line = None, None
    dss_pending_bands_combined = set()  # confirmed fix: tracked directly, not parsed back out of display text later
    if dss_raw_lines:
        dss_line = dss_raw_lines[0].replace("\t", " ")
        dss_all_bands = set(dss_raw_lines[0].split("\t")[1].split(" & ")) if "\t" in dss_raw_lines[0] else set()
        mca_scripted_locked = mcl.scripted_locked_bands(ciq_wb)
        auto_pending_bands, user_choice_bands = mcl.split_dss_bands_by_scripted_locked(
            dss_all_bands, mca_scripted_locked, market)

        if auto_pending_bands or user_choice_bands:
            with st.container(border=True):
                st.markdown(f"**DSS Activation** \u2014 detected. Pick a status (required):")
                if auto_pending_bands:
                    auto_bands_fmt = " & ".join(mcl.sort_bands_lte_first(auto_pending_bands))
                    dss_pending_auto = f"DSS Activation: {auto_bands_fmt} (AT&T)"
                    st.caption(f"\u26a0\ufe0f Scripted/locked \u2014 goes directly to Pending: {dss_pending_auto}")
                    dss_pending_line = dss_pending_auto
                    dss_pending_bands_combined |= auto_pending_bands
                if user_choice_bands:
                    user_bands_fmt = " & ".join(mcl.sort_bands_lte_first(user_choice_bands))
                    st.caption(f"Remaining band(s) \u2014 pick Completed or Pending: {user_bands_fmt}")
                    dss_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                                key="dss_status", label_visibility="collapsed")
                    if dss_choice == "Completed":
                        dss_completed_line = f"DSS Activation: {user_bands_fmt}"
                    elif dss_choice == "Pending":
                        dss_stakeholder = st.selectbox("Stakeholder", ["\u2014 Select \u2014", "AT&T", "MIC"],
                                                         key="dss_stakeholder")
                        if dss_stakeholder != "\u2014 Select \u2014":
                            user_pending_line = f"DSS Activation: {user_bands_fmt} ({dss_stakeholder})"
                            dss_pending_line = (dss_pending_line + " | " + user_pending_line) if dss_pending_line else user_pending_line
                            dss_pending_bands_combined |= user_choice_bands

    # ---- NGS activation: per-pair 3-way choice (Completed/Pending/Pre-Existing), confirmed
    # change from the old auto-Completed-only behavior. No default — same "require an active
    # pick" pattern as LKF. Pre-Existing doesn't appear in Completed/Pending at all, it only
    # adds a Notes line. ----
    ngs_raw_lines = [l for l in scope_lines if l.startswith("NGS Activation on")]
    ngs_completed_lines, ngs_pending_lines, ngs_notes_lines = [], [], []
    if ngs_raw_lines:
        with st.container(border=True):
            st.markdown(f"**NGS activation** \u2014 {len(ngs_raw_lines)} pair(s) detected. "
                        f"Pick a status for each (required):")
            for i, line in enumerate(ngs_raw_lines):
                parts = [p.strip() for p in line.split("\t") if p.strip()]
                bands = parts[1] if len(parts) > 1 else ""
                pair = parts[2] if len(parts) > 2 else ""
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.caption(f"{bands}  {pair}")
                with c2:
                    pick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending", "Pre-Existing"],
                                          key=f"ngs_{i}", label_visibility="collapsed")
                if pick == "Completed":
                    ngs_completed_lines.append(f"NGS Activation on : {bands}  {pair}")
                elif pick == "Pending":
                    ngs_pending_lines.append(f"NGS Activation on : {bands}  {pair} (MIC)")
                elif pick == "Pre-Existing":
                    ngs_notes_lines.append(f"Pre existing NGS found activated on : {pair}")
            still_needed = sum(1 for i in range(len(ngs_raw_lines))
                                if st.session_state.get(f"ngs_{i}", "\u2014 Select \u2014") == "\u2014 Select \u2014")
            if still_needed > 0:
                st.caption(f"\u26a0\ufe0f {still_needed} NGS pair(s) still need a status pick \u2014 "
                           f"they won't appear in the report until selected.")

    # ---- LKF Installation: Node(s) and Controller are independent installation points
    # (confirmed — one can be Completed while the other is Pending), so each gets its own
    # dropdown. No default on either — leaving one untouched excludes just that entity. ----
    lkf_nodes = mcl.lkf_node_triggers(new_nodes, board_swaps, precheck_text, mm_objs)
    lkf_controller_needed = mcl.lkf_controller_triggered(controller_id)
    lkf_choices = {}
    lkf_controller_choice = None
    emergency_unlock_notes = []
    if lkf_nodes or lkf_controller_needed:
        with st.container(border=True):
            st.markdown(f"**LKF Installation** \u2014 pick Completed or Pending for each "
                        f"(required — left blank = excluded from the report). Node and Controller "
                        f"are tracked independently — one can be done while the other isn't:")
            for node in lkf_nodes:
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.caption(node)
                with c2:
                    pick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                          key=f"lkf_{node}", label_visibility="collapsed")
                    if pick != "\u2014 Select \u2014":
                        lkf_choices[node] = pick
                if pick == "Pending":
                    eu = st.selectbox(f"Emergency unlock activated on {node}? LKF is pending.",
                                        ["\u2014 Select \u2014", "No", "Yes"], key=f"lkf_eu_{node}")
                    if eu == "Yes":
                        emergency_unlock_notes.append(node)
            if lkf_controller_needed:
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.caption(f"{controller_id} (controller)")
                with c2:
                    cpick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                           key="lkf_controller", label_visibility="collapsed")
                    if cpick != "\u2014 Select \u2014":
                        lkf_controller_choice = cpick

    if cascade_fires:
        # Cascade forces LKF's 6610/CONTROLLER portion specifically to Pending — confirmed:
        # this is about the controller entity, not whichever nodes the engineer marked
        # Completed. Node choices are untouched.
        lkf_controller_choice = "Pending"

    lkf_lines_by_section = mcl.lkf_lines_by_choice(lkf_choices, lkf_controller_choice, controller_id) \
        if (lkf_choices or lkf_controller_choice) else {}
    still_needed = len(lkf_nodes) - len(lkf_choices) + (1 if lkf_controller_needed and not lkf_controller_choice else 0)
    if still_needed > 0:
        st.caption(f"\u26a0\ufe0f {still_needed} LKF item(s) still need a Completed/Pending pick "
                   f"\u2014 they won't appear in the report until selected.")

    # ---- Transport SFP: 3 independent triggers now — new node OR Port Conversion OR any
    # board swap (confirmed, added this pass — a physical board replacement often needs a
    # new SFP transceiver regardless of whether port speed itself changed). Uses the same
    # broad hw-string comparison (report_detect.detect_node_board_changes) LKF's board-swap
    # trigger already uses — confirmed this catches same-generation model changes (e.g.
    # 5216->6630, both G2) that the narrower generation-based Port Conversion check misses. ----
    port_conv_nodes = sorted({l.split("MPST: ")[-1].rstrip(".") for l in scope_lines
                               if l.startswith("Port speed 1G to 10G conversion with MPST:")}
                              | {r["node"] for r in port_conv_swap_completed})
    board_swap_node_set = {n for n, _p, _q in board_swaps} if board_swaps and \
        isinstance(board_swaps[0], tuple) else set()
    sfp_trigger_nodes = sorted(set(new_nodes) | set(port_conv_nodes) | board_swap_node_set)
    sfp_models_by_node = {}
    # Confirmed fix: transport_sfp_data moved ahead of the UI block below so BBU End
    # can auto-fetch from it (ericssonprod column, matched per node) — same pattern
    # already confirmed and working for N2E/NSB. SIAD End stays fully manual.
    transport_sfp_data = {}
    if postcheck_text:
        transport_sfp_data = mcl.extract_transport_sfp(postcheck_text)
    if sfp_trigger_nodes:
        with st.container(border=True):
            st.markdown(f"**Transport SFP Installation** \u2014 {len(sfp_trigger_nodes)} node(s). SFP models are manual:")
            for node in sfp_trigger_nodes:
                c1, c2, c3 = st.columns([1, 1, 1])
                with c1:
                    st.caption(node)
                with c2:
                    _bbu_auto = transport_sfp_data.get(node, {}).get("ericssonprod", "")
                    bbu = st.text_input("BBU End", value=_bbu_auto, key=f"sfp_bbu_{node}", label_visibility="collapsed", placeholder="SFP Model (BBU End)")
                with c3:
                    siad = st.text_input("SIAD End", key=f"sfp_siad_{node}", label_visibility="collapsed", placeholder="SFP Model (SIAD End)")
                sfp_models_by_node[node] = (bbu, siad)
    transport_sfp_lines = mcl.transport_sfp_installation_lines(sfp_trigger_nodes, sfp_models_by_node) if sfp_trigger_nodes else []

    sfp_pending_extra, sfp_pre_existing_extra = [], []
    if postcheck_text:
        board_swap_names = [n for n, _p, _q in board_swaps] if board_swaps else []
        sfp_pending_extra, sfp_warnings, sfp_pre_existing_extra = mcl.transport_sfp_verification(
            ciq_wb, mm_objs, new_nodes, board_swap_names, postcheck_text, transport_sfp_data)
        warnings += sfp_warnings

    # ---- Filter the checklist to ONLY show items that actually apply (confirmed decision:
    # whatever doesn't apply is simply not shown, not an empty unchecked row). ----
    completed_items = [i for i in results if i["section"] == "completed" and i["checked_by_default"]]
    pending_items = [i for i in results if i["section"] == "pending" and i["checked_by_default"]]

    st.markdown("### Which of these apply?")
    st.caption("Only items QUICKIX actually detected are shown. Uncheck anything that doesn't apply; use the manual boxes below for anything not auto-detected.")

    choices, stakeholders = {}, {}

    with st.expander(f"Completed ({sum(1 for i in completed_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(completed_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice

        gps_c_checked, gps_c_lines = _checked_group("GPS Installation / Upgrade", gps_extra_completed, "chk_gps_c")
        sfp_c_checked, sfp_c_lines = _checked_group("Transport SFP Installation on", transport_sfp_lines, "chk_sfp_c")
        rs_c_checked, rs_c_display = _checked_group(
            "Radio Swap on", [_humanize_scope_line(l) for l in radio_swap_completed_lines], "chk_rs_c")
        fdd_c_checked, fdd_c_lines = _checked_group("FDD Renaming on", fdd_lines_fixed, "chk_fdd_c")
        lkf_c_group_lines = [lkf_lines_by_section["Completed"]] if lkf_lines_by_section.get("Completed") else []
        lkf_c_checked, lkf_c_lines = _checked_group("LKF Installation", lkf_c_group_lines, "chk_lkf_c")
        sup_c_checked, sup_c_lines = _checked_group("SUP Connections", sup_completed_lines, "chk_sup_c")
        xmu_c_checked, xmu_c_lines = _checked_group("XMU Installation", xmu_completed_lines, "chk_xmu_c")

        florida_checked = False
        if florida_rows:
            florida_checked = st.checkbox(f"Newly added Cells (Florida market) \u2014 {len(florida_cells)} cell(s)",
                                            value=True, key="chk_florida_cells")
            if florida_checked:
                for r in florida_rows:
                    st.caption(r)
        florida_active_rows = florida_rows if florida_checked else []
        additional_completed = st.text_area("\U0001F4DD Enter any additional completed information that needs to be added in report",
                                             key="rpt_add_completed", height=100)
        choices["additional_completed"] = {"text": additional_completed}

    with st.expander(f"Pending ({sum(1 for i in pending_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(pending_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice
                stakeholders[item["key"]] = stakeholder

        gps_p_checked, gps_p_lines = _checked_group("GPS-related Pending items", gps_extra_pending, "chk_gps_p")
        sfp_p_checked, sfp_p_lines = _checked_group("Transport SFP (Pending)", sfp_pending_extra, "chk_sfp_p")
        rs_p_checked, rs_p_display = _checked_group(
            "Radio Swap on (Pending)", [_humanize_scope_line(l) for l in radio_swap_pending_lines], "chk_rs_p")
        edp_group_lines = [edp_publish_text] if edp_publish_text else []
        edp_checked, edp_lines = _checked_group("EDP Publish", edp_group_lines, "chk_edp")
        lkf_p_group_lines = [lkf_lines_by_section["Pending"]] if lkf_lines_by_section.get("Pending") else []
        lkf_p_checked, lkf_p_lines = _checked_group("LKF Installation (Pending)", lkf_p_group_lines, "chk_lkf_p")
        sup_p_checked, sup_p_lines = _checked_group("SUP Connections (Pending)", sup_pending_lines, "chk_sup_p")
        xmu_p_checked, xmu_p_lines = _checked_group("XMU Installation (Pending)", xmu_pending_lines, "chk_xmu_p")

        additional_pending = st.text_area("\U0001F4DD Enter any additional pending information that needs to be reported to Market",
                                           key="rpt_add_pending", height=100)
        choices["additional_pending"] = {"text": additional_pending}

    locked_ports_list = [p for p in (controller_checks_data.get("alarm_ports", []) if controller_checks_data else [])
                          if p["admin"] == "LOCKED" and p["slogan"]]
    # Confirmed fix, same pattern as NSB: when EVERY scripted port is locked (NEA
    # pending), the simple "All external alarms are kept locked, due to NEA is
    # pending." note already covers this completely — asking the engineer to
    # individually classify all of them (Pre-existing, active alarm, loops, etc.) is
    # redundant and unnecessary busywork in that specific case.
    locked_ports_exist = bool(locked_ports_list) and testing_section != "Pending"

    bucket_pre_existing, bucket_pending, bucket_notes = [], [], []
    if locked_ports_exist:
        with st.container(border=True):
            st.markdown(f"**Locked alarm ports** \u2014 {len(locked_ports_list)} scripted port(s) detected LOCKED "
                        f"in the controller-checks file:")
            for p in locked_ports_list:
                st.caption(f"Port {p['port']} \u2014 {p['slogan']} ({p['severity']})")
            st.markdown("Classify each one below (per the confirmed 6610 Alarm Cutover reporting "
                        "standard). Leave blank whichever don't apply.")
            b1 = st.text_input("\U0001F4DD 1. Pre-existing locked \u2014 port numbers", key="lp_b1", placeholder="e.g. 1, 5, 25")
            b2 = st.text_input("\U0001F4DD 2. Pre-existing active alarm \u2014 port numbers", key="lp_b2", placeholder="e.g. 3, 6, 20")
            st.markdown("\U0001F4DD 3. Pre-Existing Loops and Bridge Clips \u2014 port numbers per category:")
            bc1, bc2, bc3 = st.columns(3)
            with bc1: b3_loops = st.text_input("Pre existing loops", key="lp_b3_loops", placeholder="e.g. 1")
            with bc2: b3_clips = st.text_input("Bridge clips", key="lp_b3_clips", placeholder="e.g. 2")
            with bc3: b3_noequip = st.text_input("No equipment end connections", key="lp_b3_noequip", placeholder="e.g. 3")
            c1, c2 = st.columns([2, 1])
            with c1:
                b4 = st.text_input("\U0001F4DD 4. Post-cutover, FE couldn't clear \u2014 port numbers", key="lp_b4")
            with c2:
                b4_owner = st.selectbox("Owner", ["Tower Crew", "AT&T"], key="lp_b4_owner", label_visibility="collapsed")
            c1, c2 = st.columns([2, 1])
            with c1:
                b5 = st.text_input("\U0001F4DD 5. Other (free entry)", key="lp_b5")
            with c2:
                b5_dest = st.selectbox("Goes to", ["Pre-Existing Issues", "Pending"], key="lp_b5_dest", label_visibility="collapsed")
            c1, c2 = st.columns([2, 1])
            with c1:
                b6 = st.text_input("\U0001F4DD 6. Other (free entry)", key="lp_b6")
            with c2:
                b6_dest = st.selectbox("Goes to", ["Pre-Existing Issues", "Pending"], key="lp_b6_dest", label_visibility="collapsed")

            # Confirmed fix: build from ALL alarm ports (not just currently-locked ones),
            # so a port that's since been unlocked still gets its slogan looked up correctly.
            all_alarm_ports = controller_checks_data.get("alarm_ports", []) if controller_checks_data else []
            port_slogan_map = {p["port"]: p["slogan"] for p in all_alarm_ports if p["slogan"]}
            t1 = mcl.locked_port_bucket_1(b1, port_slogan_map)
            t2 = mcl.locked_port_bucket_2(b2, port_slogan_map)
            b3_notes, t3_active_line = mcl.loops_bridge_clips_notes(b3_loops, b3_clips, b3_noequip, all_alarm_ports, port_slogan_map)
            t4 = mcl.locked_port_bucket_4(b4, b4_owner, port_slogan_map)
            for t in (t1, t2):
                if t:
                    bucket_pre_existing.append(t)
            if t3_active_line:
                bucket_pre_existing.append(t3_active_line)
            bucket_notes = list(b3_notes)
            if t4:
                bucket_pending.append(t4)
            if b5:
                (bucket_pre_existing if b5_dest == "Pre-Existing Issues" else bucket_pending).append(b5)
            if b6:
                (bucket_pre_existing if b6_dest == "Pre-Existing Issues" else bucket_pending).append(b6)

    if bucket_pending:
        # Pending's text_area already rendered above this point in the page — can't inject
        # into its initial value, so append onto the already-collected choices dict instead
        # (dicts are mutable; build_mca_report_text reads this at button-click time, later).
        choices["additional_pending"]["text"] = (
            (choices["additional_pending"]["text"] or "") + "\n" + "\n".join(bucket_pending)).strip()

    # Same safe-mutation pattern for every checked-gated group above — appended onto the
    # dict directly rather than relying on a widget's (staleness-prone) displayed value.
    dss_completed_lines = [dss_completed_line] if dss_completed_line else []
    dss_pending_lines = [dss_pending_line] if dss_pending_line else []
    ret_completed_lines = [ret_completed_line] if ret_completed_line else []
    ret_pending_lines = [ret_pending_line] if ret_pending_line else []
    idl_completed_lines = [idl_completed_line] if idl_completed_line else []
    idl_pending_lines = [idl_pending_line] if idl_pending_line else []
    choices["additional_completed"]["text"] = "\n".join(
        [choices["additional_completed"]["text"] or ""] + gps_c_lines + sfp_c_lines + rs_c_display
        + fdd_c_lines + lkf_c_lines + sup_c_lines + xmu_c_lines + ngs_completed_lines + florida_active_rows
        + dss_completed_lines + ret_completed_lines + idl_completed_lines
    ).strip()
    choices["additional_pending"]["text"] = "\n".join(
        [choices["additional_pending"]["text"] or ""] + gps_p_lines + sfp_p_lines + rs_p_display
        + edp_lines + lkf_p_lines + sup_p_lines + xmu_p_lines + ngs_pending_lines + dss_pending_lines
        + ret_pending_lines + idl_pending_lines
    ).strip()

    with st.expander("Pre-Existing Issues"):
        pre_existing_text = st.text_area("\U0001F4DD Enter any Pre-Existing Issues that needs to be reported to Market",
                                          key="rpt_preexisting", height=70)
        if sfp_pre_existing_extra:
            st.caption("Auto-detected (Transport SFP, existing node):")
            for l in sfp_pre_existing_extra:
                st.caption(l)
        if bucket_pre_existing:
            st.caption("From locked-port classification:")
            for l in bucket_pre_existing:
                st.caption(l)
        # Same confirmed staleness bug as the other buffer boxes — appended onto the
        # collected value directly rather than relying on this widget's `value=` default,
        # which is only honored on its very first render.
        choices["pre_existing_issues_text"] = "\n".join(
            [pre_existing_text or ""] + sfp_pre_existing_extra + bucket_pre_existing
        ).strip()

    # Confirmed: same 3-way pattern as NSB (Completed/Pending/Partially Completed),
    # reusing the already-built scope_lines (via mcl.call_test_lines, market-table
    # driven) — detection logic itself untouched, only the status/reporting UI is new.
    _psap_line = next((l.replace("\t", " ") for l in scope_lines if l.startswith("PSAP test/Speedtest/VoLTE voice calltest:")), None)
    _speed_lte_line = next((l.replace("\t", " ") for l in scope_lines if l.startswith("Speedtest/VoLTE voice calltest:")), None)
    _speed_5g_line = next((l.replace("\t", " ") for l in scope_lines if l.startswith("Speed test:")), None)
    _fnet_line = next((l.replace("\t", " ") for l in scope_lines if l.startswith("Calltest with F-NET SIM:")), None)
    _ct_row_map = {
        "psap": (63, 126, "psap_moved_lte"), "speed_lte": (64, 127, "speedtest_new_lte"),
        "speed_5g": (65, 128, "speedtest_5g"), "fnet": (66, 129, "calltest_fnet"),
    }
    _ct_items = [
        ("psap", _psap_line, "PSAP test/Speedtest/VoLTE voice calltest"),
        ("speed_lte", _speed_lte_line, "Speedtest/VoLTE voice calltest"),
        ("speed_5g", _speed_5g_line, "Speed test"),
        ("fnet", _fnet_line, "Calltest with F-NET SIM"),
    ]
    _ct_detected = [(k, l, lbl) for k, l, lbl in _ct_items if l]
    _ct_status = None
    _ct_psap_sched_id = ""
    _ct_completed_inputs, _ct_pending_inputs = {}, {}
    _ct_row_writes = []
    if _ct_detected:
        with st.container(border=True):
            st.markdown("**Call Test requirements detected (per CT sheet):**")
            for _k, _l, _lbl in _ct_detected:
                _display = _l.replace(" PSAP Schedule ID: ", "").rstrip() if _k == "psap" else _l
                st.caption(_display)
            _ct_status = st.selectbox("Call Test status", ["\u2014 Select \u2014", "Completed", "Pending", "Partially Completed"], key="rpt_ct_status")
            _psap_applies = any(k == "psap" for k, _l, _lbl in _ct_detected)

            if _ct_status == "Completed":
                if _psap_applies:
                    _ct_psap_sched_id = st.text_input("PSAP Schedule ID", key="rpt_ct_psap_sched")
                _ct_text_lines = []
                for _k, _l, _lbl in _ct_detected:
                    _comp_row, _pend_row, _row_key = _ct_row_map[_k]
                    _val = _l.split(": ", 1)[-1].replace(" PSAP Schedule ID: ", "").rstrip()
                    _extra = [(5, _ct_psap_sched_id.strip())] if (_k == "psap" and _ct_psap_sched_id.strip()) else []
                    _ct_row_writes.append((_comp_row, True, [(3, _val)] + _extra))
                    _ct_row_writes.append((_pend_row, False, []))
                    _base_l = _l.replace(" PSAP Schedule ID: ", "").rstrip() if _k == "psap" else _l
                    _display_l = _base_l + (f" (PSAP Schedule ID: {_ct_psap_sched_id.strip()})" if (_k == "psap" and _ct_psap_sched_id.strip()) else "")
                    _ct_text_lines.append(_display_l)
                    st.caption(f"\u2705 {_display_l}")
                choices["additional_completed"]["text"] = "\n".join(
                    [choices["additional_completed"]["text"] or ""] + _ct_text_lines).strip()
            elif _ct_status == "Pending":
                _ct_text_lines = []
                for _k, _l, _lbl in _ct_detected:
                    _comp_row, _pend_row, _row_key = _ct_row_map[_k]
                    _val = _l.split(": ", 1)[-1].replace(" PSAP Schedule ID: ", "").rstrip()
                    _ct_row_writes.append((_comp_row, False, []))
                    _ct_row_writes.append((_pend_row, True, [(3, _val)]))
                    _line = f"{_lbl}: {_val} (MIC PM)"
                    _ct_text_lines.append(_line)
                    st.caption(_line)
                choices["additional_pending"]["text"] = "\n".join(
                    [choices["additional_pending"]["text"] or ""] + _ct_text_lines).strip()
            elif _ct_status == "Partially Completed":
                col_c, col_p = st.columns(2)
                with col_c:
                    st.markdown("**Completed**")
                    for _k, _l, _lbl in _ct_detected:
                        _ct_completed_inputs[_k] = st.text_input(f"{_lbl} \u2014 Completed on", key=f"rpt_ct_{_k}_completed")
                        if _k == "psap" and _psap_applies and _ct_completed_inputs[_k].strip():
                            _ct_psap_sched_id = st.text_input("PSAP Schedule ID", key="rpt_ct_psap_sched")
                with col_p:
                    st.markdown("**Pending**")
                    for _k, _l, _lbl in _ct_detected:
                        _ct_pending_inputs[_k] = st.text_input(f"{_lbl} \u2014 Pending on", key=f"rpt_ct_{_k}_pending")
                _ct_result_lines = []
                for _k, _l, _lbl in _ct_detected:
                    _comp_row, _pend_row, _row_key = _ct_row_map[_k]
                    _c_val = _ct_completed_inputs.get(_k, "").strip()
                    _p_val = _ct_pending_inputs.get(_k, "").strip()
                    if _c_val:
                        _extra = [(5, _ct_psap_sched_id.strip())] if (_k == "psap" and _ct_psap_sched_id.strip()) else []
                        _ct_row_writes.append((_comp_row, True, [(3, _c_val)] + _extra))
                        _line = f"{_lbl}: {_c_val}."
                        if _k == "psap" and _ct_psap_sched_id.strip():
                            _line += f" (PSAP Schedule ID: {_ct_psap_sched_id.strip()})"
                        _ct_result_lines.append(("completed", _line))
                    if _p_val:
                        _ct_row_writes.append((_pend_row, True, [(3, _p_val)]))
                        _ct_result_lines.append(("pending", f"{_lbl}: {_p_val} (MIC PM)"))
                if _ct_result_lines:
                    st.markdown("**Result:**")
                    _ct_c_texts = [l for k, l in _ct_result_lines if k == "completed"]
                    _ct_p_texts = [l for k, l in _ct_result_lines if k == "pending"]
                    if _ct_c_texts:
                        choices["additional_completed"]["text"] = "\n".join(
                            [choices["additional_completed"]["text"] or ""] + _ct_c_texts).strip()
                    if _ct_p_texts:
                        choices["additional_pending"]["text"] = "\n".join(
                            [choices["additional_pending"]["text"] or ""] + _ct_p_texts).strip()
                    for _kind, _line in _ct_result_lines:
                        st.caption(f"\u2705 {_line}" if _kind == "completed" else _line)

    with st.expander("Notes"):
        has_5g = any(app.is_populated(row.get("gNBId")) for row in mm_objs)

        # Final Port Configuration: always included, no longer optional.
        choices["notes_final_port_config"] = {"checked": True, "text": "Final Port Configuration attached."}

        # NR configuration verified: auto-triggers when 5G is present at the site (checked
        # by default, still overridable — same pattern as every other auto-detected item).
        nr_verified_checked = st.checkbox("NR configuration has been verified.", value=has_5g, key="chk_notes_nr_verified")
        choices["notes_nr_verified"] = {"checked": nr_verified_checked, "text": "NR configuration has been verified."}
        if not has_5g:
            st.caption("(No 5G detected at this site)")

        # Confirmed fix: this is a Notes item (explicitly a "note" from
        # external_alarm_testing_placement, all-locked/NEA-pending case) — was
        # previously rendered among the Pending UI elements and concatenated into
        # additional_pending's text, showing it under Pending in both the UI and the
        # final report instead of Notes. Now rendered here directly, matching where
        # it belongs both visually and in the underlying routing.
        if testing_note:
            testing_note_checked = st.checkbox(testing_note, value=True, key="chk_testing_note")
        else:
            testing_note_checked = False
        choices["notes_testing"] = {"checked": testing_note_checked, "text": testing_note or ""}

        cpri_choice = "\u2014 Select \u2014"
        if new_nodes:
            st.markdown(f"**Area prechecks verification for CPRI/SFP check** \u2014 {len(new_nodes)} new node(s)")
            cpri_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                         key="cpri_sfp_status", label_visibility="collapsed")
        cpri_text = ("Area prechecks verification for CPRI/SFP check is completed." if cpri_choice == "Completed"
                     else "Area prechecks verification for CPRI/SFP check is pending(Node is not replicating in tool)."
                     if cpri_choice == "Pending" else "")
        choices["notes_cpri_sfp"] = {"checked": bool(cpri_choice != "\u2014 Select \u2014"), "text": cpri_text}

        # No scope of external alarms: user's own choice, unchanged.
        n_no_alarms = st.checkbox("No scope of external alarms.", key="chk_notes_no_external_alarms")
        choices["notes_no_external_alarms"] = {"checked": n_no_alarms, "text": "No scope of external alarms."}

        n_mme = st.checkbox("Pre-Existing MME configuration left as it is on nodes", key="chk_notes_mme_config")
        mme_nodes = st.text_input("\U0001F4DD Node ID(s), comma/pipe separated", key="rpt_mme_node", label_visibility="collapsed",
                                    placeholder="e.g. NodeA, NodeB") if n_mme else ""
        choices["notes_mme_config"] = {"checked": n_mme, "text": f"Pre-Existing MME configuration left as it is on nodes: {mme_nodes}"}

        # Monitored/not-monitored: auto-filled from the already-known pre-existing vs.
        # newly-added node lists — confirmed, no manual typing needed.
        pre_existing_node_names = [row.get("Node to be built as") for row in mm_objs
                                    if row.get("Node to be built as") not in new_nodes]
        c1, c2 = st.columns(2)
        with c1:
            if pre_existing_node_names:
                n_mon = st.checkbox(f"Node is in monitored state ({'|'.join(pre_existing_node_names)})",
                                      value=True, key="chk_notes_monitored")
            else:
                n_mon = False
            choices["notes_monitored"] = {"checked": n_mon, "text": f"{'|'.join(pre_existing_node_names)} is in monitored state."}
        with c2:
            if new_nodes:
                n_not_mon = st.checkbox(f"Node is in not monitored state ({'|'.join(new_nodes)})",
                                          value=True, key="chk_notes_not_monitored")
            else:
                n_not_mon = False
            choices["notes_not_monitored"] = {"checked": n_not_mon, "text": f"{'|'.join(new_nodes)} is in not monitored state."}

        # 6610 controller monitored/not-monitored state — same pattern as N2E/NSB (auto
        # "not monitored" when the cascade fires, SAU is Pending/disabled, or External alarm
        # testing is Pending; manual Monitored/Not monitored choice otherwise). MCA had the
        # node-level version above but was missing the controller-level one.
        ctrl_mon_checked, ctrl_mon_text = False, ""
        if controller_id:
            if cascade_fires or sau_disabled_on_6610 or testing_section == "Pending":
                _ctrl_mon_reason = ("no 6610 checks" if cascade_fires
                                     else ("SAU disabled" if sau_disabled_on_6610 else "External alarm testing Pending"))
                st.caption(f"{controller_id} is in not monitored state. (auto \u2014 {_ctrl_mon_reason})")
                ctrl_mon_checked = True
                ctrl_mon_text = f"{controller_id} is in not monitored state."
            else:
                ctrl_mon_choice = st.selectbox(f"{controller_id} monitored state",
                                                 ["\u2014 Select \u2014", "Monitored", "Not monitored"], key="mca_ctrl_mon")
                if ctrl_mon_choice == "Monitored":
                    ctrl_mon_checked = True
                    ctrl_mon_text = f"{controller_id} is in monitored state."
                elif ctrl_mon_choice == "Not monitored":
                    ctrl_mon_checked = True
                    ctrl_mon_text = f"{controller_id} is in not monitored state."
        choices["notes_ctrl_monitored"] = {"checked": ctrl_mon_checked, "text": ctrl_mon_text}

        # SAU enabled on point — MCA: SAU can be enabled on either the 6610 or the node.
        # When SAU is confirmed disabled on the 6610 specifically (not the full cascade),
        # it may instead be enabled on the node. Post-checks' Hardware Status table carries
        # a distinct per-node SAU row (e.g. "FCL05583 SAU-1 ... ENABLED"), so this is now
        # auto-detected via sau_enabled_nodes() the same way SUP/XMU already are — manual
        # entry stays available underneath as an override/fallback for whatever the
        # detector doesn't confidently find.
        sau_node_checked, sau_node_text = False, ""
        if sau_disabled_on_6610:
            auto_sau_nodes = mcl.sau_enabled_nodes(postcheck_text) if postcheck_text else []
            if auto_sau_nodes:
                auto_sau_str = "|".join(auto_sau_nodes)
                sau_node_on = st.checkbox(
                    f"SAU enabled on the node instead (auto-detected: {auto_sau_str})",
                    value=True, key="mca_sau_node_chk")
                if sau_node_on:
                    sau_node_checked = True
                    sau_node_text = f"SAU enabled on : {auto_sau_str}."
            else:
                sau_node_on = st.checkbox("SAU enabled on the node instead", key="mca_sau_node_chk")
                if sau_node_on:
                    st.caption("(Not auto-detected in Post-checks \u2014 enter manually)")
                    sau_node_id = st.text_input("\U0001F4DD Node ID", key="mca_sau_node_id")
                    if sau_node_id.strip():
                        sau_node_checked = True
                        sau_node_text = f"SAU enabled on : {sau_node_id.strip()}."
        choices["notes_sau_enabled"] = {"checked": sau_node_checked, "text": sau_node_text}

        emergency_unlock_lines = [f"Emergency unlock activated on the node {n}." for n in emergency_unlock_notes]
        notes_generic = st.text_area("\U0001F4DD Enter Notes that need to be reported or addressed to Market",
                                      key="rpt_notes_generic", height=70)
        if emergency_unlock_lines:
            st.caption("Auto-added (Emergency unlock confirmed):")
            for l in emergency_unlock_lines:
                st.caption(l)
        if ngs_notes_lines:
            st.caption("Auto-added (NGS marked Pre-Existing):")
            for l in ngs_notes_lines:
                st.caption(l)
        if bucket_notes:
            st.caption("Auto-added (Locked-port classification, bucket 3):")
            for l in bucket_notes:
                st.caption(l)
        scripted_locked_note = mcl.scripted_locked_bands_note(ciq_wb)
        scripted_locked_lines = []
        if scripted_locked_note:
            scripted_locked_checked = st.checkbox(scripted_locked_note, value=True, key="chk_scripted_locked")
            if scripted_locked_checked:
                scripted_locked_lines = [scripted_locked_note]
        # Confirmed: no separate auto-generated "mixed still-locked ports" note for MCA —
        # this would duplicate what the engineer already manually classifies via the
        # existing Bucket 1 (locked_port_bucket_1) "Active alarms observed on ports..."
        # mechanism, which N2E never had (hence N2E does need the automatic version).
        choices["notes_generic_text"] = "\n".join(
            [notes_generic or ""] + emergency_unlock_lines + ngs_notes_lines + bucket_notes + scripted_locked_lines).strip()

    st.markdown("---")
    node_tag = mm_objs[0].get("Node to be built as", "site") if mm_objs else "site"

    if st.button("Generate Report \u2192", type="primary", key="rpt_generate_mca"):
        if warnings:
            with st.container(border=True):
                st.markdown(
                    f"<div style='color:#c0392b; font-size:1.3em; font-weight:700;'>"
                    f"⚠️ {len(warnings)} Warning{'s' if len(warnings) != 1 else ''}</div>"
                    f"<div style='color:#c0392b; margin-bottom:0.5em;'>Informational only — these never change "
                    f"what's in the report below, they're here so nothing gets missed.</div>",
                    unsafe_allow_html=True)
                for w in warnings:
                    st.markdown(f"<div style='color:#c0392b; font-size:1.05em; padding:2px 0;'>• {w['text']}</div>",
                                unsafe_allow_html=True)

        header_fields = {
            "mic": "MIC", "market": market_subject_input, "status": status, "site_name": site_name,
            "fa_code": fa_code, "site_ids": site_ids, "sow": sow, "iwm_details": iwm_details,
            "pre_configuration": pre_line, "current_configuration": current_config,
            "post_configuration": post_line, "wll_node": wll_node, "controller_id": controller_id,
            "software_version": software_version, "gs_version": gs_version,
            "idl_build_type": idl_build_type, "idle": idle, "idly": idly, "switch": switch, "slot_port": slot_port,
        }
        report_text = mca_report_text.build_mca_report_text(mm_objs, results, choices, header_fields, stakeholder_by_key=stakeholders)
        st.success("Report generated.")
        st.text_area("Report preview", report_text, height=400, key="rpt_preview")
        st.download_button("Download report (.txt)", report_text, file_name=f"{node_tag}_Integration_Report.txt", key="rpt_dl_txt")

        row_writes = mca_glue.build_xlsm_row_writes(results, choices, ROW_MAP)
        row_writes += _ct_row_writes
        row_writes.append((3, True, [(2, "MIC"), (3, market_subject_input), (4, status), (5, site_name), (6, fa_code), (7, site_ids), (8, sow)]))
        row_writes.append((6, True, [(3, iwm_details)]))
        row_writes.append((10, True, [(3, pre_line)]))
        row_writes.append((11, bool(current_config.strip()), [(3, current_config)]))
        row_writes.append((12, True, [(3, post_line)]))
        row_writes.append((13, bool(wll_node.strip()), [(3, wll_node)]))
        # Confirmed real bug: the raw template defaults these to checked=True with
        # placeholder text baked in — anywhere we don't explicitly override with the real
        # value/False, that default silently leaks straight through to the final .xlsm.
        row_writes.append((14, bool(controller_id), [(3, controller_id)] if controller_id else []))
        row_writes.append((15, bool(software_version.strip()), [(3, software_version)] if software_version.strip() else []))
        row_writes.append((16, bool(gs_version.strip()), [(3, gs_version)] if gs_version.strip() else []))
        row_writes.append((19, bool(idl_build_type), [(3, idl_build_type)] if idl_build_type else []))

        # IDLe / IDLy (rows 20/21) — same confirmed bug: never written before this fix, so
        # the template's default (checked=True, placeholder text) always leaked through
        # regardless of whether the engineer actually entered anything.
        row_writes.append((20, bool(idle.strip()), [(3, idle)] if idle.strip() else []))
        row_writes.append((21, bool(idly.strip()), [(3, idly)] if idly.strip() else []))

        # Switch / Slot-Port (rows 23/24) — confirmed NEVER written at all before this fix,
        # so the template's default (checked=True, placeholder text) always leaked through
        # regardless of whether this site actually has Sidehaul Info data. Two paths: the
        # Sidehaul-auto-fill case, and the manual-fallback case (when no Sidehaul data
        # exists but the engineer typed into the manual Switch/Slot-Port text areas).
        sidehaul_rows_data = mcl.sidehaul_display_rows(ciq_wb)
        if sidehaul_rows_data:
            first = sidehaul_rows_data[0]
            row_writes.append((23, True, [(3, first["switch_type"]), (4, first["switch_id"])]))
            cable_pn = st.session_state.get("cable_pn_0", "")
            row_writes.append((24, True, [(3, cable_pn), (4, first["node_id"])]))
        elif switch.strip() or slot_port.strip():
            row_writes.append((23, bool(switch.strip()), [(3, switch)] if switch.strip() else []))
            row_writes.append((24, bool(slot_port.strip()), [(3, slot_port)] if slot_port.strip() else []))
        else:
            row_writes.append((23, False, []))
            row_writes.append((24, False, []))

        # Rows 25-39 (15 "Manual Feed based on NEST & CIQ" rows) — confirmed real gap:
        # overflow slots for ADDITIONAL Sidehaul connections beyond the first (real example:
        # FSL00456 had 2 connections on the same switch). Never wired at all before this fix.
        extra_sidehaul_rows = list(range(25, 40))
        extra_connections = sidehaul_rows_data[1:] if sidehaul_rows_data else []
        extra_sidehaul_lines = []
        for i, conn in enumerate(extra_connections):
            cable_pn_extra = st.session_state.get(f"cable_pn_{i+1}", "")
            line = (f"Switch type: {conn['switch_type']}  Switch ID: {conn['switch_id']}  "
                    f"Slot/Port: {conn['slot_port']}  Cable part number: {cable_pn_extra}  Node ID: {conn['node_id']}")
            extra_sidehaul_lines.append(line)
        mcl.write_buffer_with_overflow(row_writes, extra_sidehaul_rows, extra_sidehaul_lines)

        # Row 80 "Swap Sector Verification" — confirmed explicitly removed from scope
        # earlier this session, but never wrote False, so its template default (True,
        # placeholder "CBAND|DOD") was leaking through unchanged.
        row_writes.append((80, False, []))

        # Row 46 "Port speed 1G to 10G conversion with MPST:" — confirmed real, pre-existing
        # bug in mca_glue._result_to_column_values: it expects tab-separated scope lines
        # (matching most other items), but this item's line has never had tabs (confirmed
        # even in the original pre-session code) — splitting on tab produces one element,
        # then dropping "the label token" empties it, so the checkbox always correctly
        # showed True but the Node ID column always silently stayed blank. Bypasses that
        # broken path with an explicit write instead. Re-derives the node list directly
        # from scope_lines (not a variable from earlier in the function) to avoid any
        # scoping risk if postcheck_text was falsy.
        pc_lines = [l for l in scope_lines if l.startswith("Port speed 1G to 10G conversion with MPST:")]
        if pc_lines:
            pc_nodes = pc_lines[0].split("MPST:")[-1].strip().rstrip(".")
            row_writes.append((46, True, [(3, pc_nodes)]))
        else:
            row_writes.append((46, False, []))

        # DSS Activation (completed=59, pending=122) — restored toggle, confirmed real gap.
        dss_c_bands = dss_completed_line.split("DSS Activation:")[-1].strip() if dss_completed_line else None
        dss_p_bands = " & ".join(mcl.sort_bands_lte_first(dss_pending_bands_combined)) if dss_pending_bands_combined else None
        row_writes.append((59, bool(dss_completed_line), [(3, dss_c_bands)] if dss_c_bands else []))
        row_writes.append((122, bool(dss_pending_line), [(3, dss_p_bands)] if dss_p_bands else []))

        # ---- Notes section (confirmed real gap — none of this was ever written to the
        # .xlsm before). Rows 181/182/183 are the only genuinely dedicated Notes rows in
        # the real template; "CPRI/SFP", "No scope of external alarms", and "monitored/
        # not-monitored state" have NO dedicated row at all (confirmed by searching the
        # actual template text) — they route through the 8-row generic buffer (184-191)
        # alongside the free-text notes, emergency-unlock lines, NGS Pre-Existing notes,
        # and bucket-3 notes, first-come-first-served. Overflow beyond 8 -> Warnings tab. ----
        row_writes.append((181, True, [(2, "Final Port Configuration attached.")]))
        row_writes.append((182, nr_verified_checked, [(2, "NR configuration has been verified.")] if nr_verified_checked else []))
        row_writes.append((183, n_mme, [(2, f"Pre-Existing MME configuration left as it is on nodes: {mme_nodes}")] if n_mme else []))

        notes_buffer_lines = []
        if cpri_choice not in ("", "\u2014 Select \u2014"):
            notes_buffer_lines.append(cpri_text)
        if n_no_alarms:
            notes_buffer_lines.append("No scope of external alarms.")
        if n_mon:
            notes_buffer_lines.append(f"{'|'.join(pre_existing_node_names)} is in monitored state.")
        if n_not_mon:
            notes_buffer_lines.append(f"{'|'.join(new_nodes)} is in not monitored state.")
        if ctrl_mon_checked and ctrl_mon_text:
            notes_buffer_lines.append(ctrl_mon_text)
        if sau_node_checked and sau_node_text:
            notes_buffer_lines.append(sau_node_text)
        notes_buffer_lines += [l for l in (notes_generic or "").split("\n") if l.strip()]
        notes_buffer_lines += emergency_unlock_lines + ngs_notes_lines + bucket_notes + scripted_locked_lines

        notes_buffer_rows = list(range(184, 192))
        mcl.write_buffer_with_overflow(row_writes, notes_buffer_rows, notes_buffer_lines)

        # ---- Everything built this session that build_xlsm_row_writes never knew about —
        # parsed back out of the already-tested formatted strings rather than re-deriving
        # structured data, since the exact format is controlled and safe to parse. Uses the
        # CHECKBOX-GATED lines (gps_c_lines, sfp_c_lines, etc.) so unchecking any group
        # correctly excludes it from the .xlsm too, not just the plain-text report. ----
        gps_completed_groups = []
        for line in gps_c_lines[:1]:  # only the first (dedicated-row) group belongs here
            if "Version:" in line:
                nodes_part, ver_part = line.split("Version:", 1)
                nodes = nodes_part.replace("GPS Installation:", "").strip().split("|")
                gps_completed_groups.append((nodes, ver_part.strip()))
        sfp_completed_groups = []
        for line in sfp_c_lines:
            m = re.match(r"Transport SFP Installation on:\s*(.+?)\s+SFP Model \(BBU End\):\s*(.*?)\s+SFP Model \(SIAD End\):\s*(.*)", line)
            if m:
                nodes = m.group(1).split("|")
                sfp_completed_groups.append((nodes, m.group(2).strip(), m.group(3).strip()))
        radio_swap_completed_parsed = []
        for line in (radio_swap_completed_lines if rs_c_checked else []):
            parts = line.split("\t")
            if len(parts) >= 6:
                label_sectors = parts[1]
                radio_swap_completed_parsed.append((label_sectors, "", parts[3], parts[5]))
        radio_swap_pending_parsed = []
        for line in (radio_swap_pending_lines if rs_p_checked else []):
            parts = line.split("\t")
            if len(parts) >= 6:
                label_sectors = parts[1]
                radio_swap_pending_parsed.append((label_sectors, "", parts[3], parts[5]))
        fdd_parsed = []
        for line in fdd_c_lines:
            m = re.match(r"FDD Renaming on:\s*(.+?)\s+From:\s*(.+?)\s+To:\s*(.+?)\.", line)
            if m:
                fdd_parsed.append((m.group(1), m.group(2), m.group(3)))
        lkf_completed_val = (lkf_c_lines[0] if lkf_c_lines else "").replace("LKF Installation:", "").strip() or None
        lkf_pending_val = (lkf_p_lines[0] if lkf_p_lines else "").replace("LKF Installation:", "").replace("(MIC)", "").strip() or None
        ngs_completed_val = "; ".join(ngs_completed_lines) or None
        ngs_pending_val = "; ".join(ngs_pending_lines) or None

        def _split_buffer_lines(lines):
            """Buffer entries as (label, detail) — first ':' splits label from detail;
            falls back to (item text, "") if no colon found."""
            out = []
            for l in lines:
                if ":" in l:
                    label, detail = l.split(":", 1)
                    out.append((label.strip(), detail.strip()))
                else:
                    out.append((l, ""))
            return out

        # Completed buffer = whatever's left after GPS's dedicated row + Transport SFP's
        # dedicated rows + FDD's dedicated rows + LKF's dedicated row are accounted for.
        # (Port Conversion is NOT here — it merges into one line through its own dedicated
        # checklist row instead, confirmed, regardless of how many nodes.)
        # Completed buffer = only genuine OVERFLOW beyond each item's dedicated row capacity
        # (GPS: 1 completed row, Transport SFP: 3 completed rows, Radio Swap: 3 completed
        # rows) — NOT the full lists, since those already go to their own dedicated rows via
        # build_new_xlsm_row_writes below; including them again here would double-count the
        # same content in two places.
        buffer_completed_lines = gps_c_lines[1:] + sfp_c_lines[3:] + rs_c_display[3:]
        buffer_pending_lines = gps_p_lines[1:] + sfp_p_lines[4:] + rs_p_display[3:]
        buffer_pre_existing_lines = sfp_pre_existing_extra + bucket_pre_existing

        new_rw = mcl.build_new_xlsm_row_writes(
            ROW_MAP,
            current_config_text="",  # already written above at row 11 — avoid double-write
            gps_completed_groups=gps_completed_groups,
            gps_pending_lines=gps_p_lines[:1],
            sfp_completed_groups=sfp_completed_groups[:3],
            sfp_pending_lines=sfp_p_lines[:4],
            radio_swap_completed=radio_swap_completed_parsed[:3],
            radio_swap_pending=radio_swap_pending_parsed[:3],
            lkf_completed_line=lkf_completed_val,
            lkf_pending_line=lkf_pending_val,
            fdd_lines=fdd_parsed[:2],
            edp_publish=edp_publish_text,
            ngs_completed_line=ngs_completed_val,
            ngs_pending_line=ngs_pending_val,
            buffer_completed_extra=_split_buffer_lines(buffer_completed_lines)[:10],
            buffer_pending_extra=_split_buffer_lines(buffer_pending_lines)[:9],
            buffer_pre_existing_extra=buffer_pre_existing_lines[:10],
        )
        row_writes += new_rw

        # Florida-only newly added cells (rows 93-104, single column B, confirmed same
        # pattern as Pre-Existing Issues rows — not the label:detail buffer format).
        # Gated on the checkbox — unchecked means excluded from the .xlsm entirely.
        # Row 92 ("Newly added Cells" sub-header) matches the same checkbox state.
        row_writes.append((92, bool(florida_checked), []))
        florida_xlsm_rows = list(range(93, 105))
        for i, row_num in enumerate(florida_xlsm_rows):
            if i < len(florida_active_rows):
                row_writes.append((row_num, True, [(2, florida_active_rows[i])]))
            else:
                row_writes.append((row_num, False, []))

        if not TEMPLATE_PATH.exists():
            static_dir = TEMPLATE_PATH.parent
            if static_dir.exists():
                actual_files = sorted(p.name for p in static_dir.iterdir())
                st.warning(f"Template not found at {TEMPLATE_PATH}. Files actually present in {static_dir}: {actual_files}")
            else:
                st.warning(f"Template not found at {TEMPLATE_PATH} \u2014 and the folder {static_dir} doesn't exist at all.")
        else:
            xlsm_bytes = fill_legacy_mca_surgical(TEMPLATE_PATH, row_writes)
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_Legacy_MCA_Filled.xlsm", key="rpt_dl_xlsm")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"{node_tag}_Integration_Report.txt", report_text)
                zf.writestr(f"{node_tag}_Legacy_MCA_Filled.xlsm", xlsm_bytes)
            st.download_button("Download both (report + filled checklist, .zip)", zip_buffer.getvalue(),
                                file_name=f"{node_tag}_MCA_Bundle.zip", key="rpt_dl_zip")
