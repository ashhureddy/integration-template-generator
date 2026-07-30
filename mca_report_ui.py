"""
MCA Integration Report — interactive 'Generate Report' UI section.
Fixes applied: correct 6610 controller ID (Controller Info has TWO columns — 'Controller' is
just the literal string '6610' as a type marker, 'Controller ID' is the real instance name),
live preview of detected values (no longer hidden until Generate), multi-instance items show
ALL detected lines, manual-entry space + stakeholder selector for every item, high-contrast
display of auto-fetched read-only values, and a more organized bordered-card layout.
"""
import streamlit as st

import report_detect
import mca_checklist
import mca_glue
import mca_report_text
import mca_completed_logic as mcl
from mca_row_map import ROW_MAP
from mca_xlsm_fill import fill_legacy_mca

from pathlib import Path

TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "Legacy_MCA_Macro_Template_v6_1.xlsm"
STAKEHOLDER_OPTIONS = ["MIC", "MIC PM", "AT&T", "Tower Crew"]


def _get_controller_id(controller_objs):
    """Controller Info has a 'Controller' column that's just the literal string '6610' (a type
    marker) and a SEPARATE 'Controller ID' column with the real instance name (e.g.
    LSPC273360_C001) — confirmed bug: using 'Controller' directly showed the literal '6610'."""
    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    return ctrl_rows[0].get("Controller ID") if ctrl_rows else ""


def _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp):
    new_nodes, board_swaps = report_detect.detect_node_board_changes(app, ciq_wb, mm_objs, precheck_text)
    fdd_renames = report_detect.detect_fdd_renaming(app, ciq_wb)
    return {
        "scope_lines": scope_lines, "new_nodes": new_nodes, "board_swaps": board_swaps,
        "fdd_renames": fdd_renames, "controller_id": controller_id, "controller_in_edp": controller_in_edp,
        "idl_build_type": idl_build_type,
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
    st.subheader("Generate Report")

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

    calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
    if calltest_path.exists() and mm_objs:
        prefix_to_market, calltest_rules = mcl.load_calltest_table(calltest_path)
        market = mcl.determine_market(mm_objs[0].get("Node to be built as"), prefix_to_market)
        if market:
            scope_lines = scope_lines + mcl.call_test_lines(
                classification, market, calltest_rules,
                moved_bands_by_tech["lte"], added_bands_by_tech, moved_bands_by_tech)

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
    # existing same-board-only logic never recognized. ----
    port_conv_swap_completed = []
    if postcheck_text:
        port_conv_swap_completed = mcl.check_port_conversion_via_board_swap(
            ciq_wb, mm_objs, precheck_text, postcheck_text)
        # A node completed via swap shouldn't also show as a pending "still needs conversion"
        # line from the original same-board check.
        swap_nodes = {r["node"] for r in port_conv_swap_completed}
        scope_lines = [l for l in scope_lines
                       if not (l.startswith("Port speed 1G to 10G conversion with MPST:")
                               and any(n in l for n in swap_nodes))]

    # ---- 6610 cascade: if a 6610 is present/EDP-published but the controller-checks file
    # doesn't confirm alarm scripting, 4 items move to Pending together, no warning. ----
    controller_checks_data = mcl.extract_controller_checks(controller_checks_text) if controller_checks_text else {}
    cascade_fires = mcl.controller_integration_cascade(
        bool(controller_in_edp), controller_checks_data, controller_id)
    sau_placement = mcl.sau_connections_placement(controller_checks_data, controller_id) if controller_id else None
    testing_section, testing_note = mcl.external_alarm_testing_placement(controller_checks_data) \
        if controller_checks_data else (None, None)

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
            gps_extra_pending.append(gps_sync_line)

    # ---- EDP Publish fallback: 6610 present in CIQ but NOT published in EDP -> this
    # replaces the old generic message entirely; 6610 Controller Integration does NOT
    # appear in Completed at all in this case. ----
    edp_publish_text = ""
    if controller_id and not controller_in_edp:
        edp_publish_text = mcl.edp_publish_line(
            mm_objs[0].get("Node to be built as") if mm_objs else "", controller_id, "")

    # ---- Current Configuration: only populated when Post-checks actually differs from
    # the CIQ target (equipment still missing) — confirmed rule, built this pass. ----
    current_config_auto = mcl.current_configuration_line(ciq_wb, mm_objs, postcheck_text) if postcheck_text else ""

    # ---- FDD Renaming, corrected: band-label grouping instead of raw ungrouped cell
    # tuples — confirmed gap, built this pass. ----
    fdd_lines_fixed = mcl.fdd_renaming_lines(ciq_wb)

    ctx = _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp)
    results = mca_checklist.evaluate_checklist(ctx)

    if cascade_fires:
        # Force these 4 items to Pending, drop them from wherever the normal detection put
        # them, no warning per confirmed decision.
        cascade_keys = {"controller_integration", "alarm_scripting", "lkf_installation", "alarm_testing"}
        for item in results:
            if item["key"] in cascade_keys:
                item["section"] = "pending"
                item["checked_by_default"] = True
    if sau_placement and not cascade_fires:
        for item in results:
            if item["key"] == "sau_connections":
                item["section"] = "completed" if sau_placement == "Completed" else "pending"
                item["checked_by_default"] = True
    if testing_section and not cascade_fires:
        for item in results:
            if item["key"] == "alarm_testing":
                item["section"] = "completed" if testing_section == "Completed" else "pending"
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
    # Suppress the checklist's own LKF Installation item unconditionally — confirmed bug:
    # it was showing alongside the new per-node Completed/Pending dropdown section as a
    # confusing duplicate. The custom section (built this pass) fully replaces it, since it
    # covers the same 3 original triggers plus the 4th (single-tech node gaining a second
    # tech) the old item never had.
    for item in results:
        if item["key"] == "lkf_installation":
            item["checked_by_default"] = False

    # ---- Warnings tab collection: every verification function feeds here. Confirmed
    # design: warning-only for most items (never touches report placement), except
    # Radio Swap (which changes Completed/Pending placement itself) and the
    # board-swap-triggered Transport SFP case (which is both Pending AND a warning). ----
    warnings = []
    if postcheck_text:
        warnings += mcl.verify_integration_against_postcheck(classification, postcheck_text)
        warnings += mcl.verify_moved_sectors_against_postcheck(classification, postcheck_text)
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
    default_status = "STF" if has_pending else "ATP"

    with st.container(border=True):
        st.markdown("**Subject**")
        c = st.columns(7)
        with c[0]: st.markdown(f"MIC\n\n**MIC**")
        with c[1]: market = st.text_input("Market", key="rpt_market", placeholder="MNS/TILLMAN/AT&T")
        with c[2]: status = st.text_input("Status", value=default_status, key="rpt_status")
        with c[3]: site_name = st.text_input("Site Name", key="rpt_site_name")
        with c[4]: st.markdown(f"FA CODE\n\n**{fa_code or '(not found)'}**")
        with c[5]: st.markdown(f"Site ID's\n\n**{site_ids}**")
        with c[6]: sow = st.text_input("SOW", key="rpt_sow")

    with st.container(border=True):
        st.markdown("**IWM Details**")
        iwm_details = st.text_input("IWM Details", key="rpt_iwm", label_visibility="collapsed")

    with st.container(border=True):
        st.markdown("**Configuration**")
        st.markdown(f"Pre Configuration : **{pre_line}**")
        st.markdown(f"Post Configuration : **{post_line}**")
        st.markdown(f"6610 Controller : **{controller_id or '(none detected)'}**")
        c1, c2 = st.columns(2)
        with c1:
            current_config = st.text_input(
                "Current Configuration (auto — only shown when Post-checks differs from CIQ target)",
                value=current_config_auto, key="rpt_current_config")
            wll_node = st.text_input("WLL node (if applicable)", key="rpt_wll")
        with c2:
            software_version = st.text_input("Software version", key="rpt_sw")
            gs_version = st.text_input("GS Version", key="rpt_gs")

    idle = idly = switch = slot_port = ""
    if len(mm_objs) > 1:
        with st.container(border=True):
            st.markdown(f"**IDL Connections** \u2014 Build Type: **{idl_build_type or '(not detected)'}**")
            c1, c2 = st.columns(2)
            with c1:
                idle = st.text_area("IDLe cable details (manual)", key="rpt_idle", height=60)
            with c2:
                idly = st.text_area("IDLy cable details (manual)", key="rpt_idly", height=60)

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
                switch = st.text_area("Switch details (manual)", key="rpt_switch", height=60)
                slot_port = st.text_area("Slot/Port/Cable/Node ID (manual)", key="rpt_slotport", height=60)

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

    # ---- Transport SFP: trigger nodes = new nodes OR Port-Conversion-triggered nodes.
    # BBU/SIAD End models are MANUAL (confirmed), grouped by shared entered model. ----
    port_conv_nodes = sorted({l.split("MPST: ")[-1].rstrip(".") for l in scope_lines
                               if l.startswith("Port speed 1G to 10G conversion with MPST:")}
                              | {r["node"] for r in port_conv_swap_completed})
    sfp_trigger_nodes = sorted(set(new_nodes) | set(port_conv_nodes))
    sfp_models_by_node = {}
    if sfp_trigger_nodes:
        with st.container(border=True):
            st.markdown(f"**Transport SFP Installation** \u2014 {len(sfp_trigger_nodes)} node(s). SFP models are manual:")
            for node in sfp_trigger_nodes:
                c1, c2, c3 = st.columns([1, 1, 1])
                with c1:
                    st.caption(node)
                with c2:
                    bbu = st.text_input("BBU End", key=f"sfp_bbu_{node}", label_visibility="collapsed", placeholder="SFP Model (BBU End)")
                with c3:
                    siad = st.text_input("SIAD End", key=f"sfp_siad_{node}", label_visibility="collapsed", placeholder="SFP Model (SIAD End)")
                sfp_models_by_node[node] = (bbu, siad)
    transport_sfp_lines = mcl.transport_sfp_installation_lines(sfp_trigger_nodes, sfp_models_by_node) if sfp_trigger_nodes else []

    sfp_pending_extra, sfp_pre_existing_extra = [], []
    transport_sfp_data = {}
    if postcheck_text:
        transport_sfp_data = mcl.extract_transport_sfp(postcheck_text)
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

    extra_completed_text = "\n".join(
        gps_extra_completed + transport_sfp_lines + radio_swap_completed_lines
        + [r["text"] for r in port_conv_swap_completed] + fdd_lines_fixed
        + ([lkf_lines_by_section["Completed"]] if lkf_lines_by_section.get("Completed") else [])
    )
    extra_pending_text = "\n".join(
        gps_extra_pending + sfp_pending_extra + radio_swap_pending_lines
        + ([edp_publish_text] if edp_publish_text else [])
        + ([lkf_lines_by_section["Pending"]] if lkf_lines_by_section.get("Pending") else [])
    )
    if testing_note:
        extra_pending_text = (extra_pending_text + "\n" + testing_note).strip()

    with st.expander(f"Completed ({sum(1 for i in completed_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(completed_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice
        additional_completed = st.text_area("Enter any additional completed information that needs to be added in report",
                                             value=extra_completed_text, key="rpt_add_completed", height=100)
        choices["additional_completed"] = {"text": additional_completed}

    with st.expander(f"Pending ({sum(1 for i in pending_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(pending_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice
                stakeholders[item["key"]] = stakeholder
        additional_pending = st.text_area("Enter any additional pending information that needs to be reported to Market",
                                           value=extra_pending_text, key="rpt_add_pending", height=100)
        choices["additional_pending"] = {"text": additional_pending}

    locked_ports_exist = bool(controller_checks_data) and any(
        p["admin"] == "LOCKED" and p["slogan"] for p in controller_checks_data.get("alarm_ports", []))

    bucket_pre_existing, bucket_pending = [], []
    if locked_ports_exist:
        with st.container(border=True):
            st.markdown("**Locked alarm ports** \u2014 classify each locked port (per the confirmed "
                        "6610 Alarm Cutover reporting standard). Leave blank whichever don't apply.")
            b1 = st.text_input("1. Pre-existing locked \u2014 port numbers", key="lp_b1", placeholder="e.g. 1, 5, 25")
            b2 = st.text_input("2. Pre-existing active alarm \u2014 port numbers", key="lp_b2", placeholder="e.g. 3, 6, 20")
            c1, c2 = st.columns([2, 1])
            with c1:
                b3 = st.text_input("3. Pre-existing loops/bridge clips/no equipment connections \u2014 port numbers", key="lp_b3")
            with c2:
                b3_note = st.text_input("Note (optional)", key="lp_b3_note", label_visibility="collapsed", placeholder="Note (optional)")
            c1, c2 = st.columns([2, 1])
            with c1:
                b4 = st.text_input("4. Post-cutover, FE couldn't clear \u2014 port numbers", key="lp_b4")
            with c2:
                b4_owner = st.selectbox("Owner", ["Tower Crew", "AT&T"], key="lp_b4_owner", label_visibility="collapsed")
            c1, c2 = st.columns([2, 1])
            with c1:
                b5 = st.text_input("5. Other (free entry)", key="lp_b5")
            with c2:
                b5_dest = st.selectbox("Goes to", ["Pre-Existing Issues", "Pending"], key="lp_b5_dest", label_visibility="collapsed")
            c1, c2 = st.columns([2, 1])
            with c1:
                b6 = st.text_input("6. Other (free entry)", key="lp_b6")
            with c2:
                b6_dest = st.selectbox("Goes to", ["Pre-Existing Issues", "Pending"], key="lp_b6_dest", label_visibility="collapsed")

            t1 = mcl.locked_port_bucket_1(b1)
            t2 = mcl.locked_port_bucket_2(b2)
            t3 = mcl.locked_port_bucket_3(b3, b3_note)
            t4 = mcl.locked_port_bucket_4(b4, b4_owner)
            for t in (t1, t2, t3):
                if t:
                    bucket_pre_existing.append(t)
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

    with st.expander("Pre-Existing Issues"):
        pre_existing_default = "\n".join(sfp_pre_existing_extra + bucket_pre_existing)
        pre_existing_text = st.text_area("Enter any Pre-Existing Issues that needs to be reported to Market",
                                          value=pre_existing_default, key="rpt_preexisting", height=70)
        choices["pre_existing_issues_text"] = pre_existing_text

    with st.expander("Notes"):
        note_defs = [
            ("notes_final_port_config", "Final Port Configuration attached."),
            ("notes_nr_verified", "NR configuration has been verified."),
            ("notes_cpri_sfp", "Area prechecks verification for CPRI/SFP check is completed."),
            ("notes_no_external_alarms", "No scope of external alarms."),
        ]
        cols = st.columns(2)
        for i, (note_key, text) in enumerate(note_defs):
            with cols[i % 2]:
                checked = st.checkbox(text, key=f"chk_{note_key}")
                choices[note_key] = {"checked": checked, "text": text}

        c1, c2 = st.columns(2)
        with c1:
            n_mme = st.checkbox("Pre-Existing MME configuration left as it is on node", key="chk_notes_mme_config")
            mme_node = st.text_input("Node ID", key="rpt_mme_node", label_visibility="collapsed") if n_mme else ""
            choices["notes_mme_config"] = {"checked": n_mme, "text": f"Pre-Existing MME configuration left as it is on node {mme_node}"}
        with c2:
            n_mon = st.checkbox("Node is in monitored state", key="chk_notes_monitored")
            mon_node = st.text_input("Node ID (monitored)", key="rpt_mon_node", label_visibility="collapsed") if n_mon else ""
            choices["notes_monitored"] = {"checked": n_mon, "text": f"{mon_node} is in monitored state."}

        n_not_mon = st.checkbox("Node is in not monitored state", key="chk_notes_not_monitored")
        not_mon_node = st.text_input("Node ID (not monitored)", key="rpt_not_mon_node", label_visibility="collapsed") if n_not_mon else ""
        choices["notes_not_monitored"] = {"checked": n_not_mon, "text": f"{not_mon_node} is in not monitored state."}

        emergency_unlock_default = "\n".join(
            f"Emergency unlock activated on the node {n}." for n in emergency_unlock_notes)
        notes_generic = st.text_area("Enter Notes that need to be reported or addressed to Market",
                                      value=emergency_unlock_default, key="rpt_notes_generic", height=70)
        choices["notes_generic_text"] = notes_generic

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
            "mic": "MIC", "market": market, "status": status, "site_name": site_name,
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
        row_writes.append((3, True, [(2, "MIC"), (3, market), (4, status), (5, site_name), (6, fa_code), (7, site_ids), (8, sow)]))
        row_writes.append((6, True, [(3, iwm_details)]))
        row_writes.append((10, True, [(3, pre_line)]))
        row_writes.append((11, bool(current_config.strip()), [(3, current_config)]))
        row_writes.append((12, True, [(3, post_line)]))
        row_writes.append((13, bool(wll_node.strip()), [(3, wll_node)]))
        row_writes.append((14, True, [(3, controller_id)]))
        row_writes.append((15, True, [(3, software_version)]))
        row_writes.append((16, True, [(3, gs_version)]))
        if idl_build_type:
            row_writes.append((19, True, [(3, idl_build_type)]))

        if not TEMPLATE_PATH.exists():
            static_dir = TEMPLATE_PATH.parent
            if static_dir.exists():
                actual_files = sorted(p.name for p in static_dir.iterdir())
                st.warning(f"Template not found at {TEMPLATE_PATH}. Files actually present in {static_dir}: {actual_files}")
            else:
                st.warning(f"Template not found at {TEMPLATE_PATH} \u2014 and the folder {static_dir} doesn't exist at all.")
        else:
            xlsm_bytes = fill_legacy_mca(TEMPLATE_PATH, {"row_writes": row_writes})
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_Legacy_MCA_Filled.xlsm", key="rpt_dl_xlsm")
