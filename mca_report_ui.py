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

    ctx = _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp)
    results = mca_checklist.evaluate_checklist(ctx)

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

    site_ids = "/".join(r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as"))
    fa_code = ""
    if "5G Info" in ciq_wb.sheetnames:
        for row in app.sheet_objs(ciq_wb["5G Info"]):
            if app.is_populated(row.get("FA Code")):
                fa_code = row.get("FA Code")
                break
    default_status = "STF" if any(r["section"] == "pending" and r["checked_by_default"] for r in results) else "ATP"

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
            current_config = st.text_input("Current Configuration (if applicable)", key="rpt_current_config")
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
                switch = st.text_area("Switch details (manual)", key="rpt_switch", height=60)
            with c2:
                idly = st.text_area("IDLy cable details (manual)", key="rpt_idly", height=60)
                slot_port = st.text_area("Slot/Port/Cable/Node ID (manual)", key="rpt_slotport", height=60)

    st.markdown("### Report preview")
    report_placeholder = st.empty()

    st.markdown("### Which of these apply?")
    st.caption("Checked = included in the report above. Uncheck anything that doesn't apply to this site; nothing else to fill in per item — use the manual boxes below for anything not auto-detected.")

    choices, stakeholders = {}, {}
    completed_items = [i for i in results if i["section"] == "completed"]
    pending_items = [i for i in results if i["section"] == "pending"]

    with st.expander(f"Completed ({sum(1 for i in completed_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(completed_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice
        additional_completed = st.text_area("Enter any additional completed information that needs to be added in report", key="rpt_add_completed", height=70)
        choices["additional_completed"] = {"text": additional_completed}

    with st.expander(f"Pending ({sum(1 for i in pending_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(pending_items):
            with cols[i % 2]:
                choice, stakeholder = _simple_item_row(item)
                choices[item["key"]] = choice
                stakeholders[item["key"]] = stakeholder
        additional_pending = st.text_area("Enter any additional pending information that needs to be reported to Market", key="rpt_add_pending", height=70)
        choices["additional_pending"] = {"text": additional_pending}

    if warnings:
        with st.expander(f"⚠️ Warnings ({len(warnings)})", expanded=True):
            st.markdown(
                "<div style='color:#c0392b; font-weight:600;'>These are informational only "
                "— they never change what's in the report, they're here so nothing gets missed.</div>",
                unsafe_allow_html=True)
            for w in warnings:
                st.markdown(f"<div style='color:#c0392b;'>• {w['text']}</div>", unsafe_allow_html=True)

    with st.expander("Pre-Existing Issues"):
        pre_existing_text = st.text_area("Enter any Pre-Existing Issues that needs to be reported to Market", key="rpt_preexisting", height=70)
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

        notes_generic = st.text_area("Enter Notes that need to be reported or addressed to Market", key="rpt_notes_generic", height=70)
        choices["notes_generic_text"] = notes_generic

    header_fields = {
        "mic": "MIC", "market": market, "status": status, "site_name": site_name,
        "fa_code": fa_code, "site_ids": site_ids, "sow": sow, "iwm_details": iwm_details,
        "pre_configuration": pre_line, "current_configuration": current_config,
        "post_configuration": post_line, "wll_node": wll_node, "controller_id": controller_id,
        "software_version": software_version, "gs_version": gs_version,
        "idl_build_type": idl_build_type, "idle": idle, "idly": idly, "switch": switch, "slot_port": slot_port,
    }
    report_text = mca_report_text.build_mca_report_text(mm_objs, results, choices, header_fields, stakeholder_by_key=stakeholders)
    report_placeholder.text_area("Report preview (live — updates as you check/uncheck items below)",
                                  report_text, height=400, key="rpt_preview_live")

    st.markdown("---")
    node_tag = mm_objs[0].get("Node to be built as", "site") if mm_objs else "site"
    st.download_button("Download report (.txt)", report_text, file_name=f"{node_tag}_Integration_Report.txt", key="rpt_dl_txt")

    if st.button("Generate filled checklist (.xlsm) \u2192", key="rpt_generate_mca"):
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
