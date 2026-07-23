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
from mca_row_map import ROW_MAP
from mca_xlsm_fill import fill_legacy_mca

TEMPLATE_PATH = "templates/Static/Legacy_MCA_Macro_Template_v6_1.xlsm"
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


def _item_card(item):
    """One item, rendered as a compact bordered card: checkbox + live detected preview always
    visible, a manual-entry box available whenever checked (whether auto-detected or not), and
    a stakeholder selector whenever the item sits in Pending."""
    key = item["key"]
    with st.container(border=True):
        checked = st.checkbox(item["label"], value=item["checked_by_default"], key=f"chk_{key}")
        preview = _detected_preview(item)
        if preview:
            st.caption(f"\u2713 Detected: {preview}")

        section = item["section"]
        stakeholder = item.get("stakeholder", "").split("|")[0] if item.get("stakeholder") else "MIC PM"
        manual_extra = []

        if checked:
            if item.get("toggle"):
                section = st.radio("Section", ["completed", "pending"], key=f"sec_{key}", horizontal=True)
            cols = st.columns([3, 1]) if section == "pending" else [st]
            with cols[0]:
                manual_val = st.text_input("Manual entry / override (leave blank to use detected value)", key=f"manual_{key}", label_visibility="collapsed", placeholder="Manual entry / override")
                if manual_val:
                    manual_extra.append(manual_val)
            if section == "pending":
                with cols[1]:
                    stakeholder = st.selectbox("Stakeholder", STAKEHOLDER_OPTIONS,
                                                index=STAKEHOLDER_OPTIONS.index(stakeholder) if stakeholder in STAKEHOLDER_OPTIONS else 1,
                                                key=f"stake_{key}", label_visibility="collapsed")
            if item.get("manual_fields"):
                fcols = st.columns(len(item["manual_fields"]))
                for c, field_name in zip(fcols, item["manual_fields"]):
                    with c:
                        manual_extra.append(st.text_input(field_name, key=f"manualfield_{key}_{field_name}"))

    return {"checked": checked, "section": section, "manual_extra": manual_extra}, stakeholder


def render(app, ciq_wb, mm_objs, controller_objs, precheck_text, pre_line, post_line, scope_lines):
    st.subheader("Generate Report")

    idl_build_type = app.derive_idl_build_type_label(ciq_wb, mm_objs)
    controller_id = _get_controller_id(controller_objs)
    controller_in_edp = bool(controller_id)

    ctx = _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp)
    results = mca_checklist.evaluate_checklist(ctx)

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

    st.markdown("### Completed / Pending checklist")
    st.caption("Auto-detected items are pre-checked, with the detected value shown right below. Uncheck anything that doesn't apply; check anything extra that does, and use the manual-entry box to override or supply a value.")

    choices, stakeholders = {}, {}
    completed_items = [i for i in results if i["section"] == "completed"]
    pending_items = [i for i in results if i["section"] == "pending"]

    with st.expander(f"Completed ({sum(1 for i in completed_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(completed_items):
            with cols[i % 2]:
                choice, stakeholder = _item_card(item)
                choices[item["key"]] = choice
                if choice["section"] == "pending":
                    stakeholders[item["key"]] = stakeholder
        additional_completed = st.text_area("Enter any additional completed information that needs to be added in report", key="rpt_add_completed", height=70)
        choices["additional_completed"] = {"text": additional_completed}

    with st.expander(f"Pending ({sum(1 for i in pending_items if i['checked_by_default'])} auto-detected)", expanded=True):
        cols = st.columns(2)
        for i, item in enumerate(pending_items):
            with cols[i % 2]:
                choice, stakeholder = _item_card(item)
                choices[item["key"]] = choice
                stakeholders[item["key"]] = stakeholder
        additional_pending = st.text_area("Enter any additional pending information that needs to be reported to Market", key="rpt_add_pending", height=70)
        choices["additional_pending"] = {"text": additional_pending}

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

    st.markdown("---")
    if st.button("Generate Integration Report \u2192", type="primary", key="rpt_generate_mca"):
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

        node_tag = mm_objs[0].get("Node to be built as", "site") if mm_objs else "site"
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

        try:
            output_path = "/mnt/user-data/outputs/Legacy_MCA_Filled.xlsm"
            fill_legacy_mca(TEMPLATE_PATH, output_path, {"row_writes": row_writes})
            with open(output_path, "rb") as f:
                st.download_button("Download filled checklist (.xlsm)", f.read(),
                                    file_name=f"{node_tag}_Legacy_MCA_Filled.xlsm", key="rpt_dl_xlsm")
        except FileNotFoundError:
            st.warning(f"Template not found at {TEMPLATE_PATH} \u2014 upload Legacy_MCA_Macro_Template_v6_1.xlsm "
                      "to templates/Static/ in the repo to enable the .xlsm output.")
