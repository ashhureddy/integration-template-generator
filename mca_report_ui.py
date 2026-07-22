"""
MCA Integration Report — interactive 'Generate Report' UI section.
Import this module's render() function and call it inside app.py, gated to top_scope=='MCA',
passing ciq_wb, mm_objs, controller_objs, precheck_text, pre_line, post_line.
"""
import streamlit as st
import io

import report_detect
import mca_checklist
import mca_glue
import mca_report_text
from mca_row_map import ROW_MAP
from mca_xlsm_fill import fill_legacy_mca

TEMPLATE_PATH = "templates/Static/Legacy_MCA_Macro_Template_v6_1.xlsm"


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


def render(app, ciq_wb, mm_objs, controller_objs, precheck_text, pre_line, post_line, scope_lines):
    st.subheader("Generate Report")
    st.caption("Interactive Integration Report — auto-fetched fields are pre-filled and pre-checked; "
               "everything else needs manual entry. Uncheck anything that doesn't apply.")

    idl_build_type = app.derive_idl_build_type_label(ciq_wb, mm_objs)
    controller_id = controller_objs[0].get("Controller") if controller_objs else ""
    controller_in_edp = bool(controller_id)  # refined once EDP-check wiring is confirmed

    ctx = _build_ctx(app, ciq_wb, mm_objs, precheck_text, scope_lines, idl_build_type, controller_id, controller_in_edp)
    results = mca_checklist.evaluate_checklist(ctx)

    st.markdown("#### Header fields")
    c1, c2 = st.columns(2)
    with c1:
        market = st.text_input("Market", key="rpt_market")
        status = st.text_input("Status (ATP/STF \u2014 auto-suggested, override if needed)",
                                value="STF" if any(r["section"] == "pending" and r["checked_by_default"] for r in results) else "ATP",
                                key="rpt_status")
        site_name = st.text_input("Site Name", key="rpt_site_name")
        sow = st.text_input("SOW", key="rpt_sow")
        iwm_details = st.text_input("IWM Details", key="rpt_iwm")
    with c2:
        current_config = st.text_input("Current Configuration (if applicable)", key="rpt_current_config")
        wll_node = st.text_input("WLL node (if applicable)", key="rpt_wll")
        software_version = st.text_input("Software version", key="rpt_sw")
        gs_version = st.text_input("GS Version", key="rpt_gs")

    site_ids = "/".join(r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as"))
    fa_code = ""
    if "5G Info" in ciq_wb.sheetnames:
        for row in app.sheet_objs(ciq_wb["5G Info"]):
            if app.is_populated(row.get("FA Code")):
                fa_code = row.get("FA Code")
                break
    st.caption(f"Auto-fetched: Site IDs = **{site_ids}**, FA Code = **{fa_code or '(not found)'}**, "
               f"6610 Controller = **{controller_id or '(none)'}**, IDL Build Type = **{idl_build_type or '(n/a)'}**")

    idle = st.text_area("IDLe cable details (manual)", key="rpt_idle", height=60)
    idly = st.text_area("IDLy cable details (manual)", key="rpt_idly", height=60)
    switch = st.text_area("Switch details (manual)", key="rpt_switch", height=60)
    slot_port = st.text_area("Slot/Port/Cable/Node ID (manual)", key="rpt_slotport", height=60)

    st.markdown("#### Completed / Pending checklist")
    choices = {}
    for item in results:
        if item["key"] in ("radio_swap",):
            continue  # rendered in the Pending-only pass below
        key = item["key"]
        cols = st.columns([3, 2, 2])
        with cols[0]:
            checked = st.checkbox(item["label"], value=item["checked_by_default"], key=f"chk_{key}")
        section = item["section"]
        stakeholder = None
        if item.get("toggle"):
            with cols[1]:
                section = st.radio("Completed or Pending?", ["completed", "pending"], key=f"sec_{key}", horizontal=True)
            if section == "pending":
                with cols[2]:
                    stakeholder = st.selectbox("Stakeholder", ["AT&T", "MIC"], key=f"stake_{key}")
        manual_extra = []
        if item.get("manual_fields"):
            for field_name in item["manual_fields"]:
                manual_extra.append(st.text_input(f"{item['label']} \u2014 {field_name}", key=f"manual_{key}_{field_name}"))
        choices[key] = {"checked": checked, "section": section, "manual_extra": manual_extra}
        if stakeholder:
            choices.setdefault("_stakeholders", {})[key] = stakeholder

    st.markdown("#### Pending-only items")
    radio_swap_item = next(i for i in results if i["key"] == "radio_swap")
    checked = st.checkbox(radio_swap_item["label"], value=radio_swap_item["checked_by_default"], key="chk_radio_swap")
    choices["radio_swap"] = {"checked": checked, "section": "pending"}

    additional_completed = st.text_area("Enter any additional completed information that needs to be added in report", key="rpt_add_completed", height=80)
    additional_pending = st.text_area("Enter any additional pending information that needs to be report to Market", key="rpt_add_pending", height=80)
    choices["additional_completed"] = {"text": additional_completed}
    choices["additional_pending"] = {"text": additional_pending}

    st.markdown("#### Pre-Existing Issues / Notes")
    pre_existing_text = st.text_area("Enter any Pre-Existing Issues that needs to be reported to Market", key="rpt_preexisting", height=80)
    choices["pre_existing_issues_text"] = pre_existing_text

    n1 = st.checkbox("Final Port Configuration attached.", key="rpt_note1")
    n2 = st.checkbox("NR configuration has been verified.", key="rpt_note2")
    n3 = st.checkbox("Pre-Existing MME configuration left as it is on node xxNode_Idxx", key="rpt_note3")
    note3_node = st.text_input("Node ID for the MME note above (if checked)", key="rpt_note3_node") if n3 else ""
    choices["notes_final_port_config"] = {"checked": n1, "text": "Final Port Configuration attached."}
    choices["notes_nr_verified"] = {"checked": n2, "text": "NR configuration has been verified."}
    choices["notes_mme_config"] = {"checked": n3, "text": f"Pre-Existing MME configuration left as it is on node {note3_node}"}
    notes_generic = st.text_area("Enter Notes that need to be reported or addressed to Market", key="rpt_notes_generic", height=80)
    choices["notes_generic_text"] = notes_generic

    if st.button("Generate Integration Report \u2192", type="primary", key="rpt_generate_mca"):
        header_fields = {
            "mic": "MIC", "market": market, "status": status, "site_name": site_name,
            "fa_code": fa_code, "site_ids": site_ids, "sow": sow, "iwm_details": iwm_details,
            "pre_configuration": pre_line, "current_configuration": current_config,
            "post_configuration": post_line, "wll_node": wll_node, "controller_id": controller_id,
            "software_version": software_version, "gs_version": gs_version,
            "idl_build_type": idl_build_type, "idle": idle, "idly": idly, "switch": switch, "slot_port": slot_port,
        }
        report_text = mca_report_text.build_mca_report_text(mm_objs, results, choices, header_fields,
                                                              stakeholder_by_key=choices.get("_stakeholders"))
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
