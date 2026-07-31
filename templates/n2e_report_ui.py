"""
N2E (Nokia-to-Ericsson) Report UI. Mirrors mca_report_ui.py's proven architecture
(checkbox-per-item, no defaults where a real choice is needed, buffer overflow pools,
surgical .xlsm patching) but every item's actual logic is N2E-specific, confirmed
through the N2E design conversation — nothing is assumed to carry over from MCA.

Reuses app.generate_n2e() directly for what it already computes correctly (Integration,
6610 Controller Integration, DSS Activation, IDL connections, NGS activation, SA
Conversion) — those come from real, already-tested code. Everything requiring
Post-checks or controller-checks (GPS, SUP/XMU, Area test, Link failure/SFP, Active
external alarm, 6673 items) is new, built and confirmed this session.
"""

import re
import streamlit as st

import n2e_completed_logic as n2e
import mca_completed_logic as mcl
from n2e_row_map import N2E_ROW_MAP
from mca_xlsm_surgical import fill_legacy_mca_surgical
from pathlib import Path

N2E_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "N2E_Pre_IX_Macro_V2.xlsm"


def _checked_group(label, lines, key):
    """Same confirmed-safe pattern as MCA — checkbox recomputes fresh every rerun, no
    stale text_area default issue."""
    if not lines:
        return False, []
    checked = st.checkbox(label, value=True, key=key)
    if checked:
        for l in lines:
            st.caption(l)
    return checked, (lines if checked else [])


def render(app, ciq_wb, mm_objs, controller_objs, edp_index, user_id, date_str,
           postcheck_text="", controller_checks_text=""):
    st.subheader("Generate N2E Report")
    st.markdown("""
        <style>
        div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
            border: 2px solid #ff9800 !important; background-color: #fff8e1 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    n2e.set_app_module(app)

    # ---- Reuse the existing, tested generate_n2e() for Integration/6610/DSS/IDL/NGS/SA
    # Conversion — confirmed real, working code, no need to re-derive. ----
    def _log(msg):
        pass
    summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = app.generate_n2e(
        ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, _log)

    controller_id = None
    for row in controller_objs:
        if str(row.get("Controller", "")).strip() == "6610":
            controller_id = row.get("Controller ID")
            break
    controller_in_edp = bool(controller_id)

    # ---- Classification, needed for Integration band/node extraction and PSAP/Speedtest
    # band filtering (confirmed: every CIQ cell = addition, no Pre-checks). ----
    classification = {"added": {}}
    eutran_objs = app.sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    fiveg_objs = app.sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
    for row in mm_objs:
        node = row.get("Node to be built as")
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        cells = []
        for r in eutran_objs:
            c = r.get("EutranCellFDDId")
            if c and e_name and str(c).startswith(str(e_name)):
                cells.append(c)
        for r in fiveg_objs:
            c = r.get("NRCellDU")
            if c and g_name and str(c).startswith(str(g_name)):
                cells.append(c)
        if cells:
            classification["added"][node] = cells

    new_nodes = [row.get("Node to be built as") for row in mm_objs]  # confirmed: every node is "new" for N2E

    # ==================== Subject / Configuration / IDL Connections ====================
    with st.container(border=True):
        st.markdown("**Subject** (MIC / MNS / N2E / IX-STF are constants)")
        site_name = st.text_input("\U0001F4DD Site Name", key="n2e_site_name")
        sow = st.text_input("\U0001F4DD SOW", key="n2e_sow")
        fa_code = ""
        if "5G Info" in ciq_wb.sheetnames:
            for row in app.sheet_objs(ciq_wb["5G Info"]):
                if app.is_populated(row.get("FA Code")):
                    fa_code = row.get("FA Code")
                    break
        site_ids = "/".join([row.get("Node to be built as") for row in mm_objs if row.get("Node to be built as")]
                             + ([controller_id] if controller_id else []))
        st.caption(f"FA Code (auto): {fa_code or '(not detected)'}")
        st.caption(f"Site ID's (auto): {site_ids}")
        iwm_details = st.text_input("\U0001F4DD IWM Details", key="n2e_iwm")

    with st.container(border=True):
        st.markdown("**Configuration**")
        st.markdown(f"Pre Configuration : **Nokia**")
        st.markdown(f"Post Configuration : **{post_line}**")
        current_config_auto = mcl.current_configuration_line(ciq_wb, mm_objs, postcheck_text) if postcheck_text else ""
        if current_config_auto:
            current_config = st.text_input("\U0001F4DD Current Configuration \u2014 review/edit:",
                                             value=current_config_auto, key="n2e_current_config")
        else:
            current_config = ""
        c1, c2 = st.columns(2)
        with c1:
            wll_node = st.text_input("\U0001F4DD WLL node", key="n2e_wll")
            software_version = st.text_input("\U0001F4DD Software version", key="n2e_sw")
        with c2:
            gs_version = st.text_input("\U0001F4DD GS Version", key="n2e_gs")

    xmu_present_in_ciq = n2e.xmu_in_ciq(post_line)

    idl_build_type = None
    idle = idly = switch = slot_port = ""
    if len(mm_objs) > 1:
        with st.container(border=True):
            idl_build_type = app.derive_idl_build_type_label(ciq_wb, mm_objs)
            st.markdown(f"**IDL Connections** \u2014 Build Type: **{idl_build_type or '(not detected)'}**")
            c1, c2 = st.columns(2)
            with c1:
                idle = st.text_area("\U0001F4DD IDLe cable details (manual)", key="n2e_idle", height=60)
            with c2:
                idly = st.text_area("\U0001F4DD IDLy cable details (manual)", key="n2e_idly", height=60)
            sidehaul_rows = mcl.sidehaul_display_rows(ciq_wb)
            if sidehaul_rows:
                st.caption("Switch / Slot-Port \u2014 auto-filled from Sidehaul Info:")
                cable_pns = {}
                for i, srow in enumerate(sidehaul_rows):
                    sc1, sc2, sc3, sc4, sc5 = st.columns([1, 1, 1, 1, 1])
                    with sc1: st.caption(f"**{srow['switch_type']}**")
                    with sc2: st.caption(srow["switch_id"])
                    with sc3: st.caption(srow["slot_port"])
                    with sc4: cable_pns[i] = st.text_input("Cable P/N", key=f"n2e_cable_pn_{i}", label_visibility="collapsed")
                    with sc5: st.caption(srow["node_id"])
                switch = "\n".join(mcl.format_sidehaul_lines(sidehaul_rows, cable_pns))
            else:
                sidehaul_rows = []
                switch = st.text_area("\U0001F4DD Switch details (manual)", key="n2e_switch", height=60)
                slot_port = st.text_area("\U0001F4DD Slot/Port/Cable/Node ID (manual)", key="n2e_slotport", height=60)
    else:
        sidehaul_rows = []

    has_6673 = n2e.has_6673(sidehaul_rows)

    # ==================== Completed / Pending items ====================
    st.markdown("### Which of these apply?")
    choices_completed, choices_pending, stakeholders = [], [], {}
    warnings = []

    with st.expander("Completed", expanded=True):
        # Integration — from generate_n2e's own scope_lines.
        integration_line = next((l for l in scope_lines if l.startswith("Integration")), None)
        int_checked, int_lines = _checked_group("Integration", [integration_line] if integration_line else [], "n2e_int")
        if int_checked:
            choices_completed += int_lines

        # 6610 Controller Integration + cascade.
        controller_checks_data = mcl.extract_controller_checks(controller_checks_text) if controller_checks_text else {}
        cascade_fires = n2e.controller_cascade_fires(controller_in_edp, bool(controller_checks_text))
        controller_line = next((l for l in scope_lines if l.startswith("6610 Controller Integration") or l.startswith("EDP Publish")), None)
        if cascade_fires:
            st.caption("\u26a0\ufe0f Controller-checks file not uploaded \u2014 6610 Controller Integration, SAU Connections, "
                       "LKF Installation, External alarm Scripting/testing all move to Pending.")
        elif controller_line:
            ctrl_checked, ctrl_lines = _checked_group("6610 Controller Integration", [controller_line], "n2e_ctrl")
            if ctrl_checked:
                choices_completed += ctrl_lines

        # DSS Activation — manual 2-way, AT&T/MIC if Pending.
        dss_line = next((l for l in scope_lines if l.startswith("DSS Activation")), None)
        dss_completed, dss_pending = None, None
        if dss_line:
            with st.container(border=True):
                st.markdown("**DSS Activation**")
                dss_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"], key="n2e_dss")
                if dss_choice == "Completed":
                    dss_completed = dss_line.replace("\t", " ")
                elif dss_choice == "Pending":
                    dss_sh = st.selectbox("Stakeholder", ["\u2014 Select \u2014", "AT&T", "MIC"], key="n2e_dss_sh")
                    if dss_sh != "\u2014 Select \u2014":
                        dss_pending = f"{dss_line.replace(chr(9), ' ')} ({dss_sh})"
            if dss_completed:
                choices_completed.append(dss_completed)

        # NGS activation — manual 2-way, fixed MIC if Pending.
        ngs_line = next((l for l in scope_lines if l.startswith("NGS Activation")), None)
        ngs_completed, ngs_pending = None, None
        if ngs_line:
            with st.container(border=True):
                st.markdown("**NGS activation**")
                ngs_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"], key="n2e_ngs")
                if ngs_choice == "Completed":
                    ngs_completed = ngs_line.replace("\t", " ")
                elif ngs_choice == "Pending":
                    ngs_pending = f"{ngs_line.replace(chr(9), ' ')} (MIC)"
            if ngs_completed:
                choices_completed.append(ngs_completed)

        # GPS Installation — Post-checks Sync Status 2 TimeSyncIO.
        gps_completed_line, gps_pending_line = None, None
        if postcheck_text:
            post_sync = mcl.extract_sync_status_2(postcheck_text)
            post_gps = mcl.extract_gps_status(postcheck_text)
            enabled_nodes, disabled_nodes = n2e.gps_sync_status(mm_objs, post_sync)
            if enabled_nodes:
                gtype = post_gps.get(enabled_nodes[0], "")
                gps_completed_line = n2e.gps_installation_line(enabled_nodes, gtype)
                choices_completed.append(gps_completed_line)
                st.caption(f"Detected: {gps_completed_line}")
            if disabled_nodes:
                gps_pending_line = f"GPS Installation: {'|'.join(disabled_nodes)} (MIC PM)"

        # LKF Installation — per-entity, no default.
        lkf_completed, lkf_pending = None, None
        with st.container(border=True):
            st.markdown("**LKF Installation** \u2014 per node/controller, required:")
            lkf_c_nodes, lkf_p_nodes = [], []
            for row in mm_objs:
                node = row.get("Node to be built as")
                c1, c2 = st.columns([2, 1])
                with c1: st.caption(node)
                with c2:
                    pick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                          key=f"n2e_lkf_{node}", label_visibility="collapsed")
                    (lkf_c_nodes if pick == "Completed" else lkf_p_nodes if pick == "Pending" else []).append(node) \
                        if pick != "\u2014 Select \u2014" else None
            if lkf_c_nodes:
                lkf_completed = f"LKF Installation: {'|'.join(lkf_c_nodes)}|{controller_id or ''}."
                choices_completed.append(lkf_completed)
            if lkf_p_nodes:
                lkf_pending = f"LKF Installation: {'|'.join(lkf_p_nodes)}|{controller_id or ''}. (MIC)"

        # Transport SFP — new-node trigger, manual SFP model.
        sfp_completed_lines = []
        if new_nodes:
            with st.container(border=True):
                st.markdown("**Transport SFP Installation on** \u2014 manual SFP models:")
                for node in new_nodes:
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1: st.caption(node)
                    with c2: bbu = st.text_input("BBU", key=f"n2e_sfp_bbu_{node}", label_visibility="collapsed", placeholder="SFP Model (BBU End)")
                    with c3: siad = st.text_input("SIAD", key=f"n2e_sfp_siad_{node}", label_visibility="collapsed", placeholder="SFP Model (SIAD End)")
                    if bbu.strip() or siad.strip():
                        sfp_completed_lines.append(f"Transport SFP Installation on: {node} {bbu} (BBU End) & {siad} (SIAD End)")
            choices_completed += sfp_completed_lines

        # RET configuration — always Pending, never Completed.
        ret_pending = "RET configuration update. (Tower Crew)"

        # External alarm Scripting / SAU Connections.
        alarm_scripting_completed, alarm_scripting_pending = None, None
        sau_completed, sau_pending = None, None
        if controller_checks_data and not cascade_fires:
            if mcl.external_alarm_scripting_confirmed(controller_checks_data):
                alarm_scripting_completed = f"External alarm Scripting: {controller_id}"
                choices_completed.append(alarm_scripting_completed)
            sau_state = controller_checks_data.get("sau_state")
            if sau_state:
                if sau_state["oper"] == "ENABLED":
                    sau_completed = f"SAU Connection: {controller_id}."
                    choices_completed.append(sau_completed)
                else:
                    sau_pending = f"SAU Connections: {controller_id}. (MIC PM)"

        # SUP Connections / XMU Installation.
        sup_completed_lines, sup_pending_lines = [], []
        xmu_completed_lines, xmu_pending_lines = [], []
        if postcheck_text and xmu_present_in_ciq:
            sup_state = n2e.sup_connections_state(postcheck_text, xmu_present_in_ciq)
            for node, state in sup_state.items():
                (sup_completed_lines if state == "ENABLED" else sup_pending_lines).append(
                    f"SUP Connections: {node}" + ("" if state == "ENABLED" else ". (MIC PM)"))
            xmu_state = n2e.xmu_installation_state(postcheck_text, xmu_present_in_ciq)
            for node, state in xmu_state.items():
                (xmu_completed_lines if state == "ENABLED" else xmu_pending_lines).append(
                    f"XMU Installation: {node}" + ("" if state == "ENABLED" else ". (MIC PM)"))
        choices_completed += sup_completed_lines + xmu_completed_lines

        # IDL connections — 2+ nodes trigger, manual Completed/Pending.
        idl_completed, idl_pending = None, None
        if len(mm_objs) > 1:
            idl_choice = st.selectbox("IDL connections status", ["\u2014 Select \u2014", "Completed", "Pending"], key="n2e_idlconn")
            if idl_choice == "Completed":
                idl_completed = "IDL connections"
                choices_completed.append(idl_completed)
            elif idl_choice == "Pending":
                idl_pending = "IDL connections (MIC PM)"

        # SMM Triggering — manual toggle, MIC PM if Pending.
        smm_completed, smm_pending = None, None
        smm_choice = st.selectbox("SMM Triggering status", ["\u2014 Select \u2014", "Completed", "Pending"], key="n2e_smm")
        if smm_choice == "Completed":
            smm_completed = "SMM Triggering"
            choices_completed.append(smm_completed)
        elif smm_choice == "Pending":
            smm_pending = "SMM Triggering (MIC PM)"

        # Area test — new-node trigger, manual Area Lite Pass/Fail.
        area_completed, area_pending = None, None
        if new_nodes:
            area_choice = st.selectbox("Area test \u2014 Area Lite result", ["\u2014 Select \u2014", "Passed", "Failed"], key="n2e_area")
            if area_choice == "Passed":
                area_completed = n2e.area_test_line(new_nodes, "Passed")
                choices_completed.append(area_completed)
            elif area_choice == "Failed":
                area_pending = n2e.area_test_line(new_nodes, "Failed") + " (MIC PM)"

        # External alarm testing — Completed if any scripted port unlocked, else Pending+Notes.
        testing_section, testing_note = mcl.external_alarm_testing_placement(controller_checks_data) \
            if controller_checks_data and not cascade_fires else (None, None)
        testing_completed, testing_pending = None, None
        if testing_section == "Completed":
            testing_completed = f"External alarm testing: {controller_id}."
            choices_completed.append(testing_completed)
        elif testing_section == "Pending":
            testing_pending = f"External alarm testing: {controller_id}. (MIC PM)"

        # 6673 Script load — auto, Completed only.
        script_6673_completed = None
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            script_6673_completed = f"6673 Script load: {switch_id}"
            choices_completed.append(script_6673_completed)

        # Installation — manual free text.
        installation_manual = st.text_input("\U0001F4DD Installation (6673/6675/6610/SIAD/Baseband's etc)", key="n2e_install_manual")
        if installation_manual.strip():
            choices_completed.append(f"Installation: {installation_manual}")

        # SA Conversion — CIQ NR_SA tab presence, Completed only.
        sa_nodes = n2e.sa_conversion_nodes(ciq_wb, mm_objs)
        sa_completed_line = None
        if sa_nodes:
            sa_completed_line = f"SA conversion.: {'|'.join(sa_nodes)}"
            choices_completed.append(sa_completed_line)

        additional_completed = st.text_area("\U0001F4DD Enter any additional completed information",
                                             key="n2e_add_completed", height=80)
        if additional_completed.strip():
            choices_completed.append(additional_completed)

    with st.expander("Pending", expanded=True):
        choices_pending.append(ret_pending)
        st.caption(ret_pending)

        choices_pending.append("On Site Nokia cutover (Tower Crew)")
        st.caption("On Site Nokia cutover (Tower Crew)")

        if cascade_fires and controller_line:
            choices_pending.append(f"{controller_line.replace(chr(9), ' ')} (MIC PM)")
            choices_pending.append("SAU Connections. (MIC PM)")
            choices_pending.append("LKF Installation. (MIC)")
            choices_pending.append("External alarm Scripting. (MIC)")
            choices_pending.append("External alarm testing. (MIC PM)")
        else:
            for line in (dss_pending, ngs_pending, gps_pending_line, lkf_pending, sau_pending, testing_pending):
                if line:
                    choices_pending.append(line)
        choices_pending += sup_pending_lines + xmu_pending_lines
        if idl_pending: choices_pending.append(idl_pending)
        if smm_pending: choices_pending.append(smm_pending)
        if area_pending: choices_pending.append(area_pending)
        if testing_note:
            choices_pending.append(testing_note)

        # PSAP/Speedtest + Speed test — Integration bands split LTE/5G, all markets, no lookup.
        lte_bands, fiveg_bands = n2e.integration_bands_by_tech(classification)
        if lte_bands:
            psap_line = f"PSAP/Speed test/VoLTE voice call test on: {lte_bands}. (MIC PM)"
            choices_pending.append(psap_line)
        if fiveg_bands:
            speed_line = f"Speed test on: {fiveg_bands}. (MIC PM)"
            choices_pending.append(speed_line)

        # Compatible Transport SFP Installation (Pending) — manual.
        sfp_pending_manual = st.text_area("\U0001F4DD Compatible Transport SFP Installation (Pending, manual)",
                                            key="n2e_sfp_pending", height=60)
        if sfp_pending_manual.strip():
            choices_pending.append(f"Compatible Transport SFP Installation: {sfp_pending_manual}. (MIC PM)")

        # 6673 Configuration / 6673 Port Configuration in ENM — auto, always Pending.
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            choices_pending.append(f"6673 Configuration: {switch_id}. (AT&T)")
            choices_pending.append(f"6673 Port Configuration in ENM: {switch_id}. (AT&T)")

        # SIAD provisioning — manual.
        siad_manual = st.text_input("\U0001F4DD SIAD provisioning (Node ID, manual)", key="n2e_siad")
        if siad_manual.strip():
            choices_pending.append(f"SIAD provisioning: {siad_manual}. (MIC PM)")

        # Active external alarm — per-port, auto.
        if controller_checks_data:
            active_alarm_lines = n2e.active_external_alarm_lines(controller_checks_data, controller_id or "")
            choices_pending += active_alarm_lines
            for l in active_alarm_lines:
                st.caption(l)

        # Link failure / SFP Not Present — single selector, reuses Integration's full band list.
        full_bands, _ = n2e.integration_bands_and_nodes(classification)
        link_choice = st.selectbox("Link failure / SFP Not Present", ["\u2014 Select \u2014 (neither)", "Link failure", "SFP Not Present"], key="n2e_link_sfp")
        if link_choice == "Link failure":
            choices_pending.append(f"Link failure: {full_bands}. (Tower Crew)")
        elif link_choice == "SFP Not Present":
            choices_pending.append(f"SFP Not Present: {full_bands}. (Tower Crew)")

        additional_pending = st.text_area("\U0001F4DD Enter any additional pending information",
                                           key="n2e_add_pending", height=80)
        if additional_pending.strip():
            choices_pending.append(additional_pending)

    with st.expander("Notes"):
        choices_notes = ["Final Port Configuration attached."]
        has_5g = any(app.is_populated(row.get("gNBId")) for row in mm_objs)
        if has_5g:
            choices_notes.append("NR configuration has been verified.")
        if new_nodes:
            choices_notes.append(f"{'|'.join(new_nodes)} nodes is in Not monitored state.")
        if controller_id:
            ctrl_mon_choice = st.selectbox(f"{controller_id} monitored state", ["\u2014 Select \u2014", "Monitored", "Not monitored"], key="n2e_ctrl_mon")
            if ctrl_mon_choice == "Monitored":
                choices_notes.append(f"{controller_id} is in monitored state.")
            elif ctrl_mon_choice == "Not monitored":
                choices_notes.append(f"{controller_id} is in not monitored state.")
        sa_note = n2e.sa_conversion_note(sa_nodes)
        if sa_note:
            choices_notes.append(sa_note)
        cpri_choice = st.selectbox("Area prechecks verification for CPRI/SFP check", ["\u2014 Select \u2014", "Completed", "Pending"], key="n2e_cpri") \
            if new_nodes else "\u2014 Select \u2014"
        if cpri_choice == "Completed":
            choices_notes.append("Area prechecks verification for CPRI/SFP check is completed.")
        elif cpri_choice == "Pending":
            choices_notes.append("Area prechecks verification for CPRI/SFP check is pending. (Nodes not replicating in the area tool)")
        notes_manual = st.text_area("\U0001F4DD Enter Notes", key="n2e_notes_manual", height=60)
        if notes_manual.strip():
            choices_notes.append(notes_manual)
        for l in choices_notes:
            st.caption(l)

    # ==================== Report text + xlsm ====================
    report_lines = ["Subject", f"MIC | MNS | N2E | IX-STF | {site_name} | {fa_code} | {site_ids}",
                    "", "IWM Details", iwm_details,
                    "", "Configuration", f"Pre Configuration : Nokia", f"Post Configuration : {post_line}"]
    if current_config: report_lines.append(f"Current Configuration : {current_config}")
    report_lines += ["", "Completed"] + choices_completed
    report_lines += ["", "Pending"] + choices_pending
    report_lines += ["", "Notes:"] + choices_notes
    report_text = "\n".join(str(l) for l in report_lines)

    st.markdown("---")
    if st.button("Generate N2E Report \u2192", type="primary", key="n2e_generate"):
        st.text_area("Report preview", report_text, height=400, key="n2e_preview")
        node_tag = mm_objs[0].get("Node to be built as", "site") if mm_objs else "site"
        st.download_button("Download report (.txt)", report_text, file_name=f"{node_tag}_N2E_Report.txt", key="n2e_dl_txt")

        if not N2E_TEMPLATE_PATH.exists():
            st.warning(f"N2E template not found at {N2E_TEMPLATE_PATH}")
        else:
            row_writes = []
            row_writes.append((N2E_ROW_MAP["subject"], True, [(2, "MIC"), (3, "MNS"), (4, "N2E"), (5, "IX-STF"), (6, site_name), (7, fa_code), (8, site_ids)]))
            row_writes.append((N2E_ROW_MAP["iwm_details"], bool(iwm_details.strip()), [(3, iwm_details)]))
            row_writes.append((N2E_ROW_MAP["pre_configuration"], True, [(3, "Nokia")]))
            row_writes.append((N2E_ROW_MAP["current_configuration"], bool(current_config.strip()), [(3, current_config)] if current_config.strip() else []))
            row_writes.append((N2E_ROW_MAP["post_configuration"], True, [(3, post_line)]))
            row_writes.append((N2E_ROW_MAP["wll_node"], bool(wll_node.strip()), [(3, wll_node)] if wll_node.strip() else []))
            row_writes.append((N2E_ROW_MAP["controller_6610"], bool(controller_id), [(3, controller_id)] if controller_id else []))
            row_writes.append((N2E_ROW_MAP["software_version"], bool(software_version.strip()), [(3, software_version)] if software_version.strip() else []))
            row_writes.append((N2E_ROW_MAP["gs_version"], bool(gs_version.strip()), [(3, gs_version)] if gs_version.strip() else []))

            xlsm_bytes = fill_legacy_mca_surgical(N2E_TEMPLATE_PATH, row_writes)
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_N2E_Filled.xlsm", key="n2e_dl_xlsm")
