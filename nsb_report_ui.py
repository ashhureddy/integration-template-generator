"""
NSB (New Site Build) Report UI. Mirrors n2e_report_ui.py's proven architecture
(checkbox-per-item, no defaults where a real choice is needed, buffer overflow pools,
surgical .xlsm patching) but every item's actual logic is NSB-specific, confirmed
through the NSB design conversation — nothing is assumed to carry over from MCA/N2E
without being explicitly confirmed.

Key confirmed differences from N2E:
- Market and Status are user input (not fixed constants)
- Pre Configuration is always "N/A" (not "Nokia")
- 6610 cascade has 6 items (not 5) — adds Area test
- External alarm Scripting's "some ports locked" case is fully automatic (no manual
  bucket classification at all, unlike MCA/N2E)
- RET configuration is a real user choice (not always Pending like N2E)
- Link failure/SFP Not Present and SFP Installation are manual, not auto-filled
- Florida newly-added-cells reuses MCA's logic directly
"""

import re
import streamlit as st

import nsb_completed_logic as nsb
import mca_completed_logic as mcl
from nsb_row_map import NSB_ROW_MAP
from mca_xlsm_surgical import fill_legacy_mca_surgical
from pathlib import Path

NSB_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "NSB_Macro_Template_v4.xlsm"


def _checked_group(label, lines, key):
    """Same confirmed-safe pattern as MCA/N2E — checkbox recomputes fresh every rerun,
    no stale text_area default issue."""
    if not lines:
        return False, []
    checked = st.checkbox(label, value=True, key=key)
    if checked:
        for l in lines:
            st.caption(l)
    return checked, (lines if checked else [])


def render(app, ciq_wb, mm_objs, controller_objs, edp_index, user_id, date_str,
           postcheck_text="", controller_checks_text=""):
    st.subheader("Generate NSB Report")
    st.markdown("""
        <style>
        div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
            border: 3px solid #5B9BD5 !important; background-color: #CCE5FF !important;
            color: #0D1B2A !important; font-weight: 600 !important;
        }
        div[data-testid="stSelectbox"] > div > div {
            border: 3px solid #5B9BD5 !important; background-color: #CCE5FF !important;
        }
        div[data-testid="stSelectbox"] * {
            color: #0D1B2A !important; font-weight: 600 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    nsb.set_app_module(app)

    def _log(msg):
        pass

    # ---- Reuse app.py's existing scope-of-work computation. NSB doesn't have its own
    # generate_nsb() confirmed yet, so we build the classification directly here, same
    # approach as N2E when it needed CIQ-only detection (every cell = addition). ----
    classification = {"added": {}}
    eutran_objs = app.sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    fiveg_objs = app.sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
    precheck_text = ""  # NSB detection here doesn't depend on Pre-checks for Integration

    controller_id = None
    for row in controller_objs:
        if str(row.get("Controller", "")).strip() == "6610":
            controller_id = row.get("Controller ID")
            break
    controller_in_edp = bool(controller_id)

    new_nodes = [row.get("Node to be built as") for row in mm_objs]
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

    post_parts = []
    for row in mm_objs:
        node, e_name, g_name = row.get("Node to be built as"), row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(node).strip().upper() == str(e_name or "").strip().upper()
        target_row = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) if is_lte_primary else \
            app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        if not target_row:
            target_row = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) or \
                app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        hw = app.hw_string(target_row) if target_row else None
        post_parts.append(f"{node}({hw or 'NOT FOUND'})")
    post_line = " + ".join(post_parts)

    # ==================== Subject / Configuration / IDL Connections ====================
    fa_code = ""
    if "5G Info" in ciq_wb.sheetnames:
        for row in app.sheet_objs(ciq_wb["5G Info"]):
            if app.is_populated(row.get("FA Code")):
                fa_code = row.get("FA Code")
                break
    site_ids = "/".join([row.get("Node to be built as") for row in mm_objs if row.get("Node to be built as")])

    with st.container(border=True):
        st.markdown("**Subject**")
        c = st.columns(7)
        with c[0]: st.markdown(f"MIC\n\n**MIC**")
        with c[1]: market = st.text_input("\U0001F4DD Market", key="nsb_market", placeholder="MNS/TILLMAN/AT&T")
        with c[2]: status = st.text_input("\U0001F4DD Status", key="nsb_status", placeholder="IX-STF/IX-ATP")
        with c[3]: site_name = st.text_input("\U0001F4DD Site Name", key="nsb_site_name")
        with c[4]: st.markdown(f"FA CODE\n\n**{fa_code or '(not found)'}**")
        with c[5]: st.markdown(f"Site ID's\n\n**{site_ids}**")
        with c[6]: st.markdown(f"NSB\n\n**NSB**")

    with st.container(border=True):
        st.markdown("**IWM Details**")
        iwm_details = st.text_input("IWM Details", key="nsb_iwm", label_visibility="collapsed")

    with st.container(border=True):
        st.markdown("**Configuration**")
        st.markdown(f"Pre Configuration : **NSB**")
        st.markdown(f"Post Configuration : **{post_line}**")
        st.markdown(f"6610 Controller : **{controller_id or '(none detected)'}**")
        current_config_auto = mcl.current_configuration_line(ciq_wb, mm_objs, postcheck_text) if postcheck_text else ""
        if current_config_auto:
            current_config = st.text_input("\U0001F4DD Current Configuration \u2014 review/edit:",
                                             value=current_config_auto, key="nsb_current_config")
        else:
            current_config = ""
        c1, c2 = st.columns(2)
        with c1:
            wll_node = st.text_input("\U0001F4DD WLL node", key="nsb_wll")
            gs_version = st.text_input("\U0001F4DD GS Version", key="nsb_gs")
        with c2:
            software_version = st.text_input("\U0001F4DD Software version", key="nsb_sw")

    xmu_present_in_ciq = nsb.xmu_in_ciq(post_line)

    idl_build_type = None
    idle = idly = switch = slot_port = ""
    sidehaul_rows = []
    if len(mm_objs) > 1:
        with st.container(border=True):
            idl_build_type = app.derive_idl_build_type_label(ciq_wb, mm_objs)
            st.markdown(f"**IDL Connections** \u2014 Build Type: **{idl_build_type or '(not detected)'}**")
            c1, c2 = st.columns(2)
            with c1:
                idle = st.text_area("\U0001F4DD IDLe cable details (manual)", key="nsb_idle", height=60)
            with c2:
                idly = st.text_area("\U0001F4DD IDLy cable details (manual)", key="nsb_idly", height=60)
            sidehaul_rows = mcl.sidehaul_display_rows(ciq_wb)
            if sidehaul_rows:
                st.caption("Switch / Slot-Port \u2014 auto-filled from Sidehaul Info:")
                cable_pns = {}
                for i, srow in enumerate(sidehaul_rows):
                    sc1, sc2, sc3, sc4, sc5 = st.columns([1, 1, 1, 1, 1])
                    with sc1: st.caption(f"**{srow['switch_type']}**")
                    with sc2: st.caption(srow["switch_id"])
                    with sc3: st.caption(srow["slot_port"])
                    with sc4: cable_pns[i] = st.text_input("Cable P/N", key=f"nsb_cable_pn_{i}", label_visibility="collapsed")
                    with sc5: st.caption(srow["node_id"])
                switch = "\n".join(mcl.format_sidehaul_lines(sidehaul_rows, cable_pns))
            else:
                switch = st.text_area("\U0001F4DD Switch details (manual)", key="nsb_switch", height=60)
                slot_port = st.text_area("\U0001F4DD Slot/Port/Cable/Node ID (manual)", key="nsb_slotport", height=60)

    has_6673 = any(str(r.get("switch_type", "")).strip() == "6673" for r in sidehaul_rows)

    # ==================== Completed / Pending ====================
    st.markdown("### Which of these apply?")
    choices_completed, choices_pending, choices_notes = [], [], []
    controller_checks_data = mcl.extract_controller_checks(controller_checks_text) if controller_checks_text else {}
    cascade_fires = nsb.controller_cascade_fires(controller_in_edp, bool(controller_checks_text))

    with st.expander("Completed", expanded=True):
        int_pairs = []
        int_lines_display = []
        for node, cells in classification.get("added", {}).items():
            bands = sorted({app.band_label(c)[0] for c in cells if app.band_label(c)[0]})
            if bands:
                int_pairs.append(("/".join(bands), node))
                int_lines_display.append(f"Integration: {'/'.join(bands)} {node}")
        int_checked, int_lines = _checked_group("Integration", int_lines_display, "nsb_int")
        if int_checked:
            choices_completed += int_lines

        if cascade_fires:
            st.markdown(f"<div style='color:#c0392b;'>\u26a0\ufe0f Controller-checks file not uploaded \u2014 "
                        f"{', '.join(nsb.CASCADE_ITEMS)} all move to Pending.</div>", unsafe_allow_html=True)
        ctrl_checked = False
        if controller_id and not cascade_fires:
            ctrl_line = f"6610 Controller Integration: {controller_id}"
            ctrl_checked, _ = _checked_group("6610 Controller Integration", [ctrl_line], "nsb_ctrl")
            if ctrl_checked:
                choices_completed.append(ctrl_line)

        _dss_outputs, _dss_summary, dss_activation_labels = app.generate_dss(ciq_wb, mm_objs, user_id, date_str, _log)
        dss_completed, dss_pending, dss_bands = None, None, None
        if dss_activation_labels:
            dss_bands = "/".join(dss_activation_labels)
            with st.container(border=True):
                st.markdown(f"**DSS Activation** \u2014 detected: {dss_bands}")
                dss_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_dss")
                if dss_choice == "Completed":
                    dss_completed = f"DSS Activation: {dss_bands}"
                    choices_completed.append(dss_completed)
                elif dss_choice == "Pending":
                    dss_sh = st.selectbox("Stakeholder", ["\u2014 Select \u2014", "MIC", "AT&T"], key="nsb_dss_sh")
                    if dss_sh != "\u2014 Select \u2014":
                        dss_pending = f"DSS Activation: {dss_bands} ({dss_sh})"

        _ngs_summary, ngs_scope_lines = app.generate_ngs_checks(ciq_wb, mm_objs, _log)
        ngs_line = next((l for l in ngs_scope_lines if l.startswith("NGS Activation")), None)
        ngs_completed, ngs_pending, ngs_bands, ngs_node = None, None, None, None
        if ngs_line:
            parts = ngs_line.split("\t")
            ngs_bands = parts[1] if len(parts) > 1 else ""
            ngs_node = parts[2] if len(parts) > 2 else ""
            with st.container(border=True):
                st.markdown(f"**NGS activation** \u2014 detected: {ngs_bands} {ngs_node}")
                ngs_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_ngs")
                if ngs_choice == "Completed":
                    ngs_completed = f"NGS activation: {ngs_bands} {ngs_node}"
                    choices_completed.append(ngs_completed)
                elif ngs_choice == "Pending":
                    ngs_pending = f"NGS activation: {ngs_bands} {ngs_node} (MIC)"

        gps_completed_line, gps_pending_line = None, None
        enabled_nodes, disabled_nodes, gtype = [], [], ""
        if postcheck_text:
            post_sync = mcl.extract_sync_status_2(postcheck_text)
            post_gps = mcl.extract_gps_status(postcheck_text)
            for row in mm_objs:
                node = row.get("Node to be built as")
                state = post_sync.get(node)
                if state == "ENABLED":
                    enabled_nodes.append(node)
                elif state == "DISABLED":
                    disabled_nodes.append(node)
            if enabled_nodes:
                gtype = post_gps.get(enabled_nodes[0], "")
                candidate_line = f"GPS Installation: {'|'.join(enabled_nodes)}  Version: {gtype}"
                gps_checked, _ = _checked_group("GPS Installation", [candidate_line], "nsb_gps")
                if gps_checked:
                    gps_completed_line = candidate_line
                    choices_completed.append(gps_completed_line)
            if disabled_nodes:
                nsb_calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
                regional_market = None
                if nsb_calltest_path.exists() and mm_objs:
                    nsb_prefix_to_market, _ = mcl.load_calltest_table(nsb_calltest_path, tab_name="NSB")
                    regional_market = mcl.determine_market(mm_objs[0].get("Node to be built as"), nsb_prefix_to_market)
                gps_pending_line = f"GPS Installation: {'|'.join(disabled_nodes)} ({mcl.gps_pending_stakeholder(regional_market)})"
        else:
            st.caption("GPS Installation: Post-checks not uploaded \u2014 can't determine sync status.")

        lkf_completed, lkf_pending = None, None
        emergency_unlock_notes = []
        with st.container(border=True):
            st.markdown("**LKF Installation** \u2014 Node and Controller tracked independently, required:")
            lkf_node_choices = {}
            for row in mm_objs:
                node = row.get("Node to be built as")
                c1, c2 = st.columns([2, 1])
                with c1: st.caption(node)
                with c2:
                    pick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                          key=f"nsb_lkf_{node}", label_visibility="collapsed")
                    if pick != "\u2014 Select \u2014":
                        lkf_node_choices[node] = pick
                if pick == "Pending":
                    eu = st.selectbox(f"Emergency unlock activated on {node}? LKF is pending.",
                                        ["\u2014 Select \u2014", "No", "Yes"], key=f"nsb_lkf_eu_{node}")
                    if eu == "Yes":
                        emergency_unlock_notes.append(node)
            lkf_controller_choice = None
            if controller_id:
                c1, c2 = st.columns([2, 1])
                with c1: st.caption(f"{controller_id} (controller)")
                with c2:
                    cpick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                           key="nsb_lkf_controller", label_visibility="collapsed")
                    if cpick != "\u2014 Select \u2014":
                        lkf_controller_choice = cpick
            lkf_lines_by_section = mcl.lkf_lines_by_choice(lkf_node_choices, lkf_controller_choice, controller_id) \
                if (lkf_node_choices or lkf_controller_choice) else {}
            if lkf_lines_by_section.get("Completed"):
                lkf_completed = lkf_lines_by_section["Completed"]
                choices_completed.append(lkf_completed)
            if lkf_lines_by_section.get("Pending"):
                lkf_pending = lkf_lines_by_section["Pending"]

        # Call Test — same market-lookup mechanism as MCA, using NSB's own tab.
        calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
        psap_line = speed_lte_line = speed_5g_line = fnet_line = None
        lte_bands_all, fiveg_bands_all = [], []
        market_val = market
        if calltest_path.exists() and market_val:
            prefix_to_market, calltest_rules = mcl.load_calltest_table(calltest_path, tab_name="NSB")
            lte_bands_all = sorted({app.band_label(c)[0] for cells in classification.get("added", {}).values()
                                      for c in cells if app.band_label(c)[0] and not app.band_label(c)[0].startswith("5G_")
                                      and app.band_label(c)[0] not in ("CBAND", "DOD", "DOD_BWE")})
            fiveg_bands_all = sorted({app.band_label(c)[0] for cells in classification.get("added", {}).values()
                                        for c in cells if app.band_label(c)[0] and (app.band_label(c)[0].startswith("5G_")
                                        or app.band_label(c)[0] in ("CBAND", "DOD", "DOD_BWE"))})
            psap_sched_id = st.text_input("\U0001F4DD PSAP Schedule ID (if PSAP applies)", key="nsb_psap_sched")
            if lte_bands_all:
                psap_line = f"PSAP test/Speedtest/VoLTE voice calltest: {'/'.join(lte_bands_all)}"
                speed_lte_line = f"Speedtest/VoLTE voice calltest: {'/'.join(lte_bands_all)}"
            if fiveg_bands_all:
                speed_5g_line = f"Speed test: {'/'.join(fiveg_bands_all)}"
            calltest_checked, _ = _checked_group("Call Test items", [l for l in (psap_line, speed_lte_line, speed_5g_line) if l], "nsb_calltest")
            if calltest_checked:
                if psap_line: choices_completed.append(psap_line)
                if speed_lte_line: choices_completed.append(speed_lte_line)
                if speed_5g_line: choices_completed.append(speed_5g_line)

        sfp_completed_lines = []
        if new_nodes:
            with st.container(border=True):
                st.markdown("**Transport SFP Installation on** \u2014 Enter SFP models")
                for node in new_nodes:
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1: st.caption(node)
                    with c2: bbu = st.text_input("BBU", key=f"nsb_sfp_bbu_{node}", label_visibility="collapsed", placeholder="SFP Model (BBU End)")
                    with c3: siad = st.text_input("SIAD", key=f"nsb_sfp_siad_{node}", label_visibility="collapsed", placeholder="SFP Model (SIAD End)")
                    if bbu.strip() or siad.strip():
                        sfp_completed_lines.append((node, bbu, siad))
            choices_completed += [f"Transport SFP Installation on: {n} {b} (BBU End) & {s} (SIAD End)" for n, b, s in sfp_completed_lines]

        ret_completed, ret_pending = None, None
        ret_choice = st.selectbox("RET configuration", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_ret")
        if ret_choice == "Completed":
            ret_completed = "RET configuration"
            choices_completed.append(ret_completed)
        elif ret_choice == "Pending":
            ret_pending = "RET configuration (Tower Crew)"

        alarm_scripting_completed = None
        alarm_notes_line, alarm_partial_pending = None, None
        sau_completed, sau_pending = None, None
        if controller_checks_data and not cascade_fires:
            if mcl.external_alarm_scripting_confirmed(controller_checks_data):
                alarm_scripting_completed = f"External alarm Scripting on: {controller_id}"
                choices_completed.append(alarm_scripting_completed)
                st.caption(f"\u2705 {alarm_scripting_completed}")
            alarm_notes_line = nsb.external_alarm_scripting_locked_note(controller_checks_data)
            alarm_partial_pending = nsb.external_alarm_scripting_partial_pending(controller_checks_data)
            sau_state = controller_checks_data.get("sau_state")
            if sau_state:
                if sau_state["oper"] == "ENABLED":
                    sau_completed = f"SAU Connections: {controller_id}"
                    choices_completed.append(sau_completed)
                    st.caption(f"\u2705 {sau_completed}")
                else:
                    sau_pending = f"SAU Connections: {controller_id} (MIC PM)"

        sup_completed_lines, sup_pending_lines = [], []
        xmu_completed_lines, xmu_pending_lines = [], []
        if postcheck_text and xmu_present_in_ciq:
            sup_state = nsb.sup_connections_state(postcheck_text, xmu_present_in_ciq)
            for node, state in sup_state.items():
                (sup_completed_lines if state == "ENABLED" else sup_pending_lines).append(
                    f"SUP Connections: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
            xmu_state = nsb.xmu_installation_state(postcheck_text, xmu_present_in_ciq)
            for node, state in xmu_state.items():
                (xmu_completed_lines if state == "ENABLED" else xmu_pending_lines).append(
                    f"XMU Installation: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
        choices_completed += sup_completed_lines + xmu_completed_lines
        for l in sup_completed_lines + xmu_completed_lines:
            st.caption(f"\u2705 {l}")

        idl_completed, idl_pending = None, None
        if len(mm_objs) > 1:
            idl_choice = st.selectbox("IDL connections status", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_idlconn")
            if idl_choice == "Completed":
                idl_completed = "IDL connections"
                choices_completed.append(idl_completed)
            elif idl_choice == "Pending":
                idl_pending = "IDL connections (MIC PM)"

        area_pending = None
        if new_nodes:
            area_pending = f"Area test: {'|'.join(new_nodes)}: Area Lite - Failed (MIC PM)"

        testing_completed, testing_pending = None, None
        if testing_completed is None and controller_checks_data and not cascade_fires:
            testing_section, _ = mcl.external_alarm_testing_placement(controller_checks_data)
            if testing_section == "Completed":
                testing_completed = f"External alarm testing: {controller_id}"
                choices_completed.append(testing_completed)
                st.caption(f"\u2705 {testing_completed}")
            elif testing_section == "Pending":
                testing_pending = f"External alarm testing: {controller_id} (MIC PM)"

        script_6673_completed = None
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            script_6673_completed = f"6673 Script load: {switch_id}"
            choices_completed.append(script_6673_completed)

        # Installation — confirmed purely manual, handled directly in the macro, not
        # duplicated in the UI.

    with st.expander("Pending", expanded=True):
        for line in (dss_pending, ngs_pending, gps_pending_line, lkf_pending):
            if line:
                choices_pending.append(line)
                st.caption(line)
        if ret_pending:
            choices_pending.append(ret_pending)
            st.caption(ret_pending)
        if cascade_fires:
            for item_key, item_text in [
                ("6610 Controller Integration", f"6610 Controller Integration: {controller_id} (MIC PM)"),
                ("External alarm Scripting on", "External alarm Scripting on. (MIC)"),
                ("LKF Installation", "LKF Installation. (MIC)"),
                ("External alarm testing", "External alarm testing. (MIC PM)"),
                ("Area test", area_pending or "Area test. (MIC PM)"),
                ("SAU Connections", "SAU Connections. (MIC PM)"),
            ]:
                choices_pending.append(item_text)
                st.caption(item_text)
        else:
            if alarm_notes_line:
                choices_notes.append(alarm_notes_line)
            if alarm_partial_pending:
                choices_pending.append(alarm_partial_pending)
                st.caption(alarm_partial_pending)
            if sau_pending:
                choices_pending.append(sau_pending)
                st.caption(sau_pending)
            if testing_pending:
                choices_pending.append(testing_pending)
                st.caption(testing_pending)
            if area_pending:
                choices_pending.append(area_pending)
                st.caption(area_pending)

        choices_pending += sup_pending_lines + xmu_pending_lines
        for l in sup_pending_lines + xmu_pending_lines:
            st.caption(l)
        if idl_pending:
            choices_pending.append(idl_pending)
            st.caption(idl_pending)

        # SFP Installation — manual, split by BBU end / Radio end.
        sfp_bbu_manual = st.text_area("\U0001F4DD SFP Installation \u2014 BBU end (manual: SFP Type, Sector Details)", key="nsb_sfp_bbu_manual", height=60)
        if sfp_bbu_manual.strip():
            choices_pending.append(f"SFP Installation: {sfp_bbu_manual} (BBU end) (MIC PM)")
        sfp_radio_manual = st.text_area("\U0001F4DD SFP Installation \u2014 Radio end (manual: SFP Type, Sector Details)", key="nsb_sfp_radio_manual", height=60)
        if sfp_radio_manual.strip():
            choices_pending.append(f"SFP Installation: {sfp_radio_manual} (Radio end) (Tower Crew)")

        rilinks_manual = st.text_input("\U0001F4DD Rilinks Scripting (Node ID, manual)", key="nsb_rilinks")
        if rilinks_manual.strip():
            choices_pending.append(f"Rilinks Scripting: {rilinks_manual} (MIC PM)")
        siad_manual = st.text_input("\U0001F4DD SIAD provisioning (Node ID, manual)", key="nsb_siad")
        if siad_manual.strip():
            choices_pending.append(f"SIAD provisioning: {siad_manual} (AT&T)")

        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            l1 = f"6673 Configuration: {switch_id} (AT&T)"
            l2 = f"6673 Port Configuration in ENM: {switch_id} (AT&T)"
            choices_pending += [l1, l2]
            st.caption(l1); st.caption(l2)

        link_manual = st.text_input("\U0001F4DD Link failure (Radio/Node ID, manual)", key="nsb_link_failure")
        if link_manual.strip():
            choices_pending.append(f"Link failure: {link_manual} (Tower Crew)")
        sfp_np_manual = st.text_input("\U0001F4DD SFP Not Present (Radio/Node ID, manual)", key="nsb_sfp_np")
        if sfp_np_manual.strip():
            choices_pending.append(f"SFP Not Present: {sfp_np_manual} (Tower Crew)")

        # Confirmed manual-only, Pending-only alarm/link items — no auto-detection.
        mo_inc_manual = st.text_input("\U0001F4DD Mo Inconsistent configuration alarm (manual)", key="nsb_mo_inc")
        if mo_inc_manual.strip():
            choices_pending.append(f"Mo Inconsistent configuration alarm: {mo_inc_manual} (Tower Crew)")
        fiberloss1_manual = st.text_input("\U0001F4DD Fiberloss \u2014 Data Link_1 (manual)", key="nsb_fiberloss1")
        if fiberloss1_manual.strip():
            choices_pending.append(f"Fiberloss: {fiberloss1_manual} (Data Link_1) (Tower Crew)")
        fiberloss2_manual = st.text_input("\U0001F4DD Fiberloss \u2014 Data Link_2 (manual)", key="nsb_fiberloss2")
        if fiberloss2_manual.strip():
            choices_pending.append(f"Fiberloss: {fiberloss2_manual} (Data Link_2) (Tower Crew)")
        high_rssi_manual = st.text_input("\U0001F4DD High RSSI (manual)", key="nsb_high_rssi")
        if high_rssi_manual.strip():
            choices_pending.append(f"High RSSI: {high_rssi_manual} (Tower Crew)")
        low_rssi_manual = st.text_input("\U0001F4DD Low RSSI (manual)", key="nsb_low_rssi")
        if low_rssi_manual.strip():
            choices_pending.append(f"Low RSSI: {low_rssi_manual} (Tower Crew)")
        high_vswr_manual = st.text_input("\U0001F4DD High VSWR (manual)", key="nsb_high_vswr")
        if high_vswr_manual.strip():
            choices_pending.append(f"High VSWR: {high_vswr_manual} (Tower Crew)")
        low_vswr_manual = st.text_input("\U0001F4DD Low VSWR (manual)", key="nsb_low_vswr")
        if low_vswr_manual.strip():
            choices_pending.append(f"Low VSWR: {low_vswr_manual} (Tower Crew)")
        vswr_over_manual = st.text_input("\U0001F4DD VSWR overthreshold alarm on (manual)", key="nsb_vswr_over")
        if vswr_over_manual.strip():
            choices_pending.append(f"VSWR overthreshold alarm on: {vswr_over_manual} (Tower Crew)")

        additional_pending = st.text_area("\U0001F4DD Enter any additional pending information",
                                           key="nsb_add_pending", height=80)
        if additional_pending.strip():
            choices_pending.append(additional_pending)

    # Florida newly added cells — same as MCA.
    florida_cells = mcl.florida_newly_added_cells(market, classification)
    florida_rows = mcl.florida_cells_to_rows(florida_cells)
    florida_checked = False
    with st.expander("Florida (newly added CBAND/DOD)"):
        if florida_rows:
            florida_checked = st.checkbox(f"Newly added Cells (Florida market) \u2014 {len(florida_cells)} cell(s)",
                                            value=True, key="nsb_chk_florida")
            if florida_checked:
                for r in florida_rows:
                    st.caption(r)
        else:
            st.caption("Not applicable (not Florida market, or no CBAND/DOD/DOD_BWE additions detected).")
    florida_active_rows = florida_rows if florida_checked else []

    with st.expander("Notes"):
        # SAU enabled on Node/Controller — same as N2E, moved to the top.
        sau_auto_target = controller_id if sau_completed else None
        sau_notes_checked = st.checkbox("SAU enabled on the: Node or Controller", value=True, key="nsb_sau_notes")
        if sau_notes_checked:
            if sau_auto_target:
                st.caption(f"SAU enabled on the: {sau_auto_target}")
                choices_notes.append(f"SAU enabled on the: {sau_auto_target}")
            else:
                sau_manual_target = st.text_input("\U0001F4DD Node ID or Controller ID", key="nsb_sau_manual")
                if sau_manual_target.strip():
                    choices_notes.append(f"SAU enabled on the: {sau_manual_target}")

        has_5g = any(app.is_populated(row.get("gNBId")) for row in mm_objs)
        final_port_checked = st.checkbox("Final Port Configuration attached.", value=True, key="nsb_notes_finalport")
        if final_port_checked:
            choices_notes.append("Final Port Configuration attached.")
        if has_5g:
            nr_checked = st.checkbox("NR configuration has been verified.", value=True, key="nsb_notes_nr")
            if nr_checked:
                choices_notes.append("NR configuration has been verified.")

        # Nodes/controller monitored state — same as N2E.
        if new_nodes:
            mon_checked = st.checkbox(f"{'|'.join(new_nodes)} nodes is in Not monitored state.", value=True, key="nsb_notes_mon")
            if mon_checked:
                choices_notes.append(f"{'|'.join(new_nodes)} nodes is in Not monitored state.")
        if controller_id:
            ctrl_mon_choice = st.selectbox(f"{controller_id} monitored state", ["\u2014 Select \u2014", "Monitored", "Not monitored"], key="nsb_ctrl_mon")
            if ctrl_mon_choice == "Monitored":
                choices_notes.append(f"{controller_id} is in monitored state.")
            elif ctrl_mon_choice == "Not monitored":
                choices_notes.append(f"{controller_id} is in not monitored state.")

        # SA Conversion / TermPointToAmf note — same as N2E, fires only if the CIQ's
        # NR_SA tab is present AND SA Conversion is detected for at least one node.
        nsb_sa_nodes = mcl.sa_conversion_nodes(ciq_wb, mm_objs)
        nsb_sa_note = mcl.sa_conversion_note(nsb_sa_nodes)
        if nsb_sa_note:
            sa_note_checked = st.checkbox(nsb_sa_note, value=True, key="nsb_notes_sa")
            if sa_note_checked:
                choices_notes.append(nsb_sa_note)

        # Area prechecks verification for CPRI/SFP check — same as N2E.
        cpri_choice = st.selectbox("Area prechecks verification for CPRI/SFP check", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_cpri") \
            if new_nodes else "\u2014 Select \u2014"
        if cpri_choice == "Completed":
            choices_notes.append("Area prechecks verification for CPRI/SFP check is completed.")
        elif cpri_choice == "Pending":
            choices_notes.append("Area prechecks verification for CPRI/SFP check is pending. (Nodes not replicating in the area tool)")

        notes_manual = st.text_area("\U0001F4DD Enter Notes", key="nsb_notes_manual", height=60)
        if notes_manual.strip():
            choices_notes.append(notes_manual)

        emergency_unlock_lines = [f"Emergency unlock activated on the node {n}." for n in emergency_unlock_notes]
        if emergency_unlock_lines:
            st.caption("Auto-added (Emergency unlock confirmed):")
            for l in emergency_unlock_lines:
                st.caption(l)
            choices_notes += emergency_unlock_lines

        if alarm_notes_line:
            st.caption(f"Auto-added: {alarm_notes_line}")

    # ==================== Report text + xlsm ====================
    report_lines = ["Subject", f"MIC | {market} | {status} | {site_name} | {fa_code} | {site_ids} | NSB",
                    "", "IWM Details", iwm_details,
                    "", "Configuration", "Pre Configuration : NSB", f"Post Configuration : {post_line}",
                    f"6610 Controller : {controller_id or ''}"]
    if current_config: report_lines.append(f"Current Configuration : {current_config}")
    if wll_node.strip(): report_lines.append(f"WLL node : {wll_node}")
    if software_version.strip(): report_lines.append(f"Software version : {software_version}")
    if gs_version.strip(): report_lines.append(f"GS Version : {gs_version}")
    if len(mm_objs) > 1:
        report_lines += ["", "IDL Connections", f"Build Type : {idl_build_type or ''}"]
        if idle.strip(): report_lines.append(f"IDLe : {idle}")
        if idly.strip(): report_lines.append(f"IDLy : {idly}")
        if switch.strip(): report_lines.append(switch)
        if slot_port.strip(): report_lines.append(slot_port)
    report_lines += ["", "Completed"] + choices_completed
    if florida_active_rows:
        report_lines += ["", "Newly added Cells (Florida)"] + florida_active_rows
    report_lines += ["", "Pending"] + choices_pending
    report_lines += ["", "Notes:"] + choices_notes
    report_text = "\n".join(str(l) for l in report_lines)

    st.markdown("---")
    if st.button("Generate NSB Report \u2192", type="primary", key="nsb_generate"):
        st.text_area("Report preview", report_text, height=400, key="nsb_preview")
        node_tag = mm_objs[0].get("Node to be built as", "site") if mm_objs else "site"
        st.download_button("Download report (.txt)", report_text, file_name=f"{node_tag}_NSB_Report.txt", key="nsb_dl_txt")

        if not NSB_TEMPLATE_PATH.exists():
            st.warning(f"NSB template not found at {NSB_TEMPLATE_PATH}")
        else:
            row_writes = []
            row_writes.append((NSB_ROW_MAP["subject"], True, [(2, "MIC"), (3, market), (4, status), (5, site_name), (6, fa_code), (7, site_ids), (8, "NSB")]))
            row_writes.append((NSB_ROW_MAP["iwm_details"], bool(iwm_details.strip()), [(3, iwm_details)]))
            row_writes.append((NSB_ROW_MAP["pre_configuration"], True, [(3, "NSB")]))
            row_writes.append((NSB_ROW_MAP["current_configuration"], bool(current_config.strip()), [(3, current_config)] if current_config.strip() else []))
            row_writes.append((NSB_ROW_MAP["post_configuration"], True, [(3, post_line)]))
            row_writes.append((NSB_ROW_MAP["wll_node"], bool(wll_node.strip()), [(3, wll_node)] if wll_node.strip() else []))
            row_writes.append((NSB_ROW_MAP["controller_6610"], bool(controller_id), [(3, controller_id)] if controller_id else []))
            row_writes.append((NSB_ROW_MAP["software_version"], bool(software_version.strip()), [(3, software_version)] if software_version.strip() else []))
            row_writes.append((NSB_ROW_MAP["gs_version"], bool(gs_version.strip()), [(3, gs_version)] if gs_version.strip() else []))

            if len(mm_objs) > 1 and idl_build_type:
                idle_rows = NSB_ROW_MAP["idle"]
                row_writes.append((idle_rows[0], True, [(3, idl_build_type)]))
                row_writes.append((idle_rows[1], bool(idle.strip()), [(3, idle)] if idle.strip() else []))
            row_writes.append((NSB_ROW_MAP["idly"], bool(idly.strip()), [(3, idly)] if idly.strip() else []))
            row_writes.append((NSB_ROW_MAP["switch"], bool(switch.strip()), [(2, switch)] if switch.strip() else []))
            row_writes.append((NSB_ROW_MAP["slot_port"], bool(slot_port.strip()), [(2, slot_port)] if slot_port.strip() else []))

            def _rw(row_num, checked, col_value_pairs=None):
                row_writes.append((row_num, bool(checked), col_value_pairs or []))

            # ---- Completed ----
            int_rows = NSB_ROW_MAP["integration"]["completed"]
            for i, row_num in enumerate(int_rows):
                if int_checked and i < len(int_pairs):
                    _rw(row_num, True, [(3, int_pairs[i][0]), (4, int_pairs[i][1])])
                else:
                    _rw(row_num, False)
            _rw(NSB_ROW_MAP["controller_integration"]["completed"][0], ctrl_checked and controller_id, [(3, controller_id)] if ctrl_checked else None)
            _rw(NSB_ROW_MAP["dss_activation"]["completed"][0], dss_completed, [(3, dss_bands)] if dss_completed else None)
            _rw(NSB_ROW_MAP["ngs_activation"]["completed"][0], ngs_completed, [(3, ngs_bands), (4, ngs_node)] if ngs_completed else None)
            _rw(NSB_ROW_MAP["gps_installation"]["completed"][0], gps_completed_line,
                [(3, "|".join(enabled_nodes)), (5, gtype)] if gps_completed_line else None)
            _rw(NSB_ROW_MAP["lkf_installation"]["completed"][0], lkf_completed, [(3, lkf_completed)] if lkf_completed else None)
            _rw(NSB_ROW_MAP["psap_speedtest"]["completed"][0], psap_line, [(3, "/".join(lte_bands_all))] if psap_line else None)
            _rw(NSB_ROW_MAP["speedtest_lte"]["completed"][0], speed_lte_line, [(3, "/".join(lte_bands_all))] if speed_lte_line else None)
            _rw(NSB_ROW_MAP["speed_test_5g"]["completed"][0], speed_5g_line, [(3, "/".join(fiveg_bands_all))] if speed_5g_line else None)
            _rw(NSB_ROW_MAP["calltest_fnet"]["completed"][0], False)
            sfp_rows = NSB_ROW_MAP["transport_sfp"]["completed"]
            for i, row_num in enumerate(sfp_rows):
                if i < len(sfp_completed_lines):
                    node, bbu, siad = sfp_completed_lines[i]
                    _rw(row_num, True, [(3, node), (4, f"{bbu} (BBU End) & {siad} (SIAD End)")])
                else:
                    _rw(row_num, False)
            _rw(NSB_ROW_MAP["ret_configuration"]["completed"][0], ret_completed)
            _rw(NSB_ROW_MAP["external_alarm_scripting"]["completed"][0], not cascade_fires and alarm_scripting_completed, [(3, controller_id)] if alarm_scripting_completed else None)
            _rw(NSB_ROW_MAP["sau_connections"]["completed"][0], not cascade_fires and sau_completed, [(3, controller_id)] if sau_completed else None)
            _rw(NSB_ROW_MAP["sup_connections"]["completed"][0], sup_completed_lines, [(3, "|".join(s.split(": ")[-1] for s in sup_completed_lines))] if sup_completed_lines else None)
            _rw(NSB_ROW_MAP["xmu_installation"]["completed"][0], xmu_completed_lines, [(3, "|".join(s.split(": ")[-1] for s in xmu_completed_lines))] if xmu_completed_lines else None)
            _rw(NSB_ROW_MAP["idl_connections"]["completed"][0], idl_completed)
            _rw(NSB_ROW_MAP["area_test"]["completed"][0], False)
            _rw(NSB_ROW_MAP["external_alarm_testing"]["completed"][0], not cascade_fires and testing_completed, [(3, controller_id)] if testing_completed else None)
            _rw(NSB_ROW_MAP["script_load_6673"]["completed"][0], script_6673_completed,
                [(3, sidehaul_rows[0]["switch_id"])] if script_6673_completed and sidehaul_rows else None)
            _rw(NSB_ROW_MAP["installation_manual"]["completed"][0], False)

            # ---- Pending ----
            _rw(NSB_ROW_MAP["post_configuration_pending"]["pending"][0], False)
            p_int_rows = NSB_ROW_MAP["integration"]["pending"]
            for row_num in p_int_rows:
                _rw(row_num, False)
            _rw(NSB_ROW_MAP["controller_integration"]["pending"][0], cascade_fires, [(3, controller_id)] if cascade_fires else None)
            _rw(NSB_ROW_MAP["dss_activation"]["pending"][0], dss_pending, [(3, dss_bands)] if dss_pending else None)
            _rw(NSB_ROW_MAP["ngs_activation"]["pending"][0], ngs_pending, [(3, ngs_bands), (4, ngs_node)] if ngs_pending else None)
            _rw(NSB_ROW_MAP["gps_installation"]["pending"][0], gps_pending_line, [(3, "|".join(disabled_nodes))] if gps_pending_line else None)
            _rw(NSB_ROW_MAP["lkf_installation"]["pending"][0], cascade_fires or lkf_pending, None)
            _rw(NSB_ROW_MAP["psap_speedtest"]["pending"][0], False)
            _rw(NSB_ROW_MAP["speedtest_lte"]["pending"][0], False)
            _rw(NSB_ROW_MAP["speed_test_5g"]["pending"][0], False)
            _rw(NSB_ROW_MAP["calltest_fnet"]["pending"][0], False)
            _rw(NSB_ROW_MAP["sfp_installation_bbu"]["pending"][0], sfp_bbu_manual.strip(), [(3, sfp_bbu_manual)] if sfp_bbu_manual.strip() else None)
            _rw(NSB_ROW_MAP["sfp_installation_radio"]["pending"][0], sfp_radio_manual.strip(), [(3, sfp_radio_manual)] if sfp_radio_manual.strip() else None)
            for row_num in NSB_ROW_MAP["transport_sfp"]["pending"]:
                _rw(row_num, False)
            _rw(NSB_ROW_MAP["ret_configuration"]["pending"][0], ret_pending)
            _rw(NSB_ROW_MAP["external_alarm_scripting"]["pending"][0], cascade_fires or alarm_partial_pending,
                [(3, controller_id)] if cascade_fires else None)
            _rw(NSB_ROW_MAP["sau_connections"]["pending"][0], cascade_fires or sau_pending, [(3, controller_id)] if (cascade_fires or sau_pending) else None)
            _rw(NSB_ROW_MAP["sup_connections"]["pending"][0], sup_pending_lines, [(3, "|".join(s.split(":")[-1].strip() for s in sup_pending_lines))] if sup_pending_lines else None)
            _rw(NSB_ROW_MAP["xmu_installation"]["pending"][0], xmu_pending_lines, [(3, "|".join(s.split(":")[-1].strip() for s in xmu_pending_lines))] if xmu_pending_lines else None)
            _rw(NSB_ROW_MAP["rilinks_scripting"]["pending"][0], rilinks_manual.strip(), [(3, rilinks_manual)] if rilinks_manual.strip() else None)
            _rw(NSB_ROW_MAP["idl_connections"]["pending"][0], idl_pending)
            _rw(NSB_ROW_MAP["script_load_6673"]["pending"][0], False)
            _rw(NSB_ROW_MAP["siad_provisioning"]["pending"][0], siad_manual.strip(), [(3, siad_manual)] if siad_manual.strip() else None)
            _rw(NSB_ROW_MAP["area_test"]["pending"][0], cascade_fires or area_pending,
                [(3, "|".join(new_nodes)), (4, "Failed")] if area_pending else None)
            _rw(NSB_ROW_MAP["external_alarm_testing"]["pending"][0], cascade_fires or testing_pending, [(3, controller_id)] if (cascade_fires or testing_pending) else None)
            _rw(NSB_ROW_MAP["config_6673"]["pending"][0], has_6673, [(3, sidehaul_rows[0]["switch_id"])] if has_6673 and sidehaul_rows else None)
            _rw(NSB_ROW_MAP["port_config_6673_enm"]["pending"][0], has_6673, [(4, sidehaul_rows[0]["switch_id"])] if has_6673 and sidehaul_rows else None)
            _rw(NSB_ROW_MAP["link_failure"]["pending"][0], link_manual.strip(), [(3, link_manual)] if link_manual.strip() else None)
            _rw(NSB_ROW_MAP["sfp_not_present"]["pending"][0], sfp_np_manual.strip(), [(3, sfp_np_manual)] if sfp_np_manual.strip() else None)
            _rw(NSB_ROW_MAP["mo_inconsistent_config_alarm"]["pending"][0], mo_inc_manual.strip(), [(3, mo_inc_manual)] if mo_inc_manual.strip() else None)
            fiberloss_rows = NSB_ROW_MAP["fiberloss"]["pending"]
            _rw(fiberloss_rows[0], fiberloss1_manual.strip(), [(3, fiberloss1_manual)] if fiberloss1_manual.strip() else None)
            _rw(fiberloss_rows[1], fiberloss2_manual.strip(), [(3, fiberloss2_manual)] if fiberloss2_manual.strip() else None)
            _rw(NSB_ROW_MAP["high_rssi"]["pending"][0], high_rssi_manual.strip(), [(3, high_rssi_manual)] if high_rssi_manual.strip() else None)
            _rw(NSB_ROW_MAP["low_rssi"]["pending"][0], low_rssi_manual.strip(), [(3, low_rssi_manual)] if low_rssi_manual.strip() else None)
            _rw(NSB_ROW_MAP["high_vswr"]["pending"][0], high_vswr_manual.strip(), [(3, high_vswr_manual)] if high_vswr_manual.strip() else None)
            _rw(NSB_ROW_MAP["low_vswr"]["pending"][0], low_vswr_manual.strip(), [(3, low_vswr_manual)] if low_vswr_manual.strip() else None)
            _rw(NSB_ROW_MAP["vswr_overthreshold"]["pending"][0], vswr_over_manual.strip(), [(3, vswr_over_manual)] if vswr_over_manual.strip() else None)

            pending_buffer_lines = [additional_pending] if additional_pending.strip() else []
            for i, row_num in enumerate(NSB_ROW_MAP["additional_pending"]):
                if i < len(pending_buffer_lines):
                    row_writes.append((row_num, True, [(2, pending_buffer_lines[i])]))
                else:
                    row_writes.append((row_num, False, []))

            florida_xlsm_rows = NSB_ROW_MAP["florida_cells"]
            row_writes.append((NSB_ROW_MAP["florida_header"], bool(florida_checked), []))
            for i, row_num in enumerate(florida_xlsm_rows):
                if i < len(florida_active_rows):
                    row_writes.append((row_num, True, [(2, florida_active_rows[i])]))
                else:
                    row_writes.append((row_num, False, []))

            final_port_line = "Final Port Configuration attached." if "Final Port Configuration attached." in choices_notes else None
            row_writes.append((NSB_ROW_MAP["notes_final_port_config"], bool(final_port_line), [(2, final_port_line)] if final_port_line else []))
            nr_line = "NR configuration has been verified." if "NR configuration has been verified." in choices_notes else None
            row_writes.append((NSB_ROW_MAP["notes_nr_verified"], bool(nr_line), [(2, nr_line)] if nr_line else []))
            other_notes = [n for n in choices_notes if n not in ("Final Port Configuration attached.", "NR configuration has been verified.")]
            if alarm_notes_line:
                other_notes = other_notes + [alarm_notes_line]
            for i, row_num in enumerate(NSB_ROW_MAP["notes_buffer"]):
                if i < len(other_notes):
                    row_writes.append((row_num, True, [(2, other_notes[i])]))
                else:
                    row_writes.append((row_num, False, []))

            xlsm_bytes = fill_legacy_mca_surgical(NSB_TEMPLATE_PATH, row_writes)
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_NSB_Filled.xlsm", key="nsb_dl_xlsm")
