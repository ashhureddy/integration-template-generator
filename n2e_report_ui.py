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

N2E_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "N2E_Pre_IX_Macro_V2_update (1).xlsm"


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
            border: 3px solid #002D62 !important; background-color: #007FFF !important;
            color: #FFFFFF !important; font-weight: 600 !important;
        }
        div[data-testid="stSelectbox"] > div > div {
            border: 3px solid #002D62 !important; background-color: #007FFF !important;
        }
        div[data-testid="stSelectbox"] * {
            color: #FFFFFF !important; font-weight: 600 !important;
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

    all_node_names = [row.get("Node to be built as") for row in mm_objs]
    # Confirmed rule: a node whose alphabetic PREFIX ends in "L" (e.g. ALL00640 -> "ALL",
    # WAL94133 -> "WAL", ECL02586 -> "ECL") is a WLL node, not a genuinely new build — it
    # shouldn't trigger Transport SFP/Area test/etc. as if new, and its name auto-fills
    # the WLL node field instead of being manually typed. Confirmed fix: checking the
    # PREFIX specifically, not whether the whole node name ends in "L" (which would
    # never match any real example — ALL00640 ends in "0", not "L").
    def _wll_prefix_check(n):
        m = re.match(r'^([A-Za-z]+)', str(n or ""))
        return bool(m and m.group(1).upper().endswith("L"))
    wll_detected_nodes = [n for n in all_node_names if n and _wll_prefix_check(n)]
    new_nodes = [n for n in all_node_names if n not in wll_detected_nodes]  # confirmed: every OTHER node is "new" for N2E

    # Confirmed correction: Post Configuration should NOT include the WLL node's hardware
    # info at all — generate_n2e()'s own post_line includes every Mixed Mode Info row
    # (including WLL nodes), so rebuild it here filtered to new_nodes only.
    post_parts = []
    for row in mm_objs:
        node = row.get("Node to be built as")
        if node in wll_detected_nodes:
            continue
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(node).strip().upper() == str(e_name or "").strip().upper()
        target_row = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) if is_lte_primary else \
            app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        if not target_row:
            target_row = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) or \
                app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        hw = app.hw_string(target_row) if target_row else None
        post_parts.append(f"{node}({hw or 'NOT FOUND'})")
    if post_parts:
        post_line = " + ".join(post_parts)

    # ==================== Subject / Configuration / IDL Connections ====================
    fa_code = ""
    if "5G Info" in ciq_wb.sheetnames:
        for row in app.sheet_objs(ciq_wb["5G Info"]):
            if app.is_populated(row.get("FA Code")):
                fa_code = row.get("FA Code")
                break
    site_ids = "/".join(new_nodes + ([controller_id] if controller_id else []))

    with st.container(border=True):
        st.markdown("**Subject**")
        c = st.columns(7)
        with c[0]: st.markdown(f"MIC\n\n**MIC**")
        with c[1]: st.markdown(f"Market\n\n**MNS**")
        with c[2]: st.markdown(f"N2E\n\n**N2E**")
        with c[3]: st.markdown(f"Status\n\n**IX-STF**")
        with c[4]: site_name = st.text_input("\U0001F4DD Site Name", key="n2e_site_name")
        with c[5]: st.markdown(f"FA CODE\n\n**{fa_code or '(not found)'}**")
        with c[6]: st.markdown(f"Site ID's\n\n**{site_ids}**")

    with st.container(border=True):
        st.markdown("**IWM Details**")
        iwm_details = st.text_input("IWM Details", key="n2e_iwm", label_visibility="collapsed")

    with st.container(border=True):
        st.markdown("**Configuration**")
        st.markdown(f"Pre Configuration : **Nokia**")
        st.markdown(f"Post Configuration : **{post_line}**")
        st.markdown(f"6610 Controller : **{controller_id or '(none detected)'}**")
        mm_objs_no_wll = [row for row in mm_objs if row.get("Node to be built as") not in wll_detected_nodes]
        current_config_auto = mcl.current_configuration_line(ciq_wb, mm_objs_no_wll, postcheck_text) if postcheck_text else ""
        if current_config_auto:
            current_config = st.text_input("\U0001F4DD Current Configuration \u2014 review/edit:",
                                             value=current_config_auto, key="n2e_current_config")
        else:
            current_config = ""
        c1, c2 = st.columns(2)
        with c1:
            wll_node = st.text_input("\U0001F4DD WLL node", value="|".join(wll_detected_nodes), key="n2e_wll")
            gs_version = st.text_input("\U0001F4DD GS Version", key="n2e_gs")
        with c2:
            software_version = st.text_input("\U0001F4DD Software version", key="n2e_sw")

    xmu_present_in_ciq = n2e.xmu_in_ciq(post_line)

    idl_build_type = None
    idle = idly = switch = slot_port = ""
    if len(new_nodes) > 1:
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
                # Confirmed: if Switch isn't present in Sidehaul Info at all, don't show
                # any Switch UI — no manual fallback needed for N2E.
                switch = ""
    else:
        sidehaul_rows = []

    has_6673 = n2e.has_6673(sidehaul_rows)

    # ==================== Completed / Pending items ====================
    st.markdown("### Which of these apply?")
    choices_completed, choices_pending, stakeholders = [], [], {}
    warnings = []

    with st.expander("Completed", expanded=True):
        # Integration — from generate_n2e's own scope_lines. Confirmed bug: was using
        # next() which only grabbed the FIRST matching line, silently dropping every
        # additional node's Integration line when 2+ nodes are present.
        integration_lines_all = [l.replace("\t", " ") for l in scope_lines if l.startswith("Integration")]
        int_checked, int_lines = _checked_group("Integration", integration_lines_all, "n2e_int")
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

        # GPS Installation — Post-checks Sync Status 2 TimeSyncIO. Now uses the same
        # visible checkbox pattern as Integration, instead of silently appending.
        gps_completed_line, gps_pending_line = None, None
        if postcheck_text:
            post_sync = mcl.extract_sync_status_2(postcheck_text)
            post_gps = mcl.extract_gps_status(postcheck_text)
            enabled_nodes, disabled_nodes = n2e.gps_sync_status(mm_objs, post_sync)
            if enabled_nodes:
                gtype = post_gps.get(enabled_nodes[0], "")
                candidate_line = n2e.gps_installation_line(enabled_nodes, gtype)
                gps_checked, gps_lines = _checked_group("GPS Installation", [candidate_line], "n2e_gps")
                if gps_checked:
                    gps_completed_line = candidate_line
                    choices_completed.append(gps_completed_line)
            if disabled_nodes:
                gps_pending_line = f"GPS Installation: {'|'.join(disabled_nodes)} (MIC PM)"
            if not enabled_nodes and not disabled_nodes:
                st.caption(f"GPS Installation: no TimeSyncIO state found in Post-checks for "
                           f"{'|'.join(row.get('Node to be built as') for row in mm_objs)} \u2014 check Post-checks parsing.")
        else:
            st.caption("GPS Installation: Post-checks not uploaded \u2014 can't determine sync status.")

        # LKF Installation — Node(s) and Controller are independent installation points
        # (confirmed real gap: was only showing a per-node dropdown, no separate
        # Controller option at all). Reuses the same shared mcl.lkf_lines_by_choice()
        # logic already tested for MCA — no default on either.
        lkf_completed, lkf_pending = None, None
        with st.container(border=True):
            st.markdown("**LKF Installation** \u2014 Node and Controller are tracked "
                        "independently, required:")
            lkf_node_choices = {}
            for node in new_nodes:
                c1, c2 = st.columns([2, 1])
                with c1: st.caption(node)
                with c2:
                    pick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                          key=f"n2e_lkf_{node}", label_visibility="collapsed")
                    if pick != "\u2014 Select \u2014":
                        lkf_node_choices[node] = pick
            lkf_controller_choice = None
            if controller_id:
                c1, c2 = st.columns([2, 1])
                with c1: st.caption(f"{controller_id} (controller)")
                with c2:
                    cpick = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"],
                                           key="n2e_lkf_controller", label_visibility="collapsed")
                    if cpick != "\u2014 Select \u2014":
                        lkf_controller_choice = cpick

            lkf_lines_by_section = mcl.lkf_lines_by_choice(lkf_node_choices, lkf_controller_choice, controller_id) \
                if (lkf_node_choices or lkf_controller_choice) else {}
            if lkf_lines_by_section.get("Completed"):
                lkf_completed = lkf_lines_by_section["Completed"]
                choices_completed.append(lkf_completed)
            if lkf_lines_by_section.get("Pending"):
                lkf_pending = lkf_lines_by_section["Pending"]

        # Transport SFP — new-node trigger, manual SFP model per node. Confirmed
        # correction: blank fields simply produce nothing (no line at all) — no longer
        # auto-converted into a Pending "Compatible Transport SFP Installation" entry.
        sfp_completed_lines = []
        if new_nodes:
            with st.container(border=True):
                st.markdown("**Transport SFP Installation on** \u2014 Enter SFP models")
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
                st.caption(f"\u2705 {alarm_scripting_completed}")
            sau_state = controller_checks_data.get("sau_state")
            if sau_state:
                if sau_state["oper"] == "ENABLED":
                    sau_completed = f"SAU Connection: {controller_id}."
                    choices_completed.append(sau_completed)
                    st.caption(f"\u2705 {sau_completed}")
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
        for l in sup_completed_lines + xmu_completed_lines:
            st.caption(f"\u2705 {l}")

        # IDL connections — 2+ nodes trigger, manual Completed/Pending.
        idl_completed, idl_pending = None, None
        if len(new_nodes) > 1:
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

        # Area test — new-node trigger. Confirmed: Area Lite result is always "Failed",
        # so this always lands in Pending, never Completed — no choice needed.
        area_completed, area_pending = None, None
        if new_nodes:
            area_pending = n2e.area_test_line(new_nodes, "Failed") + " (MIC PM)"

        # External alarm testing — Completed if any scripted port unlocked, else Pending+Notes.
        testing_section, testing_note = mcl.external_alarm_testing_placement(controller_checks_data) \
            if controller_checks_data and not cascade_fires else (None, None)
        testing_completed, testing_pending = None, None
        if testing_section == "Completed":
            testing_completed = f"External alarm testing: {controller_id}."
            choices_completed.append(testing_completed)
            st.caption(f"\u2705 {testing_completed}")
        elif testing_section == "Pending":
            testing_pending = f"External alarm testing: {controller_id}. (MIC PM)"

        # 6673 Script load — auto, Completed only.
        script_6673_completed = None
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            script_6673_completed = f"6673 Script load: {switch_id}"
            choices_completed.append(script_6673_completed)

        # Installation and "additional completed information" — confirmed: these are
        # purely manual with no auto-detection value at all; the engineer fills them in
        # directly in the downloaded macro file itself rather than duplicating entry here.

        # SA Conversion — CIQ NR_SA tab presence, Completed only. Confirmed same
        # silent-append bug as the others — now uses the visible checkbox pattern.
        sa_nodes = n2e.sa_conversion_nodes(ciq_wb, mm_objs)
        sa_completed_line = None
        if sa_nodes:
            candidate_sa_line = f"SA conversion.: {'|'.join(sa_nodes)}"
            sa_checked, sa_lines = _checked_group("SA Conversion", [candidate_sa_line], "n2e_sa")
            if sa_checked:
                sa_completed_line = candidate_sa_line
                choices_completed.append(sa_completed_line)
        else:
            st.caption("SA Conversion: not detected (no node found in CIQ's NR_SA tab).")

    with st.expander("Pending", expanded=True):
        # Row order below matches the real template exactly: 78 On Site Nokia cutover,
        # 79 6610, 80 DSS, 81 NGS, 82 GPS, 83 LKF, 84 PSAP/Speedtest, 85 Speed test,
        # 86 F-NET, 87-89 Compatible SFP, 90 RET, 91 Scripting, 92 SAU, 93 SUP, 94 XMU,
        # 95 IDL, 97 SIAD, 98 Area test, 99 testing, 100 6673 Config, 101 6673 ENM,
        # 102-103 Link failure/SFP, 104 SA Conversion.
        choices_pending.append("On Site Nokia cutover (Tower Crew)")
        st.caption("On Site Nokia cutover (Tower Crew)")

        if cascade_fires and controller_line:
            l = f"{controller_line.replace(chr(9), ' ')} (MIC PM)"
            choices_pending.append(l)
            st.caption(l)

        for line in (dss_pending, ngs_pending, gps_pending_line):
            if line:
                choices_pending.append(line)
                st.caption(line)

        if cascade_fires:
            choices_pending.append("LKF Installation. (MIC)")
            st.caption("LKF Installation. (MIC)")
        elif lkf_pending:
            choices_pending.append(lkf_pending)
            st.caption(lkf_pending)

        # PSAP/Speedtest + Speed test + F-NET — Integration bands split LTE/5G, all
        # markets, no lookup table for N2E.
        lte_bands, fiveg_bands = n2e.integration_bands_by_tech(classification)
        if lte_bands:
            psap_line = f"PSAP/Speed test/VoLTE voice call test on: {lte_bands}. (MIC PM)"
            choices_pending.append(psap_line)
            st.caption(psap_line)
        if fiveg_bands:
            speed_line = f"Speed test on: {fiveg_bands}. (MIC PM)"
            choices_pending.append(speed_line)
            st.caption(speed_line)

        # Compatible Transport SFP Installation — confirmed removed, not needed: blank
        # Transport SFP fields simply print empty now, no Pending auto-conversion.

        choices_pending.append(ret_pending)
        st.caption(ret_pending)

        if cascade_fires:
            choices_pending.append("External alarm Scripting. (MIC)")
            choices_pending.append("SAU Connections. (MIC PM)")
            st.caption("External alarm Scripting. (MIC)")
            st.caption("SAU Connections. (MIC PM)")
        else:
            if sau_pending:
                choices_pending.append(sau_pending)
                st.caption(sau_pending)

        for l in sup_pending_lines + xmu_pending_lines:
            choices_pending.append(l)
            st.caption(l)

        if idl_pending:
            choices_pending.append(idl_pending)
            st.caption(idl_pending)

        # SIAD provisioning — confirmed purely manual, handled directly in the macro,
        # not duplicated here.

        if area_pending:
            choices_pending.append(area_pending)
            st.caption(area_pending)

        if cascade_fires:
            choices_pending.append("External alarm testing. (MIC PM)")
            st.caption("External alarm testing. (MIC PM)")
        elif testing_pending:
            choices_pending.append(testing_pending)
            st.caption(testing_pending)
        if testing_note:
            choices_pending.append(testing_note)
            st.caption(testing_note)

        # 6673 Configuration / 6673 Port Configuration in ENM — auto, always Pending.
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            l1 = f"6673 Configuration: {switch_id}. (AT&T)"
            l2 = f"6673 Port Configuration in ENM: {switch_id}. (AT&T)"
            choices_pending += [l1, l2]
            st.caption(l1); st.caption(l2)

        # Link failure / SFP Not Present — single selector, reuses Integration's full band list.
        full_bands, _ = n2e.integration_bands_and_nodes(classification)
        link_choice = st.selectbox("Link failure / SFP Not Present", ["\u2014 Select \u2014 (neither)", "Link failure", "SFP Not Present"], key="n2e_link_sfp")
        if link_choice == "Link failure":
            choices_pending.append(f"Link failure: {full_bands}. (Tower Crew)")
        elif link_choice == "SFP Not Present":
            choices_pending.append(f"SFP Not Present: {full_bands}. (Tower Crew)")

        if smm_pending:
            choices_pending.append(smm_pending)
            st.caption(smm_pending)

        # Active external alarm — confirmed removed entirely: redundant with the
        # "Locked alarm ports" section right below, which already shows the same ports
        # (with slogan/severity) as part of its own detected-ports display.

        # Locked alarm ports — confirmed N2E-specific rules from the reference doc's
        # dedicated N2E tab, genuinely different from MCA's Legacy tab: no equivalent to
        # MCA's "pre-existing locked" bucket, different Owner tags, and a new "Power
        # Plant Swap" scenario producing both a Pending AND a Note line together.
        locked_ports_list = [p for p in (controller_checks_data.get("alarm_ports", []) if controller_checks_data else [])
                              if p["admin"] == "LOCKED" and p["slogan"]]
        bucket_pending, bucket_notes = [], []
        ignore_state_notes = []
        if locked_ports_list:
            with st.container(border=True):
                st.markdown(f"**Locked alarm ports** \u2014 {len(locked_ports_list)} scripted port(s) detected LOCKED "
                            f"in the controller-checks file:")
                for p in locked_ports_list:
                    st.caption(f"Port {p['port']} \u2014 {p['slogan']} ({p['severity']})")
                st.markdown("Classify each one below (per the confirmed N2E 6610 Alarm Cutover reporting "
                            "standard). Leave blank whichever don't apply.")
                # Confirmed fix: build from ALL alarm ports (not just currently-locked
                # ones), so a port that's since been unlocked still gets its slogan.
                n2e_all_alarm_ports = controller_checks_data.get("alarm_ports", []) if controller_checks_data else []
                n2e_port_slogan_map = {p["port"]: p["slogan"] for p in n2e_all_alarm_ports if p["slogan"]}

                def _n2e_warn_missing_slogans(ports_str):
                    missing = mcl.missing_slogan_ports(ports_str, n2e_port_slogan_map)
                    if missing:
                        st.markdown(f"<div style='color:#c0392b;'>\u26a0\ufe0f Port(s) {', '.join(missing)} have "
                                    f"no detected slogan \u2014 check the number(s), a slogan is required for every "
                                    f"reported port.</div>", unsafe_allow_html=True)

                nb1 = st.text_input("\U0001F4DD 1. Pre-existing Active Alarms \u2014 port numbers", key="n2e_lp_active", placeholder="e.g. 3, 20")
                _n2e_warn_missing_slogans(nb1)
                t1 = n2e.n2e_locked_port_active_alarms(nb1, n2e_port_slogan_map)
                if t1:
                    bucket_notes.append(t1)

                st.markdown("\U0001F4DD 2. Pre-Existing Loops and Bridge Clips \u2014 port numbers per category:")
                bc1, bc2, bc3 = st.columns(3)
                with bc1: nb2_loops = st.text_input("Pre existing loops", key="n2e_lp_loops", placeholder="e.g. 1")
                with bc2: nb2_clips = st.text_input("Bridge clips", key="n2e_lp_clips", placeholder="e.g. 2")
                with bc3: nb2_noequip = st.text_input("No equipment end connections", key="n2e_lp_noequip", placeholder="e.g. 3")
                for f in (nb2_loops, nb2_clips, nb2_noequip):
                    _n2e_warn_missing_slogans(f)
                b2_notes, b2_active_line = mcl.loops_bridge_clips_notes(nb2_loops, nb2_clips, nb2_noequip,
                                                                          n2e_all_alarm_ports, n2e_port_slogan_map)
                bucket_notes += b2_notes
                if b2_active_line:
                    bucket_notes.append(b2_active_line)

                power_plant_swap_performed = st.checkbox("Power Plant Swap performed on this site", key="n2e_lp_power_performed")
                nb3 = st.text_input("\U0001F4DD 3. Power Plant Swap \u2014 port numbers (only if any alarms are currently active; leave blank otherwise)", key="n2e_lp_power")
                _n2e_warn_missing_slogans(nb3)
                nb4 = st.text_input("\U0001F4DD 4. Post-Cutover Alarms Not Cleared by FE \u2014 port numbers", key="n2e_lp_notcleared")
                _n2e_warn_missing_slogans(nb4)

                ignore_state_entries = st.text_input(
                    "\U0001F4DD Any Pre-existing External Alarms Found in 'Ignore' State? "
                    "Enter port(slogan), comma-separated for multiple", key="n2e_ignore_state",
                    placeholder="e.g. 3(RBS Temp High)")
                ignore_state_notes = n2e.ignore_state_alarm_notes(ignore_state_entries)
                for n in ignore_state_notes:
                    st.caption(n)

                # Confirmed fix: the Note ("scripted according to new power plant
                # configuration") is unconditional whenever ports were entered, but the
                # Pending "kept locked" line only includes genuinely active ports —
                # filtered here BEFORE merging with bucket 4 (which is inherently active
                # by definition, since it's "Not Cleared by FE").
                power_plant_pending, power_plant_note, power_plant_active_ports = \
                    n2e.n2e_locked_port_power_plant_swap(power_plant_swap_performed, nb3, n2e_port_slogan_map, n2e_all_alarm_ports)
                if power_plant_note:
                    bucket_notes.append(power_plant_note)

                merged_pending = n2e.n2e_merged_post_cutover_pending(power_plant_active_ports, nb4, n2e_port_slogan_map)
                if merged_pending:
                    bucket_pending.append(merged_pending)

                choices_pending += bucket_pending
                for l in bucket_pending:
                    st.caption(l)

        additional_pending = st.text_area("\U0001F4DD Enter any additional pending information",
                                           key="n2e_add_pending", height=80)
        if additional_pending.strip():
            choices_pending.append(additional_pending)

    with st.expander("Notes"):
        choices_notes = []
        choices_notes += ignore_state_notes
        final_port_checked = st.checkbox("Final Port Configuration attached.", value=True, key="n2e_notes_finalport")
        if final_port_checked:
            choices_notes.append("Final Port Configuration attached.")

        has_5g = any(app.is_populated(row.get("gNBId")) for row in mm_objs)
        if has_5g:
            nr_checked = st.checkbox("NR configuration has been verified.", value=True, key="n2e_notes_nr")
            if nr_checked:
                choices_notes.append("NR configuration has been verified.")

        if new_nodes:
            mon_checked = st.checkbox(f"{'|'.join(new_nodes)} nodes is in Not monitored state.", value=True, key="n2e_notes_mon")
            if mon_checked:
                choices_notes.append(f"{'|'.join(new_nodes)} nodes is in Not monitored state.")

        if controller_id:
            ctrl_mon_choice = st.selectbox(f"{controller_id} monitored state", ["\u2014 Select \u2014", "Monitored", "Not monitored"], key="n2e_ctrl_mon")
            if ctrl_mon_choice == "Monitored":
                choices_notes.append(f"{controller_id} is in monitored state.")
            elif ctrl_mon_choice == "Not monitored":
                choices_notes.append(f"{controller_id} is in not monitored state.")

        sa_note = n2e.sa_conversion_note(sa_nodes)
        if sa_note:
            sa_note_checked = st.checkbox(sa_note, value=True, key="n2e_notes_sa")
            if sa_note_checked:
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

        if bucket_notes:
            st.caption("Auto-added (Locked-port classification):")
            for l in bucket_notes:
                st.caption(l)
            choices_notes += bucket_notes

    # ==================== Report text + xlsm ====================
    report_lines = ["Subject", f"MIC | MNS | N2E | IX-STF | {site_name} | {fa_code} | {site_ids}",
                    "", "IWM Details", iwm_details,
                    "", "Configuration", f"Pre Configuration : Nokia", f"Post Configuration : {post_line}",
                    f"6610 Controller : {controller_id or ''}"]
    if current_config: report_lines.append(f"Current Configuration : {current_config}")
    if wll_node.strip(): report_lines.append(f"WLL node : {wll_node}")
    if software_version.strip(): report_lines.append(f"Software version : {software_version}")
    if gs_version.strip(): report_lines.append(f"GS Version : {gs_version}")
    if len(new_nodes) > 1:
        report_lines += ["", "IDL Connections", f"Build Type : {idl_build_type or ''}"]
        if idle.strip(): report_lines.append(f"IDLe : {idle}")
        if idly.strip(): report_lines.append(f"IDLy : {idly}")
        if switch.strip(): report_lines.append(switch)
        if slot_port.strip(): report_lines.append(slot_port)
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

            # IDL Connections — confirmed real gap: was never written to the .xlsm at all.
            if len(new_nodes) > 1 and idl_build_type:
                idle_rows = N2E_ROW_MAP["idle"]
                row_writes.append((idle_rows[0], True, [(3, idl_build_type)]))
                row_writes.append((idle_rows[1], bool(idle.strip()), [(3, idle)] if idle.strip() else []))
            row_writes.append((N2E_ROW_MAP["idly"], bool(idly.strip()), [(3, idly)] if idly.strip() else []))
            row_writes.append((N2E_ROW_MAP["switch"], bool(switch.strip()), [(2, switch)] if switch.strip() else []))
            row_writes.append((N2E_ROW_MAP["slot_port"], bool(slot_port.strip()), [(2, slot_port)] if slot_port.strip() else []))

            def _rw_simple(row_num, checked, value=None, col=3):
                """Writes just the raw VALUE into its one confirmed column — never the
                whole formatted sentence (label + value combined), which was the
                confirmed bug: the real template already has the label in column B, so
                writing the full sentence into C duplicated it and left the real target
                columns untouched."""
                row_writes.append((row_num, bool(checked), [(col, value)] if checked and value else []))

            # ---- Completed ----
            # Integration (42,43,44): C=bands only, D=node only, per node — confirmed
            # real column split, not the whole "Integration: bands node" sentence in one cell.
            int_rows = N2E_ROW_MAP["integration"]["completed"]
            int_pairs = []
            if int_checked:
                for node, cells in classification.get("added", {}).items():
                    bands = sorted({app.band_label(c)[0] for c in cells if app.band_label(c)[0]})
                    if bands:
                        int_pairs.append(("/".join(bands), node))
            for i, row_num in enumerate(int_rows):
                if i < len(int_pairs):
                    row_writes.append((row_num, True, [(3, int_pairs[i][0]), (4, int_pairs[i][1])]))
                else:
                    row_writes.append((row_num, False, []))

            _rw_simple(N2E_ROW_MAP["controller_integration"]["completed"][0],
                       not cascade_fires and controller_line and ctrl_checked, controller_id)
            dss_bands = dss_line.split("DSS Activation:")[-1].strip() if dss_completed and dss_line else None
            _rw_simple(N2E_ROW_MAP["dss_activation"]["completed"][0], dss_completed, dss_bands)
            ngs_bands = ngs_line.split("\t")[1] if ngs_completed and ngs_line and "\t" in ngs_line else None
            ngs_node = ngs_line.split("\t")[2] if ngs_completed and ngs_line and ngs_line.count("\t") >= 2 else None
            ngs_row = N2E_ROW_MAP["ngs_activation"]["completed"][0]
            row_writes.append((ngs_row, bool(ngs_completed),
                                [(3, ngs_bands), (4, ngs_node)] if ngs_completed else []))
            # GPS Installation (48): C=node(s), D="Version:" (literal, already in template,
            # don't touch), E=type value only.
            gps_row = N2E_ROW_MAP["gps_installation"]["completed"][0]
            if gps_completed_line:
                row_writes.append((gps_row, True, [(3, "|".join(enabled_nodes)), (5, gtype)]))
            else:
                row_writes.append((gps_row, False, []))
            _rw_simple(N2E_ROW_MAP["lkf_installation"]["completed"][0], lkf_completed, lkf_completed)  # confirmed: template's C is one combined "node | controller" column
            sfp_rows = N2E_ROW_MAP["transport_sfp"]["completed"]
            for i, row_num in enumerate(sfp_rows):
                if i < len(new_nodes) and i < len(sfp_completed_lines):
                    node = new_nodes[i]
                    m = re.search(r'on: \S+ (.+?) \(BBU End\) & (.+?) \(SIAD End\)', sfp_completed_lines[i])
                    models = f"{m.group(1)} (BBU End) & {m.group(2)} (SIAD End)" if m else ""
                    row_writes.append((row_num, True, [(3, node), (4, models)]))
                else:
                    row_writes.append((row_num, False, []))
            _rw_simple(N2E_ROW_MAP["ret_configuration"]["completed"][0], False)
            _rw_simple(N2E_ROW_MAP["external_alarm_scripting"]["completed"][0], not cascade_fires and alarm_scripting_completed, controller_id)
            _rw_simple(N2E_ROW_MAP["sau_connections"]["completed"][0], not cascade_fires and sau_completed, controller_id)
            _rw_simple(N2E_ROW_MAP["sup_connections"]["completed"][0], sup_completed_lines, "|".join(s.split(": ")[-1] for s in sup_completed_lines))
            _rw_simple(N2E_ROW_MAP["xmu_installation"]["completed"][0], xmu_completed_lines, "|".join(s.split(": ")[-1] for s in xmu_completed_lines))
            _rw_simple(N2E_ROW_MAP["idl_connections"]["completed"][0], idl_completed)
            _rw_simple(N2E_ROW_MAP["smm_triggering"]["completed"][0], smm_completed)
            _rw_simple(N2E_ROW_MAP["area_test"]["completed"][0], False)
            _rw_simple(N2E_ROW_MAP["external_alarm_testing"]["completed"][0], not cascade_fires and testing_completed, controller_id)
            _rw_simple(N2E_ROW_MAP["script_load_6673"]["completed"][0], script_6673_completed,
                       sidehaul_rows[0]["switch_id"] if script_6673_completed and sidehaul_rows else None)
            _rw_simple(N2E_ROW_MAP["installation_manual"]["completed"][0], False)
            _rw_simple(N2E_ROW_MAP["sa_conversion"]["completed"][0], sa_completed_line, "|".join(sa_nodes) if sa_completed_line else None)

            # ---- Pending ----
            _rw_simple(N2E_ROW_MAP["on_site_nokia_cutover"]["pending"][0], True)  # no value column, just checkbox + template's own fixed "(Tower Crew)"
            _rw_simple(N2E_ROW_MAP["controller_integration"]["pending"][0],
                       cascade_fires and controller_line, controller_id)
            dss_p_bands = dss_line.split("DSS Activation:")[-1].strip() if dss_pending and dss_line else None
            _rw_simple(N2E_ROW_MAP["dss_activation"]["pending"][0], dss_pending, dss_p_bands)
            ngs_p_bands = ngs_line.split("\t")[1] if ngs_pending and ngs_line and "\t" in ngs_line else None
            ngs_p_node = ngs_line.split("\t")[2] if ngs_pending and ngs_line and ngs_line.count("\t") >= 2 else None
            ngs_p_row = N2E_ROW_MAP["ngs_activation"]["pending"][0]
            row_writes.append((ngs_p_row, bool(ngs_pending), [(3, ngs_p_bands), (4, ngs_p_node)] if ngs_pending else []))
            # GPS Pending (82): C=node(s), D=type value directly (no "Version:" label
            # column here, confirmed different structure from the Completed row).
            gps_p_row = N2E_ROW_MAP["gps_installation"]["pending"][0]
            row_writes.append((gps_p_row, bool(gps_pending_line), [(3, "|".join(disabled_nodes))] if gps_pending_line else []))
            lkf_p_final = "LKF Installation. (MIC)" if cascade_fires else lkf_pending
            lkf_p_value = (lkf_p_final.replace("LKF Installation:", "").replace("(MIC)", "").strip()
                           if lkf_p_final and lkf_p_final != "LKF Installation. (MIC)" else None)
            _rw_simple(N2E_ROW_MAP["lkf_installation"]["pending"][0], lkf_p_final, lkf_p_value)
            _rw_simple(N2E_ROW_MAP["psap_speedtest"]["pending"][0], lte_bands, lte_bands)
            _rw_simple(N2E_ROW_MAP["speed_test_5g"]["pending"][0], fiveg_bands, fiveg_bands)
            _rw_simple(N2E_ROW_MAP["calltest_fnet"]["pending"][0], False)
            _rw_multi_unused = N2E_ROW_MAP["transport_sfp"]["pending"]
            for row_num in _rw_multi_unused:
                row_writes.append((row_num, False, []))
            _rw_simple(N2E_ROW_MAP["ret_configuration"]["pending"][0], True)  # no value column, template's own fixed "(Tower Crew)"
            _rw_simple(N2E_ROW_MAP["external_alarm_scripting"]["pending"][0], cascade_fires, controller_id if cascade_fires else None)
            _rw_simple(N2E_ROW_MAP["sau_connections"]["pending"][0], cascade_fires or sau_pending, controller_id if (cascade_fires or sau_pending) else None)
            _rw_simple(N2E_ROW_MAP["sup_connections"]["pending"][0], sup_pending_lines, "|".join(s.split(":")[-1].replace(". (MIC PM)", "").strip() for s in sup_pending_lines))
            _rw_simple(N2E_ROW_MAP["xmu_installation"]["pending"][0], xmu_pending_lines, "|".join(s.split(":")[-1].replace(". (MIC PM)", "").strip() for s in xmu_pending_lines))
            _rw_simple(N2E_ROW_MAP["idl_connections"]["pending"][0], idl_pending)
            _rw_simple(N2E_ROW_MAP["siad_provisioning"]["pending"][0], False)
            # Area test Pending (98): C=node(s), D="Failed" (the actual Area Lite result).
            area_p_row = N2E_ROW_MAP["area_test"]["pending"][0]
            row_writes.append((area_p_row, bool(area_pending), [(3, "|".join(new_nodes)), (4, "Failed")] if area_pending else []))
            _rw_simple(N2E_ROW_MAP["external_alarm_testing"]["pending"][0], cascade_fires or testing_pending, controller_id if (cascade_fires or testing_pending) else None)
            _rw_simple(N2E_ROW_MAP["config_6673"]["pending"][0], has_6673, sidehaul_rows[0]["switch_id"] if has_6673 and sidehaul_rows else None)
            # 6673 Port Config in ENM (101): C=Slot|Port|Node (never built, left blank), D=Switch ID.
            _rw_simple(N2E_ROW_MAP["port_config_6673_enm"]["pending"][0], has_6673,
                       sidehaul_rows[0]["switch_id"] if has_6673 and sidehaul_rows else None, col=4)
            # Link failure / SFP Not Present (102/103): C=Radio(bands), D=Node ID — confirmed
            # separate columns, not one combined sentence.
            link_rows = N2E_ROW_MAP["link_failure_or_sfp"]["pending"]
            link_nodes = "|".join([row.get("Node to be built as") for row in mm_objs if row.get("Node to be built as")])
            if link_choice == "Link failure":
                row_writes.append((link_rows[0], True, [(3, full_bands), (4, link_nodes)]))
                row_writes.append((link_rows[1], False, []))
            elif link_choice == "SFP Not Present":
                row_writes.append((link_rows[0], False, []))
                row_writes.append((link_rows[1], True, [(3, full_bands), (4, link_nodes)]))
            else:
                row_writes.append((link_rows[0], False, []))
                row_writes.append((link_rows[1], False, []))
            _rw_simple(N2E_ROW_MAP["sa_conversion"]["pending"][0], False)

            # Pending buffer (105-112): the merged Power Plant Swap/Not-Cleared-by-FE
            # line + any free-text "additional pending information".
            pending_buffer_lines = list(bucket_pending) + ([additional_pending] if additional_pending.strip() else [])
            pending_buffer_rows = N2E_ROW_MAP["additional_pending"]
            for i, row_num in enumerate(pending_buffer_rows):
                if i < len(pending_buffer_lines):
                    row_writes.append((row_num, True, [(2, pending_buffer_lines[i])]))
                else:
                    row_writes.append((row_num, False, []))

            # ---- Notes ----
            final_port_line = "Final Port Configuration attached." if "Final Port Configuration attached." in choices_notes else None
            row_writes.append((N2E_ROW_MAP["notes_final_port_config"], bool(final_port_line), [(2, final_port_line)] if final_port_line else []))
            nr_line = "NR configuration has been verified." if "NR configuration has been verified." in choices_notes else None
            row_writes.append((N2E_ROW_MAP["notes_nr_verified"], bool(nr_line), [(2, nr_line)] if nr_line else []))

            other_notes = [n for n in choices_notes
                            if n not in ("Final Port Configuration attached.", "NR configuration has been verified.")]
            notes_buffer_rows = N2E_ROW_MAP["notes_buffer"]
            for i, row_num in enumerate(notes_buffer_rows):
                if i < len(other_notes):
                    row_writes.append((row_num, True, [(2, other_notes[i])]))
                else:
                    row_writes.append((row_num, False, []))

            xlsm_bytes = fill_legacy_mca_surgical(N2E_TEMPLATE_PATH, row_writes)
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_N2E_Filled.xlsm", key="n2e_dl_xlsm")
