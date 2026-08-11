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
import io
import zipfile
import streamlit as st

import nsb_completed_logic as nsb
import mca_completed_logic as mcl
from nsb_row_map import NSB_ROW_MAP
from mca_xlsm_surgical import fill_legacy_mca_surgical
from pathlib import Path

NSB_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Static" / "NSB_Macro_Template_v4_updated (2).xlsm"


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
    # Confirmed hard block: at least one CIQ node ID must appear across every
    # uploaded document together — otherwise this is treated as a wrong/mismatched
    # site upload, and the report must not be generated at all.
    if mcl.detect_site_mismatch(mm_objs, postcheck_text, controller_checks_text):
        st.error("Wrong input given: none of this CIQ's node IDs were found together across the uploaded documents. Please confirm you've uploaded the correct files for this site.")
        st.stop()
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

    # Confirmed same fix as N2E: a node whose hardware string can't be found anywhere in
    # Post-checks at all is treated as genuinely NOT integrated. integrated_nodes
    # excludes these — used for per-node items that should only reflect nodes that
    # actually made it into Post-checks. new_nodes itself stays unfiltered, since Post
    # Configuration's own display and Pending line need to list every node.
    missing_nodes = mcl.detect_missing_nodes(postcheck_text, new_nodes)
    integrated_nodes = [n for n in new_nodes if n not in missing_nodes]
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
        # Confirmed fix: was missing the dual-identity case entirely (a node with BOTH an
        # eNodeB Name and gNodeB Name populated — i.e. co-located LTE+5G) — ported from
        # app.py's generate_n2e(), which correctly shows the secondary ID and BBU Mode.
        if app.is_populated(e_name) and app.is_populated(g_name):
            secondary = g_name if is_lte_primary else e_name
            bbu_mode = row.get("BBU Mode")
            post_parts.append(f"{node}(P)/{secondary}(S)({bbu_mode})({hw or 'NOT FOUND'})")
        else:
            post_parts.append(f"{node}({hw or 'NOT FOUND'})")
    post_line = " + ".join(post_parts)

    # Moved earlier (previously computed later, well after the Subject UI had already
    # started rendering) since the Warnings gate below needs this data before anything
    # else renders.
    controller_checks_data = mcl.extract_controller_checks(controller_checks_text) if controller_checks_text else {}
    cascade_fires = nsb.controller_cascade_fires(controller_in_edp, bool(controller_checks_text))

    # SA Conversion — moved earlier (previously computed later, inside Notes) since the
    # AMF warning check below needs it before the gate.
    nsb_sa_nodes = mcl.sa_conversion_nodes(ciq_wb, mm_objs)

    # ==================== Warnings gate ====================
    # Same confirmed 2-step flow as N2E: if there are any warnings, show ONLY the
    # warnings plus a "Continue to Report" button first — none of the rest of the UI
    # renders until the user clicks through. Acknowledgment is keyed against the actual
    # warnings content (not a plain persistent flag), so a genuinely different set of
    # warnings on a later site in the same session always re-triggers the gate.
    warnings = []
    nsb_pending_from_warnings = []

    if postcheck_text:
        transport_sfp_data = mcl.extract_transport_sfp(postcheck_text)
        sfp_warning_texts, sfp_pending_lines = mcl.transport_sfp_threshold_warnings(
            ciq_wb, mm_objs, postcheck_text, transport_sfp_data)
        warnings += sfp_warning_texts
        nsb_pending_from_warnings += sfp_pending_lines

        # Confirmed perf fix: compute these once here instead of letting each of the
        # three warning checks below independently re-parse the same sheets.
        _warnings_eutran_rows = app.sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
        _warnings_fiveg_rows = app.sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []

        warnings += mcl.lte_sector_param_warnings(ciq_wb, mm_objs, postcheck_text, _warnings_eutran_rows)
        warnings += mcl.fiveg_sector_param_warnings(ciq_wb, mm_objs, postcheck_text, _warnings_fiveg_rows)
        warnings += mcl.sctp_status_warnings(postcheck_text)
        warnings += mcl.digital_tilt_warnings(ciq_wb, mm_objs, postcheck_text, classification, _warnings_fiveg_rows)
        warnings += mcl.lte_cell_presence_warnings(ciq_wb, postcheck_text, _warnings_eutran_rows)
        warnings += mcl.fiveg_cell_presence_warnings(ciq_wb, postcheck_text, _warnings_fiveg_rows)
        warnings += mcl.sup_capacity_warning(postcheck_text, integrated_nodes)
        warnings += mcl.xmu_sup_locked_warning(postcheck_text, integrated_nodes)

    if nsb_sa_nodes and postcheck_text:
        warnings += mcl.sa_conversion_amf_warning(postcheck_text, nsb_sa_nodes)

    warnings_fingerprint = hash(tuple(sorted(warnings)))
    if warnings and st.session_state.get("nsb_warnings_ack_fingerprint") != warnings_fingerprint:
        with st.container(border=True):
            st.markdown(
                f"<div style='color:#c0392b; font-size:1.3em; font-weight:700;'>"
                f"\u26a0\ufe0f {len(warnings)} Warning{'s' if len(warnings) != 1 else ''}</div>",
                unsafe_allow_html=True)
            for w in warnings:
                st.markdown(f"<div style='color:#c0392b; font-size:1.05em; padding:2px 0;'>\u2022 {w}</div>",
                            unsafe_allow_html=True)
        if st.button("Continue to Report \u2192", type="primary", key="nsb_warnings_continue"):
            st.session_state["nsb_warnings_ack_fingerprint"] = warnings_fingerprint
            st.rerun()
        return

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
        with c[1]: market = st.selectbox("\U0001F4DD Market", ["\u2014 Select \u2014", "MNS", "AT&T"], key="nsb_market")
        with c[2]: status = st.selectbox("\U0001F4DD Status", ["\u2014 Select \u2014", "IX-STF", "IX-ATP"], key="nsb_status")
        market = "" if market == "\u2014 Select \u2014" else market
        status = "" if status == "\u2014 Select \u2014" else status
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
        if missing_nodes:
            st.markdown(f"Post Configuration (Pending) : **{post_line}(MIC PM)**")
            st.warning(f"Node(s) not found in Post-checks, treated as not yet integrated: {', '.join(missing_nodes)}")
        st.markdown(f"6610 Controller : **{controller_id or '(none detected)'}**")
        current_config_auto = mcl.current_configuration_line(ciq_wb, mm_objs, postcheck_text, missing_nodes, dual_identity=True) if postcheck_text else ""
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
    if missing_nodes:
        choices_pending.append(f"Post Configuration : {post_line}(MIC PM)")

    with st.expander("Completed", expanded=True):
        int_pairs = []
        int_lines_display = []
        for node, cells in classification.get("added", {}).items():
            if node in missing_nodes:
                continue
            bands = mcl.sort_bands_lte_first({app.band_label(c)[0] for c in cells if app.band_label(c)[0]})
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
        dss_pending_bands_combined = set()  # confirmed fix: tracked directly, since dss_bands alone only ever holds the user-choice portion
        dss_sh = None
        if dss_activation_labels:
            nsb_dss_calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
            nsb_dss_regional_market = None
            if nsb_dss_calltest_path.exists() and mm_objs:
                _dss_prefix_map, _dss_rules = mcl.load_calltest_table(nsb_dss_calltest_path, tab_name="NSB")
                nsb_dss_regional_market = mcl.determine_market(mm_objs[0].get("Node to be built as"), _dss_prefix_map)
            nsb_scripted_locked_bands = mcl.scripted_locked_bands(ciq_wb)
            # Confirmed new market override: for these 4 markets, DSS always goes
            # straight to Completed regardless of scripted/locked status — the normal
            # auto-pending/user-choice split only applies to every other market.
            _dss_always_completed_markets = {"NCSC", "Florida", "AR-OK", "STX"}
            if nsb_dss_regional_market in _dss_always_completed_markets:
                with st.container(border=True):
                    dss_bands_all_fmt = " & ".join(mcl.sort_bands_lte_first(set(dss_activation_labels)))
                    st.markdown(f"**DSS Activation** \u2014 detected: {dss_bands_all_fmt}")
                    dss_completed = f"DSS Activation: {dss_bands_all_fmt}"
                    choices_completed.append(dss_completed)
                    st.caption(f"\u2705 {dss_completed} ({nsb_dss_regional_market} market \u2014 always Completed)")
                    dss_bands = dss_bands_all_fmt
            else:
                auto_pending_bands, user_choice_bands = mcl.split_dss_bands_by_scripted_locked(
                    set(dss_activation_labels), nsb_scripted_locked_bands, nsb_dss_regional_market)

                if auto_pending_bands or user_choice_bands:
                    with st.container(border=True):
                        st.markdown(f"**DSS Activation** \u2014 detected: {' & '.join(dss_activation_labels)}")
                        if auto_pending_bands:
                            auto_bands_fmt = " & ".join(mcl.sort_bands_lte_first(auto_pending_bands))
                            dss_pending_auto = f"DSS Activation: {auto_bands_fmt} (AT&T)"
                            st.caption(f"\u26a0\ufe0f Scripted/locked \u2014 goes directly to Pending: {dss_pending_auto}")
                            dss_pending = dss_pending_auto
                            dss_pending_bands_combined |= auto_pending_bands
                        if user_choice_bands:
                            dss_bands = " & ".join(mcl.sort_bands_lte_first(user_choice_bands))
                            st.caption(f"Remaining band(s) \u2014 pick Completed or Pending: {dss_bands}")
                            dss_choice = st.selectbox("Status", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_dss")
                            if dss_choice == "Completed":
                                dss_completed = f"DSS Activation: {dss_bands}"
                                choices_completed.append(dss_completed)
                            elif dss_choice == "Pending":
                                dss_sh = st.selectbox("Stakeholder", ["\u2014 Select \u2014", "MIC", "AT&T"], key="nsb_dss_sh")
                                if dss_sh != "\u2014 Select \u2014":
                                    user_pending_line = f"DSS Activation: {dss_bands} ({dss_sh})"
                                dss_pending = (dss_pending + " | " + user_pending_line) if dss_pending else user_pending_line
                            dss_pending_bands_combined |= user_choice_bands

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
                if node in missing_nodes:
                    continue
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
                    if cascade_fires:
                        st.caption("Pending (auto \u2014 no 6610 checks)")
                        lkf_controller_choice = "Pending"
                    else:
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

        # Call Test — confirmed real fix: previously just dumped every detected band into
        # PSAP/Speedtest/Speed test unconditionally, never actually consulting the CT
        # sheet's rules table (which market/scenario combination actually requires which
        # test). Now reuses mcl.call_test_lines(), the same real function MCA uses.
        calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
        psap_line = speed_lte_line = speed_5g_line = fnet_line = None
        # Confirmed real bug (same class as GPS Installation's stakeholder fix): the
        # Subject-line "market" (MNS/TILLMAN/AT&T) is NOT the same thing as the regional
        # market (TN-KY/NCSC/etc.) that calltest_rules is actually keyed by. Passing the
        # Subject field directly meant rules.get((market, scenario)) never matched
        # anything real, so every Call Test check silently returned nothing.
        regional_market = None
        if calltest_path.exists() and mm_objs:
            _rm_prefix_to_market, _rm_calltest_rules = mcl.load_calltest_table(calltest_path, tab_name="NSB")
            regional_market = mcl.determine_market(mm_objs[0].get("Node to be built as"), _rm_prefix_to_market)
        market_val = regional_market
        if calltest_path.exists() and market_val:
            prefix_to_market, calltest_rules = mcl.load_calltest_table(calltest_path, tab_name="NSB")

            # NSB confirmed: every cell is "newly added" (no "moved" concept, same as
            # N2E) — split by tech into the three categories call_test_lines() needs.
            added_bands_by_tech = {"lte": set(), "5g": set(), "cband_dod": set()}
            for node, cells in classification.get("added", {}).items():
                if node in missing_nodes:
                    continue
                for c in cells:
                    label, _sector = app.band_label(c)
                    if not label:
                        continue
                    if label in ("CBAND", "DOD", "DOD_BWE"):
                        added_bands_by_tech["cband_dod"].add(label)
                    elif label.startswith("5G_"):
                        added_bands_by_tech["5g"].add(label)
                    else:
                        added_bands_by_tech["lte"].add(label)

            ct_lines = mcl.call_test_lines(classification, market_val, calltest_rules,
                                             set(), added_bands_by_tech, {"lte": set(), "5g": set(), "cband_dod": set()})
            psap_applies = any(l.startswith("PSAP test/Speedtest/VoLTE voice calltest") for l in ct_lines)
            for l in ct_lines:
                l_display = l.replace("\t", " ")
                if l.startswith("PSAP test/Speedtest/VoLTE voice calltest"):
                    psap_line = l_display
                elif l.startswith("Speedtest/VoLTE voice calltest"):
                    speed_lte_line = l_display
                elif l.startswith("Speed test"):
                    speed_5g_line = l_display
                elif l.startswith("Calltest with F-NET SIM"):
                    fnet_line = l_display

            # Confirmed fix: detected items now shown FIRST, before the status selector,
            # so the engineer knows what the CT sheet actually requires before choosing
            # a status. PSAP Schedule ID no longer renders unconditionally up top — it
            # only appears once genuinely relevant: Completed status with PSAP detected,
            # or Partially Completed status where the engineer has typed "PSAP" into the
            # completed field themselves.
            _ct_any_detected = bool(psap_line or speed_lte_line or speed_5g_line or fnet_line)
            if _ct_any_detected:
                st.markdown("**Call Test requirements detected (per CT sheet):**")
                for l in (psap_line, speed_lte_line, speed_5g_line, fnet_line):
                    if l:
                        st.caption(l)

            # Confirmed redesign: Call Test now has 3 possible states instead of always
            # reporting Completed. "Completed" keeps the existing auto-detect behavior
            # (including PSAP Schedule ID, now asked only within this branch).
            # "Pending" routes every detected item to Pending instead. "Partially
            # Completed" shows what the CT sheet detected as required (already shown
            # above) and lets the engineer manually type in both what was actually
            # completed and what's still pending, at whatever granularity they observed
            # in the field (band or specific sector) — since call_test_lines() only
            # tracks bands, not individual sectors within a band, the tool can't
            # reliably auto-split a partial completion itself.
            # Confirmed fix: the whole status selector (and everything below it) is now
            # skipped entirely when there's genuinely nothing detected, rather than
            # showing a dropdown with nothing meaningful to report a status for.
            ct_status = st.selectbox("Call Test status", ["\u2014 Select \u2014", "Completed", "Pending", "Partially Completed"], key="nsb_calltest_status") \
                if _ct_any_detected else None
            lte_bands_all = sorted(added_bands_by_tech["lte"])
            fiveg_bands_all = sorted(added_bands_by_tech["5g"] | added_bands_by_tech["cband_dod"])
            psap_sched_id = ""
            ct_completed_inputs, ct_pending_inputs = {}, {}

            if ct_status == "Completed":
                if psap_applies:
                    psap_sched_id = st.text_input("\U0001F4DD PSAP Schedule ID", key="nsb_psap_sched")
                    if psap_sched_id.strip():
                        psap_line = psap_line.replace("PSAP Schedule ID: ", f"PSAP Schedule ID: {psap_sched_id.strip()}")
                    else:
                        psap_line = psap_line.replace(" PSAP Schedule ID: ", "").rstrip()
                for l in (psap_line, speed_lte_line, speed_5g_line, fnet_line):
                    if l:
                        choices_completed.append(l)
                        st.caption(f"\u2705 {l}")
            elif ct_status == "Pending":
                # Confirmed: PSAP Schedule ID is never relevant while Pending — nothing
                # has been tested yet, so strip the empty label out of the line entirely
                # rather than showing a dangling "PSAP Schedule ID:" with no value.
                psap_pending_line = psap_line.replace(" PSAP Schedule ID: ", "").rstrip() if psap_line else psap_line
                for l in (psap_pending_line, speed_lte_line, speed_5g_line, fnet_line):
                    if l:
                        pending_line = f"{l} (MIC PM)"
                        choices_pending.append(pending_line)
                        st.caption(pending_line)
            elif ct_status == "Partially Completed":
                # Confirmed redesign: separate Completed/Pending input pairs PER
                # detected test type (PSAP, LTE Speed test, 5G Speed test, F-NET)
                # instead of one flat pair that lumped everything into a single
                # generic "Call Test completed/pending on" line — output preserves
                # each test type's own label and wording, matching exactly how the
                # Completed status already reports them.
                # Confirmed layout fix: all Completed inputs grouped in one column, all
                # Pending inputs grouped in the other — rather than interleaving each
                # test type's pair with its own result caption right after, which
                # looked scattered. Results now shown together in one place at the end.
                ct_items = [
                    ("psap", psap_line, "PSAP test/Speedtest/VoLTE voice calltest"),
                    ("speed_lte", speed_lte_line, "Speedtest/VoLTE voice calltest"),
                    ("speed_5g", speed_5g_line, "Speed test"),
                    ("fnet", fnet_line, "Calltest with F-NET SIM"),
                ]
                detected_items = [(k, l, lbl) for k, l, lbl in ct_items if l]
                ct_completed_inputs, ct_pending_inputs = {}, {}
                col_completed, col_pending = st.columns(2)
                with col_completed:
                    st.markdown("**Completed**")
                    for item_key, _detected_line, label in detected_items:
                        ct_completed_inputs[item_key] = st.text_input(
                            f"{label} \u2014 Completed on", key=f"nsb_ct_{item_key}_completed")
                        if item_key == "psap" and psap_applies and ct_completed_inputs[item_key].strip():
                            psap_sched_id = st.text_input("PSAP Schedule ID", key="nsb_psap_sched")
                with col_pending:
                    st.markdown("**Pending**")
                    for item_key, _detected_line, label in detected_items:
                        ct_pending_inputs[item_key] = st.text_input(
                            f"{label} \u2014 Pending on", key=f"nsb_ct_{item_key}_pending")

                ct_result_completed, ct_result_pending = [], []
                for item_key, _detected_line, label in detected_items:
                    completed_input = ct_completed_inputs.get(item_key, "")
                    pending_input = ct_pending_inputs.get(item_key, "")
                    if completed_input.strip():
                        line = f"{label}: {completed_input.strip()}."
                        if item_key == "psap" and psap_sched_id.strip():
                            line += f" (PSAP Schedule ID: {psap_sched_id.strip()})"
                        choices_completed.append(line)
                        ct_result_completed.append(line)
                    if pending_input.strip():
                        line = f"{label}: {pending_input.strip()} (MIC PM)"
                        choices_pending.append(line)
                        ct_result_pending.append(line)
                if ct_result_completed or ct_result_pending:
                    st.markdown("**Result:**")
                    for l in ct_result_completed:
                        st.caption(f"\u2705 {l}")
                    for l in ct_result_pending:
                        st.caption(l)

        # Confirmed new fix: BBU End auto-fetches from the Transport SFP table already
        # parsed from the Post-checks PDF (ericssonprod, matched by node ID). SIAD End
        # stays fully manual.
        sfp_completed_lines = []
        if integrated_nodes:
            with st.container(border=True):
                st.markdown("**Transport SFP Installation on** \u2014 Enter SFP models")
                for node in integrated_nodes:
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1: st.caption(node)
                    with c2:
                        bbu_auto = transport_sfp_data.get(node, {}).get("ericssonprod", "")
                        bbu = st.text_input("BBU", value=bbu_auto, key=f"nsb_sfp_bbu_{node}", label_visibility="collapsed", placeholder="SFP Model (BBU End)")
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
        alarm_notes_line = None
        alarm_ports_report_lines = []
        sau_completed, sau_pending = None, None
        sau_disabled = False
        if controller_checks_data and not cascade_fires:
            if mcl.external_alarm_scripting_confirmed(controller_checks_data):
                alarm_scripting_completed = f"External alarm Scripting on: {controller_id}"
                choices_completed.append(alarm_scripting_completed)
                st.caption(f"\u2705 {alarm_scripting_completed}")
            alarm_notes_line = nsb.external_alarm_scripting_locked_note(controller_checks_data)
            # Confirmed redesign: replaces both the old locked-ports-only and
            # active-ports-only checks with a single 3-category split (Active+Locked,
            # Active+Unlocked, NotActive+Locked). Confirmed Florida exception: computed
            # independently here since the market lookup used for External alarm
            # testing's own Florida check isn't computed until later in render().
            _alarm_ports_calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
            _alarm_ports_market = None
            if _alarm_ports_calltest_path.exists() and mm_objs:
                _alarm_ports_prefix_to_market, _ = mcl.load_calltest_table(_alarm_ports_calltest_path, tab_name="NSB")
                _alarm_ports_market = mcl.determine_market(mm_objs[0].get("Node to be built as"), _alarm_ports_prefix_to_market)
            # Confirmed fix: when EVERY scripted port is locked, the simple
            # "All external alarms are kept locked, due to NEA is pending." note
            # (from external_alarm_testing_placement) already covers this completely —
            # this 3-category breakdown would just redundantly re-report the same
            # locked ports individually. testing_section isn't computed until later in
            # render(), so checked independently here too, same ordering fix as above.
            _alarm_ports_testing_section, _, _ = mcl.external_alarm_testing_placement(controller_checks_data)
            if _alarm_ports_testing_section == "Pending":
                alarm_ports_report_lines = []
            else:
                alarm_ports_report_lines = nsb.external_alarm_ports_report(controller_checks_data, _alarm_ports_market)
            for l in alarm_ports_report_lines:
                nsb_pending_from_warnings.append(l)
            sau_state = controller_checks_data.get("sau_state")
            if sau_state:
                if sau_state["oper"] == "ENABLED":
                    sau_completed = f"SAU Connections: {controller_id}"
                    choices_completed.append(sau_completed)
                    st.caption(f"\u2705 {sau_completed}")
                else:
                    sau_disabled = True
                    sau_pending = f"SAU Connections: {controller_id} (MIC PM)"

        sup_completed_lines, sup_pending_lines = [], []
        xmu_completed_lines, xmu_pending_lines = [], []
        # Confirmed correction, same as N2E: SUP Connections now has its own per-node
        # trigger (5216 OR XMU present in that specific node's CIQ target).
        if postcheck_text:
            sup_expecting_nodes = mcl.nodes_expecting_sup(mm_objs, ciq_wb) & set(integrated_nodes)
            sup_state, sup_missing = nsb.sup_connections_state(postcheck_text, sup_expecting_nodes)
            for node, state in sup_state.items():
                (sup_completed_lines if state == "ENABLED" else sup_pending_lines).append(
                    f"SUP Connections: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
            for node in sorted(sup_missing):
                sup_pending_lines.append(f"SUP Connections: {node} (MIC PM)")
        if postcheck_text and xmu_present_in_ciq:
            xmu_state = nsb.xmu_installation_state(postcheck_text, xmu_present_in_ciq)
            for node, state in xmu_state.items():
                (xmu_completed_lines if state == "ENABLED" else xmu_pending_lines).append(
                    f"XMU Installation: {node}" + ("" if state == "ENABLED" else " (MIC PM)"))
            # Confirmed same fix as N2E: a node genuinely present in Post-checks but
            # whose expected XMU never appears in Post-checks' Hardware Status at all
            # was previously silently dropped entirely. Now reports Pending.
            xmu_expected_nodes = mcl.nodes_expecting_xmu(mm_objs, ciq_wb) & set(integrated_nodes)
            for node in sorted(xmu_expected_nodes - set(xmu_state.keys())):
                xmu_pending_lines.append(f"XMU Installation: {node} (MIC PM)")
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
        if integrated_nodes:
            # Confirmed fix: Area test should include the controller alongside the
            # nodes for the same 3 conditions as Controller monitored state's
            # auto-trigger — no 6610 checks, SAU disabled, or External alarm testing
            # Pending. testing_section computed independently here since it isn't
            # computed until later in render(), same ordering fix used elsewhere in
            # this file — cascade_fires/sau_disabled are already safely defined by
            # this point.
            _area_testing_section, _, _ = mcl.external_alarm_testing_placement(controller_checks_data) \
                if controller_checks_data else (None, None, None)
            _area_include_controller = cascade_fires or sau_disabled or _area_testing_section == "Pending"
            _area_targets = list(integrated_nodes) + ([controller_id] if _area_include_controller and controller_id else [])
            area_pending = f"Area test: {'|'.join(_area_targets)}: Area Lite - Failed (MIC PM)"

        testing_completed, testing_pending = None, None
        if testing_completed is None and controller_checks_data and not cascade_fires:
            testing_section, _, _ = mcl.external_alarm_testing_placement(controller_checks_data)
            if testing_section == "Completed":
                testing_completed = f"External alarm testing: {controller_id}"
                choices_completed.append(testing_completed)
                st.caption(f"\u2705 {testing_completed}")
            elif testing_section == "Pending":
                # Confirmed exception: Florida market reports to AT&T, not MIC PM.
                # Computed independently here since regional_market isn't reliably in
                # scope this early (only conditionally set inside the GPS section).
                _testing_calltest_path = Path(__file__).parent / "templates" / "Static" / "Calltest_sheet.xlsx"
                _testing_market = None
                if _testing_calltest_path.exists() and mm_objs:
                    _testing_prefix_to_market, _ = mcl.load_calltest_table(_testing_calltest_path, tab_name="NSB")
                    _testing_market = mcl.determine_market(mm_objs[0].get("Node to be built as"), _testing_prefix_to_market)
                _testing_stakeholder = "AT&T" if _testing_market == "Florida" else "MIC PM"
                testing_pending = f"External alarm testing: {controller_id} ({_testing_stakeholder})"

        script_6673_completed = None
        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            script_6673_completed = f"6673 Script load: {switch_id}"
            choices_completed.append(script_6673_completed)

        # Installation — confirmed purely manual, handled directly in the macro, not
        # duplicated in the UI.

    with st.expander("Pending", expanded=True):
        # Transport SFP threshold/BER warnings — reported here too since there's no
        # dedicated template row for this (goes through the buffer for the .xlsm).
        choices_pending += nsb_pending_from_warnings
        for l in nsb_pending_from_warnings:
            st.caption(l)

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
                ("External alarm Scripting on", f"External alarm Scripting on: {controller_id}. (MIC)"),
                ("External alarm testing", f"External alarm testing: {controller_id}. (MIC PM)"),
                ("Area test", area_pending or "Area test. (MIC PM)"),
                ("SAU Connections", f"SAU Connections: {controller_id}. (MIC PM)"),
            ]:
                choices_pending.append(item_text)
                st.caption(item_text)
        else:
            if alarm_notes_line:
                choices_notes.append(alarm_notes_line)
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

        if has_6673:
            switch_id = sidehaul_rows[0]["switch_id"] if sidehaul_rows else ""
            l1 = f"6673 Configuration: {switch_id} (AT&T)"
            l2 = f"6673 Port Configuration in ENM: {switch_id} (AT&T)"
            choices_pending += [l1, l2]
            st.caption(l1); st.caption(l2)

        additional_pending = st.text_area("\U0001F4DD Enter any additional pending information",
                                           key="nsb_add_pending", height=80)
        if additional_pending.strip():
            choices_pending.append(additional_pending)

    # Florida newly added cells — same as MCA.
    florida_cells = mcl.florida_newly_added_cells(regional_market, classification)
    florida_rows = mcl.florida_cells_to_rows(florida_cells)
    florida_checked = False
    if florida_rows:
        with st.expander("Florida (newly added CBAND/DOD)"):
            florida_checked = st.checkbox(f"Newly added Cells (Florida market) \u2014 {len(florida_cells)} cell(s)",
                                            value=True, key="nsb_chk_florida")
            if florida_checked:
                for r in florida_rows:
                    st.caption(r)
    florida_active_rows = florida_rows if florida_checked else []

    # Confirmed: previously removed as purely manual with no auto-detection, now
    # re-added as its own top-level collapsible expander (siblings with
    # Completed/Pending/Notes, not nested inside any of them — Streamlit doesn't
    # allow nesting an expander inside another expander) — all go to Pending,
    # stakeholder Tower Crew, each a checkbox + Node ID + Sector input producing
    # "{label} on: {node} {sector} (Tower Crew)".
    issue_row_data = {}  # confirmed: collected here, written to the .xlsm later once row_writes exists

    with st.expander("Issues that needs to be reported", expanded=False):

        def _issue_row(label, key_suffix, row_map_key=None, row_map_index=0, col_width=None, data_link_label=None):
            # Confirmed format: "{label} : {sectors} : {node} (Tower Crew)" — Sectors
            # input comes before Node input, both free-text. Sectors given more width
            # than Node, per confirmed proportions. Confirmed: 9 of these 12 items have
            # a dedicated .xlsm row (row_map_key given) — Sectors written to column 3,
            # Node written to column 4 (or 5 for Fiberloss specifically, since Fiberloss
            # has an extra fixed "Data Link_1"/"Data Link_2" label written to column 4
            # first). The remaining items (no dedicated row, e.g. BER) flow into the
            # shared Pending buffer instead, still as the combined text line.
            cols = st.columns([0.6, 2, 1]) if col_width is None else col_width
            c1, c2, c3 = cols
            with c1:
                checked = st.checkbox(label, key=f"nsb_issue_{key_suffix}_chk")
            with c2:
                sector_input = st.text_input("Sectors", key=f"nsb_issue_{key_suffix}_sector", label_visibility="collapsed", placeholder="Sectors")
            with c3:
                node_input = st.text_input("Node", key=f"nsb_issue_{key_suffix}_node", label_visibility="collapsed", placeholder="Node")
            if checked and node_input.strip():
                sector_part = f"{sector_input.strip()} : " if sector_input.strip() else ""
                value_only = f"{sector_part}{node_input.strip()}"
                line = f"{label} : {value_only} (Tower Crew)"
                choices_pending.append(line)
                st.caption(line)
                if row_map_key:
                    issue_row_data[(row_map_key, row_map_index)] = (sector_input.strip(), data_link_label, node_input.strip())
                else:
                    nsb_pending_from_warnings.append(line)

        _issue_row("Link failure", "link_failure", row_map_key="link_failure")
        _issue_row("SFP Not Present", "sfp_not_present", row_map_key="sfp_not_present")
        _issue_row("Mo Inconsistent configuration alarm", "mo_inconsistent", row_map_key="mo_inconsistent_config_alarm")

        # Confirmed side by side, same proportions per item for consistent alignment
        # across all paired rows (Fiber loss, RSSI, VSWR).
        f1, f2, f3, f4, f5, f6 = st.columns([0.6, 2, 1, 0.6, 2, 1])
        _issue_row("High Fiber loss", "fiberloss_high", row_map_key="fiberloss", row_map_index=0, col_width=[f1, f2, f3], data_link_label="Data Link_1")
        _issue_row("Low Fiber loss", "fiberloss_low", row_map_key="fiberloss", row_map_index=1, col_width=[f4, f5, f6], data_link_label="Data Link_2")
        c1, c2, c3, c4, c5, c6 = st.columns([0.6, 2, 1, 0.6, 2, 1])
        _issue_row("High RSSI", "high_rssi", row_map_key="high_rssi", col_width=[c1, c2, c3])
        _issue_row("Low RSSI", "low_rssi", row_map_key="low_rssi", col_width=[c4, c5, c6])
        d1, d2, d3, d4, d5, d6 = st.columns([0.6, 2, 1, 0.6, 2, 1])
        _issue_row("High VSWR", "high_vswr", row_map_key="high_vswr", col_width=[d1, d2, d3])
        _issue_row("Low VSWR", "low_vswr", row_map_key="low_vswr", col_width=[d4, d5, d6])

        _issue_row("VSWR overthreshold alarm", "vswr_overthreshold", row_map_key="vswr_overthreshold")
        # Confirmed: BER items have no dedicated row in the template — flow into the
        # shared Pending buffer instead (nsb_pending_from_warnings, no row_map_key).
        _issue_row("NZ BER reporting", "nz_ber")
        _issue_row("BER not reporting", "ber_not_reporting")

    with st.expander("Notes"):
        # SAU enabled on Node/Controller — same as N2E, moved to the top. Confirmed:
        # hidden entirely when there's no controller-checks data at all, or when SAU is
        # explicitly confirmed DISABLED.
        if not cascade_fires and not sau_disabled:
            sau_auto_target = controller_id if sau_completed else None
            sau_label = f"SAU enabled on the: {sau_auto_target}" if sau_auto_target else "SAU enabled on the: Node or Controller"
            sau_notes_checked = st.checkbox(sau_label, value=True, key="nsb_sau_notes")
            if sau_notes_checked:
                if sau_auto_target:
                    choices_notes.append(f"SAU enabled on the: {sau_auto_target}")
                else:
                    sau_manual_target = st.text_input("\U0001F4DD Node ID or Controller ID", key="nsb_sau_manual")
                    if sau_manual_target.strip():
                        choices_notes.append(f"SAU enabled on the: {sau_manual_target}")

        nsb_scripted_locked_note = mcl.scripted_locked_bands_note(ciq_wb)
        if nsb_scripted_locked_note:
            nsb_scripted_locked_checked = st.checkbox(nsb_scripted_locked_note, value=True, key="nsb_scripted_locked")
            if nsb_scripted_locked_checked:
                choices_notes.append(nsb_scripted_locked_note)

        has_5g = any(app.is_populated(row.get("gNBId")) for row in mm_objs)
        final_port_checked = st.checkbox("Final Port Configuration attached.", value=True, key="nsb_notes_finalport")
        if final_port_checked:
            choices_notes.append("Final Port Configuration attached.")
        if has_5g:
            nr_checked = st.checkbox("NR configuration has been verified.", value=True, key="nsb_notes_nr")
            if nr_checked:
                choices_notes.append("NR configuration has been verified.")

        # Nodes/controller monitored state — same as N2E.
        if integrated_nodes:
            mon_checked = st.checkbox(f"{'|'.join(integrated_nodes)} nodes is in Not monitored state.", value=True, key="nsb_notes_mon")
            if mon_checked:
                choices_notes.append(f"{'|'.join(integrated_nodes)} nodes is in Not monitored state.")
        if controller_id:
            if cascade_fires or sau_disabled or testing_section == "Pending":
                _mon_reason = "no 6610 checks" if cascade_fires else ("SAU disabled" if sau_disabled else "External alarm testing Pending")
                st.caption(f"{controller_id} is in not monitored state. (auto \u2014 {_mon_reason})")
                choices_notes.append(f"{controller_id} is in not monitored state.")
            else:
                ctrl_mon_choice = st.selectbox(f"{controller_id} monitored state", ["\u2014 Select \u2014", "Monitored", "Not monitored"], key="nsb_ctrl_mon")
                if ctrl_mon_choice == "Monitored":
                    choices_notes.append(f"{controller_id} is in monitored state.")
                elif ctrl_mon_choice == "Not monitored":
                    choices_notes.append(f"{controller_id} is in not monitored state.")

        # SA Conversion / TermPointToAmf note — same as N2E, fires only if the CIQ's
        # NR_SA tab is present AND SA Conversion is detected for at least one node.
        nsb_sa_note = mcl.sa_conversion_note(nsb_sa_nodes)
        if nsb_sa_note:
            sa_note_checked = st.checkbox(nsb_sa_note, value=True, key="nsb_notes_sa")
            if sa_note_checked:
                choices_notes.append(nsb_sa_note)

        # Area prechecks verification for CPRI/SFP check — same as N2E.
        cpri_choice = st.selectbox("Area prechecks verification for CPRI/SFP check", ["\u2014 Select \u2014", "Completed", "Pending"], key="nsb_cpri") \
            if integrated_nodes else "\u2014 Select \u2014"
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
            # Confirmed correction: the 9 "Issues that needs to be reported" items with
            # a dedicated row (link_failure, sfp_not_present, mo_inconsistent_config_alarm,
            # fiberloss x2, high/low_rssi, high/low_vswr, vswr_overthreshold) — Sectors
            # written to column 3, Node written to column 4 for most items. Fiberloss
            # specifically also writes "Data Link_1"/"Data Link_2" to column 4, shifting
            # its own Node value to column 5 instead.
            for (_row_key, _row_idx), (_sector_val, _data_link_val, _node_val) in issue_row_data.items():
                _row_num = NSB_ROW_MAP[_row_key]["pending"][_row_idx]
                _cols = [(3, _sector_val)] if _sector_val else []
                if _data_link_val:
                    _cols += [(4, _data_link_val), (5, _node_val)]
                else:
                    _cols += [(4, _node_val)]
                row_writes.append((_row_num, True, _cols))
            row_writes.append((NSB_ROW_MAP["subject"], True, [(2, "MIC"), (3, market), (4, status), (5, site_name), (6, fa_code), (7, site_ids), (8, "NSB")]))
            row_writes.append((NSB_ROW_MAP["iwm_details"], bool(iwm_details.strip()), [(3, iwm_details)]))
            row_writes.append((NSB_ROW_MAP["pre_configuration"], True, [(3, "NSB")]))
            row_writes.append((NSB_ROW_MAP["current_configuration"], bool(current_config.strip()), [(3, current_config)] if current_config.strip() else []))
            row_writes.append((NSB_ROW_MAP["post_configuration"], True, [(3, post_line)]))
            # Confirmed same fix as N2E: additional Pending point (not a replacement)
            # when any node is missing from Post-checks entirely.
            row_writes.append((NSB_ROW_MAP["post_configuration_pending"]["pending"][0], bool(missing_nodes),
                               [(3, f"{post_line}(MIC PM)")] if missing_nodes else []))
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
            mcl.write_buffer_2col_with_overflow(row_writes, int_rows, int_pairs if int_checked else [])
            _rw(NSB_ROW_MAP["controller_integration"]["completed"][0], ctrl_checked and controller_id, [(3, controller_id)] if ctrl_checked else None)
            _rw(NSB_ROW_MAP["dss_activation"]["completed"][0], dss_completed, [(3, dss_bands)] if dss_completed else None)
            _rw(NSB_ROW_MAP["ngs_activation"]["completed"][0], ngs_completed, [(3, ngs_bands), (4, ngs_node)] if ngs_completed else None)
            _rw(NSB_ROW_MAP["gps_installation"]["completed"][0], gps_completed_line,
                [(3, "|".join(enabled_nodes)), (5, gtype)] if gps_completed_line else None)
            lkf_c_value = lkf_completed.replace("LKF Installation:", "").strip() if lkf_completed else None
            _rw(NSB_ROW_MAP["lkf_installation"]["completed"][0], lkf_completed, [(3, lkf_c_value)] if lkf_c_value else None)
            # Confirmed fix: Partially Completed's manual per-item inputs now also flow
            # into these dedicated rows, not just the text report — previously these 8
            # writes were strictly gated to ct_status == "Completed"/"Pending" only, so
            # anything entered under Partially Completed never reached the .xlsm at all.
            psap_write_completed = (ct_status == "Completed" and psap_line) or \
                (ct_status == "Partially Completed" and ct_completed_inputs.get("psap", "").strip())
            _rw(NSB_ROW_MAP["psap_speedtest"]["completed"][0], psap_write_completed,
                ([(3, "/".join(lte_bands_all) if ct_status == "Completed" else ct_completed_inputs.get("psap", "").strip())]
                 + ([(5, psap_sched_id.strip())] if psap_sched_id.strip() else [])) if psap_write_completed else None)
            speed_lte_write_completed = (ct_status == "Completed" and speed_lte_line) or \
                (ct_status == "Partially Completed" and ct_completed_inputs.get("speed_lte", "").strip())
            _rw(NSB_ROW_MAP["speedtest_lte"]["completed"][0], speed_lte_write_completed,
                [(3, "/".join(lte_bands_all) if ct_status == "Completed" else ct_completed_inputs.get("speed_lte", "").strip())] if speed_lte_write_completed else None)
            speed_5g_write_completed = (ct_status == "Completed" and speed_5g_line) or \
                (ct_status == "Partially Completed" and ct_completed_inputs.get("speed_5g", "").strip())
            _rw(NSB_ROW_MAP["speed_test_5g"]["completed"][0], speed_5g_write_completed,
                [(3, "/".join(fiveg_bands_all) if ct_status == "Completed" else ct_completed_inputs.get("speed_5g", "").strip())] if speed_5g_write_completed else None)
            fnet_write_completed = (ct_status == "Completed" and fnet_line) or \
                (ct_status == "Partially Completed" and ct_completed_inputs.get("fnet", "").strip())
            _rw(NSB_ROW_MAP["calltest_fnet"]["completed"][0], fnet_write_completed,
                [(3, fnet_line if ct_status == "Completed" else ct_completed_inputs.get("fnet", "").strip())] if fnet_write_completed else None)
            sfp_rows = NSB_ROW_MAP["transport_sfp"]["completed"]
            sfp_pairs = [(node, f"{bbu} (BBU End) & {siad} (SIAD End)") for node, bbu, siad in sfp_completed_lines]
            mcl.write_buffer_2col_with_overflow(row_writes, sfp_rows, sfp_pairs)
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
            dss_stakeholder = dss_sh if dss_sh and dss_sh != "\u2014 Select \u2014" else "AT&T"
            _rw(NSB_ROW_MAP["dss_activation"]["pending"][0], dss_pending,
                [(3, " & ".join(mcl.sort_bands_lte_first(dss_pending_bands_combined))), (6, dss_stakeholder)] if dss_pending_bands_combined else None)
            _rw(NSB_ROW_MAP["ngs_activation"]["pending"][0], ngs_pending, [(3, ngs_bands), (4, ngs_node)] if ngs_pending else None)
            _rw(NSB_ROW_MAP["gps_installation"]["pending"][0], gps_pending_line, [(3, "|".join(disabled_nodes))] if gps_pending_line else None)
            lkf_p_value = (lkf_pending.replace("LKF Installation:", "").replace("(MIC)", "").strip()
                           if lkf_pending else None)
            _rw(NSB_ROW_MAP["lkf_installation"]["pending"][0], bool(lkf_pending), [(3, lkf_p_value)] if lkf_p_value else None)
            psap_write_pending = (ct_status == "Pending" and psap_line) or \
                (ct_status == "Partially Completed" and ct_pending_inputs.get("psap", "").strip())
            _rw(NSB_ROW_MAP["psap_speedtest"]["pending"][0], psap_write_pending,
                [(3, "/".join(lte_bands_all) if ct_status == "Pending" else ct_pending_inputs.get("psap", "").strip()), (6, "MIC PM")] if psap_write_pending else None)
            speed_lte_write_pending = (ct_status == "Pending" and speed_lte_line) or \
                (ct_status == "Partially Completed" and ct_pending_inputs.get("speed_lte", "").strip())
            _rw(NSB_ROW_MAP["speedtest_lte"]["pending"][0], speed_lte_write_pending,
                [(3, "/".join(lte_bands_all) if ct_status == "Pending" else ct_pending_inputs.get("speed_lte", "").strip()), (6, "MIC PM")] if speed_lte_write_pending else None)
            speed_5g_write_pending = (ct_status == "Pending" and speed_5g_line) or \
                (ct_status == "Partially Completed" and ct_pending_inputs.get("speed_5g", "").strip())
            _rw(NSB_ROW_MAP["speed_test_5g"]["pending"][0], speed_5g_write_pending,
                [(3, "/".join(fiveg_bands_all) if ct_status == "Pending" else ct_pending_inputs.get("speed_5g", "").strip()), (6, "MIC PM")] if speed_5g_write_pending else None)
            fnet_write_pending = (ct_status == "Pending" and fnet_line) or \
                (ct_status == "Partially Completed" and ct_pending_inputs.get("fnet", "").strip())
            _rw(NSB_ROW_MAP["calltest_fnet"]["pending"][0], fnet_write_pending,
                [(3, fnet_line if ct_status == "Pending" else ct_pending_inputs.get("fnet", "").strip()), (6, "MIC PM")] if fnet_write_pending else None)
            # Confirmed removed: SFP Installation, Rilinks Scripting, SIAD provisioning,
            # Link failure, SFP Not Present, Mo Inconsistent config alarm, Fiberloss,
            # High/Low RSSI, High/Low VSWR, VSWR overthreshold — all purely manual, no
            # auto-detection, left unchecked here for the engineer to fill in directly
            # in the downloaded macro.
            _rw(NSB_ROW_MAP["sfp_installation_bbu"]["pending"][0], False)
            _rw(NSB_ROW_MAP["sfp_installation_radio"]["pending"][0], False)
            for row_num in NSB_ROW_MAP["transport_sfp"]["pending"]:
                _rw(row_num, False)
            _rw(NSB_ROW_MAP["ret_configuration"]["pending"][0], ret_pending)
            _rw(NSB_ROW_MAP["external_alarm_scripting"]["pending"][0], cascade_fires or bool(alarm_ports_report_lines),
                [(3, controller_id)] if cascade_fires else None)
            _rw(NSB_ROW_MAP["sau_connections"]["pending"][0], cascade_fires or sau_pending, [(3, controller_id)] if (cascade_fires or sau_pending) else None)
            _rw(NSB_ROW_MAP["sup_connections"]["pending"][0], sup_pending_lines, [(3, "|".join(s.split(":")[-1].strip() for s in sup_pending_lines))] if sup_pending_lines else None)
            _rw(NSB_ROW_MAP["xmu_installation"]["pending"][0], xmu_pending_lines, [(3, "|".join(s.split(":")[-1].strip() for s in xmu_pending_lines))] if xmu_pending_lines else None)
            _rw(NSB_ROW_MAP["rilinks_scripting"]["pending"][0], False)
            _rw(NSB_ROW_MAP["idl_connections"]["pending"][0], idl_pending)
            _rw(NSB_ROW_MAP["script_load_6673"]["pending"][0], False)
            _rw(NSB_ROW_MAP["siad_provisioning"]["pending"][0], False)
            _rw(NSB_ROW_MAP["area_test"]["pending"][0], cascade_fires or area_pending,
                [(3, "|".join(integrated_nodes)), (4, "Failed")] if area_pending else None)
            _rw(NSB_ROW_MAP["external_alarm_testing"]["pending"][0], cascade_fires or testing_pending, [(3, controller_id)] if (cascade_fires or testing_pending) else None)
            _rw(NSB_ROW_MAP["config_6673"]["pending"][0], has_6673, [(3, sidehaul_rows[0]["switch_id"])] if has_6673 and sidehaul_rows else None)
            _rw(NSB_ROW_MAP["port_config_6673_enm"]["pending"][0], has_6673, [(4, sidehaul_rows[0]["switch_id"])] if has_6673 and sidehaul_rows else None)
            _rw(NSB_ROW_MAP["link_failure"]["pending"][0], False)
            _rw(NSB_ROW_MAP["sfp_not_present"]["pending"][0], False)
            _rw(NSB_ROW_MAP["mo_inconsistent_config_alarm"]["pending"][0], False)
            fiberloss_rows = NSB_ROW_MAP["fiberloss"]["pending"]
            _rw(fiberloss_rows[0], False)
            _rw(fiberloss_rows[1], False)
            _rw(NSB_ROW_MAP["high_rssi"]["pending"][0], False)
            _rw(NSB_ROW_MAP["low_rssi"]["pending"][0], False)
            _rw(NSB_ROW_MAP["high_vswr"]["pending"][0], False)
            _rw(NSB_ROW_MAP["low_vswr"]["pending"][0], False)
            _rw(NSB_ROW_MAP["vswr_overthreshold"]["pending"][0], False)

            pending_buffer_lines = list(nsb_pending_from_warnings) + ([additional_pending] if additional_pending.strip() else [])
            mcl.write_buffer_with_overflow(row_writes, NSB_ROW_MAP["additional_pending"], pending_buffer_lines)

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
            mcl.write_buffer_with_overflow(row_writes, NSB_ROW_MAP["notes_buffer"], other_notes)

            xlsm_bytes = fill_legacy_mca_surgical(NSB_TEMPLATE_PATH, row_writes)
            st.download_button("Download filled checklist (.xlsm)", xlsm_bytes,
                                file_name=f"{node_tag}_NSB_Filled.xlsm", key="nsb_dl_xlsm")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"{node_tag}_NSB_Report.txt", report_text)
                zf.writestr(f"{node_tag}_NSB_Filled.xlsm", xlsm_bytes)
            st.download_button("Download both (report + filled checklist, .zip)", zip_buffer.getvalue(),
                                file_name=f"{node_tag}_NSB_Bundle.zip", key="nsb_dl_zip")
