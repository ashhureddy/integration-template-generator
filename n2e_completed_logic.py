"""
N2E (Nokia-to-Ericsson) completed/pending logic. Confirmed genuinely different in kind
from MCA (vendor cutover, not carrier-add/retune) — every function here reflects an
explicitly confirmed rule from the N2E design conversation, not an assumption carried
over from MCA. Reuses mca_completed_logic's generic helpers (band_label via qx,
format_ports_with_slogans, controller-checks parsing, Sidehaul Info) directly wherever
the underlying mechanism is confirmed identical.
"""

import re
import mca_completed_logic as mcl

qx = None


def set_app_module(app_module):
    global qx
    qx = app_module
    mcl.set_app_module(app_module)


# ============================================================
# INTEGRATION — reuses app.py's own generate_n2e() Carrier ADD logic directly (every CIQ
# cell counts as an addition, no Pre-checks comparison at all — confirmed).
# ============================================================

def integration_bands_and_nodes(classification):
    """Returns (band_list_str, node_list_str) exactly matching the confirmed real example
    format: 'LTE_700/AWS_1/AWS_2/PCS_1/FNET/LTE_700_E/5G_850/CBAND' and the node(s)."""
    all_bands = set()
    all_nodes = []
    for node, cells in classification.get("added", {}).items():
        all_nodes.append(node)
        for cell in cells:
            label, _sector = qx.band_label(cell)
            if label:
                all_bands.add(label)
    return "/".join(sorted(all_bands)), "|".join(all_nodes)


def integration_bands_by_tech(classification):
    """Splits Integration's detected bands into LTE-only and 5G-only subsets — confirmed
    this is what PSAP/Speedtest and Speed test reuse directly, no separate detection and
    no market lookup at all for N2E (confirmed: report for ALL markets)."""
    lte_bands, fiveg_bands = set(), set()
    for node, cells in classification.get("added", {}).items():
        for cell in cells:
            label, _sector = qx.band_label(cell)
            if not label:
                continue
            if label.startswith("5G_") or label in ("CBAND", "DOD", "DOD_BWE"):
                fiveg_bands.add(label)
            else:
                lte_bands.add(label)
    return "/".join(sorted(lte_bands)), "/".join(sorted(fiveg_bands))


# ============================================================
# GPS INSTALLATION — confirmed: use the SAME MCA convention (node(s) then Version), not
# the reversed format seen in one real example.
# ============================================================

def gps_installation_line(nodes, gps_type):
    """Confirmed format: 'GPS Installation: {nodes}  Version: {type}' — same as MCA,
    explicitly NOT the reversed 'Version: type: node' format seen in one example."""
    return f"GPS Installation: {'|'.join(nodes)}  Version: {gps_type}"


def gps_sync_status(mm_objs, post_sync_status):
    """Confirmed: 'GPS is enabled' = Post-checks Sync Status 2 -> TimeSyncIO=ENABLED,
    applies per-node (both nodes if 2 present). Returns (enabled_nodes, disabled_nodes)."""
    enabled, disabled = [], []
    for row in mm_objs:
        node = row.get("Node to be built as")
        state = post_sync_status.get(node)
        if state == "ENABLED":
            enabled.append(node)
        elif state == "DISABLED":
            disabled.append(node)
    return enabled, disabled


# ============================================================
# SUP / XMU — confirmed real logic, verified against real ALL00640 Post-checks data
# (SUP-1 ... ENABLED, XMU03-1-1 ... ENABLED both present in the real Hardware Status table).
# ============================================================

def xmu_in_ciq(post_configuration_text):
    """Confirmed trigger: XMU appears in the CIQ target (Post Configuration string)."""
    return mcl.xmu_in_ciq(post_configuration_text)


def sup_connections_state(post_text, xmu_present_in_ciq):
    """Confirmed: SUP Connections only applies when XMU is present in the CIQ target;
    then checks Post-checks Hardware Status for SUP's own operational state."""
    if not xmu_present_in_ciq:
        return {}
    return mcl._hardware_component_state(post_text, "SUP")


def xmu_installation_state(post_text, xmu_present_in_ciq):
    """Confirmed: XMU Installation requires XMU present in BOTH the CIQ target AND
    Post-checks Hardware Status."""
    if not xmu_present_in_ciq:
        return {}
    return mcl._hardware_component_state(post_text, "XMU")


# ============================================================
# SA CONVERSION — confirmed: reuses app.py's own check_sa_conversion (CIQ NR_SA tab
# presence) directly, Completed-only per the real example ("SA conversion.\t{node}").
# ============================================================

def sa_conversion_nodes(ciq_wb, mm_objs):
    return mcl.sa_conversion_nodes(ciq_wb, mm_objs)


def sa_conversion_note(sa_nodes):
    return mcl.sa_conversion_note(sa_nodes)


# ============================================================
# 6610 CASCADE — confirmed SIMPLER than MCA: fires only when the controller-checks file
# was never uploaded at all (not the broader "uploaded but doesn't confirm scripting"
# condition MCA uses).
# ============================================================

def controller_cascade_fires(controller_present_and_edp_published, controller_checks_uploaded):
    return bool(controller_present_and_edp_published and not controller_checks_uploaded)


# ============================================================
# ACTIVE EXTERNAL ALARM — confirmed new item, per-port, reuses the same slogan-annotated
# port format already built for MCA's locked-port buckets.
# ============================================================

def active_external_alarm_lines(controller_checks_data, controller_id):
    """Confirmed real format: 'Active external alarm on Port {N}:({SLOGAN}) :{Controller
    ID}. (Tower Crew)' — confirmed real bug fixed: this used to fire on ANY port with a
    slogan (i.e., merely scripted/configured), which is wrong — confirmed against the
    real PDF table that a separate 'activeExternalAlarm' column (true/false) is the
    actual signal for a currently-firing alarm, completely independent of whether the
    port is scripted or locked."""
    lines = []
    for p in controller_checks_data.get("alarm_ports", []):
        if p.get("active") and p["slogan"]:
            lines.append(f"Active external alarm on Port {p['port']}:({p['slogan']}) :{controller_id}. (Tower Crew)")
    return lines


# ============================================================
# AREA TEST — confirmed: new-node trigger, manual Area Lite Pass/Fail choice determines
# Completed vs Pending placement.
# ============================================================

def area_test_line(nodes, area_lite_result):
    """area_lite_result: 'Passed' or 'Failed'. Passed -> Completed, Failed -> Pending
    (confirmed via the real filled example: 'Area test: DVON090779: Area Lite - Failed')."""
    node_str = "|".join(nodes)
    return f"Area test: {node_str}: Area Lite - {area_lite_result}"


# ============================================================
# 6673 CONFIGURATION / 6673 PORT CONFIGURATION IN ENM — confirmed: same "6673 present"
# trigger as 6673 Script load (Sidehaul Info Switch column), but these two ALWAYS go to
# Pending, never Completed.
# ============================================================

def has_6673(sidehaul_rows):
    return any(str(r.get("switch_type", "")).strip() == "6673" for r in sidehaul_rows)


# ============================================================
# LOCKED ALARM PORTS — N2E-specific rules, confirmed from the real "6610 Alarm Cutover
# Process & Reporting Standards" reference doc's dedicated N2E tab. Genuinely different
# from MCA's Legacy tab: no equivalent to MCA's bucket 1, different Owner tags, a new
# "Power Plant Swap" scenario that produces BOTH a Pending line and a Note together.
# ============================================================

def n2e_locked_port_active_alarms(ports, port_slogan_map=None):
    """Pre-existing Active Alarms -> Note ONLY (confirmed: not Pending/Pre-Existing at
    all for N2E). Owner is plain 'AT&T', not 'AT&T PM/OPS' like MCA's equivalent."""
    if not ports:
        return None
    ports_fmt = mcl.format_ports_with_slogans(ports, port_slogan_map or {})
    return f"Pre\u2011existing active alarms on ports {ports_fmt} are kept locked to avoid OPS tickets.(Owner:AT&T)"


def n2e_locked_port_power_plant_swap(swap_performed, ports="", port_slogan_map=None, alarm_ports_data=None):
    """Power Plant Swap — confirmed new N2E-only scenario, produces BOTH a Pending line
    and a Note line together. Returns (pending_text, note_text, active_ports_str).
    Confirmed real fix: swap_performed is now an INDEPENDENT signal (a checkbox, not
    tied to typing port numbers) — if there are no active alarms at all, the engineer
    would have nothing to type in the port field, but the Note must still fire whenever
    a Power Plant Swap genuinely happened. The Pending line still only includes ports
    genuinely marked activeExternalAlarm=true, and only appears if any exist."""
    if not swap_performed:
        return None, None, None
    port_active_map = {p["port"]: bool(p.get("active")) for p in (alarm_ports_data or [])}
    port_list = [p.strip() for p in (ports or "").split(",") if p.strip()]
    active_ports = [p for p in port_list if port_active_map.get(p)]
    active_ports_str = ",".join(active_ports) if active_ports else None

    pending = None
    if active_ports_str:
        ports_fmt = mcl.format_ports_with_slogans(active_ports_str, port_slogan_map or {})
        pending = f"Post external alarm cutover, active alarms observed on ports {ports_fmt} are kept locked (Owner: Tower crew)."
    note = "External alarms have been scripted according to the new power plant configuration."
    return pending, note, active_ports_str


def n2e_locked_port_not_cleared_by_fe(ports, port_slogan_map=None):
    """Post-Cutover Alarms Not Cleared by FE -> Pending ONLY. Confirmed matches MCA's
    bucket 4 concept, same Owner (Tower crew)."""
    if not ports:
        return None
    ports_fmt = mcl.format_ports_with_slogans(ports, port_slogan_map or {})
    return f"Post external alarm cutover, active alarms observed on ports {ports_fmt} have been kept locked (Owner: Tower crew)."


def _format_ports_ampersand(ports_str, port_slogan_map):
    """Same slogan-annotation as mcl.format_ports_with_slogans, but joins with '&' before
    the last port instead of 'and' — confirmed exact format for the merged Power Plant
    Swap / Not Cleared by FE line, matching the reference doc's own style."""
    if not ports_str:
        return ""
    port_list = [p.strip() for p in ports_str.split(",") if p.strip()]
    annotated = [f"{p}[{port_slogan_map[p]}]" if p in port_slogan_map else p for p in port_list]
    if len(annotated) == 1:
        return annotated[0]
    if len(annotated) == 2:
        return f"{annotated[0]} & {annotated[1]}"
    return ", ".join(annotated[:-1]) + f" & {annotated[-1]}"


def n2e_merged_post_cutover_pending(power_plant_ports, not_cleared_ports, port_slogan_map=None):
    """Confirmed merge: Power Plant Swap and Post-Cutover Not Cleared by FE ports combine
    into ONE Pending line (both use nearly identical wording and the same Owner), rather
    than two separate lines. Confirmed exact format: 'Post external alarm cutover, active
    alarms observed on ports 5,10 & 6 have been kept locked (Owner: Tower crew).' — '&'
    joining, 'have been kept locked' wording, slogans included per port."""
    all_ports = ",".join(p for p in (power_plant_ports, not_cleared_ports) if p)
    if not all_ports:
        return None
    ports_fmt = _format_ports_ampersand(all_ports, port_slogan_map or {})
    return f"Post external alarm cutover, active alarms observed on ports {ports_fmt} have been kept locked (Owner: Tower crew)."


def ignore_state_alarm_notes(entries_str):
    """New N2E category: pre-existing external alarms found in 'Ignore' state.
    Confirmed manual entry — port + slogan typed directly by the engineer (e.g.
    '3(RBS Temp High)'), since Ignore-state ports aren't necessarily captured by the
    standard activeExternalAlarm/LOCKED detection at all. Supports multiple entries,
    comma-separated (e.g. '3(RBS Temp High), 7(RBS INTRUSION)'), one Note line each.
    Confirmed exact format: 'Pre - existing Port {N} ({slogan}) is configured in SMM but
    is currently set to 'Ignore' state and will not be migrated to the 6610.(Owner:AT&T)'"""
    if not entries_str:
        return []
    notes = []
    for entry in entries_str.split(","):
        entry = entry.strip()
        if not entry:
            continue
        m = re.match(r'(\d+)\s*\((.+?)\)', entry)
        if not m:
            continue
        port, slogan = m.groups()
        notes.append(f"Pre - existing Port {port} ({slogan}) is configured in SMM but is "
                     f"currently set to 'Ignore' state and will not be migrated to the "
                     f"6610.(Owner:AT&T)")
    return notes


def n2e_transport_sfp_warnings(ciq_wb, mm_objs, post_text, transport_sfp_data):
    """Confirmed N2E-specific check (different wording from MCA's transport_sfp_verification):
    checks TXdBm/RXdBm against the speed-appropriate range (reusing mcl.SFP_RANGES) and
    BER against 0/0, for every node. Confirmed exact wording:
    'High/low RXdBm/TXdBm on Transport SFP: {node}.',
    'BER not reporting on the Transport port: {node}.' (BER empty), and
    'BER NZ reporting on the Transport port: {node}.' (BER present but non-zero).
    All fire independently (a node can trigger any combination, or none).
    Confirmed: all go into the Warnings tab AND get reported as Pending to MIC PM
    (via the buffer, since there's no dedicated template row for this).
    Returns (warning_texts: list[str], pending_lines: list[str])."""
    warning_texts, pending_lines = [], []
    for row in mm_objs:
        node = row.get("Node to be built as")
        gen = qx.get_node_generation(ciq_wb, row)
        if not gen:
            continue
        port_labels = qx.PORT_BY_GEN.get(gen)
        if not port_labels:
            continue
        opmode = qx.extract_transport_fiber_opmode(post_text, node, port_labels)
        if not opmode:
            continue
        speed = "10G" if "10G" in opmode.upper() else ("1G" if "1G" in opmode.upper() else None)
        if not speed:
            continue
        lo, hi = mcl.SFP_RANGES[speed]

        reading = transport_sfp_data.get(node)
        if not reading:
            continue

        out_of_range = (reading["txdbm"] > hi or reading["txdbm"] < lo
                         or reading["rxdbm"] > hi or reading["rxdbm"] < lo)
        if out_of_range:
            text = f"High/low RXdBm/TXdBm on Transport SFP: {node}."
            warning_texts.append(text)
            pending_lines.append(f"{text} (MIC PM)")

        ber = reading["ber"]
        if not ber or ber.strip() == "":
            text = f"BER not reporting on the Transport port: {node}."
            warning_texts.append(text)
            pending_lines.append(f"{text} (MIC PM)")
        elif ber.strip() != "0/0":
            text = f"BER NZ reporting on the Transport port: {node}."
            warning_texts.append(text)
            pending_lines.append(f"{text} (MIC PM)")

    return warning_texts, pending_lines


def sa_conversion_amf_warning(post_text, sa_conversion_nodes_list):
    """Confirmed check: for nodes with SA Conversion present, verify each
    TermPointToAmf entry's admin state. Confirmed real format:
    '{Node} GNBCUCPFunction=1,TermPointToAmf={amf_name} {UNLOCKED|LOCKED} {OpState}'.
    If ANY TermPointToAmf entry for the node is LOCKED (not all — any single one is
    enough), fires the warning. Confirmed exact wording:
    'TermpointtoAmf is in locked state on {node}, please unlock.'
    Returns list of warning texts (one per affected node, deduplicated)."""
    if not sa_conversion_nodes_list or not post_text:
        return []
    pattern = re.compile(r'(\S+) GNBCUCPFunction=1,TermPointToAmf=(\S+) (UNLOCKED|LOCKED) (\w+)')
    locked_nodes = set()
    for m in pattern.finditer(post_text):
        node, _amf_name, admin, _oper = m.groups()
        if node in sa_conversion_nodes_list and admin == "LOCKED":
            locked_nodes.add(node)
    return [f"TermpointtoAmf is in locked state on {node}, please unlock." for node in sorted(locked_nodes)]


# ============================================================
# LTE/5G SECTOR PARAMETER VERIFICATION — confirmed CIQ column mapping from the design
# conversation, cross-checked against real WAL94133 data. Compares Post-checks' actual
# on-air values against the CIQ's intended target values, per cell.
# ============================================================

_LTE_CELL_ROW_RE = re.compile(
    r'(\S+) (LOCKED|UNLOCKED) (\d+ \S+) (BARRED|UNBARRED) (\d+) (\d+) (\d+) '
    r'(ENABLED|DISABLED) (\d+) (true|false) (\S+) (\d+) (\d+)')


def extract_lte_cell_status(post_text):
    """Parses the real 'LTE FDD Cell Status Information' table. Confirmed real header:
    'Cells adminState availabilityStatus cellBarred dlChannelBandwidth earfcndl earfcnul
    OpState PCI PLMNStatus sectorCarrierRef tac ulChannelBandwidth' — availabilityStatus
    is a two-token value (e.g. '3 OFF_LINE'), confirmed by column-count cross-check
    against real data. Returns {cell: {field: value}}."""
    out = {}
    for m in _LTE_CELL_ROW_RE.finditer(post_text or ""):
        (cell, _admin, _avail, _barred, dlbw, earfcndl, earfcnul,
         _opstate, pci, _plmn, sector, tac, ulbw) = m.groups()
        out[cell] = {
            "dlChannelBandwidth": dlbw, "earfcndl": earfcndl, "earfcnul": earfcnul,
            "PCI": pci, "sectorCarrierRef": sector, "tac": tac, "ulChannelBandwidth": ulbw,
        }
    return out


def lte_sector_param_warnings(ciq_wb, mm_objs, post_text):
    """Confirmed CIQ mapping (cross-checked against real data, corrected from the
    original ask: PCI compares against CIQ's own PCI column, not cellId):
    dlChannelBandwidth->dlChannelBandwidth, earfcndl->earfcnDl, earfcnul->earfcnUl,
    PCI->PCI, sectorCarrierRef->sectorId, ulChannelBandwidth->ulChannelBandwidth (all in
    'eUtran Parameters'); tac->tac (in 'eNB Info', matched via eNBId, same value for
    every cell under that eNB). Returns list of warning texts, one per mismatched field."""
    warnings = []
    if not post_text or "eUtran Parameters" not in ciq_wb.sheetnames:
        return warnings
    post_cells = extract_lte_cell_status(post_text)
    if not post_cells:
        return warnings

    ciq_rows = {r.get("EutranCellFDDId"): r for r in qx.sheet_objs(ciq_wb["eUtran Parameters"])
                if r.get("EutranCellFDDId")}
    field_map = [
        ("dlChannelBandwidth", "dlChannelBandwidth"), ("earfcndl", "earfcnDl"),
        ("earfcnul", "earfcnUl"), ("PCI", "PCI"), ("sectorCarrierRef", "sectorId"),
        ("ulChannelBandwidth", "ulChannelBandwidth"),
    ]
    for cell, post_vals in post_cells.items():
        ciq_row = ciq_rows.get(cell)
        if not ciq_row:
            continue
        for post_key, ciq_key in field_map:
            post_val = str(post_vals.get(post_key, "")).strip()
            ciq_val = str(ciq_row.get(ciq_key, "")).strip()
            if ciq_val and post_val:
                ciq_parts = [p.strip() for p in ciq_val.split("/")]
                if post_val not in ciq_parts:
                    warnings.append(f"{post_key} mismatch on {cell}: Post-checks={post_val}, CIQ={ciq_val}.")

    # tac — matched via eNBId, same value expected for every cell under that eNB.
    if "eNB Info" in ciq_wb.sheetnames:
        enb_tac = {r.get("eNBId"): r.get("tac") for r in qx.sheet_objs(ciq_wb["eNB Info"]) if r.get("eNBId")}
        eutran_rows = qx.sheet_objs(ciq_wb["eUtran Parameters"])
        cell_to_enbid = {r.get("EutranCellFDDId"): r.get("eNBId") for r in eutran_rows if r.get("EutranCellFDDId")}
        checked_enbids = set()
        for cell, post_vals in post_cells.items():
            enbid = cell_to_enbid.get(cell)
            if not enbid or enbid in checked_enbids:
                continue
            checked_enbids.add(enbid)
            ciq_tac = str(enb_tac.get(enbid, "")).strip()
            post_tac = str(post_vals.get("tac", "")).strip()
            if ciq_tac and post_tac and post_tac != ciq_tac:
                warnings.append(f"tac mismatch on eNBId {enbid}: Post-checks={post_tac}, CIQ={ciq_tac}.")
    return warnings


def extract_5g_cell_du_status(post_text):
    """Parses '5G NR Cell DU Status' table. Confirmed real complication: nCI is
    sometimes empty (variable-length row), so pure positional splitting is unreliable —
    uses the fact that nRSectorCarrierRef always equals the cell name itself as an
    anchor, confirmed against real data (both cellLocalId and nRPCI verified to match
    CIQ exactly with this approach). Returns {cell: {field: value}}."""
    out = {}
    for line in (post_text or "").splitlines():
        tokens = line.split()
        if len(tokens) < 10 or tokens[1] not in ("LOCKED", "UNLOCKED"):
            continue
        cell = tokens[0]
        try:
            anchor_idx = tokens.index(cell, 1)
        except ValueError:
            continue
        before = tokens[1:anchor_idx]
        after = tokens[anchor_idx + 1:]
        if len(before) < 7 or len(after) < 3:
            continue
        local_id, cell_range = before[2], before[3]
        nrpci = before[-1]
        nrtac = after[0]
        out[cell] = {"cellLocalId": local_id, "cellRange": cell_range, "nRPCI": nrpci, "nRTAC": nrtac}
    return out


def extract_5g_sector_carrier(post_text):
    """Parses '5G NR Sector Carrier' table. Confirmed real header:
    'nrSectorCarrier adminState arfcnDL arfcnUL bSChannelBwDL bSChannelBwUL
    configuredMaxTxPower opState txDirection' — fixed structure, no variable-length
    fields. Returns {cell: {field: value}}."""
    out = {}
    pattern = re.compile(
        r'(\S+) (LOCKED|UNLOCKED) (\d+) (\d+) (\d+) (\d+) (\d+) (ENABLED|DISABLED) (\S+)')
    for m in pattern.finditer(post_text or ""):
        cell, _admin, arfcndl, arfcnul, bwdl, bwul, maxtxpower, _opstate, _txdir = m.groups()
        out[cell] = {
            "arfcnDL": arfcndl, "arfcnUL": arfcnul, "bSChannelBwDL": bwdl,
            "bSChannelBwUL": bwul, "configuredMaxTxPower": maxtxpower,
        }
    return out


def extract_ssb_frequency(post_text):
    """Parses the real 'NRCellDU={cell} ssbFrequency {value}' lines. Returns
    {cell: ssbFrequency}."""
    out = {}
    for m in re.finditer(r'NRCellDU=(\S+) ssbFrequency (\d+)', post_text or ""):
        cell, val = m.groups()
        out[cell] = val
    return out


def fiveg_sector_param_warnings(ciq_wb, mm_objs, post_text):
    """Confirmed CIQ mapping: cellLocalId, CellRange, nRPCI, arfcnDL, arfcnUL,
    bSChannelBwDL, bSChannelBwUL, configuredMaxTxPower, ssbFrequency all compared
    directly against '5G Info' (matched by NRCellDU); nrTAC compared against 'NR_SA'
    (matched by node name, same value expected for every cell on that node — same
    per-node pattern as LTE's tac/eNBId check)."""
    warnings = []
    if not post_text or "5G Info" not in ciq_wb.sheetnames:
        return warnings

    cell_du = extract_5g_cell_du_status(post_text)
    sector_carrier = extract_5g_sector_carrier(post_text)
    ssb_freq = extract_ssb_frequency(post_text)
    if not cell_du:
        return warnings

    ciq_rows = {r.get("NRCellDU"): r for r in qx.sheet_objs(ciq_wb["5G Info"]) if r.get("NRCellDU")}

    for cell, post_vals in cell_du.items():
        ciq_row = ciq_rows.get(cell)
        if not ciq_row:
            continue
        combined = dict(post_vals)
        combined.update(sector_carrier.get(cell, {}))
        if cell in ssb_freq:
            combined["ssbFrequency"] = ssb_freq[cell]

        field_map = [
            ("cellLocalId", "cellLocalId"), ("cellRange", "CellRange"), ("nRPCI", "nRPCI"),
            ("arfcnDL", "arfcnDL"), ("arfcnUL", "arfcnUL"), ("bSChannelBwDL", "bSChannelBwDL"),
            ("bSChannelBwUL", "bSChannelBwUL"), ("configuredMaxTxPower", "configuredMaxTxPower"),
            ("ssbFrequency", "ssbFrequency"),
        ]
        for post_key, ciq_key in field_map:
            post_val = str(combined.get(post_key, "")).strip()
            ciq_val = str(ciq_row.get(ciq_key, "")).strip()
            if not ciq_val or not post_val:
                continue
            ciq_parts = [p.strip() for p in ciq_val.split("/")]
            if post_val not in ciq_parts:
                warnings.append(f"{post_key} mismatch on {cell}: Post-checks={post_val}, CIQ={ciq_val}.")

    # nrTAC — matched via node name in NR_SA, same value expected for every cell on that node.
    if "NR_SA" in ciq_wb.sheetnames:
        nrsa_tac = {r.get("Node Name"): r.get("nrTAC") for r in qx.sheet_objs(ciq_wb["NR_SA"]) if r.get("Node Name")}
        checked_nodes = set()
        for cell, post_vals in cell_du.items():
            node = cell.rsplit("_N", 1)[0] if "_N" in cell else None
            if not node or node in checked_nodes:
                continue
            checked_nodes.add(node)
            ciq_tac = str(nrsa_tac.get(node, "")).strip()
            post_tac = str(post_vals.get("nRTAC", "")).strip()
            if ciq_tac and post_tac and post_tac != ciq_tac:
                warnings.append(f"nRTAC mismatch on node {node}: Post-checks={post_tac}, CIQ={ciq_tac}.")
    return warnings
