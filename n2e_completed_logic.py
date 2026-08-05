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
    return "/".join(mcl.sort_bands_lte_first(all_bands)), "|".join(all_nodes)


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
    comma-separated (e.g. '3(RBS Temp High), 7(RBS INTRUSION)').
    Confirmed real correction: all entries combine into ONE single Note line (not one
    line per port) — 'Port' becomes 'Ports' when there's more than one, e.g.:
    'Pre - existing Ports 10 (SMM IGNORE), 8 (RBS HEX FAIL) is configured in SMM but is
    currently set to 'Ignore' state and will not be migrated to the 6610.(Owner:AT&T)'
    Returns a list with either 0 or 1 note (kept as a list for compatibility with how
    callers merge it in)."""
    if not entries_str:
        return []
    parsed = []
    for entry in entries_str.split(","):
        entry = entry.strip()
        if not entry:
            continue
        m = re.match(r'(\d+)\s*\((.+?)\)', entry)
        if not m:
            continue
        port, slogan = m.groups()
        parsed.append(f"{port} ({slogan})")
    if not parsed:
        return []
    label = "Port" if len(parsed) == 1 else "Ports"
    combined = ", ".join(parsed)
    return [f"Pre - existing {label} {combined} is configured in SMM but is "
            f"currently set to 'Ignore' state and will not be migrated to the "
            f"6610.(Owner:AT&T)"]
