"""
N2E (Nokia-to-Ericsson) completed/pending logic. Confirmed genuinely different in kind
from MCA (vendor cutover, not carrier-add/retune) — every function here reflects an
explicitly confirmed rule from the N2E design conversation, not an assumption carried
over from MCA. Reuses mca_completed_logic's generic helpers (band_label via qx,
format_ports_with_slogans, controller-checks parsing, Sidehaul Info) directly wherever
the underlying mechanism is confirmed identical.
"""

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

def _hardware_component_state(post_text, component_prefix):
    """Generic Hardware Status Information row parser: '{Node} {Component} {admin} {fault}
    {steady} {oper} {description}...' — confirmed real format. Returns
    {node: oper_state} for any component whose name starts with component_prefix
    (e.g. 'SUP' matches 'SUP-1', 'XMU' matches 'XMU03-1-1')."""
    import re
    out = {}
    text = mcl._normalize(post_text) if hasattr(mcl, "_normalize") else post_text
    for m in re.finditer(
            r'(\S+) (' + component_prefix + r'\S*) (UNLOCKED|LOCKED) (ON|OFF) (\S+) (ENABLED|DISABLED)', text):
        node, _comp, _admin, _fault, _steady, oper = m.groups()
        out[node] = oper
    return out


def xmu_in_ciq(post_configuration_text):
    """Confirmed trigger: XMU appears in the CIQ target (Post Configuration string)."""
    return "XMU" in (post_configuration_text or "")


def sup_connections_state(post_text, xmu_present_in_ciq):
    """Confirmed: SUP Connections only applies when XMU is present in the CIQ target;
    then checks Post-checks Hardware Status for SUP's own operational state."""
    if not xmu_present_in_ciq:
        return {}
    return _hardware_component_state(post_text, "SUP")


def xmu_installation_state(post_text, xmu_present_in_ciq):
    """Confirmed: XMU Installation requires XMU present in BOTH the CIQ target AND
    Post-checks Hardware Status."""
    if not xmu_present_in_ciq:
        return {}
    return _hardware_component_state(post_text, "XMU")


# ============================================================
# SA CONVERSION — confirmed: reuses app.py's own check_sa_conversion (CIQ NR_SA tab
# presence) directly, Completed-only per the real example ("SA conversion.\t{node}").
# ============================================================

def sa_conversion_nodes(ciq_wb, mm_objs):
    return [row.get("Node to be built as") for row in mm_objs
            if qx.check_sa_conversion(ciq_wb, row.get("Node to be built as"))]


def sa_conversion_note(sa_nodes):
    """Confirmed Notes addition when SA Conversion is detected."""
    if not sa_nodes:
        return None
    return "Termpointtoamf is in unlocked state."


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
    ID}. (Tower Crew)' — one line per port with a real slogan, regardless of lock state
    (unlike the locked-port buckets, this is about any port carrying a live alarm)."""
    lines = []
    for p in controller_checks_data.get("alarm_ports", []):
        if p["slogan"]:
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


def n2e_locked_port_loops_bridge_clips(ports, port_slogan_map=None):
    """Pre-Existing Loops and Bridge Clips -> Note ONLY, a single combined note (loops
    line + active-alarms line together) — confirmed real format from the N2E tab."""
    if not ports:
        return None
    plain_ports_fmt = mcl.format_ports_with_slogans(ports, {})
    ports_fmt = mcl.format_ports_with_slogans(ports, port_slogan_map or {})
    return (f"Pre\u2011existing loops have been removed from alarm ports {plain_ports_fmt}.\n"
            f"Active alarms observed on ports {ports_fmt} are kept locked (Owner: AT&T).")


def n2e_locked_port_power_plant_swap(ports, port_slogan_map=None):
    """Power Plant Swap — confirmed new N2E-only scenario, produces BOTH a Pending line
    and a Note line together. Returns (pending_text, note_text)."""
    if not ports:
        return None, None
    ports_fmt = mcl.format_ports_with_slogans(ports, port_slogan_map or {})
    pending = f"Post external alarm cutover, active alarms observed on ports {ports_fmt} are kept locked (Owner: Tower crew)."
    note = "External alarms have been scripted according to the new power plant configuration."
    return pending, note


def n2e_locked_port_not_cleared_by_fe(ports, port_slogan_map=None):
    """Post-Cutover Alarms Not Cleared by FE -> Pending ONLY. Confirmed matches MCA's
    bucket 4 concept, same Owner (Tower crew)."""
    if not ports:
        return None
    ports_fmt = mcl.format_ports_with_slogans(ports, port_slogan_map or {})
    return f"Post external alarm cutover, active alarms observed on ports {ports_fmt} have been kept locked (Owner: Tower crew)."
