"""
NSB (New Site Build) completed/pending logic. Mirrors n2e_completed_logic.py's structure
exactly — genuinely different rules confirmed through the NSB design conversation live
here, while everything truly shared (band_label via qx, format_ports_with_slogans,
controller-checks parsing, Sidehaul Info, Call Test) is reused directly from
mca_completed_logic.py, never duplicated.
"""

import mca_completed_logic as mcl

qx = None


def set_app_module(app_module):
    global qx
    qx = app_module
    mcl.set_app_module(app_module)


# ============================================================
# EXTERNAL ALARM SCRIPTING/TESTING — confirmed genuinely different from MCA/N2E: no
# manual bucket classification at all. The "some ports locked" case is fully automatic,
# directly listing whichever locked+scripted ports exist.
# ============================================================

def external_alarm_scripting_locked_note(controller_checks_data):
    """Confirmed NSB wording (deliberately different from MCA's "All external alarms are
    kept locked..." — this is "All external alarm ports are kept locked..."). Fires when
    EVERY scripted port is locked."""
    scripted = [p for p in controller_checks_data.get("alarm_ports", []) if p["slogan"]]
    if scripted and all(p["admin"] == "LOCKED" for p in scripted):
        return "All external alarm ports are kept locked, due to NEA is pending."
    return None


def external_alarm_ports_report(controller_checks_data, market=None):
    """New unified NSB check, replacing both external_alarm_scripting_partial_pending
    and active_external_alarms_pending with a single 3-category split, based on the
    2x2 matrix of active/not-active x locked/unlocked (the 4th combination,
    not-active+unlocked, is the normal/healthy state and isn't reported at all):
      1. Active + LOCKED  -> "Active external alarm on ports: {ports} kept locked."
      2. Active + UNLOCKED -> "Active external alarm on: {ports}"
      3. NOT active + LOCKED -> "external alarm ports: {ports} kept locked."
    All three use the same Oxford-comma 'and' slogan format. Returns a list of 0-3
    lines (only the categories that genuinely have ports populate a line). Confirmed
    destination: Pending. Confirmed stakeholder exception: Florida market reports to
    AT&T for all three categories, same as the External alarm testing Florida
    exception — every other market keeps the standard Tower crew stakeholder."""
    scripted = [p for p in controller_checks_data.get("alarm_ports", []) if p["slogan"]]
    if not scripted:
        return []

    stakeholder = "AT&T" if market == "Florida" else "Tower crew"
    active_locked = [p for p in scripted if p.get("active") and p["admin"] == "LOCKED"]
    active_unlocked = [p for p in scripted if p.get("active") and p["admin"] != "LOCKED"]
    inactive_locked = [p for p in scripted if not p.get("active") and p["admin"] == "LOCKED"]

    def _fmt(ports):
        port_slogan_map = {p["port"]: p["slogan"] for p in ports}
        ports_str = ",".join(p["port"] for p in ports)
        return mcl.format_ports_with_slogans(ports_str, port_slogan_map)

    lines = []
    if active_locked:
        lines.append(f"Active external alarm on ports: {_fmt(active_locked)} kept locked. ({stakeholder})")
    if active_unlocked:
        lines.append(f"Active external alarm on: {_fmt(active_unlocked)} ({stakeholder})")
    if inactive_locked:
        lines.append(f"external alarm ports: {_fmt(inactive_locked)} kept locked. ({stakeholder})")
    return lines


# ============================================================
# 6610 CASCADE — confirmed 6 items (one more than MCA/N2E's 5): adds Area test to the
# forced-Pending set. Same simple trigger as N2E (controller-checks file not present).
# ============================================================

def controller_cascade_fires(controller_present_and_edp_published, controller_checks_uploaded):
    return bool(controller_present_and_edp_published and not controller_checks_uploaded)


CASCADE_ITEMS = (
    "6610 Controller Integration",
    "External alarm Scripting on",
    "LKF Installation",
    "External alarm testing",
    "Area test",
    "SAU Connections",
)


# ============================================================
# SUP / XMU — confirmed same mechanism as MCA (auto-detected), reusing the shared
# helpers directly from mca_completed_logic.py. XMU Installation is Completed only when
# BOTH the CIQ and Post-checks confirm it; SUP Connections only applies once XMU is
# present in the CIQ target.
# ============================================================

def xmu_in_ciq(post_configuration_text):
    return mcl.xmu_in_ciq(post_configuration_text)


def sup_connections_state(post_text, sup_expecting_nodes):
    """Confirmed correction, same as N2E: SUP Connections is expected per-node (not
    site-wide) on any node whose CIQ target hardware contains 5216 or XMU."""
    if not sup_expecting_nodes:
        return {}, set()
    found_state = mcl._hardware_component_state(post_text, "SUP")
    result = {node: state for node, state in found_state.items() if node in sup_expecting_nodes}
    missing = sup_expecting_nodes - set(found_state.keys())
    return result, missing


def xmu_installation_state(post_text, xmu_present_in_ciq):
    if not xmu_present_in_ciq:
        return {}
    return mcl._hardware_component_state(post_text, "XMU")
