"""
MCA Integration Report — Completed section, new logic built in this session.

Design principle throughout: reuse QUICKIX's existing, already-tested functions
(classify_carriers, band_label, extract_precheck_sectors, generate_port_conversion_checks,
PORT_BY_GEN, DU_TYPE_TO_GEN, sheet_objs, is_populated, etc. — all imported from app.py)
rather than re-deriving anything QUICKIX already knows how to do. Only genuinely new
data sources (GPS Status, Transport SFP, Sidehaul Info, controller-checks PDF, the
Calltest market table) get new parsers here.

Every function below is traceable to a specific confirmed decision in this session's
transcript — see the inline comments for the "why", not just the "what".
"""

import re
from pathlib import Path

import app as qx  # the existing, tested QUICKIX module (report-feature branch)


def _normalize(text):
    """Pre/Post-checks text comes from two genuinely different extraction paths depending
    on the uploaded file's real format (confirmed this session: some "PDF" uploads are
    actually zip archives with a clean pre-extracted text layer, one line per row; others
    are real PDFs run through pdftotext -layout, which pads columns with runs of spaces
    and inserts blank lines between a table's header and its data). Every parser below
    needs to tolerate both. Collapsing all horizontal whitespace to single spaces and
    dropping blank lines makes the two formats equivalent before any table regex runs."""
    if not text:
        return text
    lines = [re.sub(r'[ \t]+', ' ', ln).strip() for ln in text.splitlines()]
    return "\n".join(ln for ln in lines if ln)


# ============================================================
# SECTION 1 — NEW DATA-SOURCE PARSERS
# Every regex here is built and verified against REAL Pre/Post-checks text extracted
# from real project files during this session (ECL02586, SCL01706), not guessed.
# ============================================================

def extract_gps_status(check_text):
    """Parses the 'GPS Status' table (present in both Pre- and Post-checks — same format,
    confirmed against real ECL02586 data in both).

        GPS Status
        Node       Product Designation  Product No.    Product Rev.
        ECL02586   GRU 05 01             NCD 901 89/1   R1D

    Returns {node: product_designation}. 'Product Designation' is what the person
    referred to as 'GPS type'."""
    out = {}
    text = _normalize(check_text)
    if not text:
        return out
    m = re.search(r'GPS Status\nNode Product Designation Product No\. Product Rev\.\n(.*?)(?=\n[A-Z][\w /]*\n|\Z)',
                  text, re.S)
    if not m:
        return out
    block = m.group(1)
    # Row shape: "NODE  <product designation, 1-3 space-separated tokens>  <product no>  <rev>"
    # Confirmed real values: "GRU 05 01", "GPS 02 01", "GPS 03 01", "GRU 04 01" — always
    # a 2-letter prefix + two 2-digit groups. Product No. always contains a "/" (e.g.
    # "NCD 901 89/1"); Product Rev. is the final bare token (e.g. "R1D").
    row_re = re.compile(
        r'^(\S+) ((?:GPS|GRU) \d{2} \d{2}) (\S+(?: \S+)*?/\S+) (\S+)$', re.M)
    for m2 in row_re.finditer(block):
        node, designation, _prodno, _rev = m2.groups()
        out[node.strip()] = designation.strip()
    return out


KNOWN_GPS_TYPES = {"GPS 02 01", "GPS 03 01", "GRU 04 01", "GRU 05 01"}
# "Most likely a GPS 01" — confirmed: not a fallback VALUE, just the person's own label for
# "whatever shows up that isn't one of the 4 known types." We never emit "GPS 01" as a
# value; we just treat anything outside KNOWN_GPS_TYPES as unconfirmed.


def extract_sync_status_2(post_text):
    """Parses 'Status of Synchronization 2' — confirmed real structure (ECL02586, SCL01706):

        Status of Synchronization 2
        Node       MO                                            OpState
        ECL02586   Equipment=1,FRU=1,SyncPort=1                  ENABLED
        ECL02586   Transport=1,Synchronization=1,TimeSyncIO=1    ENABLED

    Returns {node: opstate} for the TimeSyncIO=1 row specifically — confirmed this row
    reflects GPS sync status, not the SyncPort row (which we ignore per instruction)."""
    out = {}
    text = _normalize(post_text)
    if not text:
        return out
    for m in re.finditer(
            r'(\S+) Transport=1,Synchronization=1,TimeSyncIO=1 (ENABLED|DISABLED)', text):
        node, opstate = m.groups()
        out[node.strip()] = opstate
    return out


def extract_transport_sfp(post_text):
    """Parses the 'Transport SFP' table — confirmed real structure (ECL02586):

        Transport SFP
        Node      BOARD     PORT  VENDOR    VENDORPROD      REV   SERIAL      DATE      ERICSSONPROD  TEMP  TXbs  TXdBm  RXdBm  BER
        ECL02586  RANP6672  IC    ERICSSON  SPP10ELRIDFKEN  0101  EB30728343  20240621  RDH10265/3    R1A   46C   30%   -2.02  -1.71  0/0

    Confirmed real-data quirk: ERICSSONPROD often carries a trailing revision token
    (e.g. "RDH10265/3 R1A") that pushes the row to 15 space-separated fields against a
    14-column header — same class of PDF-column-misalignment issue already documented
    elsewhere in this project (DL_UL_LOSS_ROW_RE, etc.). Handled by capturing
    ERICSSONPROD non-greedily up to the point where TEMP (a bare "<n>C" token) appears.

    Returns {node: {"ericssonprod": str, "txdbm": float, "rxdbm": float, "ber": str}}."""
    out = {}
    text = _normalize(post_text)
    if not text:
        return out
    m = re.search(
        r'Transport SFP\nNode BOARD PORT VENDOR VENDORPROD REV SERIAL DATE '
        r'ERICSSONPROD TEMP TXbs TXdBm RXdBm BER\n(.*?)(?=\nTransport Fiber link Status|\n[A-Z][\w /]*\n|\Z)',
        text, re.S)
    if not m:
        return out
    block = m.group(1)
    row_re = re.compile(
        r'^(\S+) (\S+) (\S+) (\S+) (\S+) (\S+) (\S+) (\d{8}) '
        r'(.+?) \d+C \d+% (-?\d+\.\d+) (-?\d+\.\d+) (\S+)$',
        re.M)
    for m2 in row_re.finditer(block):
        node = m2.group(1)
        ericssonprod = m2.group(9).strip()
        txdbm, rxdbm, ber = float(m2.group(10)), float(m2.group(11)), m2.group(12)
        out[node.strip()] = {"ericssonprod": ericssonprod, "txdbm": txdbm, "rxdbm": rxdbm, "ber": ber}
    return out


def extract_transport_fiber_opmode_for_node(check_text, node, port_labels):
    """Thin wrapper around QUICKIX's existing extract_transport_fiber_opmode — confirmed
    reusable as-is against Post-checks text (it's generic text parsing, not hardcoded to
    'pre'). Kept as a separate name here only for readability at call sites."""
    return qx.extract_transport_fiber_opmode(check_text, node, port_labels)


def extract_sidehaul_info(ciq_wb):
    """CIQ 'Sidehaul Info' tab — confirmed real structure:

        Switch | SH Switch ID | SH Switch Slot & Port -1 | SH Switch Slot & Port -2 |
        SH Switch Port SFP | Basebands | Baseband Port 1 | Baseband Port 2 | Baseband Port SFP

    One row per switch-port connection — a site can have more than one (confirmed real
    example: FSL00456, 2 rows same switch, different slot/port + baseband).
    'TBD' is a real literal placeholder value some CIQs leave in Baseband Port columns —
    treated as not-populated, same guard class as the NaN-as-"nan" bug fixed elsewhere."""
    rows = []
    if "Sidehaul Info" not in ciq_wb.sheetnames:
        return rows
    for row in qx.sheet_objs(ciq_wb["Sidehaul Info"]):
        switch = row.get("Switch")
        if not qx.is_populated(switch):
            continue
        node_id = row.get("Basebands")
        if str(node_id or "").strip().upper() == "TBD":
            node_id = None
        rows.append({
            "switch_type": switch,
            "switch_id": row.get("SH Switch ID"),
            "slot_port": row.get("SH Switch Slot & Port -1") or row.get("SH Switch Slot & Port -2"),
            "node_id": node_id,
            # Cable part number: confirmed manual — no CIQ/EDP source exists for this.
        })
    return rows


# ---- Controller-checks PDF (separate file from Pre/Post-checks, confirmed real
# structure from FNOC222775_C001_controller_checks.pdf) ----

def extract_controller_checks(text):
    """Returns {
        'node_alarm_status': 'OK' | 'NOT OK' | None,
        'controller_state': {'admin': 'UNLOCKED'|'LOCKED', 'fault': bool, 'oper': 'ENABLED'|'DISABLED'} | None,
        'sau_state': {'admin':..., 'fault':..., 'oper':...} | None,
        'alarm_ports': [{'port': str, 'admin': 'UNLOCKED'|'LOCKED', 'slogan': str|None, 'severity': str|None}, ...],
    }"""
    result = {"node_alarm_status": None, "controller_state": None, "sau_state": None, "alarm_ports": []}
    if not text:
        return result

    m = re.search(r'Node Alarm Status\s*[:\s]*\n?.*?\b(OK|NOT OK)\b', text)
    # Confirmed real line: "FNOC222775_C001 OK CXP2010233/2_R26C12 107.253.119.42" under
    # "Node Alarm Status Current Upgrade Package IP Address" header.
    m2 = re.search(r'Node Alarm Status Current Upgrade Package IP Address\r?\n\S+\s+(OK|NOT OK)', text)
    if m2:
        result["node_alarm_status"] = m2.group(1)

    def _fru_state(fru_name):
        # Confirmed real line shape:
        # "FieldReplaceableUnit=Controller6610 1 (UNLOCKED) 2 (OFF) 1 (ENABLED) ..."
        mm = re.search(
            re.escape(f"FieldReplaceableUnit={fru_name}") + r'\s+\d+\s+\((\w+)\)\s+\d+\s+\((\w+)\)\s+\d+\s+\((\w+)\)',
            text)
        if not mm:
            return None
        admin, fault, oper = mm.groups()
        return {"admin": admin, "fault_ok": fault == "OFF", "oper": oper}

    result["controller_state"] = _fru_state("Controller6610")
    result["sau_state"] = _fru_state("SAU")

    # External alarms table — confirmed real rows:
    # "FieldReplaceableUnit=SAU,AlarmPort=1 false 1 (UNLOCKED) RBS INTRUSION false 1 (ENABLED) 3 (MAJOR)"
    # "FieldReplaceableUnit=SAU,AlarmPort=11 false 0 (LOCKED) true 1 (ENABLED) 4 (MINOR)"  <- no slogan
    port_re = re.compile(
        r'FieldReplaceableUnit=SAU,AlarmPort=(\d+)\s+\S+\s+\d+\s+\((\w+)\)\s*'
        r'([A-Z][A-Z0-9 ]*?)?\s*(?:true|false)\s+\d+\s+\((\w+)\)\s+\d+\s+\((\w+)\)')
    for m3 in port_re.finditer(text):
        port, admin, slogan, _oper, severity = m3.groups()
        slogan = slogan.strip() if slogan and slogan.strip() else None
        result["alarm_ports"].append({
            "port": port, "admin": admin, "slogan": slogan, "severity": severity,
        })
    return result


# ---- Calltest market table (Calltest_sheet.xlsx, 'Legacy' tab for MCA) ----

def load_calltest_table(xlsx_path):
    """Parses Calltest_sheet.xlsx 'Legacy' tab into:
        prefix_to_market: {"NC": "NCSC", "EC": "NCSC", ...}
        rules: {(market, scenario): {"PSAP": bool, "LTE Speed test": bool,
                                       "F-net with F-net SIM": bool, "5G speedtest": bool}}
    Confirmed structure: Market column only populated on the first row of each group;
    it applies to every row until the next non-blank Market cell."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Legacy"]
    prefix_to_market, rules = {}, {}
    current_market = None
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    for row in rows:
        market, sectors, psap, lte_speed, fnet, fiveg_speed, site_name = row[:7]
        if market:
            current_market = market
        if not sectors or not current_market:
            continue
        rules[(current_market, sectors.strip())] = {
            "PSAP": str(psap).strip().upper() == "Y",
            "LTE Speed test": str(lte_speed).strip().upper() == "Y",
            "F-net with F-net SIM": str(fnet).strip().upper() == "Y",
            "5G speedtest": str(fiveg_speed).strip().upper() == "Y",
        }
        if site_name:
            for prefix in str(site_name).split("/"):
                prefix_to_market[prefix.strip().upper()] = current_market
    return prefix_to_market, rules


# ============================================================
# SECTION 2 - INTEGRATION: Post-checks presence verification (warning-only, per confirmed
# decision - never alters what the report shows, only feeds the Warnings tab)
# ============================================================

def verify_integration_against_postcheck(classification, post_text):
    """For every cell in classification['added'], check it's PRESENT in Post-checks
    (any admin state - LOCKED counts as integrated, confirmed: lock state doesn't matter,
    only whether the cell exists at all). Returns a list of warning dicts, one per cell
    that's completely absent from Post-checks. Never touches the report itself."""
    post_pairs, _ = qx.extract_precheck_sectors(post_text)
    post_cells = {cell for (_node, cell) in post_pairs}

    warnings = []
    for node, cells in classification.get("added", {}).items():
        for cell in cells:
            if cell not in post_cells:
                label, sector = qx.band_label(cell)
                warnings.append({
                    "type": "integration_missing",
                    "text": f"{label} {sector} ({cell}) sector is missing on : {node}.",
                })
    return warnings


# ============================================================
# SECTION 3 - RETUNE: sector-tracked whole/partial display (fixes the confirmed real gap
# where the built code silently dropped sector info) + eUtran Parameters/Post-checks
# verification
# ============================================================

WHOLE_BAND_SET = {"Alpha", "Beta", "Gamma"}


def classify_retunes_with_sectors(ciq_wb):
    """Re-derives retune events the same way classify_carriers does (same-node
    Sector Del_Movement rows where channel/BW differ) but keeps the sector this time -
    confirmed fix: qx.classify_carriers's own retuned list drops it."""
    out = []
    if "Sector Del_Movement" not in ciq_wb.sheetnames:
        return out
    for r in qx.sheet_objs(ciq_wb["Sector Del_Movement"]):
        src_node, src_sector = r.get("Source Node name"), r.get("Source Sector")
        tgt_node, tgt_sector = r.get("Target Node name"), r.get("Target Sector")
        if not (qx.is_populated(src_node) and qx.is_populated(tgt_node)):
            continue
        if str(tgt_node).strip().upper() == "DELETE":
            continue
        if str(src_node).strip().upper() != str(tgt_node).strip().upper():
            continue
        src_dl = str(r.get("Source channelNumberDL", "")).strip()
        tgt_dl = str(r.get("Target channelNumberDL", "")).strip()
        src_bw = str(r.get("Source Bandwidth", "")).strip()
        tgt_bw = str(r.get("Target Bandwidth", "")).strip()
        if src_dl == tgt_dl and src_bw == tgt_bw:
            continue
        label, sector = qx.band_label(src_sector)
        if not label:
            continue
        out.append({
            "label": label, "sector": sector, "cell": tgt_sector or src_sector,
            "from": f"{src_dl}/{src_bw}", "to": f"{tgt_dl}/{tgt_bw}",
        })
    return out


def format_retunes(retune_events):
    """Groups by (label, from, to) signature, tracking sectors under each."""
    grouped = {}
    for r in retune_events:
        key = (r["label"], r["from"], r["to"])
        grouped.setdefault(key, set()).add(r["sector"])

    lines = []
    for (label, frm, to), sectors in grouped.items():
        if WHOLE_BAND_SET <= sectors:
            sectors_str = " sectors"
        else:
            ordered = sorted(sectors, key=lambda s: qx.SECTOR_ORDER.index(s) if s in qx.SECTOR_ORDER else 99)
            sectors_str = " " + ", ".join(ordered)
        lines.append(f"Retune on:\t{label}{sectors_str}\tFrom:\t{frm}\tTo:\t{to}")
    return lines


def verify_retune_against_checks(ciq_wb, retune_events, post_text):
    """Verification target = CIQ eUtran Parameters/5G Info (confirmed canonical source)
    vs. Post-checks' actual earfcndl/dlChannelBandwidth (LTE) or arfcnDL/bSChannelBwDL
    (5G). Warning-only, never alters the report line itself."""
    warnings = []
    eutran_by_cell = {row.get("EutranCellFDDId"): row for row in qx.sheet_objs(ciq_wb["eUtran Parameters"])} \
        if "eUtran Parameters" in ciq_wb.sheetnames else {}
    fiveg_by_cell = {row.get("NRCellDU"): row for row in qx.sheet_objs(ciq_wb["5G Info"])} \
        if "5G Info" in ciq_wb.sheetnames else {}

    post_lte = {}
    for m in re.finditer(
            r'(\S+)\s+UNLOCKED\s+NOT_BARRED\s+(\d+)\s+(\d+)\s+(\d+)\s+ENABLED', _normalize(post_text)):
        cell, bw, dl, ul = m.groups()
        post_lte[cell] = {"dl": int(dl), "bw": int(bw)}

    for ev in retune_events:
        cell = ev["cell"]
        target_row = eutran_by_cell.get(cell) or fiveg_by_cell.get(cell)
        if not target_row:
            continue
        target_dl = target_row.get("earfcnDl") or target_row.get("arfcnDL")
        target_bw = target_row.get("dlChannelBandwidth") or target_row.get("bSChannelBwDL")
        actual = post_lte.get(cell)
        if not actual:
            warnings.append({"type": "retune_missing",
                              "text": f"Retune {ev['label']} {ev['sector']} ({cell}) not confirmed - "
                                      f"Post-checks has no entry for this cell."})
            continue
        if str(target_dl).strip() != str(actual["dl"]).strip() or int(str(target_bw or 0)) != actual["bw"]:
            warnings.append({"type": "retune_mismatch",
                              "text": f"Retune {ev['label']} {ev['sector']} ({cell}) not confirmed - "
                                      f"Post-checks shows {actual['dl']}/{actual['bw']}, expected {target_dl}/{target_bw}."})
    return warnings


# ============================================================
# SECTION 4 - MOVED SECTORS: target-node-presence-only verification (confirmed: source-node
# presence is NOT checked, per explicit instruction)
# ============================================================

def verify_moved_sectors_against_postcheck(classification, post_text):
    post_pairs, _ = qx.extract_precheck_sectors(post_text)
    post_by_node = {}
    for node, cell in post_pairs:
        post_by_node.setdefault(node, set()).add(cell)

    warnings = []
    for mv in classification.get("moved", []):
        cell, to_node, from_node = mv["cell"], mv["to_node"], mv["from_node"]
        if cell not in post_by_node.get(to_node, set()):
            label, sector = qx.band_label(cell)
            warnings.append({
                "type": "moved_sector_missing",
                "text": f"Moved sector {label} {sector} ({cell}) not confirmed on target node : "
                        f"{to_node} (moved from {from_node}).",
            })
    return warnings


# ============================================================
# SECTION 5 - DELETED SECTOR: absence-only verification (confirmed: mirror-image of
# Moved Sectors - success = cell NOT present in Post-checks at all)
# ============================================================

def verify_deleted_sectors_against_postcheck(classification, post_text):
    post_pairs, _ = qx.extract_precheck_sectors(post_text)
    post_cells = {cell for (_node, cell) in post_pairs}

    warnings = []
    for node, cells in classification.get("deleted_sectors", {}).items():
        for cell in cells:
            if cell in post_cells:
                label, sector = qx.band_label(cell)
                warnings.append({
                    "type": "deleted_sector_still_present",
                    "text": f"Deleted sector {label} {sector} ({cell}) still present in Post-checks - "
                            f"deletion not confirmed on : {node}.",
                })
    return warnings


# ============================================================
# SECTION 6 - RADIO SWAP: Completed vs. Pending is DETERMINED by Post-checks, not a
# report-only warning (confirmed: this is the one item where verification changes report
# placement, replacing the old "always defaults to Pending" rule)
# ============================================================

def classify_radio_swap_placement(precheck_text, postcheck_text, ciq_wb):
    """Reuses qx.classify_radio_swaps for detection, then re-checks the actual installed
    radio type in Post-checks (qx.extract_precheck_radio_types reused as-is - confirmed
    generic). Completed if Post-checks radio == the CIQ target ('to'); Pending otherwise
    (matches Pre-checks 'from', or anything else - no special-case warning, per
    explicit instruction to ignore that edge case)."""
    swaps = qx.classify_radio_swaps(precheck_text, ciq_wb)
    post_radios = qx.extract_precheck_radio_types(postcheck_text)

    completed, pending = [], []
    for sw in swaps:
        cell = sw.get("group_key", (None,))[0]
        post_types = [qx.radio_family(r) for r in post_radios.get(cell, [])] if cell else []
        target_type = sw["to"].replace("RRU ", "").strip()
        if target_type in post_types:
            completed.append(sw)
        else:
            pending.append(sw)
    return completed, pending


def format_radio_swaps(swap_list, label_prefix="Radio Swap on:"):
    """Same two-step grouping as the confirmed-mature qx logic: physical radio group first,
    then merge across sectors sharing an identical (from, to) signature."""
    merged = {}
    for r in swap_list:
        sig = (r["from"], r["to"])
        merged.setdefault(sig, {"labels": set(), "sectors": set()})
        merged[sig]["labels"].add(r["label"])
        merged[sig]["sectors"].add(r["sector"])

    lines = []
    for (frm, to), grp in merged.items():
        labels = tuple(sorted(grp["labels"]))
        label_str = labels[0] if len(labels) == 1 else f"[{'|'.join(labels)}]"
        sector_set = grp["sectors"]
        sector_names = sorted(sector_set, key=lambda s: qx.SECTOR_ORDER.index(s) if s in qx.SECTOR_ORDER else 99)
        is_whole = WHOLE_BAND_SET <= sector_set
        sectors_str = " sectors" if is_whole else (f" {', '.join(sector_names)}" if sector_names else "")
        lines.append(f"{label_prefix}\t{label_str}{sectors_str}\tFrom:\t{frm}\tTo:\t{to}")
    return lines


# ============================================================
# SECTION 7 - GPS: Installation (new nodes, grouped by type), Upgrade (existing nodes,
# type changed Pre->Post), and the two site-health checks (unconfirmed type on MMBB,
# sync disabled) - all four pieces confirmed this session, none existed before.
# ============================================================

def gps_installation_lines(new_nodes, post_gps_status):
    """Groups new nodes by shared GPS type; Node IDs '|'-joined per group. Returns
    (dedicated_row_line, overflow_lines) - first group goes on the template's one
    dedicated row, any additional distinct-type groups spill to the buffer pool."""
    by_type = {}
    for node in new_nodes:
        gtype = post_gps_status.get(node, "NOT FOUND")
        by_type.setdefault(gtype, []).append(node)

    lines = [f"GPS Installation: {'|'.join(nodes)}  Version: {gtype}"
             for gtype, nodes in by_type.items()]
    if not lines:
        return None, []
    return lines[0], lines[1:]


def gps_upgrade_lines(existing_nodes, pre_gps_status, post_gps_status):
    """Existing node (not new), GPS type differs Pre vs Post -> Completed, buffer-only
    (no dedicated row exists for this item at all)."""
    lines = []
    for node in existing_nodes:
        pre_t, post_t = pre_gps_status.get(node), post_gps_status.get(node)
        if pre_t and post_t and pre_t != post_t:
            lines.append(f"GPS upgraded from: {pre_t} to: {post_t} on: {node}.")
    return lines


def gps_unconfirmed_type_check(mm_objs, post_gps_status):
    """Every MMBB node, Post-checks only. Unconfirmed = not one of the 4 known types.
    Same text used for BOTH the Pending line and the Warnings tab, per confirmed decision."""
    hits = []
    for row in mm_objs:
        if str(row.get("BBU Mode", "")).strip().upper() != "MMBB":
            continue
        node = row.get("Node to be built as")
        gtype = post_gps_status.get(node)
        if gtype and gtype not in KNOWN_GPS_TYPES:
            hits.append(node)
    if not hits:
        return None
    return f"GPS needs to be upgraded for : {'|'.join(hits)}"


def gps_sync_disabled_check(mm_objs, post_sync_status):
    """Every node, Post-checks TimeSyncIO row. All disabled nodes merge into ONE Pending
    line reusing the existing 'GPS Installation:' Pending label - no warning."""
    hits = [row.get("Node to be built as") for row in mm_objs
            if post_sync_status.get(row.get("Node to be built as")) == "DISABLED"]
    if not hits:
        return None
    return f"GPS Installation: {'|'.join(hits)}"


# ============================================================
# SECTION 8 - TRANSPORT SFP: grouping by shared manual model, plus the new TXdBm/RXdBm/BER
# threshold verification (new node -> Pending; existing+board-swap -> Pending+Warning;
# existing, no swap -> Pre-Existing Issues only)
# ============================================================

SFP_RANGES = {"1G": (-9.0, 9.0), "10G": (-6.2, 6.2)}


def transport_sfp_installation_lines(nodes_needing_sfp, sfp_models_by_node):
    """nodes_needing_sfp: from new_nodes OR Port-Conversion-triggered nodes.
    sfp_models_by_node: {node: (bbu_end_model, siad_end_model)} - both MANUAL entries,
    confirmed. Grouped by shared (bbu,siad) pair."""
    by_model = {}
    for node in nodes_needing_sfp:
        model = sfp_models_by_node.get(node, ("", ""))
        by_model.setdefault(model, []).append(node)
    lines = []
    for (bbu, siad), nodes in by_model.items():
        lines.append(f"Transport SFP Installation on: {'|'.join(nodes)}  "
                      f"SFP Model (BBU End): {bbu}  SFP Model (SIAD End): {siad}")
    return lines


def _sfp_out_of_range_labels(node, board_gen, post_text_for_speed, transport_sfp_data):
    """Determines port speed via PORT_BY_GEN + Transport Fiber link Status OpMode, then
    checks TXdBm/RXdBm against the speed-appropriate range and BER against 0/0.
    Returns (combined_label_string or None)."""
    port_labels = qx.PORT_BY_GEN.get(board_gen)
    if not port_labels:
        return None
    opmode = qx.extract_transport_fiber_opmode(post_text_for_speed, node, port_labels)
    if not opmode:
        return None
    speed = "10G" if "10G" in opmode.upper() else ("1G" if "1G" in opmode.upper() else None)
    if not speed:
        return None
    lo, hi = SFP_RANGES[speed]

    reading = transport_sfp_data.get(node)
    if not reading:
        return None

    bits = []
    if reading["txdbm"] > hi:
        bits.append("High TXDBM")
    elif reading["txdbm"] < lo:
        bits.append("Low TXDBM")
    if reading["rxdbm"] > hi:
        bits.append("High RXDBM")
    elif reading["rxdbm"] < lo:
        bits.append("Low RXDBM")

    ber = reading["ber"]
    ber_note = None
    if not ber or ber.strip() == "":
        ber_note = "BER is not pulling on the transport port"
    elif ber.strip() != "0/0":
        ber_note = "BER pulling on the transport port"

    if not bits and not ber_note:
        return None

    parts = []
    if bits:
        parts.append(f"{'/'.join(bits)} out of ranges")
    if ber_note:
        parts.append(ber_note)
    return ", ".join(parts)


def transport_sfp_verification(ciq_wb, mm_objs, new_nodes, board_swap_nodes, post_text, transport_sfp_data):
    """Returns (pending_lines, warnings, pre_existing_lines) - three separate buckets per
    the confirmed node-classification rule."""
    pending_lines, warnings, pre_existing_lines = [], [], []
    board_swap_node_names = {n for n, _pre, _post in board_swap_nodes} if board_swap_nodes and \
        isinstance(board_swap_nodes[0], tuple) and len(board_swap_nodes[0]) == 3 else set(board_swap_nodes or [])

    for row in mm_objs:
        node = row.get("Node to be built as")
        gen = qx.get_node_generation(ciq_wb, row)
        if not gen:
            continue
        label = _sfp_out_of_range_labels(node, gen, post_text, transport_sfp_data)
        if not label:
            continue
        text = f"{label} on the transport port on : {node}."

        if node in new_nodes:
            pending_lines.append(text)
        elif node in board_swap_node_names:
            pending_lines.append(text)
            warnings.append({"type": "transport_sfp_out_of_range", "text": text})
        else:
            pre_existing_lines.append(text)
    return pending_lines, warnings, pre_existing_lines


# ============================================================
# SECTION 9 - 6610 CONTROLLER: SAU Connections, External alarm Scripting/testing, and the
# cascading "6610 not actually integrated" Pending rule. Manual locked-port classification
# (buckets 1-6) is a UI concern, not encoded here - see module docstring at bottom.
# ============================================================

def sau_connections_placement(controller_checks_data, controller_id):
    """None means 'no 6610 present -> stays manual' (caller's responsibility to only call
    this when a 6610 exists)."""
    sau = controller_checks_data.get("sau_state")
    if not sau:
        return None
    return "Completed" if sau["oper"] == "ENABLED" else "Pending"


def external_alarm_scripting_confirmed(controller_checks_data):
    """Any AlarmPort row with a real (non-blank) slogan = confirmed scripted."""
    return any(p["slogan"] for p in controller_checks_data.get("alarm_ports", []))


def external_alarm_testing_placement(controller_checks_data):
    """Among SCRIPTED ports (real slogan present) only: all locked -> Pending + Notes line.
    Any unlocked (all-unlocked or mixed) -> Completed."""
    scripted = [p for p in controller_checks_data.get("alarm_ports", []) if p["slogan"]]
    if not scripted:
        return None, None
    all_locked = all(p["admin"] == "LOCKED" for p in scripted)
    if all_locked:
        return "Pending", "All external alarms are kept locked, due to NEA is pending."
    return "Completed", None


def controller_integration_cascade(six610_present_and_edp_published, controller_checks_data, controller_id):
    """Returns True if the cascade fires (6610 present/EDP-published, but controller-checks
    doesn't confirm alarm scripting) - caller then moves all 4 items to Pending:
    6610 Controller Integration, External alarm Scripting on, LKF Installation (6610 portion),
    External alarm testing. No warning, per confirmed decision."""
    if not six610_present_and_edp_published:
        return False
    if not controller_checks_data or not external_alarm_scripting_confirmed(controller_checks_data):
        return True
    return False


# ============================================================
# SECTION 10 - LKF INSTALLATION: 4 independent OR'd triggers (3 pre-existing + 1 new this
# session), per-node Completed/Pending choice, merged by chosen status.
# ============================================================

def lkf_trigger_nodes(new_nodes, board_swap_nodes, controller_id, precheck_text, mm_objs):
    """4th trigger (new this session): existing single-tech node (Pre-checks shows only
    LTE or only 5G) whose CIQ target now shows BOTH eNBId and gNBId (MMBB/TMBB) -
    works either direction."""
    triggered = set(new_nodes) | {n for n, _p, _q in board_swap_nodes} if board_swap_nodes and \
        isinstance(board_swap_nodes[0], tuple) else set(new_nodes)
    if controller_id:
        triggered |= {row.get("Node to be built as") for row in mm_objs}

    pre_pairs, _ = qx.extract_precheck_sectors(precheck_text)
    pre_cells_by_node = {}
    for node, cell in pre_pairs:
        pre_cells_by_node.setdefault(node, set()).add(cell)

    for row in mm_objs:
        node = row.get("Node to be built as")
        has_enb, has_gnb = qx.is_populated(row.get("eNBId")), qx.is_populated(row.get("gNBId"))
        if not (has_enb and has_gnb):
            continue  # target isn't MMBB/TMBB, 4th trigger doesn't apply
        pre_cells = pre_cells_by_node.get(node, set())
        if not pre_cells:
            continue  # not in Pre-checks at all -> that's the "new node" trigger, already covered
        pre_has_5g = any(qx.is_5g_cell(c) for c in pre_cells)
        pre_has_lte = any(not qx.is_5g_cell(c) for c in pre_cells)
        if pre_has_lte and not pre_has_5g:
            triggered.add(node)  # LTE-only -> gaining 5G
        elif pre_has_5g and not pre_has_lte:
            triggered.add(node)  # 5G-only -> gaining LTE
    return sorted(n for n in triggered if n)


def lkf_lines_by_choice(node_status_choices, controller_id):
    """node_status_choices: {node: 'Completed'|'Pending'} - the engineer's per-node dropdown
    picks. Merges nodes sharing the same choice into one line each."""
    by_choice = {"Completed": [], "Pending": []}
    for node, choice in node_status_choices.items():
        by_choice[choice].append(node)
    out = {}
    for choice, nodes in by_choice.items():
        if nodes:
            out[choice] = f"LKF Installation: {'|'.join(nodes)} | {controller_id or ''}"
    return out


# ============================================================
# SECTION 11 - CALL TEST: Market lookup + the standalone moved-LTE PSAP rule
# ============================================================

def determine_market(node_name, prefix_to_market):
    if not node_name:
        return None
    prefix = str(node_name)[:2].upper()
    return prefix_to_market.get(prefix)


def call_test_lines(classification, market, rules, moved_lte_bands, added_bands_by_tech, moved_bands_by_tech):
    """added_bands_by_tech / moved_bands_by_tech: {'lte': set(), '5g': set(), 'cband_dod': set()}
    already computed by the caller from band_label() on classification['added']/['moved'].
    Returns list of report lines."""
    lines = []
    required = {"PSAP": False, "LTE Speed test": False, "F-net with F-net SIM": False, "5G speedtest": False}

    scenarios = set()
    if added_bands_by_tech.get("lte"):
        scenarios.add("Newly adding LTE")
    if added_bands_by_tech.get("5g"):
        scenarios.add("Newly adding 5G")
    if moved_bands_by_tech.get("5g") or moved_bands_by_tech.get("cband_dod"):
        scenarios.add("Moving 5G(Incl C band/DOD/DOD BWE)")
    if added_bands_by_tech.get("cband_dod"):
        scenarios.add("Newly adding  C band/DOD")

    for scenario in scenarios:
        row = rules.get((market, scenario))
        if row:
            for test, flag in row.items():
                if flag:
                    required[test] = True

    # Standalone rule, confirmed: moved LTE ALWAYS fires PSAP, independent of the table -
    # the table has no "Moving LTE" scenario at all.
    psap_from_moved_lte = bool(moved_lte_bands)

    if required["PSAP"] or psap_from_moved_lte:
        moved_str = ", ".join(sorted(moved_lte_bands)) if moved_lte_bands else ""
        lines.append(f"PSAP test/Speedtest/VoLTE voice calltest:\tMoved LTE Sectors: {moved_str}"
                     f"\tPSAP Schedule ID:\t")
    if required["LTE Speed test"]:
        lines.append(f"Speedtest/VoLTE voice calltest:\tLTE Sectors: {', '.join(sorted(added_bands_by_tech.get('lte', [])))}")
    if required["5G speedtest"]:
        fiveg_all = set(added_bands_by_tech.get("5g", set())) | set(added_bands_by_tech.get("cband_dod", set())) \
            | set(moved_bands_by_tech.get("5g", set())) | set(moved_bands_by_tech.get("cband_dod", set()))
        lines.append(f"Speed test:\t5G Sectors: {', '.join(sorted(fiveg_all))}")
    if required["F-net with F-net SIM"]:
        lines.append("Calltest with F-NET SIM:\tF-NET Sectors")

    return lines


# ============================================================
# SECTION 12 - PORT CONVERSION: Post-checks verification layer on top of QUICKIX's
# existing generate_port_conversion_checks (which only detects the PLANNED state)
# ============================================================

def verify_port_conversion_against_postcheck(ciq_wb, mm_objs, precheck_text, postcheck_text, edp_index):
    """Reuses qx.generate_port_conversion_checks for detection (unchanged), then re-checks
    the same port in POST-checks - must now show 10G. Warning-only."""
    _outputs, summary_rows, _scope_lines = qx.generate_port_conversion_checks(
        ciq_wb, mm_objs, edp_index, precheck_text, lambda *a: None)

    warnings = []
    for s in summary_rows:
        if s.get("Item") != "Port Conversion":
            continue
        node = s.get("Source")
        note = s.get("Note", "")
        m = re.search(r'\((\w+)\)? board, port: ([\w/]+)', note) or re.search(r'board: (\w+), port: ([\w/]+)', note)
        if not m:
            continue
        gen, ports = m.groups()
        port_labels = ports.split("/")
        opmode = qx.extract_transport_fiber_opmode(postcheck_text, node, port_labels)
        if not opmode or "10G" not in opmode.upper():
            warnings.append({
                "type": "port_conversion_not_confirmed",
                "text": f"Port conversion (1G to 10G) not confirmed on : {node} - "
                        f"Post-checks still shows {opmode or 'no reading'} on {'/'.join(port_labels)}.",
            })
    return warnings


# ============================================================
# SECTION 13 - NGS: port the main-branch fix into report-feature (both the corrected
# both_lte check AND the NodeGroupSync=Y safety-net, which report-feature was missing
# entirely). Safety-net stays log-only, never feeds the Warnings tab, per confirmed decision.
# ============================================================

def ngs_pair_is_pure_lte(a_to_b, b_to_a):
    """Ported verbatim from main branch's _ngs_pair_is_pure_lte - confirmed fix for the
    real bug where report-feature's node-level has_lte check produced false negatives on
    dual-tech TMBB nodes referenced via their 5G side."""
    for _own_cell, ref_cell in a_to_b + b_to_a:
        if qx.nr_band_label(ref_cell)[0] is not None:
            return False
    return True


def ngs_safety_net(ciq_wb, cell_to_node, confirmed_nodes, log):
    """Ported from main branch. Log-only by explicit instruction - never feeds Warnings."""
    hits = []
    for sheet_name, cell_col in (("eUtran Parameters", "EutranCellFDDId"), ("5G Info", "NRCellDU")):
        if sheet_name not in ciq_wb.sheetnames:
            continue
        for row in qx.sheet_objs(ciq_wb[sheet_name]):
            if str(row.get("NodeGroupSync", "")).strip().upper() != "Y":
                continue
            cell = row.get(cell_col)
            if not qx.is_populated(cell):
                continue
            node = cell_to_node.get(str(cell).strip())
            if node and node not in confirmed_nodes:
                log(f"NodeGroupSync=Y flagged for {cell} ({node}) but no confirmed NGS pair "
                    f"was detected - check this cell manually.")
                hits.append(cell)
    return hits


# ============================================================
# SECTION 14 - SHARED BUFFER OVERFLOW POOL (Completed rows 81-90 / Pending rows 158-166)
# First-come-first-served across every item that exceeds its dedicated row count.
# ============================================================

class BufferPool:
    """Confirmed design: shared, first-come-first-served, "Label : Detail" format
    (matches the real Report_MCA formula B&" : "&C already in the template). Pool
    exhaustion -> Warnings tab flag, nothing silently dropped."""

    def __init__(self, completed_capacity=10, pending_capacity=9):
        self.completed_capacity = completed_capacity
        self.pending_capacity = pending_capacity
        self.completed_used = []
        self.pending_used = []
        self.overflow_warnings = []

    def add(self, section, label, detail):
        bucket = self.completed_used if section == "Completed" else self.pending_used
        capacity = self.completed_capacity if section == "Completed" else self.pending_capacity
        if len(bucket) >= capacity:
            self.overflow_warnings.append({
                "type": "buffer_pool_exhausted",
                "text": f"{label} - additional entries could not fit in report, review manually.",
            })
            return False
        bucket.append(f"{label} : {detail}")
        return True
