import streamlit as st
import pandas as pd
import re
import io
import zipfile
from datetime import date
from pathlib import Path

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="Integration Template Generator", page_icon="📡", layout="wide")

TDIR = Path(__file__).parent / "templates" / "MCA"

def resolve_template(exact_name, keyword):
    """Prefer the exact expected filename; if it's missing (e.g. uploaded with a slightly
    different name), fall back to any file in templates/MCA containing `keyword`."""
    exact_path = TDIR / exact_name
    if exact_path.exists():
        return exact_path
    if TDIR.exists():
        candidates = [p for p in TDIR.glob("*.txt") if keyword.lower() in p.name.lower()]
        if candidates:
            return candidates[0]
    return exact_path  # falls through to a clear FileNotFoundError naming what was expected

TPL_MMBB = resolve_template("LTE+5G_MMBB_Integration_Pre-existing_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V11.txt", "MMBB_Integration")
TPL_TMBB = resolve_template("TRIMODE_Integration_Pre-existing_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V10.txt", "TRIMODE_Integration")
TPL_CENM = resolve_template("cENM_TRIMODE_Integration_Pre-existing_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V4.txt", "cENM_TRIMODE")
TPL_6610 = resolve_template("6610 Controller Integration Procedure_25Q3_Updated_V12.txt", "6610")
TPL_CRAN_TRIP1 = resolve_template("CRAN_TO_CRAN_Rehome_Pre-integration_Trip-1_Procedure_for_SA_Sites_V2.txt", "Trip-1")
TPL_CRAN_TRIP2 = resolve_template("CRAN_TO_CRAN_Rehome_and_6673_Sidehaul_Change_With_MPST_Trip-2_Procedure_for_SA_Sites_V1.txt", "Trip-2")
TPL_CRAN_NSA = resolve_template("CRAN_TO_CRAN_Rehome_Integration_and_Cutover_Procedure_for_NSA_Sites_V2.txt", "NSA_Sites")

def resolve_dss_template(exact_stem):
    """stand/standard were uploaded with no .txt extension — try both forms.
    Exact stem match only (no fuzzy 'contains' search) since 'stand' is a substring of
    'standard' and a fuzzy match could silently load the wrong DSS template."""
    for candidate in (TDIR / f"{exact_stem}.txt", TDIR / exact_stem):
        if candidate.exists():
            return candidate
    return TDIR / f"{exact_stem}.txt"

TPL_DSS_4SECTOR = resolve_dss_template("standard")
TPL_DSS_3SECTOR = resolve_dss_template("stand")

# ============================================================
# SHARED HELPERS
# ============================================================

def sheet_objs(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    objs = []
    for r in rows[1:]:
        if not any(str(c).strip() for c in r if c is not None):
            continue
        objs.append({headers[i]: (r[i] if i < len(r) else "") for i in range(len(headers))})
    return objs


def is_populated(v):
    if v is None:
        return False
    s = str(v).strip().upper()
    return s not in ("", "N/A")


def locate_edp_header_row(rows):
    for i, row in enumerate(rows):
        if any(str(c).strip().upper() == "EDP_SITE_ID" for c in row if c is not None):
            return i
    return -1


def build_edp_index(edp_wb):
    for sn in edp_wb.sheetnames:
        ws = edp_wb[sn]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        hidx = locate_edp_header_row(rows)
        if hidx >= 0:
            headers = rows[hidx]
            header_map = {str(h).strip().upper(): i for i, h in enumerate(headers) if h is not None}
            data_rows = [r for r in rows[hidx + 1:] if any(str(c).strip() for c in r if c is not None)]
            return {"header_map": header_map, "data_rows": data_rows}
    return None


def edp_row_for(edp_index, site_name):
    if not edp_index or not site_name:
        return None
    idx = edp_index["header_map"].get("SITE_NAME")
    if idx is None:
        return None
    for r in edp_index["data_rows"]:
        val = r[idx] if idx < len(r) else None
        if str(val or "").strip().upper() == str(site_name).strip().upper():
            return r
    return None


def edp_get(edp_index, row, header_name):
    if not edp_index or row is None:
        return None
    idx = edp_index["header_map"].get(header_name.upper())
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    return None if v is None or str(v).strip() == "" else str(v).strip()


def find_row_by_name(ciq_wb, sheet_name, name_header, name_value):
    if sheet_name not in ciq_wb.sheetnames or not name_value:
        return None
    for r in sheet_objs(ciq_wb[sheet_name]):
        if str(r.get(name_header, "")).strip().upper() == str(name_value).strip().upper():
            return r
    return None


def hw_string(row):
    if not row:
        return None
    du = row.get("DU type") or row.get("1st DU type")
    if not du:
        return None
    xmu_count = sum(1 for k in ("1st XMU", "2nd XMU") if str(row.get(k, "")).strip().upper() == "YES")
    suffix = "" if xmu_count == 0 else " + XMU" if xmu_count == 1 else f" + {xmu_count} XMU"
    return f"{du}{suffix}"


def extract_pre_hw(text, node_name):
    if not text or not node_name:
        return None
    esc = re.escape(node_name)
    m = re.search(esc + r"\s+1\s+UNLOCKED\s+OFF\s+STEADY_ON\s+ENABLED\s+([A-Za-z0-9 ]+?)\s+\d{6,8}", text, re.I)
    if not m:
        return None
    # take the LAST token as the model number — handles "Baseband 6630", "RAN Processor 6651",
    # "Baseband R503", or any future hardware family name, matching the CIQ side's bare model number
    tokens = m.group(1).strip().split()
    return tokens[-1] if tokens else None


def extract_pre_xmu_count(text, node_name):
    if not text or not node_name:
        return 0
    esc = re.escape(node_name)
    return len(re.findall(esc + r"\s+XMU\S*\s+UNLOCKED\s+OFF\s+STEADY_ON\s+ENABLED", text, re.I))


def pre_hw_string(text, node_name):
    base = extract_pre_hw(text, node_name)
    if not base:
        return None
    xmu_count = extract_pre_xmu_count(text, node_name)
    suffix = "" if xmu_count == 0 else " + XMU" if xmu_count == 1 else f" + {xmu_count} XMU"
    return f"{base}{suffix}"


def extract_pdf_text(pdf_bytes):
    import pdfplumber
    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n\n"
    return text


def push_siad_row(rows, edp_index, node_built_as):
    row = edp_row_for(edp_index, node_built_as)
    rows.append({
        "Node": node_built_as,
        "SIAD CLLI": edp_get(edp_index, row, "SIAD_CLLI") or "NOT FOUND",
        "Port Size": edp_get(edp_index, row, "SIAD_PORT_SIZE_BBU") or "NOT FOUND",
        "Port Facing BBU": edp_get(edp_index, row, "SIAD_PORT_FACING_BBU") or "NOT FOUND",
    })


def push_controller_siad_row(rows, edp_index, controller_id):
    """6610 controller rows in EDP use a different column set (ANCEQ_*) than regular BBU nodes —
    same SITE_NAME match, but the port lives in ANCEQ_SIAD_PORT, not SIAD_PORT_FACING_BBU.
    Returns True if the controller was actually found published in EDP, False otherwise."""
    row = edp_row_for(edp_index, controller_id)
    anceq_type = edp_get(edp_index, row, "ANCEQ_TYPE")
    found = row is not None and anceq_type and "6610" in str(anceq_type)
    rows.append({
        "Node": controller_id,
        "SIAD CLLI": edp_get(edp_index, row, "SIAD_CLLI") or "NOT FOUND",
        "Port Size": edp_get(edp_index, row, "SIAD_PORT_SIZE_BBU") or "NOT FOUND",
        "Port Facing BBU": (edp_get(edp_index, row, "ANCEQ_SIAD_PORT") or "NOT FOUND") if found else "EDP not published for controller",
    })
    return found


def highlight_unresolved(text):
    cands = re.findall(r"xx[A-Za-z0-9_]+xx|(?<!#)##[A-Za-z0-9_]+##(?!#)", text)
    return sorted(set(c for c in cands if not re.fullmatch(r"x+", c, re.I)))


def has_6610(controller_objs):
    return any(str(r.get("Controller", "")).strip() == "6610" for r in controller_objs)


# ============================================================
# BAND / SECTOR LABEL SYSTEM (Scope of Work display)
# Confirmed against real sites — NOT the same convention as DSS's Greek naming.
# ============================================================

SECTOR_NAME = {'A': 'Alpha', 'B': 'Beta', 'C': 'Gamma', 'D': 'Delta', 'E': 'Epsilon', 'F': 'Foxtrot'}
SECTOR_ORDER = ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon', 'Foxtrot']

def lte_band_label(cell_name):
    """e.g. ECL00043_2A_1 -> ('AWS_1', 'Alpha') ; DXL04049_7A_2_F -> ('FNET', 'Alpha')"""
    if not cell_name:
        return None, None
    m = re.search(r'_(\d)([A-F])_(\d+)(_[EF])?$', str(cell_name))
    if not m:
        return None, None
    digit, letter, carrier, suffix = m.group(1), m.group(2), m.group(3), m.group(4)
    sector = SECTOR_NAME.get(letter, letter)
    if digit == '9':
        return f"PCS_{carrier}", sector
    if digit == '2':
        return f"AWS_{carrier}", sector
    if digit == '8':
        return f"850_{carrier}", sector
    if digit == '3':
        return "WCS", sector
    if digit == '7':
        if suffix == '_F':
            return "FNET", sector
        if suffix == '_E':
            return "LTE_700_E", sector
        return "LTE_700", sector
    return f"BAND{digit}_{carrier}", sector

def nr_band_label(cell_name):
    """e.g. NCRN002376_N066A_1 -> ('5G_AWS_1', 'Alpha') ; ..._N077A_2 -> ('DOD', 'Alpha')"""
    if not cell_name:
        return None, None
    m = re.search(r'_N(\d{3})([A-F])_(\d+)$', str(cell_name))
    if not m:
        return None, None
    band, letter, carrier = m.group(1), m.group(2), m.group(3)
    sector = SECTOR_NAME.get(letter, letter)
    if band == '005':
        return "5G_850", sector
    if band == '002':
        return f"5G_PCS_{carrier}", sector
    if band == '066':
        return f"5G_AWS_{carrier}", sector
    if band == '077':
        return {'1': 'CBAND', '2': 'DOD', '3': 'DOD_BWE'}.get(carrier, f"N077_{carrier}"), sector
    return f"N{band}_{carrier}", sector

def band_label(cell_name):
    """Dispatch to LTE or 5G labeler based on whether the cell name contains an 'N0xx' 5G marker."""
    if re.search(r'_N\d{3}[A-F]_\d+$', str(cell_name or '')):
        return nr_band_label(cell_name)
    return lte_band_label(cell_name)

def is_5g_cell(cell_name):
    return bool(re.search(r'_N\d{3}[A-F]_\d+$', str(cell_name or '')))

def dedupe_labels(cell_names, lte_first=True):
    """Classify a list of cell names into unique band labels, LTE group first then 5G group,
    preserving first-seen order within each group."""
    lte_labels, fiveg_labels = [], []
    for c in cell_names:
        label, _ = band_label(c)
        if not label:
            continue
        target = fiveg_labels if is_5g_cell(c) else lte_labels
        if label not in target:
            target.append(label)
    return (lte_labels + fiveg_labels) if lte_first else (fiveg_labels + lte_labels)


def extract_precheck_sectors(text):
    """Parse the Pre-checks PDF's 'Summary Status' table: Node | Technology | Cell | ...
    Returns (set of (node, cell) tuples, set of node names)."""
    if not text:
        return set(), set()
    pairs = set()
    nodes = set()
    for m in re.finditer(r'(\S+)\s+(LTE|5G)\s+(\S+)\s+(UNLOCKED|LOCKED)', text):
        node, tech, cell = m.group(1), m.group(2), m.group(3)
        pairs.add((node, cell))
        nodes.add(node)
    return pairs, nodes


# ============================================================
# CARRIER ADD / DELETE / MOVE / RETUNE CLASSIFICATION
# ============================================================

def classify_carriers(ciq_wb, mm_objs, precheck_text):
    """Returns a dict: added (per node), moved, deleted_sectors, deleted_nodes, retuned."""
    result = {"added": {}, "moved": [], "deleted_sectors": {}, "deleted_nodes": [], "retuned": []}

    pre_pairs, pre_nodes = extract_precheck_sectors(precheck_text)
    pre_cells = {cell for (_, cell) in pre_pairs}

    ciq_nodes = {str(r.get("Node to be built as", "")).strip() for r in mm_objs if r.get("Node to be built as")}
    if pre_nodes:
        result["deleted_nodes"] = sorted(pre_nodes - ciq_nodes)

    delmove_objs = sheet_objs(ciq_wb["Sector Del_Movement"]) if "Sector Del_Movement" in ciq_wb.sheetnames else []
    handled_cells = set()

    for r in delmove_objs:
        src_node, src_sector = r.get("Source Node name"), r.get("Source Sector")
        tgt_node, tgt_sector = r.get("Target Node name"), r.get("Target Sector")
        handled_cells.add(src_sector)
        if str(tgt_node).strip().upper() == "DELETE":
            result["deleted_sectors"].setdefault(src_node, []).append(src_sector)
            continue
        handled_cells.add(tgt_sector)
        src_dl, tgt_dl = str(r.get("Source channelNumberDL", "")).strip(), str(r.get("Target channelNumberDL", "")).strip()
        src_bw, tgt_bw = str(r.get("Source Bandwidth", "")).strip(), str(r.get("Target Bandwidth", "")).strip()
        retuned = (src_dl != tgt_dl) or (src_bw != tgt_bw)
        if str(src_node).strip().upper() == str(tgt_node).strip().upper():
            if retuned:
                label, _ = lte_band_label(src_sector)
                result["retuned"].append({"label": label, "from": f"{src_dl}/{src_bw}", "to": f"{tgt_dl}/{tgt_bw}"})
        else:
            result["moved"].append({"cell": src_sector, "from_node": src_node, "to_node": tgt_node})
            if retuned:
                label, _ = lte_band_label(tgt_sector)
                result["retuned"].append({"label": label, "from": f"{src_dl}/{src_bw}", "to": f"{tgt_dl}/{tgt_bw}"})

    # ADD: any CIQ cell (LTE or 5G) not present in Pre-checks and not already accounted for as moved/deleted
    eutran_objs = sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    fiveg_objs = sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
    for r in mm_objs:
        node = r.get("Node to be built as")
        e_name, g_name = r.get("eNodeB Name"), r.get("gNodeB Name")
        added_here = []
        for row in eutran_objs:
            cell = row.get("EutranCellFDDId")
            if not cell or cell in handled_cells or cell in pre_cells:
                continue
            if e_name and str(cell).startswith(str(e_name)):
                added_here.append(cell)
        for row in fiveg_objs:
            cell = row.get("NRCellDU")
            if not cell or cell in handled_cells or cell in pre_cells:
                continue
            if g_name and str(cell).startswith(str(g_name)):
                added_here.append(cell)
        if added_here:
            result["added"][node] = added_here

    return result


def format_scope_of_work(classification, controller_objs, dss_outputs_meta=None, controller_edp_found=None):
    """Turn the classification dict into the confirmed display lines.
    controller_edp_found: dict of {controller_id: bool} — False means the 6610 shows in the CIQ
    but isn't published in EDP yet."""
    lines = []
    for node, cells in classification["added"].items():
        labels = dedupe_labels(cells)
        lines.append(f"Integration:\t{'/'.join(labels)}\t{node}")

    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    for r in ctrl_rows:
        ctrl_id = r.get('Controller ID')
        lines.append(f"6610 Controller Integration:\t{ctrl_id}")
        if controller_edp_found is not None and controller_edp_found.get(ctrl_id) is False:
            lines.append(f"6610 Controller Integration:\tEDP is not published for the controller\t{ctrl_id}")

    moved_by_pair = {}
    for m in classification["moved"]:
        key = (m["from_node"], m["to_node"])
        moved_by_pair.setdefault(key, []).append(m["cell"])
    for (from_node, to_node), cells in moved_by_pair.items():
        labels = dedupe_labels(cells)
        label_str = labels[0] if len(labels) == 1 else f"[{'/'.join(labels)}]"
        sector_names = sorted({band_label(c)[1] for c in cells if band_label(c)[1]}, key=lambda s: SECTOR_ORDER.index(s) if s in SECTOR_ORDER else 99)
        sectors_str = f" {', '.join(sector_names)}" if sector_names else ""
        lines.append(f"Moved Sectors:\t{label_str}{sectors_str}\tFrom:\t{from_node}\tTo:\t{to_node}")

    for node in classification["deleted_nodes"]:
        lines.append(f"Deleted Node from ENM:\t{node}")

    for node, cells in classification["deleted_sectors"].items():
        labels = dedupe_labels(cells)
        lines.append(f"Deleted Sector:\t{'/'.join(labels)}\t{node}")

    for r in classification["retuned"]:
        lines.append(f"Retune on:\t{r['label']}\tFrom:\t{r['from']}\tTo:\t{r['to']}")

    if dss_outputs_meta:
        lines.append(f"DSS Activation:\t{' & '.join(dss_outputs_meta)}")

    return lines


# ============================================================
# GENERATOR: shared node-template fill (used by MMBB / TMBB / CENM alike —
# they all share the identical placeholder set, confirmed against the source templates)
# ============================================================

def fill_node_template(base_tpl, row, edp_index, user_id, date_str, summary_rows, log):
    site_id = row.get("Node to be built as")
    e_name, g_name, g_id = row.get("eNodeB Name"), row.get("gNodeB Name"), row.get("gNBId")
    is_primary_lte = str(site_id).strip().upper() == str(e_name or "").strip().upper()
    lte_row = edp_row_for(edp_index, e_name)
    fiveg_row = edp_row_for(edp_index, g_name)
    primary_row = lte_row if is_primary_lte else fiveg_row
    secondary_row = fiveg_row if is_primary_lte else lte_row

    lte_bearer = edp_get(edp_index, lte_row, "IPV6_ENODEB_BEARER_IP")
    fiveg_bearer = edp_get(edp_index, fiveg_row, "IPV6_ENODEB_BEARER_IP")
    ptp_server = edp_get(edp_index, primary_row, "BBU_PTP_SERVER_IP") or edp_get(edp_index, secondary_row, "BBU_PTP_SERVER_IP")
    ptp_siad = edp_get(edp_index, primary_row, "BBU_PTP_SIAD_IP") or edp_get(edp_index, secondary_row, "BBU_PTP_SIAD_IP")

    tpl = base_tpl
    fills = [
        ("xxxSiteIdxxx", site_id, "CIQ · Mixed Mode Info · Node to be built as (triple-x variant)"),
        ("xxSiteIDxx", site_id, "CIQ · Mixed Mode Info · Node to be built as"),
        ("xxSiteIdxx", site_id, "CIQ · Mixed Mode Info · Node to be built as"),
        ("xxDatexx", date_str, "manual input"),
        ("xxUserIDxx", user_id, "manual input"),
        ("x5G_gNBIdx", g_id, "CIQ · Mixed Mode Info · gNBId"),
        ("xgNB_Namex", g_name, "CIQ · Mixed Mode Info · gNodeB Name"),
        ("xxBBU_PTP_SERVER_IPxx", ptp_server, "EDP · BBU_PTP_SERVER_IP (primary→secondary fallback)"),
        ("xxBBU_PTP_SIAD_IPxx", ptp_siad, "EDP · BBU_PTP_SIAD_IP (primary→secondary fallback)"),
        ("xLTE_IPV6_ENODEB_BEARER_IPx", lte_bearer, "EDP · IPV6_ENODEB_BEARER_IP (row matched by eNodeB Name)"),
        ("x5G_IPV6_ENODEB_BEARER_IPx", fiveg_bearer, "EDP · IPV6_ENODEB_BEARER_IP (row matched by gNodeB Name)"),
    ]
    for token, val, src in fills:
        if val:
            tpl = tpl.replace(token, str(val))
            summary_rows.append({"Item": f"{site_id} · {token}", "Source": src, "Value": val, "Note": ""})
        else:
            summary_rows.append({"Item": f"{site_id} · {token}", "Source": src, "Value": "NOT FOUND", "Note": "left as placeholder"})
        log(f"{'✓' if val else '✗'} {site_id} · {token} -> {val or 'NOT FOUND'}")
    return tpl


def generate_6610(controller_objs, user_id, date_str, log):
    """Universal add-on: generate the 6610 controller template if Controller Info shows 6610.
    Applies to ALL scopes (MCA, CENM, CRAN) per the blueprint's 'For ALL SCOPES' rule."""
    outputs, summary_rows = [], []
    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    if not ctrl_rows:
        return outputs, summary_rows
    base_tpl = TPL_6610.read_text(encoding="utf-8")
    for r in ctrl_rows:
        ctrl_id = r.get("Controller ID")
        tpl = base_tpl.replace("##Controller_id##", str(ctrl_id))
        tpl = tpl.replace("xSite_IDx", str(ctrl_id))
        tpl = tpl.replace("xxUserIDxx", user_id)
        tpl = tpl.replace("xDatex", date_str)
        outputs.append((f"{ctrl_id}_6610_Controller_Integration_Filled.txt", tpl))
        summary_rows.append({"Item": "6610 Controller ID", "Source": "CIQ · Controller Info", "Value": ctrl_id, "Note": "6610 IX template generated (applies across all scopes)"})
        summary_rows.append({"Item": "xSite_IDx", "Source": "same as Controller ID — no other node in scope for this template", "Value": ctrl_id, "Note": "VERIFY if ever wrong"})
        log(f"✓ 6610 present -> generated for Controller ID {ctrl_id}")
    return outputs, summary_rows


# ============================================================
# GENERATOR: DSS checks (ported from ashhureddy/TRYDSS — DSS Extractor Tool)
# Universal add-on across ALL scopes (MCA, CENM, CRAN), same as 6610.
# ============================================================

DSS_SECTOR_MAP = {'A': 'alpha', 'B': 'beta', 'C': 'gamma', 'D': 'delta', 'E': 'epsilon', 'F': 'zeta'}
DSS_KEEP_PARAMS = ['gNBId', 'gNB Name', 'SectorEquipmentFunction', 'cellLocalId', 'Carrier', 'ssbFrequency']
DSS_ESS_SC_LOOKUP = {
    "N066A_1": {"essScPairId": 2222, "essScLocalId": 20}, "N066B_1": {"essScPairId": 2223, "essScLocalId": 21},
    "N066C_1": {"essScPairId": 2224, "essScLocalId": 22}, "N066D_1": {"essScPairId": 2225, "essScLocalId": 23},
    "N066A_2": {"essScPairId": 2226, "essScLocalId": 24}, "N066B_2": {"essScPairId": 2227, "essScLocalId": 25},
    "N066C_2": {"essScPairId": 2228, "essScLocalId": 26}, "N066D_2": {"essScPairId": 2229, "essScLocalId": 27},
    "N002A_1": {"essScPairId": 3322, "essScLocalId": 30}, "N002B_1": {"essScPairId": 3323, "essScLocalId": 31},
    "N002C_1": {"essScPairId": 3324, "essScLocalId": 32}, "N002D_1": {"essScPairId": 3325, "essScLocalId": 33},
    "N002A_2": {"essScPairId": 3326, "essScLocalId": 34}, "N002B_2": {"essScPairId": 3327, "essScLocalId": 35},
    "N002C_2": {"essScPairId": 3328, "essScLocalId": 36}, "N002D_2": {"essScPairId": 3329, "essScLocalId": 37},
    "N005A_1": {"essScPairId": 1122, "essScLocalId": 10}, "N005B_1": {"essScPairId": 1123, "essScLocalId": 11},
    "N005C_1": {"essScPairId": 1124, "essScLocalId": 12}, "N005D_1": {"essScPairId": 1125, "essScLocalId": 13},
    "N005A_2": {"essScPairId": 1126, "essScLocalId": 14}, "N005B_2": {"essScPairId": 1127, "essScLocalId": 15},
    "N005C_2": {"essScPairId": 1128, "essScLocalId": 16}, "N005D_2": {"essScPairId": 1129, "essScLocalId": 17},
}
DSS_PLACEHOLDERS = {
    "primary_node": "xxMMBB_Primary_Node_Namexx", "lte_site_id": "xxLTE_Site_IDxx",
    "nr_node_name": "xx5G_NR_Node_Namexx", "lte_enbid": "xxLTE_eNBIDxx", "nr_gnbid": "xx5G_NR_gNBIDxx",
    "lte_cellid_a": "LTE_cellidA", "lte_cellid_b": "LTE_cellidB", "lte_cellid_c": "LTE_cellidC", "lte_cellid_d": "LTE_cellidD",
    "nr_celllocalid_a": "xx5G_celllocalidAxx", "nr_celllocalid_b": "xx5G_celllocalidBxx",
    "nr_celllocalid_c": "xx5G_celllocalidCxx", "nr_celllocalid_d": "xx5G_celllocalidDxx",
    "nr_ssbfrequency_a": "xx5G_ssbfrequencyAxx",
    "nr_sector_carrier_alpha": "xx5G_NRSectorCarrier_Alphaxx", "nr_sector_carrier_beta": "xx5G_NRSectorCarrier_Betaxx",
    "nr_sector_carrier_gamma": "xx5G_NRSectorCarrier_Gammaxx", "nr_sector_carrier_delta": "xx5G_NRSectorCarrier_Deltaxx",
    "lte_sector_carrier_alpha": "xxLTE_SectorCarrier_No_Alphaxx", "lte_sector_carrier_beta": "xxLTE_SectorCarrier_No_Betaxx",
    "lte_sector_carrier_gamma": "xxLTE_SectorCarrier_No_Gammaxx", "lte_sector_carrier_delta": "xxLTE_SectorCarrier_No_Deltaxx",
    "lte_site_xa_1": "xxLTE_Site_IDxx_XA_1", "lte_site_xb_1": "xxLTE_Site_IDxx_XB_1",
    "lte_site_xc_1": "xxLTE_Site_IDxx_XC_1", "lte_site_xd_1": "xxLTE_Site_IDxx_XD_1",
    "nr_node_n00xa_1": "xx5G_NR_Node_Namexx_N00XA_1", "nr_node_n00xb_1": "xx5G_NR_Node_Namexx_N00XB_1",
    "nr_node_n00xc_1": "xx5G_NR_Node_Namexx_N00XC_1", "nr_node_n00xd_1": "xx5G_NR_Node_Namexx_N00XD_1",
    "n00xa": "N00XA", "n00xb": "N00XB", "n00xc": "N00XC", "n00xd": "N00XD", "n00x": "N00X",
    "ess_sc_pair_id_a": "essScPairId_A", "ess_sc_pair_id_b": "essScPairId_B",
    "ess_sc_pair_id_c": "essScPairId_C", "ess_sc_pair_id_d": "essScPairId_D",
    "ess_sc_local_id_a": "essScLocalId_A", "ess_sc_local_id_b": "essScLocalId_B",
    "ess_sc_local_id_c": "essScLocalId_C", "ess_sc_local_id_d": "essScLocalId_D",
    "nr_node_n00x": "xx5G_NR_Node_Namexx_N00X",
}

def dss_extract_band_carrier_pattern(nrcelldu):
    if not nrcelldu:
        return "UNKNOWN"
    parts = str(nrcelldu).strip().split('_')
    if len(parts) < 3:
        return "UNKNOWN"
    middle_part, carrier_num = parts[1], parts[2]
    m = re.match(r'^([A-Z]\d+)[A-Z]?$', middle_part)
    return f"{m.group(1)}_{carrier_num}" if m else "UNKNOWN"

def dss_extract_sector(value):
    if not value or str(value) == 'nan':
        return None
    s = str(value)
    m = re.search(r'_([A-Z]\d+)([A-Z])_', s)
    if m:
        return m.group(2)
    m = re.search(r'_(\d+)([A-Z])_', s)
    return m.group(2) if m else None

def dss_get_greek_name(sector, counts):
    if sector not in DSS_SECTOR_MAP:
        return f"sector_{sector.lower()}"
    greek = DSS_SECTOR_MAP[sector]
    if sector in counts:
        counts[sector] += 1
        return f"{greek}{counts[sector]}"
    counts[sector] = 0
    return greek

def dss_filter_row(row):
    row_map = {str(k).strip().upper(): k for k in row.keys()}
    out = {}
    for p in DSS_KEEP_PARAMS:
        ku = p.strip().upper()
        if ku in row_map:
            out[p] = row[row_map[ku]]
    return out

def dss_get_primary_node_info(mm_objs, gnb_name, gnb_id):
    for r in mm_objs:
        if r.get("gNodeB Name") == gnb_name and r.get("gNBId") == gnb_id:
            out = {}
            if r.get("Node to be built as"):
                out["primary_node"] = r.get("Node to be built as")
            if is_populated(r.get("eNBId")):
                out["eNBId"] = r.get("eNBId")
            if r.get("eNodeB Name"):
                out["lte_siteID"] = r.get("eNodeB Name")
            return out
    return {}

def dss_get_sector_cell_ids(eutran_objs, dss_value, greek_name):
    search_val = str(dss_value).strip()
    for row in eutran_objs:
        if str(row.get("EutranCellFDDId", row.get("EUtranCellFDDId", ""))).strip() == search_val:
            out = {}
            if is_populated(row.get("sectorId")) or is_populated(row.get("SectorId")):
                out[f"{greek_name}_sectorId"] = row.get("sectorId", row.get("SectorId"))
            if is_populated(row.get("cellId")) or is_populated(row.get("CellId")):
                out[f"{greek_name}_cellId"] = row.get("cellId", row.get("CellId"))
            return out
    return {}

def dss_extract_pattern_for_ess(nr_value):
    if not nr_value:
        return None
    m = re.search(r'_(N\d{3}[A-D]_\d)$', str(nr_value))
    return m.group(1) if m else None

def dss_extract_n00x_from_node(nr_node_value):
    if not nr_node_value:
        return None
    m = re.match(r'^(.+_N\d{3})[A-D]_\d$', str(nr_node_value))
    return m.group(1) if m else None


def generate_dss(ciq_wb, mm_objs, user_id, date_str, log):
    """Ported from ashhureddy/TRYDSS DSS Extractor Tool — 6-step pipeline:
    extract -> group -> clean -> populate -> map -> generate.
    Also returns dss_activation_labels: ["5G_PCS_1|PCS_1", ...] for the Scope of Work summary."""
    outputs, summary_rows, dss_activation_labels = [], [], []

    if "5G Info" not in ciq_wb.sheetnames:
        return outputs, summary_rows, dss_activation_labels
    fiveg_objs = sheet_objs(ciq_wb["5G Info"])
    dss_col = next((k for k in (fiveg_objs[0].keys() if fiveg_objs else []) if k.strip().upper() == "DSS"), None)
    if not dss_col:
        return outputs, summary_rows, dss_activation_labels

    # Step 1: extract rows where DSS != "NO"
    dss_rows = [r for r in fiveg_objs if r.get(dss_col) is not None and str(r.get(dss_col)).strip().upper() != "NO"]
    if not dss_rows:
        summary_rows.append({"Item": "DSS checks", "Source": "CIQ · 5G Info · DSS column", "Value": "no DSS-active cells found", "Note": ""})
        return outputs, summary_rows, dss_activation_labels
    log(f"✓ DSS: found {len(dss_rows)} DSS-active cell(s) in 5G Info")

    # Step 2: group by band+carrier pattern
    groups = {}
    for row in dss_rows:
        pattern = dss_extract_band_carrier_pattern(row.get("NRCellDU"))
        groups.setdefault(pattern, []).append(row)

    eutran_objs = sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []

    outputs_count = {"4_sector": 0, "3_sector": 0}
    for i, (pattern, rows) in enumerate(sorted(groups.items()), start=1):
        var_name = pattern if pattern != "UNKNOWN" else f"DSS{i}"

        # Step 3: clean — Greek sector names + filtered rows
        cleaned = {}
        dss_counts, nr_counts = {}, {}
        for row in rows:
            sector = dss_extract_sector(row.get(dss_col))
            if sector:
                greek = dss_get_greek_name(sector, dss_counts)
                cleaned[f"DSS_{greek}"] = row.get(dss_col)
        for row in rows:
            sector = dss_extract_sector(row.get("NRCellDU"))
            if sector:
                greek = dss_get_greek_name(sector, nr_counts)
                cleaned[f"NR_{greek}"] = row.get("NRCellDU")
        cleaned["rows"] = [dss_filter_row(r) for r in rows]

        # Step 4: populate — Mixed Mode Info + eUtran Parameters
        if cleaned["rows"]:
            gnb_name = cleaned["rows"][0].get("gNB Name")
            gnb_id = cleaned["rows"][0].get("gNBId")
            if gnb_name and gnb_id:
                cleaned.update(dss_get_primary_node_info(mm_objs, gnb_name, gnb_id))
        for dss_key in sorted(k for k in cleaned if k.startswith("DSS_")):
            greek = dss_key.replace("DSS_", "")
            cleaned.update(dss_get_sector_cell_ids(eutran_objs, cleaned[dss_key], greek))
        for idx, row in enumerate(rows, start=1):
            sec_eq = row.get("SectorEquipmentFunction")
            if sec_eq:
                parts = str(sec_eq).split("_")
                if len(parts) >= 2:
                    cleaned[f"row{idx}"] = parts[-1]

        # Step 5: map to placeholders
        r = cleaned.get("rows", [])
        mapped = {}
        mapped[DSS_PLACEHOLDERS["primary_node"]] = cleaned.get("primary_node")
        mapped[DSS_PLACEHOLDERS["lte_site_id"]] = cleaned.get("lte_siteID")
        mapped[DSS_PLACEHOLDERS["nr_node_name"]] = r[0].get("gNB Name") if r else None
        mapped[DSS_PLACEHOLDERS["lte_enbid"]] = cleaned.get("eNBId")
        mapped[DSS_PLACEHOLDERS["nr_gnbid"]] = r[0].get("gNBId") if r else None
        for i2, letter in enumerate(["a", "b", "c", "d"]):
            mapped[DSS_PLACEHOLDERS[f"nr_celllocalid_{letter}"]] = r[i2].get("cellLocalId") if len(r) > i2 else None
        mapped[DSS_PLACEHOLDERS["nr_ssbfrequency_a"]] = r[0].get("ssbFrequency") if r else None
        nr_vals = {}
        for letter, key in [("a", "nr_alpha"), ("b", "nr_beta"), ("c", "nr_gamma"), ("d", "nr_delta")]:
            greek = DSS_SECTOR_MAP[letter.upper()]
            nr_vals[key] = cleaned.get(f"NR_{greek}")
        mapped[DSS_PLACEHOLDERS["nr_sector_carrier_alpha"]] = nr_vals["nr_alpha"]
        mapped[DSS_PLACEHOLDERS["nr_sector_carrier_beta"]] = nr_vals["nr_beta"]
        mapped[DSS_PLACEHOLDERS["nr_sector_carrier_gamma"]] = nr_vals["nr_gamma"]
        mapped[DSS_PLACEHOLDERS["nr_sector_carrier_delta"]] = nr_vals["nr_delta"]
        for letter, greek in [("a", "alpha"), ("b", "beta"), ("c", "gamma"), ("d", "delta")]:
            mapped[DSS_PLACEHOLDERS[f"lte_sector_carrier_{greek}"]] = cleaned.get(f"{greek}_sectorId")
            mapped[DSS_PLACEHOLDERS[f"lte_cellid_{letter}"]] = cleaned.get(f"{greek}_cellId")
            mapped[DSS_PLACEHOLDERS[f"lte_site_x{letter}_1"]] = cleaned.get(f"DSS_{greek}")
            mapped[DSS_PLACEHOLDERS[f"nr_node_n00x{letter}_1"]] = cleaned.get(f"NR_{greek}")
        mapped[DSS_PLACEHOLDERS["n00xa"]] = cleaned.get("row1")
        mapped[DSS_PLACEHOLDERS["n00xb"]] = cleaned.get("row2")
        mapped[DSS_PLACEHOLDERS["n00xc"]] = cleaned.get("row3")
        mapped[DSS_PLACEHOLDERS["n00xd"]] = cleaned.get("row4")
        for letter, greek in [("a", "alpha"), ("b", "beta"), ("c", "gamma"), ("d", "delta")]:
            ess = DSS_ESS_SC_LOOKUP.get(dss_extract_pattern_for_ess(nr_vals.get(f"nr_{greek}")) or "", {})
            mapped[DSS_PLACEHOLDERS[f"ess_sc_pair_id_{letter}"]] = ess.get("essScPairId")
            mapped[DSS_PLACEHOLDERS[f"ess_sc_local_id_{letter}"]] = ess.get("essScLocalId")
        nr_node_ref = nr_vals.get("nr_gamma") or nr_vals.get("nr_delta")
        mapped[DSS_PLACEHOLDERS["nr_node_n00x"]] = dss_extract_n00x_from_node(nr_node_ref)
        mapped[DSS_PLACEHOLDERS["n00x"]] = pattern.split('_')[0] if pattern != "UNKNOWN" else None
        mapped["xxDatexx"] = date_str

        # Step 6: pick template (4-sector if Delta present, else 3-sector) and generate
        has_delta = any(mapped.get(DSS_PLACEHOLDERS[k]) is not None for k in
                         ["lte_cellid_d", "nr_celllocalid_d", "nr_sector_carrier_delta", "ess_sc_pair_id_d"])
        tpl_path = TPL_DSS_4SECTOR if has_delta else TPL_DSS_3SECTOR
        tpl_key = "4_sector" if has_delta else "3_sector"
        outputs_count[tpl_key] += 1
        tpl_text = tpl_path.read_text(encoding="utf-8")

        for placeholder, val in sorted(mapped.items(), key=lambda x: len(x[0]), reverse=True):
            if val is not None:
                tpl_text = tpl_text.replace(placeholder, str(val))
                summary_rows.append({"Item": f"{var_name} · {placeholder}", "Source": "DSS pipeline", "Value": val, "Note": ""})
            else:
                summary_rows.append({"Item": f"{var_name} · {placeholder}", "Source": "DSS pipeline", "Value": "NOT FOUND", "Note": "left as placeholder"})

        summary_rows.append({"Item": f"{var_name} · noOfRxAntennas/noOfTxAntennas", "Source": "not computed by source tool", "Value": "n/a", "Note": "gap in the original DSS Extractor — never mapped there either"})
        outputs.append((f"{var_name}_DSS_output.txt", tpl_text))
        log(f"✓ DSS group {var_name} -> {tpl_key} template, {sum(1 for v in mapped.values() if v is not None)} placeholders resolved")

        nr_cell_for_label = nr_vals.get("nr_alpha")
        lte_cell_for_label = cleaned.get("DSS_alpha")
        nr_label, _ = nr_band_label(nr_cell_for_label) if nr_cell_for_label else (None, None)
        lte_label, _ = lte_band_label(lte_cell_for_label) if lte_cell_for_label else (None, None)
        if nr_label and lte_label:
            dss_activation_labels.append(f"{nr_label}|{lte_label}")

    return outputs, summary_rows, dss_activation_labels


# ============================================================
# FINAL CONNECTIONS (universal — same across MCA/CENM/CRAN/N2E/NSB)
# ============================================================

def sheet_objs_dedup_first(ws):
    """Like sheet_objs but keeps the FIRST occurrence of a duplicate header name (eUtran Parameters
    has DUS/XMU columns repeated near the end — the two sets hold identical values, first wins)."""
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    seen, header_idx = {}, []
    for i, h in enumerate(headers):
        hs = str(h).strip() if h is not None else ''
        if hs and hs not in seen:
            seen[hs] = i
            header_idx.append((hs, i))
    objs = []
    for r in rows[1:]:
        if not any(str(c).strip() for c in r if c is not None):
            continue
        objs.append({h: (r[i] if i < len(r) else None) for h, i in header_idx})
    return objs


def generate_final_connections(ciq_wb, mm_objs):
    """One Excel file per CIQ (not per node) — Mixed Mode Info + conditional per-node XMU rows,
    then all 5G Info rows, then all eUtran Parameters rows. Styled: yellow/red title, blue/white headers, borders."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side

    fiveg_objs = sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
    eutran_objs = sheet_objs_dedup_first(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    enb_objs = sheet_objs(ciq_wb["eNB Info"]) if "eNB Info" in ciq_wb.sheetnames else []
    gnb_objs = sheet_objs(ciq_wb["gNB Info"]) if "gNB Info" in ciq_wb.sheetnames else []

    MM_COLS = ['Node to be built as', 'eNBId', 'eNodeB Name', 'gNBId', 'gNodeB Name', 'IDLA', 'Connected To Node', 'Connected From Port']
    XMU_ENB_COLS = ['eNBId', 'eNodeB Name', '1st DU type', '1st XMU', '1st XMU Port 1', '1st XMU Port 2', '1st XMU Port 3', '2nd DU type', '2nd XMU', '2nd XMU Port 1', '2nd XMU Port 2', '2nd XMU Port 3']
    XMU_GNB_COLS = ['gNBId', 'gNodeB Name', 'DU type', '1st XMU', '1st XMU Port 1', '1st XMU Port 2', '1st XMU Port 3', '2nd XMU', '2nd XMU Port 1', '2nd XMU Port 2', '2nd XMU Port 3']
    FIVEG_COLS = ['gNBId', 'gNB Name', 'NRCellDU', 'Operating Band', 'RRU Type', 'BB/XMU', 'Port 1', 'Port 2', 'Port 3', 'Port 4', 'Radio Port', 'Cascaded From Radio', 'BBU/XMU End SFP', 'Radio End SFP']
    LTE_COLS = ['EutranCellFDDId', 'eUTRA operating band', 'RRU type', 'DUS / XMU', 'DUS / XMU Port', 'DUS / XMU Port Expansion', 'Cascaded From Radio', 'Radio Port', 'BBU/XMU End SFP', 'Radio End SFP']
    NCOLS = 14

    TITLE_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    TITLE_FONT = Font(bold=True, color="FF0000", size=14)
    HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="000000")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_header_row(ws, r, cols, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = HEADER_FILL
            cell.border = BORDER
            if c <= len(cols):
                cell.value = cols[c - 1]
                cell.font = HEADER_FONT

    def write_data_row(ws, r, row, cols, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if c <= len(cols):
                cell.value = row.get(cols[c - 1])

    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "Sheet1"
    r = 1
    title_cell = ws.cell(row=r, column=2, value="Final connections")
    title_cell.font = TITLE_FONT
    for c in range(2, 9):
        ws.cell(row=r, column=c).fill = TITLE_FILL
    r += 2

    write_header_row(ws, r, MM_COLS, len(MM_COLS)); r += 1
    xmu_rows_to_add = []
    for row in mm_objs:
        write_data_row(ws, r, row, MM_COLS, len(MM_COLS)); r += 1
        primary = row.get("Node to be built as")
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(primary).strip().upper() == str(e_name or "").strip().upper()
        if is_lte_primary:
            match = next((x for x in enb_objs if str(x.get("eNodeB Name")) == str(e_name)), None)
            if match and (str(match.get("1st XMU")).strip().upper() == "YES" or str(match.get("2nd XMU")).strip().upper() == "YES"):
                xmu_rows_to_add.append((XMU_ENB_COLS, match))
        else:
            match = next((x for x in gnb_objs if str(x.get("gNodeB Name")) == str(g_name)), None)
            if match and (str(match.get("1st XMU")).strip().upper() == "YES" or str(match.get("2nd XMU")).strip().upper() == "YES"):
                xmu_rows_to_add.append((XMU_GNB_COLS, match))
    for cols, match in xmu_rows_to_add:
        r += 1
        write_header_row(ws, r, cols, len(cols)); r += 1
        write_data_row(ws, r, match, cols, len(cols)); r += 1

    r += 2
    write_header_row(ws, r, FIVEG_COLS, NCOLS); r += 1
    for row in fiveg_objs:
        write_data_row(ws, r, row, FIVEG_COLS, NCOLS); r += 1

    r += 2
    write_header_row(ws, r, LTE_COLS, NCOLS); r += 1
    for row in eutran_objs:
        write_data_row(ws, r, row, LTE_COLS, NCOLS); r += 1

    widths = [14, 10, 12, 10, 14, 10, 10, 10, 10, 10, 10, 16, 16, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# GENERIC PRE/POST CONFIGURATION (MCA / CENM — any node set, not CRAN's fixed roles)
# ============================================================

def generate_generic_pre_post(ciq_wb, mm_objs, precheck_text, precheck_node_names):
    """Pre = nodes actually found in Pre-checks. Post = nodes actually found in CIQ Mixed Mode Info.
    Each shown only on the side it's actually present — no 'vacated'/'new' padding.
    Dual-identity nodes (MMBB/TMBB) get the (P)/(S) pairing, tagged with their real BBU Mode."""
    def node_label(row):
        primary = row.get('Node to be built as')
        e_name, g_name = row.get('eNodeB Name'), row.get('gNodeB Name')
        bbu_mode = row.get('BBU Mode')
        if is_populated(e_name) and is_populated(g_name):
            is_lte_primary = str(primary).strip().upper() == str(e_name).strip().upper()
            secondary = g_name if is_lte_primary else e_name
            return f"{primary}(P)/{secondary}(S)({bbu_mode})"
        return str(primary)

    post_nodes, labels = {}, {}
    ciq_order = []
    for row in mm_objs:
        primary = row.get('Node to be built as')
        ciq_order.append(primary)
        labels[primary] = node_label(row)
        e_name, g_name = row.get('eNodeB Name'), row.get('gNodeB Name')
        is_lte_primary = str(primary).strip().upper() == str(e_name or '').strip().upper()
        r = find_row_by_name(ciq_wb, 'eNB Info', 'eNodeB Name', e_name) if is_lte_primary else find_row_by_name(ciq_wb, 'gNB Info', 'gNodeB Name', g_name)
        if not r:
            r = find_row_by_name(ciq_wb, 'eNB Info', 'eNodeB Name', e_name) or find_row_by_name(ciq_wb, 'gNB Info', 'gNodeB Name', g_name)
        post_nodes[primary] = hw_string(r) or 'NOT FOUND'

    # order: CIQ order first (so Pre and Post always list shared nodes in the same sequence),
    # then any Pre-only nodes (e.g. a fully vacated node) appended after
    ordered_names = list(ciq_order) + [n for n in precheck_node_names if n not in ciq_order]

    pre_nodes = {}
    for name in ordered_names:
        hw = pre_hw_string(precheck_text, name)
        if hw:
            pre_nodes[name] = hw

    def lbl(n):
        return labels.get(n, n)

    pre_parts = [f"{lbl(n)}({hw})" for n, hw in pre_nodes.items()]
    post_parts = [f"{lbl(n)}({hw})" for n, hw in post_nodes.items()]
    return " + ".join(pre_parts), " + ".join(post_parts)


def push_all_controller_siad_rows(siad_rows, edp_index, controller_objs):
    """For every 6610 controller in the CIQ's Controller Info, add its SIAD row (ANCEQ_* columns)
    and track whether EDP actually has it published. Returns {controller_id: bool}."""
    found_status = {}
    for r in controller_objs:
        if str(r.get("Controller", "")).strip() == "6610":
            ctrl_id = r.get("Controller ID")
            found_status[ctrl_id] = push_controller_siad_row(siad_rows, edp_index, ctrl_id)
    return found_status



def generate_mca(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, precheck_text, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_mmbb = TPL_MMBB.read_text(encoding="utf-8")
    tpl_tmbb = TPL_TMBB.read_text(encoding="utf-8")

    for row in mm_objs:
        bbu_mode = str(row.get("BBU Mode", "")).strip()
        site_id = row.get("Node to be built as")
        if bbu_mode == "MMBB":
            tpl = fill_node_template(tpl_mmbb, row, edp_index, user_id, date_str, summary_rows, log)
            outputs.append((f"{site_id}_MMBB_Integration_Filled.txt", tpl))
            push_siad_row(siad_rows, edp_index, site_id)
        elif bbu_mode == "TMBB":
            tpl = fill_node_template(tpl_tmbb, row, edp_index, user_id, date_str, summary_rows, log)
            outputs.append((f"{site_id}_TMBB_Integration_Filled.txt", tpl))
            push_siad_row(siad_rows, edp_index, site_id)
        else:
            summary_rows.append({"Item": f"Node: {site_id}", "Source": f"BBU Mode = {bbu_mode}", "Value": "skipped", "Note": "not MMBB or TMBB"})
            log(f"· {site_id}: BBU Mode = {bbu_mode}, skipped")

    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []

    _, pre_nodes_found = extract_precheck_sectors(precheck_text)
    ciq_node_names = {r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as")}
    pre_line, post_line = generate_generic_pre_post(ciq_wb, mm_objs, precheck_text, pre_nodes_found | ciq_node_names)

    classification = classify_carriers(ciq_wb, mm_objs, precheck_text)
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found)

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# GENERATOR: CENM (always cENM_TRIMODE template, for TMBB-mode nodes)
# ============================================================

def generate_cenm(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, precheck_text, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_cenm = TPL_CENM.read_text(encoding="utf-8")

    tmbb_rows = [r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "TMBB"]
    if not tmbb_rows:
        summary_rows.append({"Item": "Node identification", "Source": "CIQ · Mixed Mode Info", "Value": "NOT FOUND", "Note": "CENM expects a BBU Mode = TMBB row"})
        return summary_rows, None, None, siad_rows, outputs, [], []

    for row in tmbb_rows:
        site_id = row.get("Node to be built as")
        tpl = fill_node_template(tpl_cenm, row, edp_index, user_id, date_str, summary_rows, log)
        outputs.append((f"{site_id}_cENM_TMBB_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, site_id)

    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []

    _, pre_nodes_found = extract_precheck_sectors(precheck_text)
    ciq_node_names = {r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as")}
    pre_line, post_line = generate_generic_pre_post(ciq_wb, mm_objs, precheck_text, pre_nodes_found | ciq_node_names)

    classification = classify_carriers(ciq_wb, mm_objs, precheck_text)
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found)

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# GENERATOR: CRAN (Trip 1 / Trip 2 / NSA — shared logic, per-variant options)
# ============================================================

def generate_cran(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, precheck_text, log, tpl_path, include_source_poles, needs_6673, out_name):
    summary_rows, siad_rows, outputs = [], [], []
    tpl = tpl_path.read_text(encoding="utf-8")

    macro = next((r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "MMBB"), None)
    lte = next((r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "SMBB"
                and is_populated(r.get("eNBId")) and not is_populated(r.get("gNBId"))), None)
    target = next((r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "SMBB"
                   and is_populated(r.get("gNBId")) and not is_populated(r.get("eNBId"))
                   and str(r.get("gNodeB Name", "")).strip().upper().endswith("F")), None)

    if not (macro and lte and target):
        summary_rows.append({"Item": "Node identification", "Source": "CIQ · Mixed Mode Info", "Value": "incomplete",
                              "Note": f"MMBB={bool(macro)} LTE={bool(lte)} CRAN={bool(target)}"})
        return summary_rows, None, None, siad_rows, outputs, [], []

    if "Sector Del_Movement" not in ciq_wb.sheetnames:
        summary_rows.append({"Item": "Sector Del_Movement tab", "Source": "CIQ", "Value": "NOT FOUND", "Note": "required for Source CRAN"})
        return summary_rows, None, None, siad_rows, outputs, [], []

    delmove = sheet_objs(ciq_wb["Sector Del_Movement"])
    source_id = delmove[0].get("Source Node name") if delmove else None

    target_poles, source_poles = {}, {}
    if "5G Info" in ciq_wb.sheetnames:
        for r in sheet_objs(ciq_wb["5G Info"]):
            if str(r.get("gNB Name", "")).strip().upper() == str(target.get("gNodeB Name", "")).strip().upper():
                cell = r.get("NRCellDU")
                m = re.search(r"([A-C])_([12])$", str(cell or ""))
                if m:
                    target_poles[f"{m.group(1).upper()}_{m.group(2)}"] = cell
    for r in delmove:
        cell = r.get("Source Sector")
        m = re.search(r"([A-C])_([12])$", str(cell or ""))
        if m:
            source_poles[f"{m.group(1).upper()}_{m.group(2)}"] = cell

    for key in ["A_1", "A_2", "B_1", "B_2", "C_1", "C_2"]:
        t_token = f"xxTarget_SiteIdxx_Pole_N077{key}"
        t_val = target_poles.get(key)
        if t_val:
            tpl = tpl.replace(t_token, t_val)
            summary_rows.append({"Item": t_token, "Source": "CIQ · 5G Info · NRCellDU", "Value": t_val, "Note": ""})
        else:
            summary_rows.append({"Item": t_token, "Source": "CIQ · 5G Info · NRCellDU", "Value": "NOT PRESENT", "Note": ""})
        log(f"{'✓' if t_val else '·'} {t_token} -> {t_val or 'not present'}")

        if include_source_poles:
            s_token = f"xxSource_SiteIdxx_Pole_N077{key}"
            s_val = source_poles.get(key)
            if s_val:
                tpl = tpl.replace(s_token, s_val)
                summary_rows.append({"Item": s_token, "Source": "CIQ · Sector Del_Movement · Source Sector", "Value": s_val, "Note": ""})
            else:
                summary_rows.append({"Item": s_token, "Source": "CIQ · Sector Del_Movement · Source Sector", "Value": "NOT PRESENT", "Note": ""})
            log(f"{'✓' if s_val else '·'} {s_token} -> {s_val or 'not present'}")

    target_5g_row = edp_row_for(edp_index, target.get("gNodeB Name"))
    target_bearer = edp_get(edp_index, target_5g_row, "IPV6_ENODEB_BEARER_IP")

    fills = [
        ("xxMacro_MMBB_SiteIdxx", macro.get("Node to be built as"), "CIQ · Mixed Mode Info (MMBB)"),
        ("xxMacro_MMBB_gNB_Namexx", macro.get("gNodeB Name"), "CIQ · Mixed Mode Info (MMBB)"),
        ("xxMacro_MMBB_gNBIdxx", macro.get("gNBId"), "CIQ · Mixed Mode Info (MMBB)"),
        ("xxTarget_5G_IPV6_ENODEB_BEARER_IPxx", target_bearer, "EDP · IPV6_ENODEB_BEARER_IP"),
        ("xxTarget_5G_gNBIdxx", target.get("gNBId"), "CIQ · Mixed Mode Info (Target CRAN)"),
        ("xxTarget_SiteIdxx", target.get("Node to be built as"), "CIQ · Mixed Mode Info (Target CRAN)"),
        ("xxLTE_SiteIDxx", lte.get("Node to be built as"), "CIQ · Mixed Mode Info (LTE)"),
        ("xxLTE_SiteIdxx", lte.get("Node to be built as"), "CIQ · Mixed Mode Info (LTE)"),
        ("xxSiteIDxx", target.get("Node to be built as"), "ambiguous — defaulted to Target, VERIFY"),
        ("xxSource_SiteIdxx", source_id, "CIQ · Sector Del_Movement · Source Node name"),
        ("xxUserIDxx", user_id, "manual input"),
        ("xxDATExx", date_str, "manual input"),
        ("xxDatexx", date_str, "manual input"),
        ("xxdatexx", date_str, "manual input"),
    ]
    if needs_6673:
        switch_id = None
        if "Sidehaul Info" in ciq_wb.sheetnames:
            for r in sheet_objs(ciq_wb["Sidehaul Info"]):
                if str(r.get("Switch", "")).strip() == "6673":
                    switch_id = r.get("SH Switch ID")
                    break
        fills.append(("xx6673_switch_idxx", switch_id, "CIQ · Sidehaul Info (Switch = 6673 -> SH Switch ID)"))

    for token, val, src in fills:
        if val:
            tpl = tpl.replace(token, str(val))
            summary_rows.append({"Item": token, "Source": src, "Value": val, "Note": ""})
        else:
            summary_rows.append({"Item": token, "Source": src, "Value": "NOT FOUND", "Note": "left as placeholder"})
        log(f"{'✓' if val else '✗'} {token} -> {val or 'NOT FOUND'}")
    tpl = tpl.replace("xDatex", date_str)

    if "NSA" in out_name:
        summary_rows.append({"Item": "xx5G_Cell_namexx / xxFDD_namexx", "Source": "not in header legend", "Value": "n/a", "Note": "confirmed manual RF-judgment field — left untouched"})

    # Pre/Post configuration — compact line format
    lte_row = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", lte.get("eNodeB Name"))
    lte_hw_post = hw_string(lte_row) or "NOT FOUND"
    lte_hw_pre = pre_hw_string(precheck_text, lte.get("eNodeB Name")) or "NOT FOUND"

    macro_primary = macro.get("Node to be built as")
    macro_is_primary_lte = str(macro_primary).strip().upper() == str(macro.get("eNodeB Name", "")).strip().upper()
    macro_secondary = macro.get("gNodeB Name") if macro_is_primary_lte else macro.get("eNodeB Name")
    macro_row = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", macro.get("eNodeB Name")) or \
                find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", macro.get("gNodeB Name"))
    macro_hw_post = hw_string(macro_row) or "NOT FOUND"
    macro_hw_pre = pre_hw_string(precheck_text, macro.get("eNodeB Name")) or "NOT FOUND"

    target_row = find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", target.get("gNodeB Name"))
    target_hw = hw_string(target_row) or "NOT FOUND"
    source_hw = pre_hw_string(precheck_text, source_id) or "NOT FOUND (no Pre-checks match)"

    pre_line = f"{lte.get('Node to be built as')}({lte_hw_pre}) + {macro_primary}(P)/{macro_secondary}(S)(MMBB)({macro_hw_pre}) + {source_id}({source_hw})"
    post_line = f"{lte.get('Node to be built as')}({lte_hw_post}) + {macro_primary}(P)/{macro_secondary}(S)(MMBB)({macro_hw_post}) + {target.get('Node to be built as')}({target_hw})"

    push_siad_row(siad_rows, edp_index, macro.get("Node to be built as"))
    push_siad_row(siad_rows, edp_index, lte.get("Node to be built as"))
    push_siad_row(siad_rows, edp_index, target.get("Node to be built as"))

    outputs.append((f"{target.get('Node to be built as')}_{out_name}_Filled.txt", tpl))

    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)

    binary_outputs = [(f"Final_Connections_{target.get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))]

    # CRAN has no Carrier ADD/Delete/Move "checks" per the blueprint — only 6610 and DSS ride along here
    scope_of_work_lines = format_scope_of_work({"added": {}, "moved": [], "deleted_sectors": {}, "deleted_nodes": []}, controller_objs, dss_labels, controller_edp_found)

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# UI
# ============================================================

st.markdown("<h1 style='text-align:center;color:#5b4fe0;margin-bottom:0;'>MASTEC</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align:center;'>📡 Integration Template Generator</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:gray;'>CIQ + EDP → filled AMOS/CMCLI templates, per scope of work</p>", unsafe_allow_html=True)
st.divider()

col_inputs, col_instructions = st.columns([3, 2])

with col_inputs:
    st.subheader("📤 Inputs")
    top_scope = st.radio("Scope of work", ["CRAN", "MCA", "CENM"], horizontal=True)
    cran_sub = None
    if top_scope == "CRAN":
        cran_sub = st.selectbox("CRAN scope", ["CRAN SA Rehome Trip 1", "CRAN SA Rehome Trip 2", "CRAN NSA Rehome"])

    ciq_file = st.file_uploader("CIQ (.xlsx)", type=["xlsx", "xls"])
    edp_file = st.file_uploader("EDP (.xlsx / .xls)", type=["xlsx", "xls"])
    pre_file = st.file_uploader("Pre-checks (.pdf) — optional", type=["pdf"])
    c1, c2 = st.columns(2)
    with c1:
        user_id = st.text_input("User ID", placeholder="e.g. pr970b")
    with c2:
        date_str = st.text_input("Execution date (mmddyyyy)", value=date.today().strftime("%m%d%Y"))
    run = st.button("Generate templates →", type="primary", disabled=not (ciq_file and edp_file))

with col_instructions:
    st.subheader("📋 Instructions")
    st.markdown("""
    1. **Pick** CRAN / MCA / CENM (CRAN needs a further Trip-1/Trip-2/NSA choice)
    2. **Upload** the CIQ and EDP for the site (Pre-checks optional)
    3. **Enter** your User ID and execution date
    4. Click **Generate templates**
    5. **Review**, then **download**
    """)
    st.caption("MCA auto-detects MMBB vs TMBB per node from the CIQ. 6610 controller template auto-generates alongside any scope if Controller Info shows 6610.")

st.divider()
st.subheader("📊 Processing Log")
log_box = st.empty()
log_lines = []

def log(msg):
    log_lines.append(msg)
    log_box.code("\n".join(log_lines) or "Waiting for file upload...", language=None)

log_box.code("Waiting for file upload...", language=None)

if run:
    _all_templates = {
        "MMBB": TPL_MMBB, "TMBB": TPL_TMBB, "cENM": TPL_CENM, "6610": TPL_6610,
        "CRAN Trip-1": TPL_CRAN_TRIP1, "CRAN Trip-2": TPL_CRAN_TRIP2, "CRAN NSA": TPL_CRAN_NSA,
        "DSS 4-sector": TPL_DSS_4SECTOR, "DSS 3-sector": TPL_DSS_3SECTOR,
    }
    _missing = [f"{label}  (expected: `{path.name}`)" for label, path in _all_templates.items() if not path.exists()]
    if _missing:
        st.error(
            "Some template files aren't in `templates/MCA/` in the repo — check the exact filenames match "
            "(GitHub sometimes changes spacing/characters on manual upload):\n\n"
            + "\n".join(f"- {m}" for m in _missing)
        )
        st.stop()

    import openpyxl
    log("Reading CIQ workbook...")
    ciq_wb = openpyxl.load_workbook(io.BytesIO(ciq_file.read()), data_only=True)
    if "Mixed Mode Info" not in ciq_wb.sheetnames:
        st.error('Could not find a "Mixed Mode Info" tab in the CIQ.')
        st.stop()
    mm_objs = sheet_objs(ciq_wb["Mixed Mode Info"])
    controller_objs = sheet_objs(ciq_wb["Controller Info"]) if "Controller Info" in ciq_wb.sheetnames else []

    log("Reading EDP workbook...")
    edp_bytes = edp_file.read()
    try:
        edp_wb = openpyxl.load_workbook(io.BytesIO(edp_bytes), data_only=True)
    except Exception:
        st.error("This EDP couldn't be read as .xlsx. If it's an old .xls file, make sure `xlrd` is in requirements.txt, or re-save it as .xlsx and re-upload.")
        st.stop()
    edp_index = build_edp_index(edp_wb)
    if not edp_index:
        st.error('Could not locate the EDP header row (expected a column "EDP_SITE_ID" and "SITE_NAME").')
        st.stop()

    precheck_text = ""
    if pre_file:
        log("Extracting Pre-checks PDF text...")
        precheck_text = extract_pdf_text(pre_file.read())

    pre_line = post_line = None
    uid = user_id or "xxUserIDxx"
    dstr = date_str or "xxDatexx"

    if top_scope == "MCA":
        summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = generate_mca(
            ciq_wb, edp_index, controller_objs, mm_objs, uid, dstr, precheck_text, log)
    elif top_scope == "CENM":
        summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = generate_cenm(
            ciq_wb, edp_index, controller_objs, mm_objs, uid, dstr, precheck_text, log)
    else:  # CRAN
        cran_opts = {
            "CRAN SA Rehome Trip 1": (TPL_CRAN_TRIP1, False, False, "CRAN_Trip1"),
            "CRAN SA Rehome Trip 2": (TPL_CRAN_TRIP2, True, False, "CRAN_Trip2"),
            "CRAN NSA Rehome": (TPL_CRAN_NSA, True, True, "CRAN_NSA"),
        }
        tpl_path, inc_src, need_6673, out_name = cran_opts[cran_sub]
        summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = generate_cran(
            ciq_wb, edp_index, controller_objs, mm_objs, uid, dstr, precheck_text, log,
            tpl_path, inc_src, need_6673, out_name)

    log("Done.")

    st.divider()
    st.subheader("✅ Extraction summary")
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    if pre_line and post_line:
        st.subheader("🔀 Pre / Post configuration")
        st.code(f"Pre Configuration:  {pre_line}\nPost Configuration: {post_line}", language=None)

    if siad_rows:
        st.subheader("🔌 SIAD port assignment")
        st.dataframe(pd.DataFrame(siad_rows), use_container_width=True, hide_index=True)

    if scope_lines:
        st.subheader("📋 Scope of work summary")
        st.code("\n".join(scope_lines), language=None)
        st.caption("Tab-separated — paste directly into Excel/Notepad and it will land in columns.")

    if outputs:
        st.subheader("📄 Generated output")
        for name, text in outputs:
            unresolved = highlight_unresolved(text)
            with st.expander(f"{name}  ({'⚠ ' + str(len(unresolved)) + ' unresolved' if unresolved else '✓ fully resolved'})"):
                st.text_area("Preview", text, height=300, key=name)
                st.download_button("⬇ Download .txt", text, file_name=name, key=f"dl_{name}")

    if binary_outputs:
        st.subheader("📊 Final connections")
        for name, data in binary_outputs:
            st.download_button(f"⬇ Download {name}", data, file_name=name, key=f"dl_bin_{name}")

    if outputs or binary_outputs:
        if len(outputs) + len(binary_outputs) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for name, text in outputs:
                    zf.writestr(name, text)
                for name, data in binary_outputs:
                    zf.writestr(name, data)
            st.download_button("⬇ Download all as .zip", zip_buf.getvalue(), file_name="generated_templates.zip")

st.divider()
st.caption("Made by **AKSHATHA KALLUR** | Powered by **MASTEC** | © 2026")
