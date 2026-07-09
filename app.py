import streamlit as st
import pandas as pd
import re
import io
import time
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
TPL_CENM_MMBB = resolve_template("cENM_MMBB_Integration_Pre-existing_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V4.txt", "cENM_MMBB")
TPL_6610 = resolve_template("6610 Controller Integration Procedure_25Q3_Updated_V12.txt", "6610")
TPL_PORT_CONVERSION = resolve_template("Template_Port_Conversion_1G_to_10G_BBU_V1_1.txt", "Port_Conversion")
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

TDIR_N2E = Path(__file__).parent / "templates" / "N2E"
TPL_N2E_LTE = TDIR_N2E / "N2E_LTE_Integration_Procedure_with_LTE_Node_as_Primary_V4.txt"
TPL_N2E_5G = TDIR_N2E / "N2E_5G_Integration_Procedure_with_5G_Node_as_Primary_V4.txt"
TPL_N2E_MMBB = TDIR_N2E / "MMBB_N2E_Integration_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V6.txt"
TPL_N2E_TRIMODE = TDIR_N2E / "N2E_TRIMODE_Integration_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V6.txt"

TDIR_NSB = Path(__file__).parent / "templates" / "NSB"
TDIR_STATIC = Path(__file__).parent / "templates" / "Static"
TPL_NSB_MMBB = TDIR_NSB / "LTE+5G_MMBB_Integration_NSB_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V13.txt"
TPL_NSB_TRIMODE = TDIR_NSB / "TRIMODE_Integration_NSB_Procedure_with_LTE_or_5G_Node_as_Primary_CMCLI_Updated_V6.txt"

# ============================================================
# SHARED HELPERS
# ============================================================

def load_workbook_any(file_bytes, filename):
    """openpyxl can only read real .xlsx (zip-based OOXML) — legacy .xls (OLE2/CFB binary) needs xlrd.
    Some files have a mismatched extension (e.g. an old .xls saved/renamed with a .xlsx name), so this
    doesn't trust the filename alone: it tries the format the extension suggests first, then the other
    one on failure. It also repairs a common 'could not read stylesheet' crash — some non-Microsoft
    export tools produce a malformed xl/styles.xml even though the actual cell data is fine — by
    swapping in a minimal valid stylesheet and retrying (verified against a deliberately-corrupted
    styles.xml: openpyxl's read_only mode does NOT sidestep this, but replacing the styles part does)."""
    import openpyxl, zipfile

    def via_xlrd():
        import pandas as pd
        all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="xlrd", header=None)
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)
        for sheet_name, df in all_sheets.items():
            ws = out_wb.create_sheet(title=str(sheet_name)[:31])  # Excel sheet name limit
            for row in df.itertuples(index=False, name=None):
                ws.append(list(row))
        return out_wb

    def via_openpyxl():
        return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    def via_repaired_styles():
        MINIMAL_STYLES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>'''
        zin = zipfile.ZipFile(io.BytesIO(file_bytes), "r")
        repaired_buf = io.BytesIO()
        zout = zipfile.ZipFile(repaired_buf, "w")
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = MINIMAL_STYLES
            zout.writestr(item, data)
        zout.close()
        return openpyxl.load_workbook(io.BytesIO(repaired_buf.getvalue()), data_only=True)

    looks_like_xls = filename.lower().endswith(".xls") and not filename.lower().endswith(".xlsx")
    attempts = [via_xlrd, via_openpyxl, via_repaired_styles] if looks_like_xls else [via_openpyxl, via_repaired_styles, via_xlrd]

    first_error = None
    for attempt in attempts:
        try:
            return attempt()
        except Exception as e:
            if first_error is None:
                first_error = e
    raise first_error  # surface the first (most likely relevant) error

def sheet_objs(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    objs = []
    for r in rows[1:]:
        if not any(str(c).strip() for c in r if c is not None):
            continue
        objs.append({headers[i]: (r[i].strip() if isinstance(r[i], str) else r[i]) if i < len(r) else "" for i in range(len(headers))})
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
    if v is None or (isinstance(v, float) and v != v):  # v != v is the standard NaN check
        return None
    s = str(v).strip()
    return None if s == "" or s.lower() == "nan" else s


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
    # some Pre-checks PDF versions insert an extra ISO-timestamp token between ENABLED and the
    # actual product name (seen on newer AAS/5216-style hardware rows) — (?:\S+\s+)? skips it if present
    # some Pre-checks PDF versions insert an extra isSharedWithExternalMe column (true/false)
    # between faultIndicator and operationalIndicator — (?:(?:true|false)\s+)? skips it if present
    m = re.search(esc + r"\s+1\s+UNLOCKED\s+OFF\s+(?:(?:true|false)\s+)?STEADY_ON\s+ENABLED\s+(?:\S+\s+)?([A-Za-z0-9 ]+?)\s+\d{6,8}", text, re.I)
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
    return len(re.findall(esc + r"\s+XMU\S*\s+UNLOCKED\s+OFF\s+(?:(?:true|false)\s+)?STEADY_ON\s+ENABLED", text, re.I))


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
        "SIAD CLLI": (edp_get(edp_index, row, "SIAD_CLLI") or "NOT FOUND") if found else "NOT FOUND",
        "Port Size": "1G" if found else "NOT FOUND",
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

def build_node_alias_map(mm_objs):
    """A node's secondary identity (eNodeB or gNodeB name) can appear in Sector Del_Movement's
    Source/Target columns instead of its Primary ID — happens specifically when the moving cell's
    own technology matches the secondary identity (e.g. a 5G cell moving into a dual-identity node
    records the target using that node's gNodeB name, not its Primary 'Node to be built as')."""
    alias = {}
    for row in mm_objs:
        primary = row.get("Node to be built as")
        if not primary:
            continue
        for secondary in (row.get("eNodeB Name"), row.get("gNodeB Name")):
            if secondary and str(secondary).strip() and str(secondary).strip() != str(primary).strip():
                alias[str(secondary).strip()] = str(primary).strip()
    return alias


def classify_carriers(ciq_wb, mm_objs, precheck_text):
    """Returns a dict: added (per node), moved, deleted_sectors, deleted_nodes, retuned."""
    result = {"added": {}, "moved": [], "deleted_sectors": {}, "deleted_nodes": [], "retuned": [], "node_band_sectors": {}}
    alias_map = build_node_alias_map(mm_objs)

    def normalize(name):
        return alias_map.get(str(name).strip(), name) if name else name

    pre_pairs, pre_nodes = extract_precheck_sectors(precheck_text)
    pre_cells = {cell for (_, cell) in pre_pairs}

    # per (node, band label) sector inventory — used to tell "whole band moved" from "partial move"
    node_band_sectors = {}
    for (node, cell) in pre_pairs:
        label, sector = band_label(cell)
        if label and sector:
            node_band_sectors.setdefault((node, label), set()).add(sector)

    ciq_nodes = {str(r.get("Node to be built as", "")).strip() for r in mm_objs if r.get("Node to be built as")}
    if pre_nodes:
        result["deleted_nodes"] = sorted(pre_nodes - ciq_nodes)

    delmove_objs = sheet_objs(ciq_wb["Sector Del_Movement"]) if "Sector Del_Movement" in ciq_wb.sheetnames else []
    handled_cells = set()

    for r in delmove_objs:
        src_node, src_sector = normalize(r.get("Source Node name")), r.get("Source Sector")
        tgt_node_raw, tgt_sector = r.get("Target Node name"), r.get("Target Sector")
        tgt_node = tgt_node_raw if str(tgt_node_raw).strip().upper() == "DELETE" else normalize(tgt_node_raw)
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
                if not label:
                    label, _ = nr_band_label(src_sector)
                result["retuned"].append({"label": label, "from": f"{src_dl}/{src_bw}", "to": f"{tgt_dl}/{tgt_bw}"})
        else:
            result["moved"].append({"cell": src_sector, "from_node": src_node, "to_node": tgt_node})
            if retuned:
                label, _ = lte_band_label(tgt_sector)
                if not label:
                    label, _ = nr_band_label(tgt_sector)
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

    result["node_band_sectors"] = node_band_sectors
    return result


def format_scope_of_work(classification, controller_objs, dss_outputs_meta=None, controller_edp_found=None, radio_swaps=None):
    """Turn the classification dict into the confirmed display lines.
    controller_edp_found: dict of {controller_id: bool} — False means the 6610 shows in the CIQ
    but isn't published in EDP yet."""
    lines = []
    for node, cells in classification.get("added", {}).items():
        labels = dedupe_labels(cells)
        lines.append(f"Integration:\t{'/'.join(labels)}\t{node}")

    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    for r in ctrl_rows:
        ctrl_id = r.get('Controller ID')
        if controller_edp_found is not None and controller_edp_found.get(ctrl_id) is False:
            lines.append(f"EDP is not published for the controller — {ctrl_id}")
        else:
            lines.append(f"6610 Controller Integration:\t{ctrl_id}")

    moved_by_pair = {}
    for m in classification.get("moved", []):
        key = (m["from_node"], m["to_node"])
        moved_by_pair.setdefault(key, []).append(m["cell"])
    WHOLE_BAND_SET = {"Alpha", "Beta", "Gamma"}
    for (from_node, to_node), cells in moved_by_pair.items():
        if not is_populated(to_node) or not any(is_populated(c) for c in cells):
            # Malformed Sector Del_Movement row (missing Target Node name / Source Sector) —
            # flag it plainly instead of emitting a garbled "Moved Sectors: [] ... To: None" line.
            lines.append(f"Moved Sectors:\tCHECK CIQ — incomplete Sector Del_Movement row\tFrom:\t{from_node or 'NOT FOUND'}\tTo:\t{to_node or 'NOT FOUND'}")
            continue
        labels = dedupe_labels(cells)
        label_str = labels[0] if len(labels) == 1 else f"[{'/'.join(labels)}]"
        per_label_moved = {}
        for c in cells:
            label, sector = band_label(c)
            if label and sector:
                per_label_moved.setdefault(label, set()).add(sector)
        # "whole" = every band in this move brought all of Alpha+Beta+Gamma together — not
        # about what the source node happened to have historically, just this move itself
        is_whole = bool(per_label_moved) and all(WHOLE_BAND_SET <= sset for sset in per_label_moved.values())
        sector_names = sorted({s for sset in per_label_moved.values() for s in sset}, key=lambda s: SECTOR_ORDER.index(s) if s in SECTOR_ORDER else 99)
        sectors_str = "" if is_whole else (f" {', '.join(sector_names)}" if sector_names else "")
        lines.append(f"Moved Sectors:\t{label_str}{sectors_str}\tFrom:\t{from_node}\tTo:\t{to_node}")

    for node in classification.get("deleted_nodes", []):
        lines.append(f"Deleted Node from ENM:\t{node}")

    for node, cells in classification.get("deleted_sectors", {}).items():
        labels = dedupe_labels(cells)
        lines.append(f"Deleted Sector:\t{'/'.join(labels)}\t{node}")

    retune_seen = set()
    for r in classification.get("retuned", []):
        sig = (r["label"], r["from"], r["to"])
        if sig in retune_seen:
            continue
        retune_seen.add(sig)
        lines.append(f"Retune on:\t{r['label']}\tFrom:\t{r['from']}\tTo:\t{r['to']}")

    # Group by the actual swap SIGNATURE (From -> To), not by physical co-location. Co-located
    # bands (e.g. PCS_1 and AWS_1 sharing one antenna group) can have genuinely different radio
    # compositions (e.g. PCS_1 alone has a daisy-chained secondary radio) — blending them by
    # co-location silently combined unrelated radio tokens. Only bands that share the identical
    # From/To signature get combined into one bracketed line; sectors merge the same way.
    merged = {}
    for r in (radio_swaps or []):
        sig = (r["from"], r["to"])
        merged.setdefault(sig, {"labels": set(), "sectors": set()})
        merged[sig]["labels"].add(r["label"])
        merged[sig]["sectors"].add(r["sector"])

    for (from_radio, to_radio), grp in merged.items():
        labels = tuple(sorted(grp["labels"]))
        sector_set = grp["sectors"]

        label_str = labels[0] if len(labels) == 1 else f"[{'|'.join(labels)}]"
        sector_names = sorted(sector_set, key=lambda s: SECTOR_ORDER.index(s) if s in SECTOR_ORDER else 99)
        is_whole = WHOLE_BAND_SET <= sector_set
        sectors_str = " sectors" if is_whole else (f" {', '.join(sector_names)}" if sector_names else "")
        lines.append(f"Radio Swap on:\t{label_str}{sectors_str}\tFrom:\t{from_radio}\tTo:\t{to_radio}")

    if dss_outputs_meta:
        lines.append(f"DSS Activation:\t{' & '.join(dss_outputs_meta)}")

    return lines


def scope_lines_to_table(scope_lines):
    """Parse the tab-separated Scope of Work lines into a clean table: Category | Details | From | To.
    Plain monospace text can't align cleanly since 'Integration:' and '6610 Controller Integration:'
    are very different lengths — a real table sidesteps that entirely."""
    rows = []
    for line in scope_lines:
        parts = line.split("\t")
        category = parts[0].rstrip(":")
        if "From:" in parts:
            fi = parts.index("From:")
            details = " ".join(p for p in parts[1:fi] if p)
            from_val = parts[fi + 1] if fi + 1 < len(parts) else ""
            ti = parts.index("To:") if "To:" in parts else None
            to_val = parts[ti + 1] if ti is not None and ti + 1 < len(parts) else ""
            rows.append({"Category": category, "Details": details, "From": from_val, "To": to_val})
        else:
            details = " — ".join(p for p in parts[1:] if p)
            rows.append({"Category": category, "Details": details, "From": "", "To": ""})
    return rows


def scope_lines_to_readable_text(scope_lines):
    """Same parsed fields as scope_lines_to_table, but rendered as compact readable sentences —
    single spaces, no raw tab characters (which jump to wide tab-stops in monospace display)."""
    out = []
    for row in scope_lines_to_table(scope_lines):
        if row["From"] or row["To"]:
            out.append(f"{row['Category']}: {row['Details']}  From: {row['From']}  To: {row['To']}")
        else:
            out.append(f"{row['Category']}: {row['Details']}" if row["Details"] else row["Category"])
    return out


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


def generate_6610(controller_objs, user_id, date_str, log, edp_found=None):
    """Universal add-on: generate the 6610 controller template if Controller Info shows 6610.
    Applies to ALL scopes (MCA, CENM, CRAN) per the blueprint's 'For ALL SCOPES' rule.
    edp_found: {controller_id: bool} — a controller not confirmed published in EDP gets no
    IX template at all (nothing reliable to fill it with), just a summary note explaining why."""
    outputs, summary_rows = [], []
    ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
    if not ctrl_rows:
        return outputs, summary_rows
    base_tpl = TPL_6610.read_text(encoding="utf-8")
    for r in ctrl_rows:
        ctrl_id = r.get("Controller ID")
        if edp_found is not None and edp_found.get(ctrl_id) is False:
            summary_rows.append({"Item": "6610 Controller ID", "Source": "CIQ · Controller Info", "Value": ctrl_id, "Note": "EDP not published — 6610 IX template skipped"})
            log(f"✗ 6610 present but EDP not published for Controller ID {ctrl_id} — IX template skipped")
            continue
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
        objs.append({h: ((r[i].strip() if isinstance(r[i], str) else r[i]) if i < len(r) else None) for h, i in header_idx})
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
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = HEADER_FILL
            cell.border = BORDER
            cell.value = cols[c - 1]
            cell.font = HEADER_FONT

    def write_data_row(ws, r, row, cols, ncols):
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
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
        clean_row = dict(row)
        for k in ("eNBId", "eNodeB Name", "gNBId", "gNodeB Name"):
            if not is_populated(clean_row.get(k)):
                clean_row[k] = None
        write_data_row(ws, r, clean_row, MM_COLS, len(MM_COLS)); r += 1
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
# IDL CONNECTIONS (shared by MCA / CENM / NSB — CIQ-only, no Pre-checks PDF)
#
# Confirmed logic:
#  - Trigger: site has 2+ BBU rows in Mixed Mode Info.
#  - Per node: type (MMBB/TMBB vs LTE-standalone vs 5G-standalone) is read off eNBId/gNBId
#    presence in Mixed Mode Info; board generation comes from the "DU type" column in
#    eNB Info (LTE-standalone), gNB Info (5G-standalone), or either (MMBB/TMBB) —
#    "1st DU type"/"2nd DU type" are explicitly ignored.
#  - DU type -> generation: 6630/5216 -> G2, 6648/6651 -> G3, 6672 -> G4.
#  - The site's generation combination (order-independent) is matched against the confirmed
#    15-template registry. Combinations with both a Preferred and Alternate variant
#    (G2+G3, G3+G3, G3+G3+G3) generate BOTH files. The 4 known-unsupported 3-BBU
#    combinations (G2+G2+G4, G2+G3+G4, G3+G3+G4, G3+G4+G4) return "IDL template not found".
#  - Node ordering for same-generation nodes follows CIQ row order, top = 1st.
#  - Placeholder filling is generic rather than hardcoded per template: for each node we build
#    a set of candidate slot-prefixes (global row-position ordinal, per-generation-group
#    ordinal, and the plain generation label when that generation is a singleton at the site)
#    and only replace whichever placeholder actually appears in that specific template file —
#    per your instruction to "just fill whatever placeholders the template has", since the 15
#    files don't all share one exact naming convention (e.g. G2+G4+G4 uses Node_ID/ENBID
#    instead of NODE_ID/Node_eNBId).
# ============================================================

TDIR_IDL = Path(__file__).parent / "templates" / "IDL"
TDIR_N2E_IDL = Path(__file__).parent / "templates" / "N2E" / "IDL"

DU_TYPE_TO_GEN = {"6630": "G2", "5216": "G2", "6648": "G3", "6651": "G3", "6672": "G4"}

# combo (sorted tuple of generations) -> list of (filename, variant label)
IDL_TEMPLATE_REGISTRY = {
    ("G2", "G2"): [("G2+G2_RPM 777 417.txt", "")],
    ("G2", "G3"): [("G2+ G3_RPM 777 544.txt", "Preferred"), ("G2+ G3_RPM 777 098.txt", "Alternate")],
    ("G2", "G4"): [("G4+G2_RPM 777 543.txt", "IDLe")],
    ("G3", "G3"): [("G3+G3_RPM 777 052.txt", "Preferred"), ("G3+G3_RPM 777 053.txt", "Alternate")],
    ("G3", "G4"): [("G4+G3_RPM 777 052.txt", "IDLe")],
    ("G4", "G4"): [("G4+G4_RPM 777 052.txt", "Preferred")],
    ("G2", "G2", "G2"): [("G2+ G2+G2_RPM 77 417.txt", "")],
    ("G2", "G2", "G3"): [("G2+ G2+G3_RPM 77 417_098.txt", "")],
    ("G2", "G3", "G3"): [("G2+G3+G3_RPM 777 053_544.txt", "")],
    ("G3", "G3", "G3"): [("G3+ G3+ G3_RPM 777 052.txt", "Preferred"), ("G3+ G3+ G3_RPM 777 053.txt", "Alternate")],
    ("G2", "G4", "G4"): [("G2+G4+G4_RPM_777_053_543.txt", "")],
    ("G4", "G4", "G4"): [("G4+G4+G4_RPM 777 052.txt", "")],
    ("G3", "G4", "G4"): [("G3 + G4 + G4_RPM_777_052.txt", "")],
    # ("G2","G2","G4"), ("G2","G3","G4"), ("G3","G3","G4") -> no template exists;
    # falls through to the "IDL Template not found" branch below.
}

# N2E confirmed to support only these 2 combinations (not the full 15) — reuses the same
# file content/naming as the shared set, just from its own templates/N2E/IDL/ folder.
N2E_IDL_TEMPLATE_REGISTRY = {
    ("G3", "G3"): [("G3+G3_RPM 777 052.txt", "Preferred"), ("G3+G3_RPM 777 053.txt", "Alternate")],
    ("G4", "G4"): [("G4+G4_RPM 777 052.txt", "Preferred")],
    # every other combination -> "IDL Template not found" for N2E specifically, even though
    # MCA/CENM/NSB support it via the full registry above.
}

IDL_SUFFIX_CANDIDATES = {
    "NODE_ID": ["NODE_ID", "Node_ID", "BBU_Node_ID"],
    "5G_NODE_ID": ["5G_NODE_ID", "5G_NodeID"],
    "GNB_ID": ["NODE_GNB_ID", "GNBID"],
    "eNBId": ["Node_eNBId", "ENBID", "BBU_ENBID"],
}


def _ordinal(n):
    return {1: "1st", 2: "2nd", 3: "3rd"}.get(n, f"{n}th")


def get_node_generation(ciq_wb, row):
    """Board generation (G2/G3/G4) for one Mixed Mode Info row, per confirmed rule."""
    has_enb = is_populated(row.get("eNBId"))
    has_gnb = is_populated(row.get("gNBId"))
    e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")

    du_type = None
    if has_enb and has_gnb:  # MMBB/TMBB — either tab carries it
        r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name)
        du_type = r.get("DU type") if r else None
        if not is_populated(du_type):
            r = find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
            du_type = r.get("DU type") if r else None
    elif has_enb:  # LTE standalone
        r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name)
        du_type = r.get("DU type") if r else None
    elif has_gnb:  # 5G standalone
        r = find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        du_type = r.get("DU type") if r else None

    if not is_populated(du_type):
        return None
    return DU_TYPE_TO_GEN.get(str(du_type).strip())


def _idl_node_values(row):
    return {
        "NODE_ID": row.get("Node to be built as"),
        "5G_NODE_ID": row.get("gNodeB Name"),
        "GNB_ID": row.get("gNBId"),
        "eNBId": row.get("eNBId"),
    }


def fill_idl_template(template_text, node_slots, summary_rows, log, template_name):
    """node_slots: list of (candidate_prefixes, row). For each node/concept, tries every
    candidate-prefix x suffix-variant combination and fills whichever placeholder actually
    exists in this template — templates only get the placeholders they actually reference.

    Divider/hash handling (per confirmed instruction: placeholders are filled ALONG WITH their
    surrounding hashes, leaving just the bare value):
    - 4-hash forms like ####G2_NODE_ID#### are replaced entirely with the value.
    - Bare ##<prefix>_NODE## tokens are replaced entirely with the node's ID."""
    tpl = template_text
    for prefixes, row in node_slots:
        values = _idl_node_values(row)
        node_label = row.get("Node to be built as")
        for concept, value in values.items():
            for prefix in prefixes:
                for suffix in IDL_SUFFIX_CANDIDATES[concept]:
                    placeholder = f"##{prefix}_{suffix}##"
                    divider_form = f"####{prefix}_{suffix}####"
                    if divider_form in tpl and is_populated(value):
                        tpl = tpl.replace(divider_form, str(value))
                    if placeholder in tpl:
                        if is_populated(value):
                            tpl = tpl.replace(placeholder, str(value))
                            summary_rows.append({"Item": f"IDL · {node_label} · {placeholder}", "Source": template_name, "Value": value, "Note": ""})
                            log(f"✓ IDL {template_name}: {placeholder} -> {value}")
                        else:
                            summary_rows.append({"Item": f"IDL · {node_label} · {placeholder}", "Source": template_name, "Value": "NOT FOUND", "Note": ""})
                            log(f"✗ IDL {template_name}: {placeholder} -> NOT FOUND")
        # bare "_NODE" tokens (e.g. ##1st_G3_NODE##) — filled entirely with the node's ID
        if is_populated(node_label):
            for prefix in prefixes:
                node_divider = f"##{prefix}_NODE##"
                if node_divider in tpl:
                    tpl = tpl.replace(node_divider, str(node_label))
    return tpl


def generate_idl_connections(ciq_wb, mm_objs, user_id, date_str, log, template_dir=None, registry=None):
    """Returns (outputs, summary_rows, scope_lines) — same shape as the other generate_* add-ons.
    Shared by MCA / CENM / NSB (default template_dir/registry). N2E passes its own restricted
    registry (only G3+G3, G4+G4) and its own templates/N2E/IDL/ folder. No-ops for single-BBU sites."""
    template_dir = template_dir if template_dir is not None else TDIR_IDL
    registry = registry if registry is not None else IDL_TEMPLATE_REGISTRY
    outputs, summary_rows, scope_lines = [], [], []

    if len(mm_objs) < 2:
        return outputs, summary_rows, scope_lines

    nodes = [{"row": row, "gen": get_node_generation(ciq_wb, row)} for row in mm_objs]

    unresolved = [n["row"].get("Node to be built as") for n in nodes if not n["gen"]]
    if unresolved:
        note = f"Could not determine board generation (DU type) for: {', '.join(str(u) for u in unresolved)}"
        summary_rows.append({"Item": "IDL Connections", "Source": "DU type lookup", "Value": "NOT FOUND", "Note": note})
        log(f"✗ IDL Connections: {note}")
        scope_lines.append(f"IDL Connections:\tCould not determine board generation\t{', '.join(str(u) for u in unresolved)}")
        return outputs, summary_rows, scope_lines

    combo = tuple(sorted(n["gen"] for n in nodes))
    matches = registry.get(combo)

    if not matches:
        summary_rows.append({"Item": "IDL Connections", "Source": f"combination {'+'.join(combo)}", "Value": "IDL Template not found", "Note": ""})
        log(f"✗ IDL Connections: IDL Template not found for combination {'+'.join(combo)}")
        scope_lines.append(f"IDL Connections:\tIDL Template not found\t{'+'.join(combo)}")
        return outputs, summary_rows, scope_lines

    gen_counts = {}
    for n in nodes:
        gen_counts[n["gen"]] = gen_counts.get(n["gen"], 0) + 1
    group_seen = {}
    for i, n in enumerate(nodes, start=1):
        g = n["gen"]
        group_seen[g] = group_seen.get(g, 0) + 1
        candidates = [f"{_ordinal(i)}_{g}", f"{_ordinal(group_seen[g])}_{g}"]
        if gen_counts[g] == 1:
            candidates.append(g)
        n["prefixes"] = list(dict.fromkeys(candidates))  # dedupe, preserve order

    site_id = mm_objs[0].get("Node to be built as", "site")
    node_slots = [(n["prefixes"], n["row"]) for n in nodes]

    for fname, variant in matches:
        tpl_path = template_dir / fname
        if not tpl_path.exists():
            summary_rows.append({"Item": "IDL Connections", "Source": f"template {fname}", "Value": "NOT FOUND", "Note": f"expected file not in {template_dir}/: {fname}"})
            log(f"✗ IDL Connections: template file not found: {fname}")
            scope_lines.append(f"IDL Connections:\ttemplate file missing from repo\t{fname}")
            continue
        tpl_text = tpl_path.read_text(encoding="utf-8")
        filled = fill_idl_template(tpl_text, node_slots, summary_rows, log, fname)
        label = "+".join(combo) + (f"_{variant}" if variant else "")
        outputs.append((f"{site_id}_IDL_Connections_{label}.txt", filled))
        scope_lines.append(f"IDL Connections:\t{'+'.join(combo)}" + (f" ({variant})" if variant else "") + f"\t{fname}")

    return outputs, summary_rows, scope_lines


# ============================================================
# NGS CHECKS (all scopes — CIQ-only, no template output)
#
# Confirmed logic:
#  - Trigger: site has 2+ BBUs (same as IDL Connections).
#  - Data source: eUtran Parameters tab, "Co-Located Technology Cell" column (comma-separated
#    list of cell names sharing the same physical radio as that row's cell).
#  - Detection: build a cellname -> owning-node map (from eUtran Parameters' EutranCellFDDId and
#    5G Info's cell id, both matched back to a node via eNBId/gNBId), then for every pair of
#    different nodes at the site, check whether a cell on Node A references a cell on Node B
#    AND a cell on Node B references a cell on Node A (bidirectional, per your original
#    description — "mapped for the BBU2 sectors... & vice versa"). If both directions are
#    confirmed, the two nodes are sharing a physical radio -> NGS applies.
#  - No template file is generated — this only ever contributes a line to the Scope of Work
#    (and, by extension, the Checks Performed panel, since that reads Scope of Work lines).
# ============================================================

def _ngs_build_cell_node_map(ciq_wb, mm_objs):
    """cell name (LTE or 5G) -> owning node's 'Node to be built as', via eNBId/gNBId match."""
    enbid_to_node = {str(r.get("eNBId")).strip(): r.get("Node to be built as") for r in mm_objs if is_populated(r.get("eNBId"))}
    gnbid_to_node = {str(r.get("gNBId")).strip(): r.get("Node to be built as") for r in mm_objs if is_populated(r.get("gNBId"))}

    cell_to_node = {}
    if "eUtran Parameters" in ciq_wb.sheetnames:
        for row in sheet_objs(ciq_wb["eUtran Parameters"]):
            cell_name = row.get("EutranCellFDDId")
            node = enbid_to_node.get(str(row.get("eNBId")).strip())
            if cell_name and node:
                cell_to_node[str(cell_name).strip()] = node
    if "5G Info" in ciq_wb.sheetnames:
        for row in sheet_objs(ciq_wb["5G Info"]):
            cell_name = row.get("NRCellDU") or row.get("gNodeB Name")
            node = gnbid_to_node.get(str(row.get("gNBId")).strip())
            if cell_name and node:
                cell_to_node[str(cell_name).strip()] = node
    return cell_to_node


def _ngs_cell_band(cell_name):
    """Band label for either an LTE or a 5G cell name, whichever pattern matches."""
    label, _sector = lte_band_label(cell_name)
    if label:
        return label
    label, _sector = nr_band_label(cell_name)
    return label


def generate_ngs_checks(ciq_wb, mm_objs, log):
    """Returns (summary_rows, scope_lines). No file outputs — pure detection."""
    summary_rows, scope_lines = [], []

    if len(mm_objs) < 2 or "eUtran Parameters" not in ciq_wb.sheetnames:
        return summary_rows, scope_lines

    cell_to_node = _ngs_build_cell_node_map(ciq_wb, mm_objs)
    node_names = [r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as")]
    # Whether each node has an LTE side at all — only eUtran Parameters carries the
    # "Co-Located Technology Cell" column, so a 5G-only node can never declare a reference
    # back; requiring bidirectional confirmation for an LTE<->5G pair would be structurally
    # impossible to satisfy and produce a false negative on every genuinely-shared-radio site.
    has_lte = {r.get("Node to be built as"): is_populated(r.get("eNBId")) for r in mm_objs}

    # directional_refs[(from_node, to_node)] = list of every (from_cell, to_cell) pair seen
    directional_refs = {}
    for row in sheet_objs(ciq_wb["eUtran Parameters"]):
        own_cell = row.get("EutranCellFDDId")
        raw = row.get("Co-Located Technology Cell")
        if not is_populated(own_cell) or not is_populated(raw) or str(raw).strip().upper() in ("NA", "N/A"):
            continue
        own_node = cell_to_node.get(str(own_cell).strip())
        if not own_node:
            continue
        for ref_cell in str(raw).split(","):
            ref_cell = ref_cell.strip()
            ref_node = cell_to_node.get(ref_cell)
            if ref_node and ref_node != own_node:
                directional_refs.setdefault((own_node, ref_node), []).append((own_cell, ref_cell))

    checked_pairs = set()
    for i, node_a in enumerate(node_names):
        for node_b in node_names[i + 1:]:
            pair_key = frozenset((node_a, node_b))
            if pair_key in checked_pairs:
                continue
            checked_pairs.add(pair_key)
            a_to_b = directional_refs.get((node_a, node_b), [])
            b_to_a = directional_refs.get((node_b, node_a), [])
            both_lte = has_lte.get(node_a) and has_lte.get(node_b)
            confirmed = (a_to_b and b_to_a) if both_lte else (a_to_b or b_to_a)
            if confirmed:
                bands = set()
                for own_cell, ref_cell in a_to_b + b_to_a:
                    for c in (own_cell, ref_cell):
                        band = _ngs_cell_band(c)
                        if band:
                            bands.add(band)
                band_list = ", ".join(sorted(bands)) if bands else "band not determined"
                summary_rows.append({
                    "Item": "NGS Checks", "Source": f"{node_a} <-> {node_b}",
                    "Value": "radio shared", "Note": f"bands: {band_list}",
                })
                log(f"\u2713 NGS Checks: {node_a} <-> {node_b} share a radio (bands: {band_list})")
                scope_lines.append(f"NGS Activation on :\t{band_list}\t{node_a} <-> {node_b}")

    return summary_rows, scope_lines


# ============================================================
# PORT CONVERSION (MCA / CENM / CRAN — CIQ + Pre-checks + EDP)
#
# Confirmed logic:
#  - Rule 1: no board swap. Pre-checks' Hardware Status Information reports the node's actual
#    baseband model (via extract_pre_hw, already used elsewhere for Pre/Post configuration) —
#    if its generation doesn't match the CIQ's DU-type-derived generation, a board swap is
#    already in progress and Port Conversion doesn't apply (the OpMode difference is explained
#    by the swap, not a pure port-speed change).
#  - Rule 2: Pre-checks' "Transport Fiber link Status" table shows the board's relevant port
#    (below) at OpMode = 1G_FULL, while the EDP's SIAD_PORT_SIZE_BBU already shows 10G for that
#    same node — i.e. the port hasn't been converted yet, but the EDP already calls for it.
#  - Rule 3: which port to check depends on board generation:
#      G2 -> TN_A or TN_B      G3 -> TN_IDL_B      G4 -> TN_IDL_C
#  - Output: "Port speed 1G to 10G conversion with MPST: <NodeID>." — plain sentence, not
#    tab-separated like the other Scope of Work lines (confirmed).
# ============================================================

PORT_BY_GEN = {"G2": ["TN_A", "TN_B"], "G3": ["TN_IDL_B"], "G4": ["TN_IDL_C"]}

TRANSPORT_FIBER_ROW_RE = re.compile(
    r'(\S+)\s+(\S+)\s+(\d+)\s+(TN_A|TN_B|TN_IDL_B|TN_IDL_C)\s+\S+\s+\d+\s+(\S+)\s+(?:true|false)'
)


def extract_transport_fiber_opmode(precheck_text, node, port_labels):
    """OpMode string for the first Transport Fiber link Status row matching this node and
    one of its generation's relevant port labels, or None if no such row exists."""
    if not precheck_text or not node:
        return None
    node_u = str(node).strip().upper()
    for m in TRANSPORT_FIBER_ROW_RE.finditer(precheck_text):
        row_node, board, lnh, port, opmode = m.groups()
        if row_node.strip().upper() == node_u and port in port_labels:
            return opmode
    return None


def generate_port_conversion_checks(ciq_wb, mm_objs, edp_index, precheck_text, log):
    """Returns (outputs, summary_rows, scope_lines). Shared by MCA / CENM / CRAN.
    Generation is read from Pre-checks' Hardware Status (the board that CURRENTLY exists),
    not the CIQ's target/post board — the template applies to whatever board is actually in
    Pre-checks right now, regardless of what it's being swapped to (confirmed: this template is
    for the board that's in Pre, not the board in the CIQ's post state). G4 is excluded outright
    since it's the newest board (can never be a Pre-checks-side board here) and the template has
    no G4/TN_IDL_C content."""
    outputs, summary_rows, scope_lines = [], [], []
    if not precheck_text:
        return outputs, summary_rows, scope_lines

    tpl_text = TPL_PORT_CONVERSION.read_text(encoding="utf-8") if TPL_PORT_CONVERSION.exists() else None

    for row in mm_objs:
        node = row.get("Node to be built as")
        if not node:
            continue

        pre_model = extract_pre_hw(precheck_text, node)
        pre_gen = DU_TYPE_TO_GEN.get(str(pre_model).strip()) if pre_model else None
        if pre_gen not in ("G2", "G3"):
            continue  # G4 (or undetectable) in Pre-checks — template doesn't apply here at all
        port_labels = PORT_BY_GEN[pre_gen]

        opmode = extract_transport_fiber_opmode(precheck_text, node, port_labels)
        if not opmode or "1G" not in opmode.upper():
            continue  # not currently 1G in Pre-checks — nothing pending

        edp_row = edp_row_for(edp_index, node)
        siad_port_size = edp_get(edp_index, edp_row, "SIAD_PORT_SIZE_BBU") if edp_row else None
        if not is_populated(siad_port_size) or "10G" not in str(siad_port_size).upper():
            continue  # EDP doesn't call for 10G — nothing pending

        # Confirmed mismatch on a G2/G3 Pre-checks board — always show display line and always
        # generate the template, regardless of what the CIQ's target board ends up being.
        summary_rows.append({
            "Item": "Port Conversion", "Source": node,
            "Value": "1G -> 10G pending", "Note": f"Pre-checks board: {pre_gen}, port: {'/'.join(port_labels)}, EDP SIAD_PORT_SIZE_BBU: {siad_port_size}",
        })
        log(f"\u2713 Port Conversion: {node} — 1G in Pre-checks ({pre_gen} board), EDP calls for 10G")
        scope_lines.append(f"Port speed 1G to 10G conversion with MPST: {node}.")

        if tpl_text is None:
            summary_rows.append({"Item": "Port Conversion", "Source": f"template {TPL_PORT_CONVERSION.name}", "Value": "NOT FOUND", "Note": f"expected file not in templates/MCA/: {TPL_PORT_CONVERSION.name}"})
            log(f"\u2717 Port Conversion: template file not found for {node}")
            continue

        filled = tpl_text.replace("xxSiteIdxx", str(node)).replace("xSiteIDx", str(node))
        outputs.append((f"{node}_Port_Conversion_1G_to_10G.txt", filled))

    return outputs, summary_rows, scope_lines


# ============================================================
# PRE FIBERS (universal — pulled from Pre-checks' DL/UL Loss table)
# ============================================================

DL_UL_LOSS_ROW_RE = re.compile(
    r'(\S+)\s+(?:Up|Down)\s+\d+\s+(?:(?!\S+\s+(?:Up|Down)\s+\d+).)*?'
    r'((?:Baseband|XMU|:\s*RRU)\S*(?:(?!\S+\s+(?:Up|Down)\s+\d+).)*Port\s+D\d)', re.DOTALL
)

def extract_dl_ul_loss_rows(precheck_text):
    if not precheck_text:
        return []
    seen, rows = set(), []
    for m in DL_UL_LOSS_ROW_RE.finditer(precheck_text):
        cell, dus_xmu_rru = m.group(1), m.group(2).strip()
        key = (cell, dus_xmu_rru)  # dedupe by cell+description, not cell alone — a dual-band
        if key not in seen:        # radio cell can legitimately have two distinct entries (Port D1/D2)
            seen.add(key)
            rows.append({"Cells": cell, "DUS/XMU (S.No) - RRU": dus_xmu_rru})
    return rows


# ============================================================
# UNIVERSAL STATIC OUTPUTS (all scopes — no filling, pure passthrough)
# ============================================================

STATIC_OUTPUT_FILES = [
    "Integration_Checklist_v3.xlsx",
    "Global Local Script Execution Order.xlsx",
]


def get_universal_static_outputs(log):
    """Returns a list of (filename, bytes) for the static reference files that ship alongside
    Final Connections / Pre Fibers for every scope, unmodified — no CIQ/EDP data goes into these."""
    outputs = []
    for fname in STATIC_OUTPUT_FILES:
        fpath = TDIR_STATIC / fname
        if fpath.exists():
            outputs.append((fname, fpath.read_bytes()))
            log(f"\u2713 Static output attached: {fname}")
        else:
            log(f"\u2717 Static output not found: templates/Static/{fname}")
    return outputs


def generate_pre_fibers(precheck_text):
    """One Excel file per CIQ: Cells + DUS/XMU (S.No) - RRU from Pre-checks' DL/UL Loss table,
    plus a blank 'Pre fibers' column for manual fill-in."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side
    rows = extract_dl_ul_loss_rows(precheck_text)
    if not rows:
        return None
    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "Sheet1"
    HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="000000")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(["Cells", "DUS/XMU (S.No) - RRU", "Pre fibers"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    for r in rows:
        ws.append([r["Cells"], r["DUS/XMU (S.No) - RRU"], None])
        for cell in ws[ws.max_row]:
            cell.border = BORDER
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 14
    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# RADIO SWAP (universal — compares Pre-checks' DL/UL Loss radio type against the CIQ's RRU Type)
# ============================================================

RADIO_TYPE_RE = re.compile(r'RRU[-\w]*\s*\(([^,)]+),')

def extract_precheck_radio_types(precheck_text):
    """Reuses the same DL/UL Loss row match as Pre Fibers, but pulls the radio product type
    out of the captured DUS/XMU description instead. A cell can legitimately appear more than
    once (a dual-band radio unit serving one sector through two physical radios/ports) — every
    distinct radio type seen for that cell is kept, not just the first or last."""
    out = {}
    if not precheck_text:
        return out
    for m in DL_UL_LOSS_ROW_RE.finditer(precheck_text):
        cell, desc = m.group(1), m.group(2)
        rm = RADIO_TYPE_RE.search(desc)
        if rm:
            radio = rm.group(1).strip()
            out.setdefault(cell, [])
            if radio not in out[cell]:
                out[cell].append(radio)
    return out


def ciq_radio_types(ciq_wb):
    """CIQ-side radio type per cell: 5G Info's 'RRU Type' for 5G cells, eUtran Parameters'
    'RRU type' for LTE cells."""
    out = {}
    if "5G Info" in ciq_wb.sheetnames:
        for r in sheet_objs(ciq_wb["5G Info"]):
            cell, rru = r.get("NRCellDU"), r.get("RRU Type")
            if cell and is_populated(rru):
                out[cell] = str(rru).strip()
    if "eUtran Parameters" in ciq_wb.sheetnames:
        for r in sheet_objs(ciq_wb["eUtran Parameters"]):
            cell, rru = r.get("EutranCellFDDId"), r.get("RRU type")
            if cell and is_populated(rru):
                out[cell] = str(rru).strip()
    return out


def radio_family(radio_string):
    """Extract just the RRU type — the token right after 'RRUS'/'Radio', ignoring the band suffix
    entirely (e.g. 'RRUS A2 B4' -> 'A2', 'RRUS 12 B4' -> '12', 'Radio 4890HP 48B2/B25 48B66 M01' ->
    '4890', 'RRUS 4890' -> '4890'). Handles both digit-leading types (strips trailing letters like
    'HP') and letter-leading alphanumeric types (kept as-is, e.g. 'A2')."""
    s = str(radio_string or "").strip()
    tokens = s.split()
    if len(tokens) >= 2 and tokens[0].upper() in ("RRUS", "RADIO"):
        type_token = tokens[1]
        m = re.match(r"^(\d+)", type_token)
        return m.group(1) if m else type_token.upper()
    m = re.search(r"\d{2,5}", s)
    return m.group(0) if m else s.upper()


def build_colocation_groups(ciq_wb):
    """Cell -> canonical co-location group key, from eUtran Parameters' 'Co-Located Technology
    Cell' column (lists peer cell names sharing the same physical radio, spans LTE+5G together).
    'NA' or blank means the cell isn't co-located with anything else. The group is registered
    under EVERY member's name (the LTE cell itself plus all its listed peers, including 5G
    cells) — a 5G cell only ever appears as a peer inside an LTE row, never as its own row, so
    without this it could never find the group it's actually listed in."""
    groups = {}
    if "eUtran Parameters" not in ciq_wb.sheetnames:
        return groups
    for r in sheet_objs(ciq_wb["eUtran Parameters"]):
        cell = r.get("EutranCellFDDId")
        colo_raw = r.get("Co-Located Technology Cell")
        if not cell:
            continue
        if colo_raw and str(colo_raw).strip().upper() != "NA":
            peers = {p.strip() for p in str(colo_raw).split(",") if p.strip()}
            peers.add(cell)
            group_key = tuple(sorted(peers))
            for member in peers:
                groups[member] = group_key
        else:
            groups[cell] = (cell,)
    return groups


def classify_radio_swaps(precheck_text, ciq_wb):
    """Cells present in both Pre-checks and the CIQ where the radio family genuinely differs
    (compared by RRU type only — ignoring band suffix — to avoid false positives from naming-
    format differences between the two sources). Follows Sector Del_Movement's rename mapping —
    a moved cell can be renamed, so the Pre-checks value must be compared against the CIQ using
    the cell's NEW name. A cell with multiple distinct Pre-checks radios (dual-band radio unit)
    shows all of them combined with '+' in the From field."""
    pre_radios = extract_precheck_radio_types(precheck_text)
    post_radios = ciq_radio_types(ciq_wb)
    colo_groups = build_colocation_groups(ciq_wb)

    rename_map = {}
    if "Sector Del_Movement" in ciq_wb.sheetnames:
        for r in sheet_objs(ciq_wb["Sector Del_Movement"]):
            src_sector, tgt_sector = r.get("Source Sector"), r.get("Target Sector")
            tgt_node = r.get("Target Node name")
            if src_sector and tgt_sector and str(tgt_node).strip().upper() != "DELETE":
                rename_map[src_sector] = tgt_sector

    swaps = []
    for cell, pre_radio_list in pre_radios.items():
        ciq_cell = rename_map.get(cell, cell)
        post_radio = post_radios.get(ciq_cell)
        if not post_radio:
            continue
        post_type = radio_family(post_radio)
        pre_types = []
        for r in pre_radio_list:
            t = radio_family(r)
            if t not in pre_types:
                pre_types.append(t)
        if post_type not in pre_types:
            label, sector = band_label(cell)
            pre_types_sorted = sorted(pre_types, key=lambda t: (not t[0].isdigit(), t))
            from_str = "RRU " + "+".join(pre_types_sorted)
            to_str = f"RRU {post_type}"
            group_key = colo_groups.get(ciq_cell, (ciq_cell,))
            swaps.append({"label": label, "sector": sector, "from": from_str, "to": to_str, "group_key": group_key})
    return swaps


# ============================================================
# GENERIC PRE/POST CONFIGURATION (MCA / CENM — any node set, not CRAN's fixed roles)
# ============================================================

def pre_node_label(precheck_text, node_name):
    """Determine a node's (P)/(S) identity pairing as it existed in Pre-checks — independent from
    the CIQ, since a node can genuinely convert LTE-only <-> MMBB/TMBB between Pre and Post (5G
    sectors moving onto or off of it as part of the same scope)."""
    pre_pairs, _ = extract_precheck_sectors(precheck_text)
    node_cells = [cell for (node, cell) in pre_pairs if node == node_name]
    if not node_cells:
        return node_name
    fiveg_cells = [c for c in node_cells if is_5g_cell(c)]
    lte_cells = [c for c in node_cells if not is_5g_cell(c)]
    if fiveg_cells and lte_cells:
        m = re.match(r"^(.+?)_N\d{3}[A-F]_\d+$", fiveg_cells[0])
        secondary = m.group(1) if m else fiveg_cells[0]
        return f"{node_name}(P)/{secondary}(S)"
    return node_name

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
    _, pre_nodes_set = extract_precheck_sectors(precheck_text)
    ordered_names = list(ciq_order) + [n for n in precheck_node_names if n not in ciq_order]

    pre_nodes = {}
    for name in ordered_names:
        if name in pre_nodes_set:  # only include nodes actually confirmed present in Pre-checks
            pre_nodes[name] = pre_hw_string(precheck_text, name) or "NOT FOUND"

    def lbl(n):
        return labels.get(n, n)

    pre_parts = [f"{pre_node_label(precheck_text, n)}({hw})" for n, hw in pre_nodes.items()]
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


# ============================================================
# GENERATOR: N2E (Nokia-to-Ericsson) — no Pre-checks at all, greenfield-style build.
# Pre Configuration is always the fixed string "Nokia"; Post derived from CIQ as usual.
# ============================================================

def n2e_node_type(row):
    """LTE-only (no gNBId) / 5G-only (no eNBId) / MMBB / TMBB (TRIMODE), from Mixed Mode Info."""
    has_lte = is_populated(row.get("eNBId")) or is_populated(row.get("eNodeB Name"))
    has_5g = is_populated(row.get("gNBId")) or is_populated(row.get("gNodeB Name"))
    if has_lte and not has_5g:
        return "LTE"
    if has_5g and not has_lte:
        return "5G"
    bbu_mode = str(row.get("BBU Mode", "")).strip().upper()
    if bbu_mode == "MMBB":
        return "MMBB"
    if bbu_mode == "TMBB":
        return "TRIMODE"
    return None


def fill_node_template_n2e(template_text, row, edp_index, user_id, date_str, controller_objs, summary_rows, log):
    """N2E placeholder fill — confirmed mapping, distinct from MCA's fill_node_template.
    xxOAMIPAddressxx always from EDP ipv6_enodeb_oam_ip matched by Primary ID, regardless of node type.
    5G-side EDP values (bearer/SIAD/Vlan) always matched by gNodeB Name, not by whichever identity
    is Primary — confirmed explicitly for MMBB/TRIMODE, and holds trivially for 5G-only too."""
    node_type = n2e_node_type(row)
    primary = row.get("Node to be built as")
    lte_name, gnb_name = row.get("eNodeB Name"), row.get("gNodeB Name")
    tpl = template_text

    def sub(placeholder, value, note=""):
        nonlocal tpl
        if is_populated(value):
            tpl = tpl.replace(placeholder, str(value))
            summary_rows.append({"Item": f"{primary} · {placeholder}", "Source": note, "Value": value, "Note": ""})
            log(f"✓ {primary} · {placeholder} -> {value}")
        else:
            summary_rows.append({"Item": f"{primary} · {placeholder}", "Source": note, "Value": "NOT FOUND", "Note": ""})
            log(f"✗ {primary} · {placeholder} -> NOT FOUND")

    for ph in ("xSite_IDx", "xSITE_IDx", "xxSiteIdxx"):
        sub(ph, primary, "Primary ID")
    sub("xxUserIDxx", user_id, "manual")
    sub("xDatex", date_str, "manual")

    oam_row = edp_row_for(edp_index, primary)
    sub("xxOAMIPAddressxx", edp_get(edp_index, oam_row, "ipv6_enodeb_oam_ip"), "EDP · ipv6_enodeb_oam_ip (by Primary ID)")

    if node_type == "LTE":
        lte_row = edp_row_for(edp_index, primary)  # LTE-only: Primary == eNodeB Name
        sub("xsecondary_IPV6_ENODEB_BEARER_IPx", edp_get(edp_index, lte_row, "ipv6_enodeb_bearer_ip"), "EDP · ipv6_enodeb_bearer_ip (LTE site ID)")
        sub("xLTE_IPV6_SIAD_BEARER_IPx", edp_get(edp_index, lte_row, "ipv6_siad_bearer_ip_def_router"), "EDP · ipv6_siad_bearer_ip_def_router")
        sub("xLTE_Vlan_IDx", edp_get(edp_index, lte_row, "bearer_enodeb_sb_vlan_id"), "EDP · bearer_enodeb_sb_vlan_id")

    elif node_type == "5G":
        sub("xgNBIdx", row.get("gNBId"), "CIQ · Mixed Mode Info")
        sub("xgNB_Namex", gnb_name, "CIQ · Mixed Mode Info")
        gnb_row = edp_row_for(edp_index, primary)  # 5G-only: Primary == gNodeB Name
        bearer_ip = edp_get(edp_index, gnb_row, "ipv6_enodeb_bearer_ip")
        sub("xSecondary_IPV6_ENODEB_BEARER_IPx", bearer_ip, "EDP · ipv6_enodeb_bearer_ip (5G Primary ID)")
        sub("xsecondary_IPV6_ENODEB_BEARER_IPx", bearer_ip, "EDP · ipv6_enodeb_bearer_ip (5G Primary ID)")
        sub("x5G_IPV6_SIAD_BEARER_IPx", edp_get(edp_index, gnb_row, "ipv6_siad_bearer_ip_def_router"), "EDP · ipv6_siad_bearer_ip_def_router")
        sub("x5G_Vlan_IDx", edp_get(edp_index, gnb_row, "bearer_enodeb_sb_vlan_id"), "EDP · bearer_enodeb_sb_vlan_id")

    elif node_type in ("MMBB", "TRIMODE"):
        sub("xgNBIdx", row.get("gNBId"), "CIQ · Mixed Mode Info")
        sub("xgNB_Namex", gnb_name, "CIQ · Mixed Mode Info")
        gnb_row = edp_row_for(edp_index, gnb_name)  # always matched by gNodeB Name, regardless of Primary/Secondary
        bearer_ip = edp_get(edp_index, gnb_row, "ipv6_enodeb_bearer_ip")
        sub("xSecondary_IPV6_ENODEB_BEARER_IPx", bearer_ip, "EDP · ipv6_enodeb_bearer_ip (by gNodeB Name)")
        sub("xsecondary_IPV6_ENODEB_BEARER_IPx", bearer_ip, "EDP · ipv6_enodeb_bearer_ip (by gNodeB Name)")
        sub("x5G_IPV6_SIAD_BEARER_IPx", edp_get(edp_index, gnb_row, "ipv6_siad_bearer_ip_def_router"), "EDP · ipv6_siad_bearer_ip_def_router")
        sub("x5G_Vlan_IDx", edp_get(edp_index, gnb_row, "bearer_enodeb_sb_vlan_id"), "EDP · bearer_enodeb_sb_vlan_id")
        ctrl_rows = [r for r in controller_objs if str(r.get("Controller", "")).strip() == "6610"]
        if ctrl_rows:
            sub("xController_IDX", ctrl_rows[0].get("Controller ID"), "CIQ · Controller Info")

    return tpl


def check_sa_conversion(ciq_wb, node_id):
    """SA Conversion: is this node's ID present anywhere in the CIQ's NR_SA tab?"""
    if "NR_SA" not in ciq_wb.sheetnames or not node_id:
        return False
    for row in ciq_wb["NR_SA"].iter_rows(values_only=True):
        if any(str(c).strip() == str(node_id).strip() for c in row if c is not None):
            return True
    return False


def generate_n2e(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_paths = {"LTE": TPL_N2E_LTE, "5G": TPL_N2E_5G, "MMBB": TPL_N2E_MMBB, "TRIMODE": TPL_N2E_TRIMODE}
    sa_conversion_nodes = []

    for row in mm_objs:
        node_type = n2e_node_type(row)
        primary = row.get("Node to be built as")
        if node_type is None:
            summary_rows.append({"Item": f"Node: {primary}", "Source": "node type detection", "Value": "skipped", "Note": "couldn't determine LTE/5G/MMBB/TRIMODE"})
            log(f"· {primary}: could not determine node type, skipped")
            continue
        tpl_path = tpl_paths[node_type]
        if not tpl_path.exists():
            summary_rows.append({"Item": f"Node: {primary}", "Source": f"N2E {node_type} template", "Value": "NOT FOUND", "Note": f"expected file not in templates/N2E/: {tpl_path.name} — this node type's template hasn't been uploaded yet"})
            log(f"✗ {primary}: N2E {node_type} template file not found, skipped")
            continue
        tpl_text = tpl_path.read_text(encoding="utf-8")
        tpl = fill_node_template_n2e(tpl_text, row, edp_index, user_id, date_str, controller_objs, summary_rows, log)
        outputs.append((f"{primary}_N2E_{node_type}_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, primary)
        if check_sa_conversion(ciq_wb, primary):
            sa_conversion_nodes.append(primary)

    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)
    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log, controller_edp_found)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    idl_outputs, idl_summary, idl_scope_lines = generate_idl_connections(
        ciq_wb, mm_objs, user_id, date_str, log, template_dir=TDIR_N2E_IDL, registry=N2E_IDL_TEMPLATE_REGISTRY)
    outputs += idl_outputs
    summary_rows += idl_summary

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []

    pre_line = "Nokia"
    post_parts = []
    for row in mm_objs:
        primary = row.get("Node to be built as")
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(primary).strip().upper() == str(e_name or "").strip().upper()
        r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) if is_lte_primary else find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        if not r:
            r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) or find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        hw = hw_string(r) or "NOT FOUND"
        if is_populated(e_name) and is_populated(g_name):
            secondary = g_name if is_lte_primary else e_name
            bbu_mode = row.get("BBU Mode")
            post_parts.append(f"{primary}(P)/{secondary}(S)({bbu_mode})({hw})")
        else:
            post_parts.append(f"{primary}({hw})")
    post_line = " + ".join(post_parts)

    # Carrier ADD — no Pre-checks for N2E, so every cell in the CIQ counts as an addition
    added = {}
    eutran_objs = sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    fiveg_objs = sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
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
            added[node] = cells

    classification = {"added": added, "moved": [], "deleted_sectors": {}, "deleted_nodes": [], "retuned": []}
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found)
    for node in sa_conversion_nodes:
        scope_of_work_lines.append(f"SA conversion.\t{node}")
    scope_of_work_lines += idl_scope_lines
    ngs_summary, ngs_scope_lines = generate_ngs_checks(ciq_wb, mm_objs, log)
    summary_rows += ngs_summary
    scope_of_work_lines += ngs_scope_lines

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# GENERATOR: NSB — no Pre-checks (Pre Configuration is always the fixed string "NA").
# Only 2 templates (MMBB, TRIMODE) — no LTE-only/5G-only variants, per the blueprint.
# Same confirmed placeholder mapping as N2E's MMBB/TRIMODE, minus the controller ID field
# (NSB templates don't fill xController_IDX directly — 6610 is purely the universal add-on here too).
# ============================================================

def nsb_node_type(row):
    bbu_mode = str(row.get("BBU Mode", "")).strip().upper()
    if bbu_mode == "MMBB":
        return "MMBB"
    if bbu_mode == "TMBB":
        return "TRIMODE"
    return None


def fill_node_template_nsb(template_text, row, edp_index, user_id, date_str, summary_rows, log):
    """NSB placeholder fill — confirmed identical to N2E's MMBB/TRIMODE mapping, minus xController_IDX."""
    primary = row.get("Node to be built as")
    gnb_name = row.get("gNodeB Name")
    tpl = template_text

    def sub(placeholder, value, note=""):
        nonlocal tpl
        if is_populated(value):
            tpl = tpl.replace(placeholder, str(value))
            summary_rows.append({"Item": f"{primary} · {placeholder}", "Source": note, "Value": value, "Note": ""})
            log(f"✓ {primary} · {placeholder} -> {value}")
        else:
            summary_rows.append({"Item": f"{primary} · {placeholder}", "Source": note, "Value": "NOT FOUND", "Note": ""})
            log(f"✗ {primary} · {placeholder} -> NOT FOUND")

    sub("xxSiteIdxx", primary, "Primary ID")
    sub("xxUserIDxx", user_id, "manual")
    sub("xDatex", date_str, "manual")
    sub("xgNBIdx", row.get("gNBId"), "CIQ · Mixed Mode Info")
    sub("xgNB_Namex", gnb_name, "CIQ · Mixed Mode Info")

    gnb_row = edp_row_for(edp_index, gnb_name)  # always matched by gNodeB Name, regardless of Primary/Secondary
    bearer_ip = edp_get(edp_index, gnb_row, "ipv6_enodeb_bearer_ip")
    sub("xsecondary_IPV6_ENODEB_BEARER_IPx", bearer_ip, "EDP · ipv6_enodeb_bearer_ip (by gNodeB Name)")
    sub("x5G_IPV6_SIAD_BEARER_IPx", edp_get(edp_index, gnb_row, "ipv6_siad_bearer_ip_def_router"), "EDP · ipv6_siad_bearer_ip_def_router")
    sub("x5G_Vlan_IDx", edp_get(edp_index, gnb_row, "bearer_enodeb_sb_vlan_id"), "EDP · bearer_enodeb_sb_vlan_id")

    return tpl


def generate_nsb(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_paths = {"MMBB": TPL_NSB_MMBB, "TRIMODE": TPL_NSB_TRIMODE}

    for row in mm_objs:
        node_type = nsb_node_type(row)
        primary = row.get("Node to be built as")
        if node_type is None:
            summary_rows.append({"Item": f"Node: {primary}", "Source": "node type detection", "Value": "skipped", "Note": "NSB only supports MMBB/TMBB — not LTE-only or 5G-only"})
            log(f"· {primary}: BBU Mode not MMBB/TMBB, skipped")
            continue
        tpl_path = tpl_paths[node_type]
        if not tpl_path.exists():
            summary_rows.append({"Item": f"Node: {primary}", "Source": f"NSB {node_type} template", "Value": "NOT FOUND", "Note": f"expected file not in templates/NSB/: {tpl_path.name}"})
            log(f"✗ {primary}: NSB {node_type} template file not found, skipped")
            continue
        tpl_text = tpl_path.read_text(encoding="utf-8")
        tpl = fill_node_template_nsb(tpl_text, row, edp_index, user_id, date_str, summary_rows, log)
        outputs.append((f"{primary}_NSB_{node_type}_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, primary)

    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)
    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log, controller_edp_found)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    idl_outputs, idl_summary, idl_scope_lines = generate_idl_connections(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += idl_outputs
    summary_rows += idl_summary
    ngs_summary, ngs_scope_lines = generate_ngs_checks(ciq_wb, mm_objs, log)
    summary_rows += ngs_summary

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []

    pre_line = "NA"
    post_parts = []
    for row in mm_objs:
        primary = row.get("Node to be built as")
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(primary).strip().upper() == str(e_name or "").strip().upper()
        r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) if is_lte_primary else find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        if not r:
            r = find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) or find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        hw = hw_string(r) or "NOT FOUND"
        if is_populated(e_name) and is_populated(g_name):
            secondary = g_name if is_lte_primary else e_name
            bbu_mode = row.get("BBU Mode")
            post_parts.append(f"{primary}(P)/{secondary}(S)({bbu_mode})({hw})")
        else:
            post_parts.append(f"{primary}({hw})")
    post_line = " + ".join(post_parts)

    # Carrier ADD — no Pre-checks for NSB, so every cell in the CIQ counts as an addition (same rule as N2E)
    added = {}
    eutran_objs = sheet_objs(ciq_wb["eUtran Parameters"]) if "eUtran Parameters" in ciq_wb.sheetnames else []
    fiveg_objs = sheet_objs(ciq_wb["5G Info"]) if "5G Info" in ciq_wb.sheetnames else []
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
            added[node] = cells

    classification = {"added": added, "moved": [], "deleted_sectors": {}, "deleted_nodes": [], "retuned": []}
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found)
    scope_of_work_lines += idl_scope_lines
    scope_of_work_lines += ngs_scope_lines

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


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
            push_siad_row(siad_rows, edp_index, site_id)

    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)
    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log, controller_edp_found)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    idl_outputs, idl_summary, idl_scope_lines = generate_idl_connections(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += idl_outputs
    summary_rows += idl_summary
    ngs_summary, ngs_scope_lines = generate_ngs_checks(ciq_wb, mm_objs, log)
    summary_rows += ngs_summary

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []
    pre_fibers_bytes = generate_pre_fibers(precheck_text)
    if pre_fibers_bytes and mm_objs:
        binary_outputs.append((f"Pre_Fibers_{mm_objs[0].get('Node to be built as','site')}.xlsx", pre_fibers_bytes))

    _, pre_nodes_found = extract_precheck_sectors(precheck_text)
    ciq_node_names = {r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as")}
    pre_line, post_line = generate_generic_pre_post(ciq_wb, mm_objs, precheck_text, pre_nodes_found | ciq_node_names)

    classification = classify_carriers(ciq_wb, mm_objs, precheck_text)
    radio_swaps = classify_radio_swaps(precheck_text, ciq_wb)
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found, radio_swaps)
    scope_of_work_lines += idl_scope_lines
    scope_of_work_lines += ngs_scope_lines
    pc_outputs, pc_summary, pc_scope_lines = generate_port_conversion_checks(ciq_wb, mm_objs, edp_index, precheck_text, log)
    outputs += pc_outputs
    summary_rows += pc_summary
    scope_of_work_lines += pc_scope_lines

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# GENERATOR: CENM (always cENM_TRIMODE template, for TMBB-mode nodes)
# ============================================================

def generate_cenm(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, precheck_text, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_cenm_tmbb = TPL_CENM.read_text(encoding="utf-8")
    tpl_cenm_mmbb = TPL_CENM_MMBB.read_text(encoding="utf-8") if TPL_CENM_MMBB.exists() else None

    tmbb_rows = [r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "TMBB"]
    mmbb_rows = [r for r in mm_objs if str(r.get("BBU Mode", "")).strip() == "MMBB"]
    if not tmbb_rows and not mmbb_rows:
        summary_rows.append({"Item": "Node identification", "Source": "CIQ · Mixed Mode Info", "Value": "NOT FOUND", "Note": "CENM expects a BBU Mode = TMBB or MMBB row"})
        return summary_rows, None, None, siad_rows, outputs, [], []

    for row in tmbb_rows:
        site_id = row.get("Node to be built as")
        tpl = fill_node_template(tpl_cenm_tmbb, row, edp_index, user_id, date_str, summary_rows, log)
        outputs.append((f"{site_id}_cENM_TMBB_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, site_id)

    for row in mmbb_rows:
        site_id = row.get("Node to be built as")
        if tpl_cenm_mmbb is None:
            summary_rows.append({"Item": f"Node: {site_id}", "Source": "CENM MMBB template", "Value": "NOT FOUND", "Note": f"expected file not in templates/MCA/: {TPL_CENM_MMBB.name}"})
            log(f"✗ {site_id}: CENM MMBB template file not found, skipped")
            push_siad_row(siad_rows, edp_index, site_id)
            continue
        tpl = fill_node_template(tpl_cenm_mmbb, row, edp_index, user_id, date_str, summary_rows, log)
        outputs.append((f"{site_id}_cENM_MMBB_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, site_id)

    for row in mm_objs:
        if str(row.get("BBU Mode", "")).strip() not in ("TMBB", "MMBB"):
            site_id = row.get("Node to be built as")
            summary_rows.append({"Item": f"Node: {site_id}", "Source": f"BBU Mode = {row.get('BBU Mode')}", "Value": "skipped for template", "Note": "not TMBB/MMBB — still included in Pre/Post and SIAD"})
            push_siad_row(siad_rows, edp_index, site_id)

    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)
    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log, controller_edp_found)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary
    idl_outputs, idl_summary, idl_scope_lines = generate_idl_connections(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += idl_outputs
    summary_rows += idl_summary
    ngs_summary, ngs_scope_lines = generate_ngs_checks(ciq_wb, mm_objs, log)
    summary_rows += ngs_summary

    binary_outputs = [(f"Final_Connections_{mm_objs[0].get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))] if mm_objs else []
    pre_fibers_bytes = generate_pre_fibers(precheck_text)
    if pre_fibers_bytes and mm_objs:
        binary_outputs.append((f"Pre_Fibers_{mm_objs[0].get('Node to be built as','site')}.xlsx", pre_fibers_bytes))

    _, pre_nodes_found = extract_precheck_sectors(precheck_text)
    ciq_node_names = {r.get("Node to be built as") for r in mm_objs if r.get("Node to be built as")}
    pre_line, post_line = generate_generic_pre_post(ciq_wb, mm_objs, precheck_text, pre_nodes_found | ciq_node_names)

    classification = classify_carriers(ciq_wb, mm_objs, precheck_text)
    radio_swaps = classify_radio_swaps(precheck_text, ciq_wb)
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found, radio_swaps)
    scope_of_work_lines += idl_scope_lines
    scope_of_work_lines += ngs_scope_lines
    pc_outputs, pc_summary, pc_scope_lines = generate_port_conversion_checks(ciq_wb, mm_objs, edp_index, precheck_text, log)
    outputs += pc_outputs
    summary_rows += pc_summary
    scope_of_work_lines += pc_scope_lines

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
    target_name = target.get("Node to be built as")
    source_row = next((r for r in delmove if str(r.get("Target Node name", "")).strip().upper() == str(target_name).strip().upper()), None)
    source_id = source_row.get("Source Node name") if source_row else (delmove[0].get("Source Node name") if delmove else None)

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

    lte_node = lte.get("Node to be built as")
    already_listed = {str(lte_node).strip().upper(), str(macro_primary).strip().upper()}
    source_is_distinct = is_populated(source_id) and str(source_id).strip().upper() not in already_listed

    pre_line = f"{lte_node}({lte_hw_pre}) + {macro_primary}(P)/{macro_secondary}(S)(MMBB)({macro_hw_pre})"
    if source_is_distinct:
        pre_line += f" + {source_id}({source_hw})"

    post_line = f"{lte_node}({lte_hw_post}) + {macro_primary}(P)/{macro_secondary}(S)(MMBB)({macro_hw_post}) + {target.get('Node to be built as')}({target_hw})"

    push_siad_row(siad_rows, edp_index, macro.get("Node to be built as"))
    push_siad_row(siad_rows, edp_index, lte.get("Node to be built as"))
    push_siad_row(siad_rows, edp_index, target.get("Node to be built as"))

    outputs.append((f"{target.get('Node to be built as')}_{out_name}_Filled.txt", tpl))

    controller_edp_found = push_all_controller_siad_rows(siad_rows, edp_index, controller_objs)
    add_outputs, add_summary = generate_6610(controller_objs, user_id, date_str, log, controller_edp_found)
    outputs += add_outputs
    summary_rows += add_summary
    dss_outputs, dss_summary, dss_labels = generate_dss(ciq_wb, mm_objs, user_id, date_str, log)
    outputs += dss_outputs
    summary_rows += dss_summary

    binary_outputs = [(f"Final_Connections_{target.get('Node to be built as','site')}.xlsx", generate_final_connections(ciq_wb, mm_objs))]
    pre_fibers_bytes = generate_pre_fibers(precheck_text)
    if pre_fibers_bytes:
        binary_outputs.append((f"Pre_Fibers_{target.get('Node to be built as','site')}.xlsx", pre_fibers_bytes))

    # CRAN's checklist now matches MCA/CENM (Carrier ADD/Delete/Move, Retune) — run the same
    # classification off the same Sector Del_Movement tab CRAN already reads for role detection.
    # Pre/Post Configuration stays CRAN's own distinct role-based format, untouched above.
    radio_swaps = classify_radio_swaps(precheck_text, ciq_wb)
    classification = classify_carriers(ciq_wb, mm_objs, precheck_text)
    classification["deleted_nodes"] = []  # every CRAN rehome vacates a source node — not a noteworthy anomaly here, unlike MCA/CENM
    scope_of_work_lines = format_scope_of_work(classification, controller_objs, dss_labels, controller_edp_found, radio_swaps)
    ngs_summary, ngs_scope_lines = generate_ngs_checks(ciq_wb, mm_objs, log)
    summary_rows += ngs_summary
    scope_of_work_lines += ngs_scope_lines
    pc_outputs, pc_summary, pc_scope_lines = generate_port_conversion_checks(ciq_wb, mm_objs, edp_index, precheck_text, log)
    outputs += pc_outputs
    summary_rows += pc_summary
    scope_of_work_lines += pc_scope_lines

    return summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_of_work_lines


# ============================================================
# CHECKS-PERFORMED PANEL (per-scope checklist, matched against the blueprint) —
# derives pass/fail per check from the already-computed scope_lines rather than
# threading new return values through every generate_* function.
# ============================================================

SCOPE_CHECKLIST = {
    "CRAN": ["Carrier ADD", "Carrier delete", "Carrier moving", "DSS checks", "Radio swap", "Retune", "6610 Present", "NGS Checks", "Port Conversion"],
    "MCA": ["Carrier ADD", "Carrier delete", "Carrier moving", "IDL Connections", "DSS checks", "Radio swap", "Retune", "6610 Present", "NGS Checks", "Port Conversion"],
    "CENM": ["Carrier ADD", "Carrier delete", "Carrier moving", "IDL Connections", "DSS checks", "Radio swap", "Retune", "6610 Present", "NGS Checks", "Port Conversion"],
    "N2E": ["Carrier ADD", "IDL Connections", "DSS checks", "6610 Present", "SA Conversion", "NGS Checks"],
    "NSB": ["Carrier ADD", "IDL Connections", "DSS checks", "NGS Checks", "6610 Present"],
}

# label -> function(line) -> True if that scope_lines entry counts as a "found/applicable" hit for this check
CHECK_MATCHERS = {
    "Carrier ADD": lambda l: l.startswith("Integration:"),
    "Carrier delete": lambda l: l.startswith("Deleted Node from ENM:") or l.startswith("Deleted Sector:"),
    "Carrier moving": lambda l: l.startswith("Moved Sectors:") and "CHECK CIQ" not in l,
    "IDL Connections": lambda l: l.startswith("IDL Connections:") and "not found" not in l.lower() and "could not determine" not in l.lower() and "missing" not in l.lower(),
    "DSS checks": lambda l: l.startswith("DSS Activation:"),
    "Radio swap": lambda l: l.startswith("Radio Swap on:"),
    "Retune": lambda l: l.startswith("Retune on:"),
    "6610 Present": lambda l: l.startswith("6610 Controller Integration:") or l.startswith("EDP is not published for the controller"),
    "SA Conversion": lambda l: l.startswith("SA conversion."),
    "NGS Checks": lambda l: l.startswith("NGS Activation on :"),
    "Port Conversion": lambda l: l.startswith("Port speed 1G to 10G conversion with MPST:"),
}

# (scope, check label) pairs that aren't wired into the tool yet — shown as not-run rather than a
# misleading "fail", since a fail here would otherwise look identical to "this site has none of these"
NOT_BUILT_YET = set()


def derive_check_status(top_scope, scope_lines):
    checklist = SCOPE_CHECKLIST.get(top_scope, [])
    lines = scope_lines or []
    results = []
    for label in checklist:
        matcher = CHECK_MATCHERS.get(label, lambda l: False)
        found = any(matcher(line) for line in lines)
        results.append({"label": label, "found": found, "not_built": (top_scope, label) in NOT_BUILT_YET})
    return results


def render_checks_panel(container, top_scope, scope_lines):
    """Reveals each check row in sequence, pauses, then fades out and is fully removed
    so the results below can take its place (per the requested checks -> disappear -> reveal outputs flow).
    `container` must be an st.empty() placeholder so it can be cleared afterward."""
    statuses = derive_check_status(top_scope, scope_lines)
    with container.container():
        with st.container(border=True):
            st.subheader("Checks Performed")
            rows_ph = st.empty()
            html_rows = []
            for s in statuses:
                if s["not_built"]:
                    icon_cls, icon_char, label_cls = "fail", "\u2717", "dim"
                elif s["found"]:
                    icon_cls, icon_char, label_cls = "pass", "\u2713", ""
                else:
                    icon_cls, icon_char, label_cls = "fail", "\u2717", "dim"
                suffix = " (not built yet)" if s["not_built"] else ""
                html_rows.append(
                    f'<div class="qkx-check-row"><div class="qkx-check-icon {icon_cls}">{icon_char}</div>'
                    f'<div class="qkx-check-label {label_cls}">{s["label"]}{suffix}</div></div>'
                )
                rows_ph.markdown('<div class="qkx-checklist">' + "".join(html_rows) + "</div>", unsafe_allow_html=True)
                time.sleep(0.18)
            time.sleep(0.6)  # let the completed checklist register before it fades away
            rows_ph.markdown(
                '<div class="qkx-checklist qkx-checks-fadeout">' + "".join(html_rows) + "</div>",
                unsafe_allow_html=True,
            )
        time.sleep(0.55)  # match the fade-out animation duration
    container.empty()  # fully removed — the results below then animate in via the global 3D reveal


# ============================================================
# UI
# ============================================================

if "qkx_page" not in st.session_state:
    st.session_state.qkx_page = "home"
if "qkx_scope" not in st.session_state:
    st.session_state.qkx_scope = None

def _qkx_go(page, scope=None):
    st.session_state.qkx_page = page
    if scope is not None:
        st.session_state.qkx_scope = scope
    st.rerun()

# ---- sticky top bar + shared styling (stays put on scroll — every page) ----
# MasTec brand palette used as ACCENTS only now: Prussian Blue #00284e, Endeavour #024ea4, Orange #ff5b24.
# Main content area is a light background — this is a deliberate fix: the previous dark theme fought
# with Streamlit's native (light) widget chrome and, combined with a buggy fill-mode:both entrance
# animation, left several result sections stuck at near-zero opacity (confirmed via screenshots).
st.markdown("""
<style>
  .stApp {
      background: linear-gradient(180deg, #eef3fa 0%, #f7f9fc 100%);
  }
  .qkx-topbar {
      position: sticky; top: 0; z-index: 999;
      display: flex; justify-content: space-between; align-items: center;
      padding: 0.9rem 1.75rem; margin: -1rem -1rem 1.5rem -1rem;
      background: linear-gradient(90deg, #011b36 0%, #012a4e 100%);
      border-bottom: 1px solid rgba(255,91,36,0.55);
      box-shadow: 0 4px 18px rgba(0,0,0,0.2);
  }
  .qkx-topbar .qkx-logo { font-size: 1.4rem; font-weight: 900; color: #ffffff; letter-spacing: 1px; }
  .qkx-topbar .qkx-logo span { color: #ffffff; }
  .qkx-topbar .qkx-credit { font-size: 0.78rem; color: #cfe0f5; text-align: right; line-height: 1.3; }

  .qkx-hero { text-align: center; margin: 1rem 0 2.5rem 0; }
  .qkx-hero h1 {
      font-size: 3.2rem; font-weight: 900; letter-spacing: 3px; margin-bottom: 0.3rem;
      color: #012a4e;
  }
  .qkx-hero p { color: #4a5b70; font-size: 1.02rem; }

  div[data-testid="stButton"] button {
      border-radius: 10px; font-weight: 700; border: 1.5px solid #013a6b;
      background: linear-gradient(135deg, #024ea4, #013a6b); color: #ffffff;
      box-shadow: 0 3px 8px rgba(1,42,78,0.25);
      transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
  }
  div[data-testid="stButton"] button:hover {
      border-color: #ff5b24; color: #ffffff; transform: translateY(-1px);
      box-shadow: 0 6px 14px rgba(255,91,36,0.35);
  }
  div[data-testid="stButton"] button:active { transform: translateY(0); }

  /* Bordered containers (st.container(border=True)) — clean light cards, not translucent-on-dark */
  div[data-testid="stVerticalBlockBorderWrapper"] {
      background: #ffffff !important;
      border: 1px solid #dde5ef !important;
      border-radius: 12px !important;
      box-shadow: 0 2px 10px rgba(1,42,78,0.06);
  }

  .qkx-checklist { margin: 0.5rem 0 0.5rem 0; }
  .qkx-check-row {
      display: flex; align-items: center; gap: 0.75rem;
      padding: 0.55rem 0.9rem; margin-bottom: 0.4rem; border-radius: 10px;
      background: #f3f6fb; border: 1px solid #e2e8f2;
      opacity: 0; transform: translateX(-14px);
      animation: qkxRowIn 0.4s ease forwards;
  }
  @keyframes qkxRowIn { to { opacity: 1; transform: translateX(0); } }
  .qkx-checks-fadeout { animation: qkxFadeOut 0.55s ease forwards; }
  @keyframes qkxFadeOut { to { opacity: 0; transform: translateY(-8px); } }
  .qkx-check-icon {
      width: 26px; height: 26px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
      font-weight: 900; font-size: 0.95rem; flex-shrink: 0;
  }
  .qkx-check-icon.pass { background: linear-gradient(135deg, #024ea4, #17c3a2); color: #ffffff; }
  .qkx-check-icon.fail { background: #d6dee8; color: #6b7c91; }
  .qkx-check-label { font-weight: 600; color: #1a2c40; }
  .qkx-check-label.dim { color: #7c8ba0; }

  /* Subtle, safe entrance motion for results — never hides content: no opacity/fill-mode tricks,
     so if the animation doesn't fire for any reason the element is simply static and fully visible. */
  @keyframes qkxSettle {
      0%   { transform: translateY(10px); }
      100% { transform: translateY(0); }
  }
  div[data-testid="stVerticalBlockBorderWrapper"] {
      animation: qkxSettle 0.35s ease-out;
  }
</style>
<div class="qkx-topbar">
  <div class="qkx-logo">MAS<span>TEC</span></div>
  <div class="qkx-credit">Made by <b>AKSHATHA KALLUR</b><br>Powered by <b>MASTEC</b></div>
</div>
""", unsafe_allow_html=True)

# ---- HOME ----
if st.session_state.qkx_page == "home":
    st.markdown("""
    <div class="qkx-hero">
      <h1>QUICKIX</h1>
      <p>SOW analysis and Integration templates generator</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption("Pre-existing sites — MCA, CENM, or CRAN rehome")
        if st.button("MCA", use_container_width=True, key="qkx_card_mca"):
            _qkx_go("family")
    with c2:
        st.caption("Nokia to Ericsson site integration")
        if st.button("N2E", use_container_width=True, key="qkx_card_n2e"):
            _qkx_go("input", "N2E")
    with c3:
        st.caption("New site build")
        if st.button("NSB", use_container_width=True, key="qkx_card_nsb"):
            _qkx_go("input", "NSB")

    st.divider()
    st.subheader("Instructions")
    st.markdown("""
    1. **Please select your SOW to continue.**
    2. **Upload** the CIQ and EDP & Pre-checks (optional) for the site
    3. **Enter** your User ID
    4. Click **Generate templates**
    5. **Review**, then **download**
    """)

# ---- FAMILY CHOICE (MCA / CENM / CRAN) ----
elif st.session_state.qkx_page == "family":
    if st.button("← Back"):
        _qkx_go("home")
    st.subheader("Choose scope")
    f1, f2, f3 = st.columns(3)
    with f1:
        if st.button("MCA", use_container_width=True, key="qkx_fam_mca"):
            _qkx_go("input", "MCA")
    with f2:
        if st.button("CENM", use_container_width=True, key="qkx_fam_cenm"):
            _qkx_go("input", "CENM")
    with f3:
        if st.button("CRAN", use_container_width=True, key="qkx_fam_cran"):
            _qkx_go("input", "CRAN")

# ---- INPUT PAGE (all scopes land here — same form + results as before) ----
elif st.session_state.qkx_page == "input":
    top_scope = st.session_state.qkx_scope
    if not top_scope:
        st.warning("No scope selected.")
        if st.button("← Back to home"):
            _qkx_go("home")
        st.stop()

    back_target = "family" if top_scope in ("MCA", "CENM", "CRAN") else "home"

    col_left, col_right = st.columns([2, 3])

    with col_left:
        if st.button("← Back"):
            _qkx_go(back_target)

        st.subheader(f"Inputs — {top_scope}")
        with st.container(border=True):
            cran_sub = None
            if top_scope == "CRAN":
                cran_sub = st.selectbox("CRAN scope", ["CRAN SA Rehome Trip 1", "CRAN SA Rehome Trip 2", "CRAN NSA Rehome"])

            ciq_file = st.file_uploader("CIQ (.xlsx / .xls)", type=["xlsx", "xls"])
            edp_file = st.file_uploader("EDP (.xlsx / .xls)", type=["xlsx", "xls"])
            pre_file = None
            if top_scope not in ("N2E", "NSB"):
                pre_file = st.file_uploader("Pre-checks (.pdf) — optional", type=["pdf"])
            c1, c2 = st.columns(2)
            with c1:
                user_id = st.text_input("User ID", placeholder="e.g. pr970b")
            with c2:
                date_str = st.text_input("Execution date (mmddyyyy)", value=date.today().strftime("%m%d%Y"))
            run = st.button("Generate templates →", type="primary", disabled=not (ciq_file and edp_file))

    if run:
        with col_left:
            log_card = st.container(border=True)
            with log_card:
                ph_log = st.empty()

        with col_right:
            ph_checks = st.empty()
            ph_prepost = st.container(border=True)
            ph_sow = st.container(border=True)
            ph_siad = st.container(border=True)
            ph_summary = st.container(border=True)
            ph_outputs = st.container(border=True)

        log_lines = []

        def log(msg):
            log_lines.append(msg)
            ph_log.code("\n".join(log_lines) or "Processing...", language=None)

        log("Starting...")

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

        log("Reading CIQ workbook...")
        try:
            ciq_wb = load_workbook_any(ciq_file.read(), ciq_file.name)
        except Exception as e:
            st.error(f"This CIQ couldn't be read as either .xlsx or legacy .xls. "
                      f"It may be corrupted, or its content doesn't match its extension — try re-saving "
                      f"it as .xlsx in Excel and re-uploading. Error detail: {e}")
            st.stop()
        if "Mixed Mode Info" not in ciq_wb.sheetnames:
            st.error('Could not find a "Mixed Mode Info" tab in the CIQ.')
            st.stop()
        mm_objs = sheet_objs(ciq_wb["Mixed Mode Info"])
        controller_objs = sheet_objs(ciq_wb["Controller Info"]) if "Controller Info" in ciq_wb.sheetnames else []

        log("Reading EDP workbook...")
        edp_bytes = edp_file.read()
        try:
            edp_wb = load_workbook_any(edp_bytes, edp_file.name)
        except Exception as e:
            st.error(f"This EDP couldn't be read (tried both .xlsx and legacy .xls handling). "
                      f"Try re-saving it as .xlsx in Excel and re-uploading. Error detail: {e}")
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
        elif top_scope == "N2E":
            summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = generate_n2e(
                ciq_wb, edp_index, controller_objs, mm_objs, uid, dstr, log)
        elif top_scope == "NSB":
            summary_rows, pre_line, post_line, siad_rows, outputs, binary_outputs, scope_lines = generate_nsb(
                ciq_wb, edp_index, controller_objs, mm_objs, uid, dstr, log)
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

        binary_outputs += get_universal_static_outputs(log)

        render_checks_panel(ph_checks, top_scope, scope_lines)

        with ph_prepost:
            if pre_line and post_line:
                st.subheader("Pre / Post configuration")
                st.code(f"Pre Configuration:  {pre_line}\nPost Configuration: {post_line}", language=None)

        with ph_sow:
            if scope_lines:
                st.subheader("Scope of work summary")
                st.code("\n".join(scope_lines_to_readable_text(scope_lines)), language=None)
                with st.expander("Copy tab-separated version for Excel/Notepad"):
                    st.text_area("Tab-separated (select all, copy, paste into Excel — lands in columns)",
                                  "\n".join(scope_lines), height=150, key="sow_raw")

        with ph_siad:
            if siad_rows:
                st.subheader("SIAD port assignment")
                st.dataframe(pd.DataFrame(siad_rows), use_container_width=True, hide_index=True)

        with ph_summary:
            st.subheader("Extraction summary")
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        with ph_outputs:
            if outputs:
                st.subheader("Generated output")
                for name, text in outputs:
                    unresolved = highlight_unresolved(text)
                    with st.expander(f"{name}  ({str(len(unresolved)) + ' unresolved' if unresolved else 'fully resolved'})"):
                        st.text_area("Preview", text, height=300, key=name)
                        st.download_button("Download .txt", text, file_name=name, key=f"dl_{name}")

            if binary_outputs:
                st.subheader("Excel outputs")
                for name, data in binary_outputs:
                    st.download_button(f"Download {name}", data, file_name=name, key=f"dl_bin_{name}")

            if outputs or binary_outputs:
                if len(outputs) + len(binary_outputs) > 1:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as zf:
                        for name, text in outputs:
                            zf.writestr(name, text)
                        for name, data in binary_outputs:
                            zf.writestr(name, data)
                    st.download_button("Download all as .zip", zip_buf.getvalue(), file_name="generated_templates.zip")
