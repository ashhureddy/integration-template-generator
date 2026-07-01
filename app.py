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

TEMPLATES_DIR = Path(__file__).parent / "templates"

SCOPE_MAP = {
    "MCA → MMBB Pre-existing": {
        "template": TEMPLATES_DIR / "MCA" / "MMBB_Pre-existing.txt",
        "generator": "mmbb",
    },
    "MCA → CRAN SA Rehome Trip 1": {
        "template": TEMPLATES_DIR / "MCA" / "CRAN_Trip1.txt",
        "generator": "cran_trip1",
    },
}

# ============================================================
# SHARED HELPERS (ported 1:1 from the validated browser-JS logic)
# ============================================================

def sheet_objs(ws):
    """openpyxl worksheet -> list of dicts, first row = headers."""
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
    """Find the EDP header row (has EDP_SITE_ID) and return {header_map, data_rows}."""
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
    objs = sheet_objs(ciq_wb[sheet_name])
    for r in objs:
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
    return re.sub(r"^baseband\s+", "", m.group(1).strip(), flags=re.I)


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


def highlight_unresolved(text):
    """Return a set of tokens still unresolved (for display), excluding decorative filler."""
    cands = re.findall(r"xx[A-Za-z0-9_]+xx|(?<!#)##[A-Za-z0-9_]+##(?!#)", text)
    return sorted(set(c for c in cands if not re.fullmatch(r"x+", c, re.I)))


# ============================================================
# GENERATOR: MCA -> MMBB Pre-existing
# ============================================================

def generate_mmbb(ciq_wb, edp_index, controller_objs, mm_objs, user_id, date_str, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_path = SCOPE_MAP["MCA → MMBB Pre-existing"]["template"]
    base_tpl = tpl_path.read_text(encoding="utf-8")
    has_6610 = any(str(r.get("Controller", "")).strip() == "6610" for r in controller_objs)

    for row in mm_objs:
        site_id = row.get("Node to be built as")
        bbu_mode = str(row.get("BBU Mode", "")).strip()
        if bbu_mode != "MMBB":
            summary_rows.append({"Item": f"Node: {site_id}", "Source": f"BBU Mode = {bbu_mode}", "Value": "skipped", "Note": "only MMBB wired up"})
            continue

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
                summary_rows.append({"Item": token, "Source": src, "Value": val, "Note": ""})
            else:
                summary_rows.append({"Item": token, "Source": src, "Value": "NOT FOUND", "Note": "left as placeholder"})
            log(f"{'✓' if val else '✗'} {token} -> {val or 'NOT FOUND'}")

        outputs.append((f"{site_id}_MMBB_Integration_Filled.txt", tpl))
        push_siad_row(siad_rows, edp_index, site_id)

    if has_6610:
        summary_rows.append({"Item": "6610", "Source": "CIQ · Controller Info", "Value": "present", "Note": "6610 template not bundled in this minimal pilot yet"})

    return summary_rows, None, None, siad_rows, outputs


# ============================================================
# GENERATOR: MCA -> CRAN SA Rehome Trip 1
# ============================================================

def generate_cran_trip1(ciq_wb, edp_index, mm_objs, user_id, date_str, precheck_text, log):
    summary_rows, siad_rows, outputs = [], [], []
    tpl_path = SCOPE_MAP["MCA → CRAN SA Rehome Trip 1"]["template"]
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
        return summary_rows, None, None, siad_rows, outputs

    if "Sector Del_Movement" not in ciq_wb.sheetnames:
        summary_rows.append({"Item": "Sector Del_Movement tab", "Source": "CIQ", "Value": "NOT FOUND", "Note": "required for Source CRAN"})
        return summary_rows, None, None, siad_rows, outputs

    delmove = sheet_objs(ciq_wb["Sector Del_Movement"])
    source_id = delmove[0].get("Source Node name") if delmove else None

    target_poles = {}
    if "5G Info" in ciq_wb.sheetnames:
        for r in sheet_objs(ciq_wb["5G Info"]):
            if str(r.get("gNB Name", "")).strip().upper() == str(target.get("gNodeB Name", "")).strip().upper():
                cell = r.get("NRCellDU")
                m = re.search(r"([A-C])_([12])$", str(cell or ""))
                if m:
                    target_poles[f"{m.group(1).upper()}_{m.group(2)}"] = cell

    for key in ["A_1", "A_2", "B_1", "B_2", "C_1", "C_2"]:
        token = f"xxTarget_SiteIdxx_Pole_N077{key}"
        val = target_poles.get(key)
        if val:
            tpl = tpl.replace(token, val)
            summary_rows.append({"Item": token, "Source": "CIQ · 5G Info · NRCellDU", "Value": val, "Note": ""})
        else:
            summary_rows.append({"Item": token, "Source": "CIQ · 5G Info · NRCellDU", "Value": "NOT PRESENT", "Note": "no cell at this sector/band"})
        log(f"{'✓' if val else '·'} {token} -> {val or 'not present'}")

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
        ("xxSiteIDxx", target.get("Node to be built as"), "ambiguous — defaulted to Target (generic log names only), VERIFY"),
        ("xxSource_SiteIdxx", source_id, "CIQ · Sector Del_Movement · Source Node name"),
        ("xxUserIDxx", user_id, "manual input"),
        ("xxDATExx", date_str, "manual input"),
        ("xxDatexx", date_str, "manual input"),
        ("xxdatexx", date_str, "manual input"),
    ]
    for token, val, src in fills:
        if val:
            tpl = tpl.replace(token, str(val))
            summary_rows.append({"Item": token, "Source": src, "Value": val, "Note": ""})
        else:
            summary_rows.append({"Item": token, "Source": src, "Value": "NOT FOUND", "Note": "left as placeholder"})
        log(f"{'✓' if val else '✗'} {token} -> {val or 'NOT FOUND'}")
    tpl = tpl.replace("xDatex", date_str)

    summary_rows.append({"Item": "xx5G_Cell_namexx / xxFDD_namexx", "Source": "not in this template", "Value": "n/a", "Note": "Trip-1 doesn't use these"})

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

    outputs.append((f"{target.get('Node to be built as')}_CRAN_Trip1_Filled.txt", tpl))
    return summary_rows, pre_line, post_line, siad_rows, outputs


# ============================================================
# UI — styled to match the existing MASTEC DSS Extractor tool
# ============================================================

st.markdown("<h1 style='text-align:center;color:#5b4fe0;margin-bottom:0;'>MASTEC</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align:center;'>📡 Integration Template Generator</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:gray;'>CIQ + EDP → filled AMOS/CMCLI templates, per scope of work</p>", unsafe_allow_html=True)
st.divider()

col_inputs, col_instructions = st.columns([3, 2])

with col_inputs:
    st.subheader("📤 Inputs")
    scope = st.selectbox("Scope of work", list(SCOPE_MAP.keys()))
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
    1. **Pick** the scope of work
    2. **Upload** the CIQ and EDP for the site (Pre-checks optional)
    3. **Enter** your User ID and execution date
    4. Click **Generate templates**
    5. **Review** the extraction summary, then **download**
    """)
    st.caption("Note: template files live in the `templates/` folder in this repo — updating a template there updates the live tool automatically.")

st.divider()
st.subheader("📊 Processing Log")
log_box = st.empty()
log_lines = []

def log(msg):
    log_lines.append(msg)
    log_box.code("\n".join(log_lines) or "Waiting for file upload...", language=None)

log_box.code("Waiting for file upload...", language=None)

if run:
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
        import xlrd
        # legacy .xls fallback via pandas -> in-memory openpyxl-like structure not needed;
        # simplest: read with pandas and wrap minimally
        st.error("This EDP is an old .xls file — legacy .xls support needs the `xlrd` package (already in requirements.txt). If this still fails, re-save the EDP as .xlsx and re-upload.")
        st.stop()
    edp_index = build_edp_index(edp_wb)
    if not edp_index:
        st.error('Could not locate the EDP header row (expected a column "EDP_SITE_ID" and "SITE_NAME").')
        st.stop()

    precheck_text = ""
    if pre_file:
        log("Extracting Pre-checks PDF text...")
        precheck_text = extract_pdf_text(pre_file.read())

    gen = SCOPE_MAP[scope]["generator"]
    if gen == "mmbb":
        summary_rows, pre_line, post_line, siad_rows, outputs = generate_mmbb(
            ciq_wb, edp_index, controller_objs, mm_objs, user_id or "xxUserIDxx", date_str or "xxDatexx", log)
    else:
        summary_rows, pre_line, post_line, siad_rows, outputs = generate_cran_trip1(
            ciq_wb, edp_index, mm_objs, user_id or "xxUserIDxx", date_str or "xxDatexx", precheck_text, log)

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

    if outputs:
        st.subheader("📄 Generated output")
        for name, text in outputs:
            unresolved = highlight_unresolved(text)
            with st.expander(f"{name}  ({'⚠ ' + str(len(unresolved)) + ' unresolved' if unresolved else '✓ fully resolved'})"):
                st.text_area("Preview", text, height=300, key=name)
                st.download_button("⬇ Download .txt", text, file_name=name, key=f"dl_{name}")

        if len(outputs) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for name, text in outputs:
                    zf.writestr(name, text)
            st.download_button("⬇ Download all as .zip", zip_buf.getvalue(), file_name="generated_templates.zip")

st.divider()
st.caption("Made by **AKSHATHA KALLUR** | Powered by **MASTEC** | © 2026")
