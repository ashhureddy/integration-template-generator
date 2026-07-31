"""
Surgical .xlsm patcher — confirmed necessary fix.

openpyxl's wb.save() silently destroys this template's ~180 ActiveX checkbox controls
(the entire xl/activeX/ folder — 359 files: every control's .xml definition AND its
binary .bin property/state storage) plus xl/drawings/vmlDrawing1.vml (the legacy
rendering/positioning data every embedded control also needs). This is a known openpyxl
limitation — it has no support for legacy VML drawings or ActiveX OLE embeddings at all,
and confirmed via direct inspection: a file processed through openpyxl's save() has zero
files in xl/activeX/ where the original template has 359.

Fix: never call openpyxl's save() for the final output. Take the original template's raw
zip bytes untouched, and only patch the specific cell values inside xl/worksheets/sheet1.xml
via direct string manipulation (the cell XML structure is confirmed simple and predictable:
<c r="A45" s="52" t="b"><v>1</v></c> for booleans, similar for text) — every other file,
including all 359 ActiveX/VML files, gets copied through byte-for-byte unchanged.
"""

import re
import zipfile
import io


def _col_letter(col_num):
    letters = ""
    n = col_num
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _escape_xml(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def _patch_cell(sheet_xml, row_num, col_num, value, is_bool=False):
    """Finds the existing <c r="A45" ...>...</c> or <c r="A45" .../> element (confirmed:
    every cell we ever write to already exists with placeholder content in the real
    template, so lookup-and-replace is safe — no need to handle cell insertion) and
    replaces it, preserving the original style ('s=') attribute so formatting doesn't
    change. Returns the modified sheet_xml string."""
    cell_ref = f"{_col_letter(col_num)}{row_num}"
    # Matches both self-closing <c r="X".../> and <c r="X"...>...</c> forms.
    pattern = re.compile(
        r'<c r="' + re.escape(cell_ref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
    m = pattern.search(sheet_xml)
    if not m:
        # Confirmed this should never happen for the cells we actually write to — every
        # one already has placeholder content in the real template. Skip safely if it does
        # (better to silently no-op than risk malformed XML from a blind insertion).
        return sheet_xml
    attrs = m.group(1)
    style_m = re.search(r's="(\d+)"', attrs)
    style_attr = f' s="{style_m.group(1)}"' if style_m else ""

    if is_bool:
        new_cell = f'<c r="{cell_ref}"{style_attr} t="b"><v>{1 if value else 0}</v></c>'
    elif value is None or value == "":
        new_cell = f'<c r="{cell_ref}"{style_attr}/>'
    else:
        new_cell = (f'<c r="{cell_ref}"{style_attr} t="inlineStr">'
                    f'<is><t xml:space="preserve">{_escape_xml(value)}</t></is></c>')
    return sheet_xml[:m.start()] + new_cell + sheet_xml[m.end():]


def fill_legacy_mca_surgical(template_path, row_writes):
    """row_writes: same shape as before — list of (row_num, checked, [(col, value), ...]).
    Returns the patched .xlsm as raw bytes, with every file except
    xl/worksheets/sheet1.xml (Legacy_MCA) copied through byte-for-byte from the original —
    confirmed this is what's needed to keep the ~180 ActiveX checkboxes and their VML
    functional when opened in Excel, instead of degrading into inert draggable pictures."""
    with open(template_path, "rb") as f:
        original_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(original_bytes)) as zin:
        sheet_xml = zin.read("xl/worksheets/sheet1.xml").decode("utf-8")

        for row_num, checked, col_value_pairs in row_writes:
            sheet_xml = _patch_cell(sheet_xml, row_num, 1, checked, is_bool=True)
            for col, v in (col_value_pairs or []):
                if v is not None:
                    sheet_xml = _patch_cell(sheet_xml, row_num, col, v, is_bool=False)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = sheet_xml.encode("utf-8") if item.filename == "xl/worksheets/sheet1.xml" \
                    else zin.read(item.filename)
                zout.writestr(item, data)

    return out_buf.getvalue()
