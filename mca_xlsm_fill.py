import io
import openpyxl
from mca_row_map import ROW_MAP


def fill_legacy_mca(template_path, values):
    """values: dict of row_key -> either a plain value (single-column rows) or a list of
    column values [colB, colC, colD, ...] (multi-column rows), or for multi-slot items,
    a list of such value-lists (one per used slot).
    checked: dict of row_number -> True/False, explicit checkbox states to set.
    Always preserves the VBA project (keep_vba=True) so the existing macro (Report_MCA
    generation + DeleteMatchingRows cleanup) still works when the user opens the file and
    runs it themselves — confirmed: Report_MCA should NOT be pre-written by this function;
    the macro reads Legacy_MCA's cell VALUES (not the checkbox shapes' own rendering state)
    to decide what to include, so as long as Legacy_MCA's data is correct, running the
    macro produces the correct Report_MCA regardless of anything else.
    Returns the filled workbook as raw bytes (in-memory) — no filesystem output path at all,
    since a hardcoded sandbox-only path previously caused a misleading 'template not found'
    error on deployment (the actual failure was writing the OUTPUT, not reading the template)."""
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    ws = wb["Legacy_MCA"]

    def set_row(row_num, checked, col_value_pairs=None):
        ws.cell(row=row_num, column=1, value=checked)
        if col_value_pairs:
            for col, v in col_value_pairs:
                if v is not None:
                    ws.cell(row=row_num, column=col, value=v)

    for row_num, checked, col_value_pairs in values.get("row_writes", []):
        set_row(row_num, checked, col_value_pairs)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
