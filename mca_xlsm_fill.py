import openpyxl
from mca_row_map import ROW_MAP


def fill_legacy_mca(template_path, output_path, values):
    """values: dict of row_key -> either a plain value (single-column rows) or a list of
    column values [colB, colC, colD, ...] (multi-column rows), or for multi-slot items,
    a list of such value-lists (one per used slot).
    checked: dict of row_number -> True/False, explicit checkbox states to set.
    Always preserves the VBA project (keep_vba=True) so the existing macro (Report_MCA
    generation + DeleteMatchingRows cleanup) still works when the user opens the file."""
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

    wb.save(output_path)
    return output_path
