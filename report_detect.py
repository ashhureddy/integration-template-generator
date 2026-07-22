"""
Standalone test module for MCA Integration Report detection logic.
Reuses existing app.py functions (pre_hw_string, hw_string, extract_precheck_sectors,
find_row_by_name) without duplicating their internals — only NEW logic lives here.
"""


def detect_node_board_changes(app, ciq_wb, mm_objs, precheck_text):
    """Returns (new_nodes, board_swaps).
    new_nodes: [node_name, ...] — present in CIQ (Post) but not confirmed anywhere in Pre-checks.
    board_swaps: [(node_name, pre_board, post_board), ...] — present in both, but the board
    model string differs between Pre and Post (same node, hardware swapped)."""
    _, pre_nodes_set = app.extract_precheck_sectors(precheck_text)
    new_nodes = []
    board_swaps = []
    for row in mm_objs:
        primary = row.get("Node to be built as")
        e_name, g_name = row.get("eNodeB Name"), row.get("gNodeB Name")
        is_lte_primary = str(primary).strip().upper() == str(e_name or "").strip().upper()
        r = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) if is_lte_primary else app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        if not r:
            r = app.find_row_by_name(ciq_wb, "eNB Info", "eNodeB Name", e_name) or app.find_row_by_name(ciq_wb, "gNB Info", "gNodeB Name", g_name)
        post_board = app.hw_string(r)
        if primary not in pre_nodes_set:
            new_nodes.append(primary)
        else:
            pre_board = app.pre_hw_string(precheck_text, primary)
            if pre_board and post_board and pre_board.strip() != post_board.strip():
                board_swaps.append((primary, pre_board, post_board))
    return new_nodes, board_swaps


def detect_fdd_renaming(app, ciq_wb):
    """Returns [(node, old_cell_name, new_cell_name), ...] — Sector Del_Movement rows where
    Source Node == Target Node (same physical node), but the cell name itself changed."""
    out = []
    if "Sector Del_Movement" not in ciq_wb.sheetnames:
        return out
    for row in app.sheet_objs(ciq_wb["Sector Del_Movement"]):
        src_node, tgt_node = row.get("Source Node name"), row.get("Target Node name")
        src_cell, tgt_cell = row.get("Source Sector"), row.get("Target Sector")
        if not (app.is_populated(src_node) and app.is_populated(tgt_node) and app.is_populated(src_cell) and app.is_populated(tgt_cell)):
            continue
        if str(src_node).strip() == str(tgt_node).strip() and str(src_cell).strip() != str(tgt_cell).strip():
            out.append((str(src_node).strip(), str(src_cell).strip(), str(tgt_cell).strip()))
    return out
