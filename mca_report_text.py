def _line_from_result(item, choice_manual_extra=None):
    """Builds the prose Completed/Pending line(s) for one checked item, reusing whatever
    detect() already found. Multi-instance items (Radio Swap, Moved Sectors, Retune, etc.)
    render EVERY detected instance, not just the first — confirmed bug: a site with 5 separate
    Radio Swap events was only showing 1."""
    result = item.get("result") or {}
    label = item["label"]
    if result.get("lines"):
        out_lines = []
        for raw_line in result["lines"]:
            parts = [p.strip() for p in raw_line.split("\t") if p.strip()]
            rest = " ".join(parts[1:])
            out_lines.append(f"{parts[0]} {rest}.".replace("  ", " "))
        return "\n".join(out_lines)
    if result.get("fdd"):
        return "\n".join(f"FDD Renaming on: {node} From: {old_name} To: {new_name}." for node, old_name, new_name in result["fdd"])
    fill = result.get("fill", {})
    bits = []
    if fill.get("nodes"):
        bits.append("|".join(fill["nodes"]))
    if fill.get("controller"):
        bits.append(fill["controller"])
    if fill.get("bands"):
        bits.append(fill["bands"])
    if choice_manual_extra:
        bits.extend([b for b in choice_manual_extra if b])
    return f"{label} {' | '.join(bits)}.".strip()


def build_mca_report_text(mm_objs, checklist_results, choices, header_fields, stakeholder_by_key=None):
    """header_fields: {mic, market, status, site_name, fa_code, site_ids, sow, iwm_details,
    pre_configuration, current_configuration, post_configuration, wll_node, controller_id,
    software_version, gs_version, idl_build_type, idle, idly, switch, slot_port}.
    stakeholder_by_key: item key -> chosen stakeholder tag, for items placed in Pending."""
    lines = []
    lines.append("Subject")
    lines.append(f"{header_fields.get('mic','MIC')} | {header_fields.get('market','')} | "
                  f"{header_fields.get('status','')} | {header_fields.get('site_name','')} | "
                  f"{header_fields.get('fa_code','')} | {header_fields.get('site_ids','')} | "
                  f"{header_fields.get('sow','')}")
    lines.append("")
    lines.append("IWM Details:")
    lines.append(header_fields.get("iwm_details", ""))
    lines.append("")
    lines.append("Configuration")
    lines.append(f"Pre Configuration : {header_fields.get('pre_configuration','')}")
    if (header_fields.get("current_configuration") or "").strip():
        lines.append(f"Current Configuration : {header_fields['current_configuration']}")
    lines.append(f"Post Configuration : {header_fields.get('post_configuration','')}")
    if (header_fields.get("wll_node") or "").strip():
        lines.append(f"WLL  node : {header_fields['wll_node']}")
    lines.append(f"6610 Controller : {header_fields.get('controller_id','')}")
    lines.append(f"Software version: {header_fields.get('software_version','')}")
    lines.append(f"GS Version: {header_fields.get('gs_version','')}")
    lines.append("")
    lines.append("IDL Connections")
    if header_fields.get("idl_build_type"):
        lines.append(f"Build Type : {header_fields['idl_build_type']}")
    if (header_fields.get("idle") or "").strip():
        lines.append(header_fields["idle"])
    if (header_fields.get("idly") or "").strip():
        lines.append(header_fields["idly"])
    if (header_fields.get("switch") or "").strip():
        lines.append(header_fields["switch"])
    if (header_fields.get("slot_port") or "").strip():
        lines.append(header_fields["slot_port"])
    lines.append("")

    completed_lines, pending_lines = [], []
    for item in checklist_results:
        key = item["key"]
        choice = choices.get(key, {})
        checked = choice.get("checked", item["checked_by_default"])
        if not checked:
            continue
        section = choice.get("section", item["section"])
        text = _line_from_result(item, choice.get("manual_extra"))
        if section == "pending":
            stakeholder = (stakeholder_by_key or {}).get(key, item.get("stakeholder", ""))
            if stakeholder:
                tagged = "\n".join(f"{ln} ({stakeholder})" for ln in text.split("\n"))
                pending_lines.append(tagged)
            else:
                pending_lines.append(text)
        else:
            completed_lines.append(text)

    lines.append("Completed:")
    lines.extend(completed_lines)
    if (choices.get("additional_completed", {}).get("text") or "").strip():
        lines.append(choices["additional_completed"]["text"])
    lines.append("")
    lines.append("Pending:")
    lines.extend(pending_lines)
    if (choices.get("additional_pending", {}).get("text") or "").strip():
        lines.append(choices["additional_pending"]["text"])
    lines.append("")
    lines.append("Pre-Existing Issues:")
    lines.append(choices.get("pre_existing_issues_text", ""))
    lines.append("")
    lines.append("Notes:")
    for note_key in ("notes_final_port_config", "notes_nr_verified", "notes_cpri_sfp",
                     "notes_no_external_alarms", "notes_mme_config", "notes_monitored", "notes_not_monitored"):
        if choices.get(note_key, {}).get("checked"):
            text = choices[note_key].get("text", "")
            lines.append(text)
    if (choices.get("notes_generic_text") or "").strip():
        lines.append(choices["notes_generic_text"])

    return "\n".join(lines)
