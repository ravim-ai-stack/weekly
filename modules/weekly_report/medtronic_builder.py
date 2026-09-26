"""Builds the Medtronic weekly status Word report, which has a completely
different format from the PowerPoint-based templates: a single growing
.docx where every generation PREPENDS a brand-new dated page (Activities
Completed This Week / Activities In Process + Next Action / Activities To
Be Started Next Week) ahead of whatever was generated before, separated by
an explicit page break. Nothing from earlier weeks is ever edited or
removed - source_path (the previous week's output, or the original template
on the very first run) is never modified; the result is always written to a
new output_path. Every item handed in is kept in full - none are dropped or
truncated to fit available space."""

import copy
from datetime import datetime

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _mmddyyyy(iso_date: str) -> str:
    return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%m/%d/%Y")


def _find_title_paragraph(document):
    for p in document.paragraphs:
        if "WEEKLY STATUS REPORT" in p.text.upper():
            return p
    raise RuntimeError("Could not find the 'WEEKLY STATUS REPORT' title paragraph.")


def _validate_table_headers(completed_table, process_table, next_week_table):
    def header(table, r, c):
        return table.cell(r, c).text.strip().upper()

    if "ACTIVITIES COMPLETED" not in header(completed_table, 0, 0):
        raise RuntimeError("Unexpected header for the Activities Completed This Week table.")
    if "ACTIVITIES IN PROCESS" not in header(process_table, 0, 0) or "NEXT ACTION" not in header(process_table, 0, 1):
        raise RuntimeError("Unexpected headers for the Activities In Process / Next Action table.")
    if "STARTED NEXT WEEK" not in header(next_week_table, 0, 0):
        raise RuntimeError("Unexpected header for the Activities To Be Started Next Week table.")


def _collect_block_elements(start_el, end_el):
    """Walks document-order siblings from start_el (the title paragraph)
    through end_el (the 3rd table) inclusive, so the whole week "block" -
    title, spacers, date, and all 3 tables, whatever the exact spacing - is
    captured as one ordered unit to clone."""
    elements = []
    el = start_el
    while el is not None:
        elements.append(el)
        if el is end_el:
            return elements
        el = el.getnext()
    raise RuntimeError("Could not walk from the title paragraph to the 3rd table - unexpected document structure.")


def _find_date_paragraph(elements):
    for el in elements:
        if el.tag == qn("w:p"):
            text = "".join(t.text or "" for t in el.iter(qn("w:t")))
            if text.strip().startswith("Date:"):
                return el
    raise RuntimeError("Could not find the 'Date:' paragraph in the cloned block.")


def _set_paragraph_text(p_element, new_text: str):
    """Keeps the first <w:r> in a paragraph (and its formatting), sets its
    text to new_text, removes any other runs - mirrors _clear_extra_runs in
    pptx_builder.py but for a docx <w:p>."""
    runs = p_element.findall(qn("w:r"))
    if not runs:
        return
    first_run = runs[0]
    t = first_run.find(qn("w:t"))
    if t is None:
        t = OxmlElement("w:t")
        first_run.append(t)
    t.set(qn("xml:space"), "preserve")
    t.text = new_text
    for extra in runs[1:]:
        p_element.remove(extra)


def _nth_tc(tbl_element, row_idx: int, col_idx: int):
    trs = tbl_element.findall(qn("w:tr"))
    tcs = trs[row_idx].findall(qn("w:tc"))
    return tcs[col_idx]


def _set_bullet_list_cell(tc_element, items: list):
    """Rebuilds a table cell's content as one bulleted paragraph per item,
    reusing the first existing bulleted paragraph's <w:pPr> (which carries
    the ListParagraph style + bullet numbering reference) as the formatting
    template. Every item is kept in full - none are dropped or truncated."""
    items = [str(i).strip() for i in items if str(i).strip()]
    if not items:
        items = ["None"]

    paragraphs = tc_element.findall(qn("w:p"))
    template_pPr = None
    for p in paragraphs:
        if p.findall(qn("w:r")):
            template_pPr = p.find(qn("w:pPr"))
            break

    for p in paragraphs:
        tc_element.remove(p)

    for text in items:
        new_p = OxmlElement("w:p")
        if template_pPr is not None:
            new_p.append(copy.deepcopy(template_pPr))
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        run.append(t)
        new_p.append(run)
        tc_element.append(new_p)


def _make_page_break_paragraph():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def _milestone_line(milestone: dict) -> str:
    area = str(milestone.get("area", "")).strip()
    task = str(milestone.get("task", "")).strip()
    remarks = str(milestone.get("remarks", "")).strip()
    label = " - ".join(p for p in (area, task) if p)
    if label and remarks:
        return f"{label}: {remarks}"
    return remarks or label


def _genuine_upcoming_activities(report: dict) -> list:
    """"Activities To Be Started Next Week" should only hold genuinely new/
    not-yet-started work - excludes the "<area> - Continue to Support."
    lines that reclassify_milestones_by_status/_carry_continuing_milestones_
    forward append to upcoming_activities for still-ongoing work, since that
    belongs under "Next Action" instead."""
    return [
        u for u in report.get("upcoming_activities", [])
        if str(u).strip() and not str(u).rstrip().endswith("Continue to Support.")
    ]


def build_medtronic_report(source_path: str, output_path: str, week: dict, report: dict):
    """Opens source_path (the previous week's generated output, or the
    original template on the first run), clones its topmost week "block" as
    a formatting template, fills the clone with this week's data, and
    prepends it - followed by an explicit page break - ahead of everything
    that was already there. source_path is never modified; the result is
    always written to output_path."""
    document = Document(source_path)

    title_p = _find_title_paragraph(document)
    tables = document.tables
    if len(tables) < 3:
        raise RuntimeError(
            "Expected at least 3 tables (Activities Completed / In Process+Next Action / "
            "To Be Started Next Week) in the Medtronic document."
        )
    completed_table, process_table, next_week_table = tables[0], tables[1], tables[2]
    _validate_table_headers(completed_table, process_table, next_week_table)

    block_elements = _collect_block_elements(title_p._p, next_week_table._tbl)
    new_elements = [copy.deepcopy(el) for el in block_elements]

    new_date_p = _find_date_paragraph(new_elements)
    _set_paragraph_text(new_date_p, f"Date: {_mmddyyyy(week['start'])} - {_mmddyyyy(week['end'])}")

    new_tables = [el for el in new_elements if el.tag == qn("w:tbl")]
    if len(new_tables) != 3:
        raise RuntimeError("Could not locate the 3 cloned tables in the new block.")
    new_completed_tbl, new_process_tbl, new_next_week_tbl = new_tables

    _set_bullet_list_cell(_nth_tc(new_completed_tbl, 1, 0), report.get("completed_tasks", []))

    process_lines = [_milestone_line(m) for m in report.get("milestones", [])]
    _set_bullet_list_cell(_nth_tc(new_process_tbl, 1, 0), process_lines)
    _set_bullet_list_cell(_nth_tc(new_process_tbl, 1, 1), report.get("next_action", []))

    _set_bullet_list_cell(_nth_tc(new_next_week_tbl, 1, 0), _genuine_upcoming_activities(report))

    anchor = title_p._p
    for el in new_elements:
        anchor.addprevious(el)
    anchor.addprevious(_make_page_break_paragraph())

    document.save(output_path)
    return output_path
