"""Edits ONLY slide 1 (engagement date) and the status-report slide of a
DataPattern weekly status report template: title date range, Completed
Tasks, Milestone & Supporting Activities table, Upcoming Activities and
Risk & Issues. Every other slide is left untouched because we only ever
touch the XML nodes for these specific shapes. Every item handed in is kept
in full - none are dropped or truncated to fit available space.

Completed Tasks / Milestone / Upcoming Activities can hold any number of
items. Each of those three boxes first tries to fit everything by shrinking
its font down to MIN_FONT_PT; if it still doesn't fit even at that floor
size, the overflow is moved onto a "(continued)" slide cloned right after
the current one (holding only that one box, so it gets the whole page),
repeating for as many continuation slides as needed. Available height for a
box is computed from the real slide geometry - the gap to whatever shape
sits below it, or the slide edge - so nothing ever overlaps."""

import copy
import re

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt

SLIDE_INDEX = 8  # slide 9, 0-indexed
TITLE_SLIDE_INDEX = 0  # slide 1, 0-indexed

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

SLIDE1_DATE_RE = re.compile(r"\d{1,2}-[A-Za-z]+-\d{4}")

EMU_PER_PT = 12700
MIN_FONT_PT = 8.0
FONT_STEP_PT = 0.5
CELL_H_PADDING = 2 * 91440  # default left+right cell margins
CELL_V_PADDING = 2 * 45720  # default top+bottom cell margins

# header_contains strings used to locate each of the status-slide's four
# boxes, shared between the normal single-slide path and the
# continuation-slide overflow path so both agree on which shape is which.
SECTION_HEADERS = {
    "completed": "Highlights & Accomplishments",
    "milestone": "Milestone & Supporting Activities",
    "upcoming": "Upcoming Activities",
    "risk": "Risk & Issues",
}


def _find_table_shape(slide, header_contains):
    for shape in slide.shapes:
        if shape.has_table:
            try:
                header_text = shape.table.cell(0, 0).text.strip()
            except Exception:
                continue
            if header_contains.lower() in header_text.lower():
                return shape
    raise RuntimeError(f"Could not find a table containing header '{header_contains}' on the status report slide.")


def _find_title_shape(slide):
    for shape in slide.shapes:
        if shape.has_text_frame and "Weekly Status Report" in shape.text_frame.text:
            return shape
    raise RuntimeError("Could not find the title placeholder on the status report slide.")


def _set_date_range(slide, date_range_text: str):
    shape = _find_title_shape(slide)
    txBody = shape.text_frame._txBody
    for r in txBody.iter(qn("a:r")):
        t = r.find(qn("a:t"))
        if t is not None and t.text and "Week of" in t.text:
            t.text = re.sub(r"Week of.*$", f"Week of {date_range_text}", t.text)
            return
    raise RuntimeError("Could not find the 'Week of ...' run in the title placeholder.")


def _replace_across_runs(paragraph_el, pattern, replacement_text) -> bool:
    """Some templates (e.g. ACT's) split a date like '14-August-2026' across
    two <a:r> runs (PowerPoint splits runs wherever formatting was tweaked by
    hand), so the pattern can't be found inside any single run's text. This
    searches the paragraph's full concatenated text instead, then writes the
    replacement into the first run touched by the match and blanks out the
    matched portion of any other runs it spans, leaving surrounding text
    (and each run's own formatting) untouched."""
    runs = paragraph_el.findall(qn("a:r"))
    texts = []
    for r in runs:
        t = r.find(qn("a:t"))
        texts.append(t.text or "" if t is not None else "")
    match = pattern.search("".join(texts))
    if not match:
        return False

    start, end = match.span()
    pos = 0
    replaced = False
    for r, text in zip(runs, texts):
        run_start, run_end = pos, pos + len(text)
        pos = run_end
        overlap_start = max(start, run_start)
        overlap_end = min(end, run_end)
        if overlap_start >= overlap_end:
            continue
        prefix = text[: overlap_start - run_start]
        suffix = text[overlap_end - run_start :]
        new_text = prefix + (replacement_text if not replaced else "") + suffix
        replaced = True
        t = r.find(qn("a:t"))
        if t is None:
            t = r.makeelement(qn("a:t"), {})
            r.append(t)
        t.text = new_text
    return replaced


def _set_title_slide_date(prs, slide1_date_text: str):
    """Updates the 'Engagement Report - DD-Month-YYYY' date on slide 1 only."""
    slide = prs.slides[TITLE_SLIDE_INDEX]
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        txBody = shape.text_frame._txBody
        for p in txBody.findall(qn("a:p")):
            if _replace_across_runs(p, SLIDE1_DATE_RE, slide1_date_text):
                return
    raise RuntimeError("Could not find the engagement date on slide 1.")


def _clear_extra_runs(paragraph_el, keep_text: str):
    """Keeps the first <a:r> in a paragraph, sets its text, removes the rest."""
    runs = paragraph_el.findall(qn("a:r"))
    if not runs:
        # A genuinely empty template cell (a fresh, never-typed-into
        # PowerPoint cell only has <a:endParaRPr>, no <a:r> to reuse) -
        # synthesize one so the text isn't silently dropped, carrying over
        # whatever font size/formatting that endParaRPr specifies.
        end_pr = paragraph_el.find(qn("a:endParaRPr"))
        run = paragraph_el.makeelement(qn("a:r"), {})
        if end_pr is not None:
            rpr = copy.deepcopy(end_pr)
            rpr.tag = qn("a:rPr")
            run.append(rpr)
        t = run.makeelement(qn("a:t"), {})
        t.text = keep_text
        run.append(t)
        if end_pr is not None:
            end_pr.addprevious(run)
        else:
            paragraph_el.append(run)
        return
    first_run = runs[0]
    t = first_run.find(qn("a:t"))
    if t is None:
        t = first_run.makeelement(qn("a:t"), {})
        first_run.append(t)
    t.text = keep_text
    for extra in runs[1:]:
        paragraph_el.remove(extra)


def set_cell_simple_text(cell, text: str):
    """Collapses a cell to its first paragraph / first run and sets the full
    text, preserving that run's formatting (font, size, bold, color). The
    text is never truncated - every point must reach the deck intact, even
    if the table grows into the space below it."""
    txBody = cell._tc.find(qn("a:txBody"))
    paragraphs = txBody.findall(qn("a:p"))
    if not paragraphs:
        return
    first_p = paragraphs[0]
    _clear_extra_runs(first_p, text)
    for extra_p in paragraphs[1:]:
        txBody.remove(extra_p)


def _set_run_font_pt(run_el, font_pt):
    rpr = run_el.find(qn("a:rPr"))
    if rpr is not None:
        rpr.set("sz", str(int(round(font_pt * 100))))


def _set_cell_font_pt(tc_el, font_pt):
    txBody = tc_el.find(qn("a:txBody"))
    for p in txBody.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            _set_run_font_pt(r, font_pt)


TASK_PREFIX_RE = re.compile(r"^(Task\s*(?:No\.?|#)?\s*[-–:]?\s*\d+\s*:?\s*)(.*)$", re.IGNORECASE)


def _split_prefix(item_text: str):
    """Splits 'Task No - 82: did the thing' into ('Task No - 82: ', 'did the thing').
    Returns (None, item_text) if no task-number prefix is present."""
    match = TASK_PREFIX_RE.match(item_text.strip())
    if match:
        prefix, rest = match.groups()
        return prefix.strip() + " ", rest.strip()
    return None, item_text.strip()


def _make_run(template_run, text, font_pt=None):
    new_run = copy.deepcopy(template_run)
    t = new_run.find(qn("a:t"))
    if t is None:
        t = new_run.makeelement(qn("a:t"), {})
        new_run.append(t)
    t.text = text
    if font_pt is not None:
        _set_run_font_pt(new_run, font_pt)
    return new_run


def set_list_cell(cell, items: list, header_text: str = None, body_pt: float = None, header_pt: float = None):
    """Rebuilds a cell that holds a bulleted list of items (Completed Tasks /
    Upcoming Activities style: optional bold header paragraph, then one
    paragraph per item with an optional bold 'Task No - N:' prefix run,
    separated by blank spacer paragraphs), reusing the existing formatting
    found in the template cell as a pattern. Every item is kept in full -
    none are dropped or truncated for space. body_pt/header_pt optionally
    override the template's font size (used to shrink-to-fit long lists)."""
    items = [str(i).strip() for i in items if str(i).strip()]

    txBody = cell._tc.find(qn("a:txBody"))
    paragraphs = txBody.findall(qn("a:p"))

    header_p = None
    body_paragraphs = paragraphs
    if header_text is not None and paragraphs:
        first_text = "".join(t.text or "" for t in paragraphs[0].iter(qn("a:t")))
        if first_text.strip() == header_text.strip():
            header_p = paragraphs[0]
            body_paragraphs = paragraphs[1:]

    if header_p is not None and header_pt is not None:
        for r in header_p.findall(qn("a:r")):
            _set_run_font_pt(r, header_pt)

    # Find template runs/paragraphs to clone formatting from.
    bold_run_template = None
    normal_run_template = None
    item_ppr_template = None
    blank_p_template = None
    for p in body_paragraphs:
        runs = p.findall(qn("a:r"))
        if runs and item_ppr_template is None:
            ppr = p.find(qn("a:pPr"))
            item_ppr_template = ppr
        for r in runs:
            rpr = r.find(qn("a:rPr"))
            is_bold = rpr is not None and rpr.get("b") == "1"
            if is_bold and bold_run_template is None:
                bold_run_template = r
            if not is_bold and normal_run_template is None:
                normal_run_template = r
        if not runs and blank_p_template is None:
            blank_p_template = p

    if normal_run_template is None:
        # Fall back to any run we can find anywhere in the cell.
        any_run = txBody.find(f".//{qn('a:r')}")
        normal_run_template = any_run
    if normal_run_template is None:
        # The cell has no run anywhere (a fully empty template content cell,
        # separate from whatever header row carries actual runs) -
        # synthesize a minimal one so items still render instead of
        # silently vanishing, carrying over font size from endParaRPr where
        # one is available.
        end_pr = None
        if blank_p_template is not None:
            end_pr = blank_p_template.find(qn("a:endParaRPr"))
        if end_pr is None:
            end_pr = txBody.find(f".//{qn('a:endParaRPr')}")
        normal_run_template = txBody.makeelement(qn("a:r"), {})
        if end_pr is not None:
            rpr = copy.deepcopy(end_pr)
            rpr.tag = qn("a:rPr")
            normal_run_template.append(rpr)
        t = normal_run_template.makeelement(qn("a:t"), {})
        t.text = ""
        normal_run_template.append(t)
    if bold_run_template is None:
        bold_run_template = normal_run_template

    # Remove all existing paragraphs except the header (which we keep as-is).
    for p in paragraphs:
        if p is not header_p:
            txBody.remove(p)

    def new_blank_paragraph():
        if blank_p_template is not None:
            p = copy.deepcopy(blank_p_template)
            if body_pt is not None:
                for epr in p.findall(qn("a:endParaRPr")):
                    epr.set("sz", str(int(round(body_pt * 100))))
            return p
        return txBody.makeelement(qn("a:p"), {})

    def new_item_paragraph(item_text):
        p = txBody.makeelement(qn("a:p"), {})
        if item_ppr_template is not None:
            p.append(copy.deepcopy(item_ppr_template))
        prefix, rest = _split_prefix(item_text)
        if prefix and bold_run_template is not None:
            p.append(_make_run(bold_run_template, prefix, body_pt))
            if normal_run_template is not None:
                p.append(_make_run(normal_run_template, rest, body_pt))
        elif normal_run_template is not None:
            p.append(_make_run(normal_run_template, item_text.strip(), body_pt))
        return p

    if not items:
        txBody.append(new_item_paragraph("None") if not header_p else new_blank_paragraph())
        return

    for idx, item_text in enumerate(items):
        if not str(item_text).strip():
            continue
        txBody.append(new_item_paragraph(str(item_text)))
        if idx != len(items) - 1:
            txBody.append(new_blank_paragraph())


def _col_widths(tbl):
    grid = tbl.find(qn("a:tblGrid"))
    return [int(col.get("w")) for col in grid.findall(qn("a:gridCol"))]


def _chars_per_line(width_emu, font_pt):
    avg_char_width = max(int(font_pt * EMU_PER_PT * 0.52), 1)
    usable = max(width_emu - CELL_H_PADDING, avg_char_width)
    return max(1, usable // avg_char_width)


def _line_count(text, width_emu, font_pt):
    chars_per_line = _chars_per_line(width_emu, font_pt)
    lines = 0
    for line in str(text).split("\n"):
        lines += -(-max(len(line), 1) // chars_per_line)
    return max(lines, 1)


def _line_height(font_pt):
    return int(font_pt * EMU_PER_PT * 1.25)


def _estimate_row_height(texts, col_widths, font_pt=12):
    """Estimates the row height (EMU) needed so wrapped text in the widest
    cell isn't clipped, since the cloned row keeps the template's original
    (short-content) height otherwise and some viewers don't auto-grow it."""
    max_lines = 1
    for text, width in zip(texts, col_widths):
        max_lines = max(max_lines, _line_count(text, width, font_pt))
    return max_lines * _line_height(font_pt) + CELL_V_PADDING


def _list_cell_height(header_text, items, header_pt, body_pt, col_width):
    """Estimated stacked height (EMU) of a Completed-Tasks/Upcoming-Activities
    style cell: optional header paragraph, then one line-wrapped paragraph
    per item with a blank spacer line between items."""
    total = 0
    if header_text is not None:
        total += _line_count(header_text, col_width, header_pt) * _line_height(header_pt)
    for idx, item in enumerate(items):
        total += _line_count(item, col_width, body_pt) * _line_height(body_pt)
        if idx != len(items) - 1:
            total += _line_height(body_pt)  # blank spacer paragraph
    return total + CELL_V_PADDING


def _cell_font_sizes(cell, header_text):
    """Peeks at a list cell's template runs (without mutating it) to read the
    header/body font sizes it was authored with, in points."""
    txBody = cell._tc.find(qn("a:txBody"))
    paragraphs = txBody.findall(qn("a:p"))
    header_pt = None
    body_pt = None
    start_idx = 0

    if header_text is not None and paragraphs:
        first_text = "".join(t.text or "" for t in paragraphs[0].iter(qn("a:t")))
        if first_text.strip() == header_text.strip():
            for r in paragraphs[0].findall(qn("a:r")):
                rpr = r.find(qn("a:rPr"))
                if rpr is not None and rpr.get("sz"):
                    header_pt = int(rpr.get("sz")) / 100
                    break
            start_idx = 1

    for p in paragraphs[start_idx:]:
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            if rpr is not None and rpr.get("sz"):
                body_pt = int(rpr.get("sz")) / 100
                break
        if body_pt is not None:
            break

    return header_pt or 14.0, body_pt or 12.0


def _available_height(shape, slide, slide_height):
    """How tall `shape` can render before it would overlap another shape on
    the same slide that sits below it (and shares horizontal space with it),
    or the slide's bottom edge if nothing does."""
    left, top, width = shape.left, shape.top, shape.width
    right = left + width
    limit = slide_height
    shape_id = shape.shape_id

    for other in slide.shapes:
        if other.shape_id == shape_id:
            continue
        o_top, o_left, o_width = other.top, other.left, other.width
        if o_top is None or o_left is None or o_width is None or o_top <= top:
            continue
        if o_left + o_width <= left or o_left >= right:
            continue
        limit = min(limit, o_top)

    return max(limit - top, 0)


def _fit_list_cell(cell, items, header_text, col_width, available_height, allow_shrink=True):
    """Fits as many items as possible into `cell`. When allow_shrink is True
    (Upcoming Activities), it first tries shrinking the font down to
    MIN_FONT_PT before giving up any items. When False (Completed Tasks -
    the font must stay at its template size), it never changes the font and
    simply keeps as many whole items as fit at that fixed size, leaving the
    rest as overflow. Writes the fitting subset into the cell and returns the
    leftover items (empty if everything fit) plus the height used."""
    items = [str(i).strip() for i in items if str(i).strip()]
    header_pt0, body_pt0 = _cell_font_sizes(cell, header_text)
    ratio = header_pt0 / body_pt0 if body_pt0 else 1

    def needed(body, header, subset):
        return _list_cell_height(header_text, subset, header, body, col_width)

    fitting = items
    chosen_body, chosen_header = body_pt0, header_pt0

    if items and needed(body_pt0, header_pt0, items) > available_height:
        if allow_shrink:
            pt = body_pt0 - FONT_STEP_PT
            fits = False
            while pt >= MIN_FONT_PT - 1e-6:
                if needed(pt, pt * ratio, items) <= available_height:
                    fits = True
                    break
                pt -= FONT_STEP_PT

            if fits:
                chosen_body, chosen_header = pt, pt * ratio
            else:
                chosen_body, chosen_header = MIN_FONT_PT, MIN_FONT_PT * ratio
                fitting = []
                for i in range(len(items)):
                    candidate = items[: i + 1]
                    if needed(chosen_body, chosen_header, candidate) <= available_height:
                        fitting = candidate
                    else:
                        break
                if not fitting:
                    # Never silently drop the very first item - worst case it
                    # renders slightly beyond the estimated box.
                    fitting = items[:1]
        else:
            fitting = []
            for i in range(len(items)):
                candidate = items[: i + 1]
                if needed(chosen_body, chosen_header, candidate) <= available_height:
                    fitting = candidate
                else:
                    break
            if not fitting:
                fitting = items[:1]

    set_list_cell(cell, fitting, header_text, body_pt=chosen_body, header_pt=chosen_header)
    used_height = needed(chosen_body, chosen_header, fitting) if fitting or header_text is not None else 0
    remainder = items[len(fitting):]
    return remainder, used_height


def _apply_milestone_rows(table, milestones: list, font_pt: float):
    """Writes `milestones` into the table's data rows at `font_pt`, adding or
    removing rows as needed and growing each row's height to fit its actual
    (wrapped) content."""
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    data_trs = trs[2:]
    template_tr = data_trs[-1]

    n_needed = len(milestones)
    n_have = len(data_trs)

    if n_needed > n_have:
        for _ in range(n_needed - n_have):
            new_tr = copy.deepcopy(template_tr)
            template_tr.addnext(new_tr)
            template_tr = new_tr
            data_trs.append(new_tr)
    elif n_needed < n_have:
        for extra_tr in data_trs[n_needed:]:
            tbl.remove(extra_tr)
        data_trs = data_trs[:n_needed]

    col_widths = _col_widths(tbl)

    for tr, item in zip(data_trs, milestones):
        tcs = tr.findall(qn("a:tc"))
        if len(tcs) < 4:
            continue
        texts = [
            str(item.get("area", "")),
            str(item.get("task", "")),
            str(item.get("status", "")),
            str(item.get("remarks", "")),
        ]
        for tc, text in zip(tcs[:4], texts):
            _set_tc_simple_text(tc, text, font_pt)

        needed_h = _estimate_row_height(texts, col_widths, font_pt)
        current_h = int(tr.get("h", "0") or 0)
        tr.set("h", str(max(current_h, needed_h)))


def _fit_milestone_rows(table, milestones: list, available_height: float, allow_shrink=True):
    """Fits as many milestone rows as possible below the table's fixed title
    + column-header rows. When allow_shrink is True it shrinks the font down
    to MIN_FONT_PT before giving up any rows; when False the font stays at
    its template size and rows simply overflow instead. Writes the fitting
    subset and returns (leftover milestones, height used)."""
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    if len(trs) < 3:
        raise RuntimeError("Milestone table does not have the expected header rows.")

    title_h = int(trs[0].get("h", "0") or 0)
    header_h = int(trs[1].get("h", "0") or 0)
    data_available = max(available_height - title_h - header_h, 0)

    placeholder_used = False
    if not milestones:
        milestones = [{"area": "-", "task": "-", "status": "-", "remarks": "No milestone updates this week."}]
        placeholder_used = True

    col_widths = _col_widths(tbl)

    def row_texts(item):
        return [str(item.get("area", "")), str(item.get("task", "")), str(item.get("status", "")), str(item.get("remarks", ""))]

    def total_height(font_pt, subset):
        return sum(_estimate_row_height(row_texts(m), col_widths, font_pt) for m in subset)

    template_pt = 12.0
    fitting = milestones
    chosen_pt = template_pt

    if total_height(template_pt, milestones) > data_available:
        if allow_shrink:
            pt = template_pt - FONT_STEP_PT
            fits = False
            while pt >= MIN_FONT_PT - 1e-6:
                if total_height(pt, milestones) <= data_available:
                    fits = True
                    break
                pt -= FONT_STEP_PT

            if fits:
                chosen_pt = pt
            else:
                chosen_pt = MIN_FONT_PT
                fitting = []
                for i in range(len(milestones)):
                    candidate = milestones[: i + 1]
                    if total_height(chosen_pt, candidate) <= data_available:
                        fitting = candidate
                    else:
                        break
                if not fitting:
                    fitting = milestones[:1]
        else:
            fitting = []
            for i in range(len(milestones)):
                candidate = milestones[: i + 1]
                if total_height(chosen_pt, candidate) <= data_available:
                    fitting = candidate
                else:
                    break
            if not fitting:
                fitting = milestones[:1]

    _apply_milestone_rows(table, fitting, chosen_pt)
    used_height = title_h + header_h + total_height(chosen_pt, fitting)
    if placeholder_used:
        return [], used_height
    return milestones[len(fitting):], used_height


def _set_tc_simple_text(tc_el, text: str, font_pt: float = None):
    txBody = tc_el.find(qn("a:txBody"))
    paragraphs = txBody.findall(qn("a:p"))
    if not paragraphs:
        return
    first_p = paragraphs[0]
    _clear_extra_runs(first_p, text)
    # A template row's original sample content may have wrapped onto a
    # second line via a manual <a:br/>, left over here since it's a
    # sibling of the runs rather than one itself, so _clear_extra_runs()
    # never touches it. Left in place, it renders as a blank trailing line
    # under this cell's new (usually shorter, single-line) text, growing
    # the row and leaving it visibly uneven against its neighbors.
    for br in first_p.findall(qn("a:br")):
        first_p.remove(br)
    if font_pt is not None:
        for r in first_p.findall(qn("a:r")):
            _set_run_font_pt(r, font_pt)
    for extra_p in paragraphs[1:]:
        txBody.remove(extra_p)


def _fit_simple_cell(shape, slide, slide_height, text):
    """Risk & Issues is a single free-form paragraph, not a list of discrete
    items, so it only shrinks to fit (down to MIN_FONT_PT) - there's nothing
    sensible to split onto a continuation slide."""
    table = shape.table
    cell = table.cell(1, 0)
    col_width = table.columns[0].width
    available = _available_height(shape, slide, slide_height)
    title_h = int(table._tbl.findall(qn("a:tr"))[0].get("h", "0") or 0)
    data_available = max(available - title_h, 0)

    _, body_pt0 = _cell_font_sizes(cell, None)
    pt = body_pt0
    needed = _line_count(text, col_width, pt) * _line_height(pt) + CELL_V_PADDING
    while needed > data_available and pt > MIN_FONT_PT - 1e-6:
        pt -= FONT_STEP_PT
        needed = _line_count(text, col_width, pt) * _line_height(pt) + CELL_V_PADDING

    set_cell_simple_text(cell, text)
    _set_cell_font_pt(cell._tc, pt)

    content_tr = table._tbl.findall(qn("a:tr"))[1]
    current_h = int(content_tr.get("h", "0") or 0)
    content_tr.set("h", str(max(current_h, int(needed))))


def _clone_slide(prs, source_slide):
    """Deep-copies every shape on source_slide onto a brand-new slide (same
    layout). Safe here because this slide's shapes are all text/table shapes
    with no image relationships to re-wire."""
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in source_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shp._element))
    return new_slide


def _move_slide(prs, from_index, to_index):
    sld_id_lst = prs.slides._sldIdLst
    slides = list(sld_id_lst)
    sld = slides[from_index]
    sld_id_lst.remove(sld)
    sld_id_lst.insert(to_index, sld)


def _remove_shape_if_present(slide, header):
    """Removes the box named by `header` from `slide` if it's there; a no-op
    if it was already stripped off (or never present)."""
    try:
        shp = _find_table_shape(slide, header)
    except RuntimeError:
        return
    shp._element.getparent().remove(shp._element)


def _shape_present(slide, header):
    try:
        _find_table_shape(slide, header)
        return True
    except RuntimeError:
        return False


def _new_continuation_slide(prs, source_slide, keep_headers, insert_at):
    """Clones source_slide, moves the clone to `insert_at`, and strips
    everything off it except placeholders (title, slide number, ...) and the
    box(es) named in keep_headers - so the overflowing section(s) get the
    whole page (their available height then extends all the way to the slide
    edge since nothing sits below them any more) with no leftover unrelated
    content (e.g. a Workstream table sharing that slide) dragged along onto
    the clone. Marks the title as continued, if there is one."""
    new_slide = _clone_slide(prs, source_slide)
    _move_slide(prs, len(prs.slides) - 1, insert_at)

    for shp in list(new_slide.shapes):
        if shp.is_placeholder:
            continue
        if shp.has_table:
            try:
                header_text = shp.table.cell(0, 0).text.strip()
            except Exception:
                header_text = None
            if header_text and any(header_text.lower() == kh.lower() for kh in keep_headers):
                continue
        shp._element.getparent().remove(shp._element)

    try:
        title_shape = _find_title_shape(new_slide)
    except RuntimeError:
        title_shape = None
    if title_shape is not None:
        p = title_shape.text_frame.add_paragraph()
        run = p.add_run()
        run.text = "(continued)"
        run.font.size = Pt(14)
        run.font.italic = True

    return new_slide


def _layout_section(prs, first_slide, header, kind, header_text, items, insert_at):
    """Lays `items` into the box named by `header` on first_slide, spilling
    onto as many "(continued)" slides as needed so every item is visible and
    nothing overlaps. Returns the insert_at index the next section should use."""
    current_slide = first_slide
    shape = _find_table_shape(current_slide, header)
    remaining = items

    while True:
        available = _available_height(shape, current_slide, prs.slide_height)

        if kind == "list":
            table = shape.table
            cell = table.cell(1, 0)
            remaining, used_height = _fit_list_cell(cell, remaining, header_text, table.columns[0].width, available)
            content_tr = table._tbl.findall(qn("a:tr"))[1]
            current_h = int(content_tr.get("h", "0") or 0)
            content_tr.set("h", str(max(current_h, int(used_height))))
        else:
            remaining, _ = _fit_milestone_rows(shape.table, remaining, available)

        if not remaining:
            return insert_at

        current_slide = _new_continuation_slide(prs, first_slide, {header}, insert_at)
        shape = _find_table_shape(current_slide, header)
        insert_at += 1


def _find_slide_with_table(prs, header):
    """Scans every slide in the deck for one containing a table whose header
    matches - used to locate Upcoming Activities / Risk & Issues when a
    template keeps them on a different slide than Completed Tasks/Milestone
    (e.g. a template that's been hand-edited to give Completed/Milestone a
    full page of their own, moving Upcoming/Risk onto a later page)."""
    for sl in prs.slides:
        if _shape_present(sl, header):
            return sl
    raise RuntimeError(f"Could not find any slide containing a table with header '{header}'.")


def _layout_completed_and_milestone_standalone(prs, first_slide, completed_items, milestones, insert_at):
    """Simpler counterpart to _layout_completed_and_milestone for templates
    where Completed Tasks / Milestone already have first_slide entirely to
    themselves (Upcoming Activities / Risk & Issues live elsewhere, so there
    is nothing to move out of the way or reinstate). Still never shrinks the
    font - just fits as much as the current box height allows and spills the
    rest onto cloned "(continued)" slides, holding only whichever of the two
    boxes still has leftover content, until both are fully placed."""
    completed_header = SECTION_HEADERS["completed"]
    milestone_header = SECTION_HEADERS["milestone"]

    remaining_completed = [str(i).strip() for i in completed_items if str(i).strip()]
    remaining_milestones = list(milestones)
    completed_done = False
    milestone_done = False
    current_slide = first_slide

    while True:
        completed_shape = None if completed_done else _find_table_shape(current_slide, completed_header)
        milestone_shape = None if milestone_done else _find_table_shape(current_slide, milestone_header)

        if completed_shape is not None:
            completed_cell = completed_shape.table.cell(1, 0)
            completed_col_w = completed_shape.table.columns[0].width
            avail_c = _available_height(completed_shape, current_slide, prs.slide_height)
            remaining_completed, _ = _fit_list_cell(
                completed_cell, remaining_completed, "Completed Tasks", completed_col_w, avail_c, allow_shrink=False
            )
            completed_done = not remaining_completed

        if milestone_shape is not None:
            avail_m = _available_height(milestone_shape, current_slide, prs.slide_height)
            remaining_milestones, _ = _fit_milestone_rows(
                milestone_shape.table, remaining_milestones, avail_m, allow_shrink=False
            )
            milestone_done = not remaining_milestones

        if completed_done and milestone_done:
            return insert_at

        keep_headers = set()
        if not completed_done:
            keep_headers.add(completed_header)
        if not milestone_done:
            keep_headers.add(milestone_header)
        current_slide = _new_continuation_slide(prs, first_slide, keep_headers, insert_at)
        insert_at += 1


def _layout_completed_and_milestone(prs, first_slide, completed_items, milestones, insert_at):
    """Lays Completed Tasks and Milestone & Supporting Activities into
    first_slide at their template font size - never shrunk. As long as both
    fit in their normal (template-sized) boxes alongside Upcoming Activities
    and Risk & Issues, nothing changes from a plain single slide.

    The moment either box would overflow at that fixed size, Upcoming
    Activities and Risk & Issues are pulled off that slide entirely and
    Completed Tasks / Milestone are re-laid-out using the freed space (their
    boxes now extend all the way to the slide edge instead of stopping above
    Upcoming/Risk). This repeats on freshly cloned "(continued)" slides,
    holding only whichever of the two boxes still has leftover content, until
    both are fully placed.

    Once both are fully placed, Upcoming Activities and Risk & Issues are
    reinstated - on that same final slide if the Completed/Milestone content
    that actually landed there would still fit within the ORIGINAL template
    box size (so it can't collide with Upcoming/Risk sitting at their fixed
    template position), otherwise on one more fresh slide of their own - so
    they only ever appear once Completed Tasks and Milestone are completely
    done.

    Returns (insert_at, host_slide): host_slide is where Upcoming Activities
    and Risk & Issues now live (with their boxes present and ready to fill),
    and insert_at is the next index a further continuation slide should use.
    """
    completed_header = SECTION_HEADERS["completed"]
    milestone_header = SECTION_HEADERS["milestone"]
    upcoming_header = SECTION_HEADERS["upcoming"]
    risk_header = SECTION_HEADERS["risk"]

    # Fixed reference: how much room Completed/Milestone have in the
    # template's own (untouched, Upcoming/Risk-present) layout. Used later to
    # decide whether Upcoming/Risk can safely go back at their fixed
    # position without overlapping whatever ended up in these boxes.
    template_c_avail = _available_height(
        _find_table_shape(first_slide, completed_header), first_slide, prs.slide_height
    )
    template_m_avail = _available_height(
        _find_table_shape(first_slide, milestone_header), first_slide, prs.slide_height
    )

    # Pristine copies from the untouched first_slide, kept around in case
    # Upcoming/Risk get stripped off a slide and need to be reinstated later
    # (on that same slide, or on a fresh one of their own).
    upcoming_template_el = copy.deepcopy(_find_table_shape(first_slide, upcoming_header)._element)
    risk_template_el = copy.deepcopy(_find_table_shape(first_slide, risk_header)._element)

    remaining_completed = [str(i).strip() for i in completed_items if str(i).strip()]
    remaining_milestones = list(milestones)
    completed_done = False
    milestone_done = False
    current_slide = first_slide

    while True:
        completed_shape = None if completed_done else _find_table_shape(current_slide, completed_header)
        milestone_shape = None if milestone_done else _find_table_shape(current_slide, milestone_header)

        completed_cell = completed_col_w = None
        used_c = used_m = 0

        if completed_shape is not None:
            completed_cell = completed_shape.table.cell(1, 0)
            completed_col_w = completed_shape.table.columns[0].width
            avail_c = _available_height(completed_shape, current_slide, prs.slide_height)
            remaining_completed, used_c = _fit_list_cell(
                completed_cell, remaining_completed, "Completed Tasks", completed_col_w, avail_c, allow_shrink=False
            )

        if milestone_shape is not None:
            avail_m = _available_height(milestone_shape, current_slide, prs.slide_height)
            remaining_milestones, used_m = _fit_milestone_rows(
                milestone_shape.table, remaining_milestones, avail_m, allow_shrink=False
            )

        if remaining_completed or remaining_milestones:
            # Doesn't fit in the room this slide currently offers - free up
            # space by dropping Upcoming/Risk (if they're still here) and
            # retry once at the same fixed font size before spilling over.
            _remove_shape_if_present(current_slide, upcoming_header)
            _remove_shape_if_present(current_slide, risk_header)

            if completed_shape is not None:
                avail_c = _available_height(completed_shape, current_slide, prs.slide_height)
                remaining_completed, used_c = _fit_list_cell(
                    completed_cell, remaining_completed, "Completed Tasks", completed_col_w, avail_c, allow_shrink=False
                )
            if milestone_shape is not None:
                avail_m = _available_height(milestone_shape, current_slide, prs.slide_height)
                remaining_milestones, used_m = _fit_milestone_rows(
                    milestone_shape.table, remaining_milestones, avail_m, allow_shrink=False
                )

        if completed_shape is not None:
            completed_done = not remaining_completed
        if milestone_shape is not None:
            milestone_done = not remaining_milestones

        if completed_done and milestone_done:
            fits_normally = used_c <= template_c_avail and used_m <= template_m_avail

            if fits_normally:
                if not _shape_present(current_slide, upcoming_header):
                    current_slide.shapes._spTree.append(copy.deepcopy(upcoming_template_el))
                if not _shape_present(current_slide, risk_header):
                    current_slide.shapes._spTree.append(copy.deepcopy(risk_template_el))
            else:
                # Build this dedicated slide from the preserved pristine
                # copies, not by cloning first_slide - if overflow already
                # stripped Upcoming/Risk off first_slide itself (because
                # first_slide was the slide being extended), first_slide no
                # longer has them to clone from.
                new_slide = _clone_slide(prs, first_slide)
                _move_slide(prs, len(prs.slides) - 1, insert_at)
                for header in SECTION_HEADERS.values():
                    _remove_shape_if_present(new_slide, header)
                new_slide.shapes._spTree.append(copy.deepcopy(upcoming_template_el))
                new_slide.shapes._spTree.append(copy.deepcopy(risk_template_el))
                title_shape = _find_title_shape(new_slide)
                p = title_shape.text_frame.add_paragraph()
                run = p.add_run()
                run.text = "(continued)"
                run.font.size = Pt(14)
                run.font.italic = True
                current_slide = new_slide
                insert_at += 1
            return insert_at, current_slide

        keep_headers = set()
        if not completed_done:
            keep_headers.add(completed_header)
        if not milestone_done:
            keep_headers.add(milestone_header)
        current_slide = _new_continuation_slide(prs, first_slide, keep_headers, insert_at)
        insert_at += 1


def build_report(template_path: str, output_path: str, week: dict, report: dict, slide_index: int = SLIDE_INDEX):
    """Opens template_path, edits ONLY slide 1's date and the status-report
    slide's content (using week['range_text'] and week['slide1_date_text'])
    and saves the result to output_path. template_path is never modified.
    slide_index (0-indexed) picks which slide holds the status report tables
    - different projects' templates place it at a different slide number.
    Every other slide, and every other element of slide 1 and that slide, is
    left untouched.

    Completed Tasks and Milestone & Supporting Activities always render at
    their template's font size - never shrunk. If they don't both fit
    alongside Upcoming Activities / Risk & Issues, those two boxes are moved
    off the slide and Completed/Milestone's boxes extend into the freed
    space, spilling onto cloned "(continued)" slides as needed until both are
    fully shown; Upcoming Activities and Risk & Issues then reappear right
    after, on that same slide if there's room or on one more slide of their
    own otherwise. (If the template already keeps Upcoming/Risk on a
    separate slide of their own - e.g. a hand-edited layout - Completed Tasks
    and Milestone just get the whole of first_slide from the start, no
    reshuffling needed.) Upcoming Activities keeps the old shrink-then-spill
    behavior for any overflow of its own; Risk & Issues only ever shrinks to
    fit (there's nothing sensible to split it onto another page)."""
    prs = Presentation(template_path)
    slide = prs.slides[slide_index]

    _set_title_slide_date(prs, week["slide1_date_text"])
    _set_date_range(slide, week["range_text"])

    upcoming_header = SECTION_HEADERS["upcoming"]

    if _shape_present(slide, upcoming_header):
        insert_at, host_slide = _layout_completed_and_milestone(
            prs, slide, report.get("completed_tasks", []), report.get("milestones", []), slide_index + 1
        )
    else:
        insert_at = _layout_completed_and_milestone_standalone(
            prs, slide, report.get("completed_tasks", []), report.get("milestones", []), slide_index + 1
        )
        host_slide = _find_slide_with_table(prs, upcoming_header)

    risk_shape = _find_table_shape(host_slide, SECTION_HEADERS["risk"])
    _fit_simple_cell(risk_shape, host_slide, prs.slide_height, report.get("risks_issues") or "None")

    _layout_section(
        prs, host_slide, upcoming_header, "list", None, report.get("upcoming_activities", []), insert_at
    )

    prs.save(output_path)
    return output_path
