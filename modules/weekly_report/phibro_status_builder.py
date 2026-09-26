"""Builds the newer Phibro Animal Healthcare "Project Status Report"
template (template/Phibro_Weekly_Status_Sep11_1.pptx), a different shape
from the other "standard" 4-box templates:

- Slide 1 (cover) carries the engagement date inline, same as the standard
  template.
- Slide 2 is dynamic: a small "Project:" table (Reporting Period / Reporting
  Date get updated; the Project name row is left as-is), a "Project Status"
  narrative paragraph (the LLM's project_status_summary), and one unified
  Milestone / Status / Value table that holds EVERY point from the week -
  completed work, ongoing milestones, upcoming activities, and any called-out
  risk - there's no separate Completed Tasks / Upcoming Activities / Risk &
  Issues box like the other templates. Nothing is ever dropped or truncated;
  whatever doesn't fit spills onto cloned "(continued)" slides that give the
  Milestone table the whole page.
- Slide 3 (Workstream) keeps its "Daily CRM & Database Operations" table
  static, but its "Issues Resolved Last Week" # / Issue / Resolution /
  Status table is filled from the step3 review screen's resolved_issues
  rows, spilling onto cloned "(continued)" slides the same way.

The report dict handed in is the exact same shape the "standard" builder
takes (completed_tasks/milestones/upcoming_activities/risks_issues/
project_status_summary) - this template just renders it differently."""

import copy
import re
from datetime import datetime

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .pptx_builder import (
    _clone_slide,
    _col_widths,
    _estimate_row_height,
    _move_slide,
    _replace_across_runs,
    _set_tc_simple_text,
    set_cell_simple_text,
)

COVER_SLIDE_INDEX = 0  # slide 1, 0-indexed
STATUS_SLIDE_INDEX = 1  # slide 2, 0-indexed
WORKSTREAM_SLIDE_INDEX = 2  # slide 3, 0-indexed

COVER_DATE_RE = re.compile(r"\d{1,2}-[A-Za-z]+-\d{4}")

PROJECT_TABLE_HEADER = "Project:"
MILESTONE_TABLE_HEADER = "Milestone"
STATUS_NARRATIVE_MARKER = "Project Status"
TITLE_TEXT = "Project Status Report"
ISSUES_TITLE_TEXT = "Issues Resolved Last Week"
WORKSTREAM_TABLE_HEADER = "Workstream"
ISSUES_BOTTOM_MARGIN = Emu(274320)  # 0.3in
CONTINUATION_GAP = Emu(150000)
DEFAULT_ROW_FONT_PT = 11.0


def _find_table_shape(slide, header_contains):
    for shape in slide.shapes:
        if shape.has_table:
            try:
                header_text = shape.table.cell(0, 0).text.strip()
            except Exception:
                continue
            if header_contains.lower() in header_text.lower():
                return shape
    raise RuntimeError(f"Could not find a table containing header '{header_contains}' on the Phibro status slide.")


def _is_status_narrative_shape(shape):
    """True for the real Project Status narrative box, matched structurally
    rather than by a loose text prefix: the slide's own title box ('Project
    Status Report') also starts with the word "Project Status", so a plain
    startswith() check would grab that instead. The real narrative box's
    first paragraph is exactly one run reading "Project Status" immediately
    followed by a manual line break."""
    if not shape.has_text_frame:
        return False
    txBody = shape.text_frame._txBody
    p = txBody.find(qn("a:p"))
    if p is None:
        return False
    children = list(p)
    if len(children) < 2 or children[0].tag != qn("a:r") or children[1].tag != qn("a:br"):
        return False
    t = children[0].find(qn("a:t"))
    return t is not None and (t.text or "").strip() == STATUS_NARRATIVE_MARKER


def _find_status_narrative_shape(slide):
    for shape in slide.shapes:
        if _is_status_narrative_shape(shape):
            return shape
    raise RuntimeError("Could not find the Project Status narrative text box.")


def _find_progress_title_shape(slide):
    """The 'Last Week's Progress' heading above the Milestone table. Matched
    by "contains 'progress' but not 'status'" so it doesn't also match the
    Project Status narrative box, which itself contains the word
    "progressing"."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip().lower()
        if "progress" in text and "status" not in text:
            return shape
    return None


def _find_title_shape(slide):
    """The 'Project Status Report' heading in the top-left corner."""
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip() == TITLE_TEXT:
            return shape
    return None


def _bbox(shape):
    return (shape.left, shape.top, shape.left + shape.width, shape.top + shape.height)


def _overlap_fraction(inner_bbox, outer_bbox):
    """Fraction of inner_bbox's area that falls inside outer_bbox."""
    ix1, iy1, ix2, iy2 = inner_bbox
    ox1, oy1, ox2, oy2 = outer_bbox
    dx = min(ix2, ox2) - max(ix1, ox1)
    dy = min(iy2, oy2) - max(iy1, oy1)
    if dx <= 0 or dy <= 0:
        return 0.0
    inner_area = max((ix2 - ix1) * (iy2 - iy1), 1)
    return (dx * dy) / inner_area


def _find_decorative_background(slide, target_bbox, max_area_ratio=3.0):
    """A plain fill rectangle sitting behind a table/textbox (e.g. the
    shaded panel behind the Project: table, or the tinted highlight behind
    the Project Status narrative) - matched by footprint overlap with
    target_bbox rather than by shape name, since names aren't a stable
    contract. max_area_ratio excludes the full-slide background rectangle,
    which would otherwise also "overlap" everything."""
    tx1, ty1, tx2, ty2 = target_bbox
    target_area = max((tx2 - tx1) * (ty2 - ty1), 1)
    best, best_overlap = None, 0.5
    for shape in slide.shapes:
        if shape.has_table or (shape.has_text_frame and shape.text_frame.text.strip()):
            continue
        bbox = _bbox(shape)
        area = max((bbox[2] - bbox[0]) * (bbox[3] - bbox[1]), 1)
        if area > target_area * max_area_ratio:
            continue
        overlap = _overlap_fraction(target_bbox, bbox)
        if overlap > best_overlap:
            best, best_overlap = shape, overlap
    return best


def _set_cover_date(prs, date_text: str):
    slide = prs.slides[COVER_SLIDE_INDEX]
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        txBody = shape.text_frame._txBody
        for p in txBody.findall(qn("a:p")):
            if _replace_across_runs(p, COVER_DATE_RE, date_text):
                return
    raise RuntimeError("Could not find the engagement date on the cover slide.")


def _set_project_info(slide, reporting_period_text: str, reporting_date_text: str):
    shape = _find_table_shape(slide, PROJECT_TABLE_HEADER)
    table = shape.table
    for r in range(len(table.rows)):
        label = table.cell(r, 0).text.strip().lower()
        if label.startswith("reporting period"):
            set_cell_simple_text(table.cell(r, 1), reporting_period_text)
        elif label.startswith("reporting date"):
            set_cell_simple_text(table.cell(r, 1), reporting_date_text)
    return shape


def _tc_text(tc_el):
    return "".join(t.text or "" for t in tc_el.iter(qn("a:t")))


def _table_run_font_pt(table, default=DEFAULT_ROW_FONT_PT):
    tbl = table._tbl
    for tr in tbl.findall(qn("a:tr")):
        for tc in tr.findall(qn("a:tc")):
            for r in tc.iter(qn("a:r")):
                rpr = r.find(qn("a:rPr"))
                if rpr is not None and rpr.get("sz"):
                    return int(rpr.get("sz")) / 100
    return default


def _grow_table_to_fit(table, font_pt):
    """Recomputes each row's height from its actual text and applies it,
    returning the table's new total height. PowerPoint always grows a table
    row to fit wrapped text regardless of the row's stored height, so
    anything positioned below the table using its stored (pre-wrap) height
    will end up rendered on top of it unless that growth is accounted for
    here first."""
    tbl = table._tbl
    col_widths = _col_widths(tbl)
    total = 0
    for tr in tbl.findall(qn("a:tr")):
        texts = [_tc_text(tc) for tc in tr.findall(qn("a:tc"))]
        needed_h = _estimate_row_height(texts, col_widths, font_pt)
        current_h = int(tr.get("h", "0") or 0)
        new_h = max(current_h, needed_h)
        tr.set("h", str(new_h))
        total += new_h
    return total


def _fit_status_slide_header(slide, progress_title, milestone_shape, info_shape):
    """The Project: table's Reporting Period/Date values can wrap onto a
    second line once real dates are filled in - the template's label column
    ("Reporting Period:") is only wide enough for the label on one line -
    pushing the table's actual rendered bottom edge past the fixed-position
    'Last Week's Progress' banner and Milestone table below it, so the
    wrapped text renders on top of them. Grows the table (and its shaded
    background panel) to the height its text actually needs, then shifts
    the banner and Milestone table down by the same amount."""
    font_pt = _table_run_font_pt(info_shape.table)
    background = _find_decorative_background(slide, _bbox(info_shape))

    new_height = _grow_table_to_fit(info_shape.table, font_pt)
    delta = new_height - info_shape.height
    if delta <= 0:
        return

    info_shape.height = new_height
    if background is not None:
        background.height += delta

    if progress_title is not None:
        progress_title.top += delta
    milestone_shape.top += delta
    milestone_shape.height -= delta


def _set_project_status_narrative(shape, text: str):
    """The narrative box is one paragraph: a bold 'Project Status' header
    run, a manual line break, then the narrative run(s). Keeps the header
    and line break untouched and replaces everything after the break with a
    single new run holding `text`, reusing the last narrative run's
    formatting."""
    txBody = shape.text_frame._txBody
    p = txBody.find(qn("a:p"))
    if p is None:
        return

    br = p.find(qn("a:br"))
    if br is None:
        # No header/line-break structure to preserve - fall back to keeping
        # just the first run's formatting, like the other simple-text setters.
        runs = p.findall(qn("a:r"))
        if not runs:
            return
        first = runs[0]
        t = first.find(qn("a:t"))
        if t is None:
            t = first.makeelement(qn("a:t"), {})
            first.append(t)
        t.text = text
        for extra in runs[1:]:
            p.remove(extra)
        return

    children = list(p)
    idx = children.index(br)
    after = children[idx + 1 :]
    template_run = next((el for el in after if el.tag == qn("a:r")), None)
    for el in after:
        p.remove(el)

    new_run = copy.deepcopy(template_run) if template_run is not None else p.makeelement(qn("a:r"), {})
    t = new_run.find(qn("a:t"))
    if t is None:
        t = new_run.makeelement(qn("a:t"), {})
        new_run.append(t)
    t.text = text
    rpr = new_run.find(qn("a:rPr"))
    if rpr is not None and rpr.get("err"):
        del rpr.attrib["err"]
    p.append(new_run)


def _milestone_label(milestone: dict) -> str:
    area = str(milestone.get("area", "")).strip()
    task = str(milestone.get("task", "")).strip()
    return " - ".join(p for p in (area, task) if p)


_LABEL_PREFIX_RE = re.compile(r"^([^:]{1,60}):\s*(.+)$")

# Coordinating conjunctions/prepositions that normally open a trailing
# clause ("... to review", "... and shared it", "... with Melanie"). A
# truncated label is only ever cut *before* one of these, never after -
# cutting at a plain word count instead routinely left a label dangling
# on exactly one of these words ("Presented the automation demo to…"),
# which reads as a sentence cut off mid-thought rather than a short name.
_LABEL_CONNECTORS = {
    "and", "or", "nor", "to", "for", "with", "in", "on", "at", "of",
    "from", "into", "onto", "by", "as", "&",
}
_MIN_LABEL_WORDS = 3
_MAX_LABEL_WORDS = 9


def _find_label_cut(words):
    """Index of the first connector word at or after _MIN_LABEL_WORDS,
    within the first _MAX_LABEL_WORDS - everything before it becomes the
    label. None if no such word exists, meaning there's no clean place to
    shorten this sentence."""
    limit = min(len(words), _MAX_LABEL_WORDS)
    for i in range(_MIN_LABEL_WORDS, limit):
        if words[i].strip(",;:").lower() in _LABEL_CONNECTORS:
            return i
    return None


def _split_label_value(text: str, max_words: int = _MAX_LABEL_WORDS):
    """Splits one flat completed-task/upcoming-activity sentence into a
    short Milestone label and its full Value description, without ever
    inventing new words - only ever a substring or truncation of what's
    already there:
    1. Reuses an embedded "Label: detail" split when llm_consolidate.py
       already produced one (from a reclassified in-progress milestone).
    2. Otherwise cuts at the first natural clause break (a period, a
       semicolon, or a " - "/"–" dash) if there is one early in the
       sentence.
    3. Otherwise, only if the sentence is long enough to be worth
       shortening at all, cuts right before the first coordinating
       conjunction/preposition so the label reads as a short, complete
       name on its own - not a fragment with a trailing ellipsis, which
       still reads as changing what the sentence says by implying it
       continues into something unstated.
    Returns the text unchanged for both when it's already short enough
    that there's nothing worth shortening, or when no clean cut point
    exists - a duplicated Milestone/Value is preferable to a Milestone
    that reads as an incomplete sentence. The Value column is always the
    untouched original text either way."""
    text = text.strip()

    m = _LABEL_PREFIX_RE.match(text)
    if m and m.group(2).strip():
        return m.group(1).strip(), text

    for sep in (". ", "; ", " - ", " – "):
        idx = text.find(sep)
        if 0 < idx <= 70:
            return text[:idx].strip(), text

    words = text.split()
    if len(words) <= max_words:
        return text, text

    cut = _find_label_cut(words)
    if cut is None:
        return text, text
    return " ".join(words[:cut]).rstrip(",;:"), text


STATUS_ICONS = {
    "completed": "🟢",
    "in progress": "🟡",
    "at risk": "🔴",
    "not started": "⚪",
}


def _status_with_icon(status: str) -> str:
    """Prefixes a status with the same colored-circle convention "In
    Progress" already used everywhere else in this app's milestone tables
    (🟡), so Completed/At Risk/Not Started get their own matching green/
    red/white circle instead of being the only plain-text values in the
    column. Strips any icon a value already carries first (e.g. a status
    coming from the older "standard" review UI's emoji-prefixed dropdown,
    or a re-generated report), so a value is never double-prefixed."""
    status = str(status or "").strip()
    if not status:
        return status
    bare = status
    for icon in STATUS_ICONS.values():
        if bare.startswith(icon):
            bare = bare[len(icon):].strip()
            break
    icon = STATUS_ICONS.get(bare.lower())
    return f"{icon} {bare}" if icon else bare


def _add_status_icons(rows):
    return [(milestone, _status_with_icon(status), value) for milestone, status, value in rows]


def _status_rows_from_review(report: dict):
    """The step3 review screen for Phibro edits one unified Milestone/
    Status/Value table directly (matching this template's own layout)
    instead of the separate Completed Tasks/Milestones/Upcoming Activities
    fields every other project uses, and submits it as report["status_rows"]
    - a list of {"milestone", "status", "value"} dicts already reviewed and
    approved by the user. When present, those rows are used verbatim (no
    re-derivation, so nothing the user edited gets overwritten); returns
    None so the caller falls back to _build_status_rows() for any older
    caller that still submits the plain completed_tasks/milestones/
    upcoming_activities/risks_issues shape instead."""
    raw_rows = report.get("status_rows")
    if not raw_rows:
        return None
    rows = []
    for r in raw_rows:
        milestone = str(r.get("milestone", "")).strip()
        value = str(r.get("value", "")).strip()
        if not milestone and not value:
            continue
        status = str(r.get("status", "")).strip() or "In Progress"
        rows.append((milestone or value, status, value or milestone))
    return rows or None


def _build_status_rows(report: dict):
    """Every completed task, ongoing milestone, upcoming activity, and any
    explicitly-called-out risk becomes one (Milestone, Status, Value) row -
    nothing from the week's update is left out. Order: completed work first,
    then ongoing work, then upcoming/not-yet-started work, then risks.

    Milestone is always a short label and Value the fuller description of
    what that milestone actually represents - never identical placeholders
    for each other, and never fabricated: every Value is either the
    original completed-task/upcoming-activity sentence in full, or the
    milestone's own LLM-provided remarks."""
    rows = []

    for text in report.get("completed_tasks", []):
        text = str(text).strip()
        if text:
            milestone_text, value_text = _split_label_value(text)
            rows.append((milestone_text, "Completed", value_text))

    for m in report.get("milestones", []):
        label = _milestone_label(m)
        remarks = str(m.get("remarks", "")).strip()
        status = str(m.get("status", "")).strip() or "In Progress"
        if label and remarks and label.lower() != remarks.lower():
            milestone_text, value_text = label, remarks
        else:
            milestone_text, value_text = _split_label_value(remarks or label or "Update in progress.")
        rows.append((milestone_text, status, value_text))

    for text in report.get("upcoming_activities", []):
        text = str(text).strip()
        if text:
            milestone_text, value_text = _split_label_value(text)
            rows.append((milestone_text, "Not Started", value_text))

    risks_text = str(report.get("risks_issues") or "").strip()
    if risks_text and risks_text.lower() != "none":
        rows.append(("Risks & Issues", "At Risk", risks_text))

    if not rows:
        rows = [("No updates to report this week.", "-", "-")]

    return rows


def _table_font_pt(table, default=DEFAULT_ROW_FONT_PT):
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    if len(trs) < 2:
        return default
    for tc in trs[1].findall(qn("a:tc")):
        txBody = tc.find(qn("a:txBody"))
        for p in txBody.findall(qn("a:p")):
            for r in p.findall(qn("a:r")):
                rpr = r.find(qn("a:rPr"))
                if rpr is not None and rpr.get("sz"):
                    return int(rpr.get("sz")) / 100
    return default


def _set_row_vertical_anchor(tr_el, anchor="t"):
    """Forces every cell in the row to the same vertical text anchor (top,
    by default). The template's own original sample rows mix "top" and
    "center" anchors cell-by-cell, and since each logical row reuses
    whichever physical row-slot it lands in, real content ends up centered
    in one column and top-aligned in the next within the very same row -
    a visibly crooked table where the Value text sits lower than its own
    Milestone/Status text."""
    for tc in tr_el.findall(qn("a:tc")):
        tcPr = tc.find(qn("a:tcPr"))
        if tcPr is None:
            tcPr = tc.makeelement(qn("a:tcPr"), {})
            tc.append(tcPr)
        tcPr.set("anchor", anchor)


def _set_row_text_alignment(tr_el, algn="l"):
    """Forces every paragraph in the row to the same horizontal alignment
    (left, by default), for the same reason as _set_row_vertical_anchor:
    the template's own sample rows mix "left" and "justify" cell-by-cell,
    which is invisible on a single line but visibly uneven once real
    (usually longer) content wraps to more than one line."""
    for tc in tr_el.findall(qn("a:tc")):
        txBody = tc.find(qn("a:txBody"))
        for p in txBody.findall(qn("a:p")):
            pPr = p.find(qn("a:pPr"))
            if pPr is None:
                pPr = p.makeelement(qn("a:pPr"), {})
                p.insert(0, pPr)
            pPr.set("algn", algn)


def _apply_status_rows(table, rows: list, font_pt: float):
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    data_trs = trs[1:]  # single header row: Milestone | Status | Value (or # | Issue | Resolution | Status)
    if not data_trs:
        raise RuntimeError("Milestone table has no data row to use as a template.")
    template_tr = data_trs[-1]

    n_needed = len(rows)
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

    for tr, row in zip(data_trs, rows):
        tcs = tr.findall(qn("a:tc"))
        texts = list(row)
        if len(tcs) < len(texts):
            continue
        for tc, text in zip(tcs[: len(texts)], texts):
            _set_tc_simple_text(tc, text, font_pt)
        _set_row_vertical_anchor(tr)
        _set_row_text_alignment(tr)
        needed_h = _estimate_row_height(texts, col_widths, font_pt)
        # Always the freshly-estimated height for THIS row's content, never
        # maxed with whatever height happens to already be on the row
        # (leftover from the template's own unrelated sample rows, or - on a
        # continuation slide, which is cloned from the already-populated
        # first slide - from a previous page's content in that same slot).
        # _fit_status_rows below decides how many rows fit using this exact
        # same _estimate_row_height() call with no such floor, so applying a
        # taller stale height here would silently let more rows "fit" than
        # actually do, pushing real content off the bottom of the slide.
        tr.set("h", str(needed_h))


def _fit_status_rows(table, rows: list, available_height: float, font_pt: float):
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    if len(trs) < 2:
        raise RuntimeError("Milestone table does not have the expected header row.")
    header_h = int(trs[0].get("h", "0") or 0)
    data_available = max(available_height - header_h, 0)

    col_widths = _col_widths(tbl)

    def total_height(subset):
        return sum(_estimate_row_height(list(r), col_widths, font_pt) for r in subset)

    fitting = rows
    if rows and total_height(rows) > data_available:
        fitting = []
        for i in range(len(rows)):
            candidate = rows[: i + 1]
            if total_height(candidate) <= data_available:
                fitting = candidate
            else:
                break
        if not fitting:
            # Never silently drop the very first row - worst case it renders
            # slightly beyond the estimated box.
            fitting = rows[:1]

    _apply_status_rows(table, fitting, font_pt)
    used_height = header_h + total_height(fitting)
    return rows[len(fitting) :], used_height


def _milestone_available_height(prs, shape, bottom_margin):
    """How tall the Milestone table can render before running off the slide.
    Deliberately NOT the generic shape-overlap _available_height() helper:
    this template has a stray empty textbox sitting inside the Milestone
    table's own footprint, which that heuristic mistakes for something
    blocking it from below, collapsing the available height to almost
    nothing. Milestone is always the last real content on this slide, so its
    room is simply whatever's left above the template's own bottom margin."""
    return max(prs.slide_height - shape.top - bottom_margin, 0)


def _new_status_continuation_slide(prs, first_slide, insert_at, ref_gap, bottom_margin):
    """Clones first_slide, drops the project-info table and the Project
    Status narrative (repeating them on every overflow page isn't useful),
    along with the shaded background panel each of them sits on (otherwise
    that panel is left behind, floating over the shifted-up content), and
    moves the 'Last Week's Progress' heading + Milestone table up to
    reclaim that freed space, marking the heading as continued."""
    new_slide = _clone_slide_with_images(prs, first_slide)
    _move_slide(prs, len(prs.slides) - 1, insert_at)

    removed_bboxes = []
    for shape in list(new_slide.shapes):
        if shape.has_table and shape.table.cell(0, 0).text.strip() == PROJECT_TABLE_HEADER:
            removed_bboxes.append(_bbox(shape))
            shape._element.getparent().remove(shape._element)
        elif _is_status_narrative_shape(shape):
            removed_bboxes.append(_bbox(shape))
            shape._element.getparent().remove(shape._element)

    for target_bbox in removed_bboxes:
        background = _find_decorative_background(new_slide, target_bbox)
        if background is not None:
            background._element.getparent().remove(background._element)

    milestone_shape = _find_table_shape(new_slide, MILESTONE_TABLE_HEADER)
    progress_title = _find_progress_title_shape(new_slide)
    title_shape = _find_title_shape(new_slide)

    top_start = CONTINUATION_GAP
    if title_shape is not None:
        top_start = max(top_start, title_shape.top + title_shape.height + CONTINUATION_GAP)

    if progress_title is not None:
        progress_title.top = top_start
        p = progress_title.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = " (continued)"
        run.font.size = Pt(10)
        run.font.italic = True
        milestone_top = top_start + progress_title.height + ref_gap
    else:
        milestone_top = top_start

    milestone_shape.top = milestone_top
    milestone_shape.height = max(prs.slide_height - milestone_top - bottom_margin, Emu(500000))

    return new_slide


def _layout_status_rows(prs, first_slide, rows: list, insert_at: int):
    ref_progress_title = _find_progress_title_shape(first_slide)
    ref_milestone = _find_table_shape(first_slide, MILESTONE_TABLE_HEADER)
    ref_gap = (
        ref_milestone.top - (ref_progress_title.top + ref_progress_title.height)
        if ref_progress_title is not None
        else 0
    )
    bottom_margin = prs.slide_height - (ref_milestone.top + ref_milestone.height)

    current_slide = first_slide
    remaining = rows

    while True:
        shape = _find_table_shape(current_slide, MILESTONE_TABLE_HEADER)
        font_pt = _table_font_pt(shape.table)
        available = _milestone_available_height(prs, shape, bottom_margin)
        remaining, _ = _fit_status_rows(shape.table, remaining, available, font_pt)

        if not remaining:
            return insert_at

        current_slide = _new_status_continuation_slide(prs, first_slide, insert_at, ref_gap, bottom_margin)
        insert_at += 1


def _find_issues_table_shape(slide):
    """The # / Issue / Resolution / Status table. Matched on its second
    header cell, since the first one ("#") is too generic to search on."""
    for shape in slide.shapes:
        if shape.has_table and len(shape.table.columns) >= 4:
            if shape.table.cell(0, 1).text.strip().lower() == "issue":
                return shape
    raise RuntimeError("Could not find the Issue / Resolution / Status table on the Workstream slide.")


def _find_text_shape(slide, text):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip() == text:
            return shape
    return None


def _resolved_issue_rows(report: dict):
    """report["resolved_issues"] is the step3 review screen's list of
    {"issue", "resolution", "status"} dicts, used verbatim and numbered in
    order. Falls back to a single placeholder row so the template's own
    stale sample issue is never carried over into a new week's report."""
    rows = []
    for r in report.get("resolved_issues") or []:
        issue = str(r.get("issue", "")).strip()
        resolution = str(r.get("resolution", "")).strip()
        if not issue and not resolution:
            continue
        status = str(r.get("status", "")).strip() or "Resolved"
        rows.append((str(len(rows) + 1), issue or "-", resolution or "-", status))
    return rows or [("-", "No issues reported this week.", "-", "-")]


def _clone_slide_with_images(prs, source_slide):
    """_clone_slide() plus re-pointing every picture's r:embed at a fresh
    image relationship on the new slide - this slide carries the DataPattern
    logo, which would otherwise reference a relationship that doesn't exist
    on the clone."""
    new_slide = _clone_slide(prs, source_slide)
    for blip in new_slide.shapes._spTree.iter(qn("a:blip")):
        old_rid = blip.get(qn("r:embed"))
        if not old_rid:
            continue
        image_part = source_slide.part.related_part(old_rid)
        blip.set(qn("r:embed"), new_slide.part.relate_to(image_part, RT.IMAGE))
    return new_slide


def _new_issues_continuation_slide(prs, first_slide, insert_at):
    """Clones the Workstream slide, drops the Daily CRM & Database
    Operations heading + table (static, no need to repeat them), and moves
    the Issues heading + table up into that freed space, marking the
    heading as continued."""
    new_slide = _clone_slide_with_images(prs, first_slide)
    _move_slide(prs, len(prs.slides) - 1, insert_at)

    ops_title = next(
        (s for s in new_slide.shapes if s.has_text_frame and "database operations" in s.text_frame.text.lower()),
        None,
    )
    ops_table = next(
        (s for s in new_slide.shapes if s.has_table and s.table.cell(0, 0).text.strip() == WORKSTREAM_TABLE_HEADER),
        None,
    )
    issues_title = _find_text_shape(new_slide, ISSUES_TITLE_TEXT)
    issues_shape = _find_issues_table_shape(new_slide)

    gap = issues_shape.top - (issues_title.top + issues_title.height) if issues_title is not None else 0
    top_start = ops_title.top if ops_title is not None else (ops_table.top if ops_table is not None else issues_shape.top)
    for shape in (ops_title, ops_table):
        if shape is not None:
            shape._element.getparent().remove(shape._element)

    if issues_title is not None:
        issues_title.top = top_start
        run = issues_title.text_frame.paragraphs[0].add_run()
        run.text = " (continued)"
        run.font.size = Pt(10)
        run.font.italic = True
        issues_shape.top = top_start + issues_title.height + gap
    else:
        issues_shape.top = top_start
    return new_slide


def _layout_issue_rows(prs, first_slide, rows: list, insert_at: int):
    current_slide = first_slide
    remaining = rows
    while True:
        shape = _find_issues_table_shape(current_slide)
        font_pt = _table_font_pt(shape.table)
        available = max(prs.slide_height - shape.top - ISSUES_BOTTOM_MARGIN, 0)
        remaining, used_height = _fit_status_rows(shape.table, remaining, available, font_pt)
        shape.height = used_height
        if not remaining:
            return insert_at
        current_slide = _new_issues_continuation_slide(prs, first_slide, insert_at)
        insert_at += 1


def build_phibro_status_report(template_path: str, output_path: str, week: dict, report: dict):
    """Opens template_path, updates the cover date, the Project: table's
    Reporting Period/Date, the Project Status narrative, and the unified
    Milestone/Status/Value table on slide 2 - spilling onto cloned
    "(continued)" slides as needed so every point from the week is shown -
    and slide 3's Issues Resolved Last Week table the same way. The rest of
    slide 3 (Workstream table) and everything else is left untouched.
    template_path is never modified."""
    prs = Presentation(template_path)

    _set_cover_date(prs, week["slide1_date_text"])

    slide = prs.slides[STATUS_SLIDE_INDEX]

    start = datetime.strptime(week["start"], "%Y-%m-%d")
    end = datetime.strptime(week["end"], "%Y-%m-%d")
    reporting_period_text = f"{start.strftime('%m/%d/%Y')} – {end.strftime('%m/%d/%Y')}"
    reporting_date_text = end.strftime("%m/%d/%Y")
    info_shape = _set_project_info(slide, reporting_period_text, reporting_date_text)

    narrative_shape = _find_status_narrative_shape(slide)
    narrative_text = str(report.get("project_status_summary") or "").strip()
    if not narrative_text:
        narrative_text = "No project status summary was provided for this reporting period."
    _set_project_status_narrative(narrative_shape, narrative_text)

    progress_title = _find_progress_title_shape(slide)
    milestone_shape = _find_table_shape(slide, MILESTONE_TABLE_HEADER)
    _fit_status_slide_header(slide, progress_title, milestone_shape, info_shape)

    # Slide 3's Issues table is laid out before slide 2's Milestone table so
    # its "(continued)" slides are inserted while it's still at its template
    # index; slide 2's own continuation slides then push them all down
    # together, keeping every continuation right after its source slide.
    _layout_issue_rows(prs, prs.slides[WORKSTREAM_SLIDE_INDEX], _resolved_issue_rows(report), WORKSTREAM_SLIDE_INDEX + 1)

    rows = _status_rows_from_review(report) or _build_status_rows(report)
    rows = _add_status_icons(rows)
    _layout_status_rows(prs, slide, rows, STATUS_SLIDE_INDEX + 1)

    prs.save(output_path)
    return output_path
