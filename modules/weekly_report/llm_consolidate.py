"""Calls the Databricks-served LLM to consolidate raw weekly notes into the
structured sections used on slide 9 of the status report template."""

import json
import os
import re
from difflib import SequenceMatcher

import requests

SYSTEM_PROMPT = """You are an assistant that turns a consultant's raw, messy weekly \
status notes into a structured JSON weekly status report for a client-facing \
PowerPoint slide.

THE SINGLE MOST IMPORTANT RULE: every field you output must be grounded ONLY in what \
the user's raw notes actually say. Never invent, guess, embellish, or fill in \
plausible-sounding workstream names, milestone areas, task labels, activities, or risks \
that are not stated (or very directly implied) in the notes. You have no knowledge of any \
prior report, prior week's milestones, or "typical" project workstreams - use ONLY the \
text the user gives you in this request. If a section has nothing to report, return it \
empty (see per-field rules below) rather than making something up to fill the section.

The raw notes you receive are often several short update blocks, one per team member, each \
prefixed with "<Name>: ". Read every block as part of the same week's notes and pull \
completed tasks, milestones, and upcoming activities from ALL of them together - not just \
the first or longest block. A team member's update being brief (a sentence or two) does not \
make it "just a flat list of tasks" - a short update can still name a real workstream.

Return ONLY valid JSON (no markdown fences, no commentary) matching this exact schema:

{
  "completed_tasks": [
    "<short description>"   // plain bullet sentence, no task-number prefix at all, e.g. "Migrated the shipment ingestion pipeline to the new Bronze schema."
    // ONLY if the notes explicitly state a task number/ticket ID (e.g. "task 82", "#109", "ticket 112"), reuse it verbatim as a prefix, e.g. "Task No - 82: Implemented changes to display only available reps."
  ],
  "milestones": [
    {
      "area": "<milestone/workstream name, at most 3-4 words, e.g. 'ECC Migration', 'CRM (OpFocus)'>",
      "task": "<task label, at most 2-4 words, e.g. 'ECC support', 'Service account', 'Bronze access'>",
      "status": "Completed" | "In Progress" | "At Risk" | "Not Started",
      "remarks": "<1 short sentence remark, based only on what the notes say>"
    }
  ],
  "upcoming_activities": [
    "<short bullet describing brand-new work planned to START next week, only if stated in the notes>"
  ],
  "next_action": [
    "<short bullet describing the immediate next step for a CURRENTLY ONGOING (milestones) item, only if stated in the notes>"
  ],
  "risks_issues": "None",  // or a short description of real risks/issues explicitly mentioned in the notes
  "project_status_summary": "<2-4 sentence flowing narrative (not bullets) summarizing the week's overall engagement status and progress, e.g. 'The engagement is progressing as planned. During the reporting period, we focused on ...'>"
}

Rules:
- Keep each bullet/remark short (1-2 sentences), matching the tone of a client status report.
- "milestones": every area/task/remark must come from a specific workstream, system, feature, or \
  piece of ongoing work the user actually named in the notes - naming it can be as short as a \
  project/system/feature name (e.g. "ECC Migration", "OpFocus", "Profitability Report", "Bronze \
  access rollout"); it does not need a long description to count as a workstream. Only return an \
  EMPTY list when the notes are truly generic, unnamed tasks with no project/system/feature \
  reference anywhere in any block - do not invent a workstream name that isn't in the notes, but \
  don't demand an elaborate description either before treating something as a milestone.
- Keep "area" and "task" SHORT LABELS, not sentences (this table column is narrow) - a short noun \
  phrase like "ECC Migration" or "Bronze access", never a full sentence. Put any detail/explanation \
  in "remarks" instead, which also must stay to one short sentence.
- "remarks" must NEVER be left empty for any milestone - every milestone table cell is shown to the \
  client, so blank cells are not acceptable (the only field allowed to read "None" is "risks_issues"). \
  Always give at least a short concrete detail grounded in the notes; if the notes genuinely give no \
  more detail than the area/task name itself, write a plain sentence restating it (e.g. area "ECC \
  Migration" -> remarks "Continuing the ECC Migration workstream.") rather than leaving it blank.
- A given piece of work belongs in EXACTLY ONE place - it must never be described in more than one \
  of "completed_tasks", "milestones", and "upcoming_activities". The "milestones" table is reserved \
  for a named workstream that is actively ONGOING this week - prefer giving it status "In Progress" \
  and put any detail in "remarks". If that same workstream is actually finished, write it as a \
  "completed_tasks" bullet instead (never also add a "Completed" milestone row for it). If it hasn't \
  started yet or is blocked/at risk, describe it under "upcoming_activities" or "risks_issues" \
  instead (never also add a "Not Started"/"At Risk" milestone row for it). In short: only use \
  "milestones" for work you would mark "In Progress" - never restate the same item there that you \
  already captured elsewhere.
- "upcoming_activities" vs "next_action" - both are forward-looking, but keep them distinct: \
  "upcoming_activities" is for work that hasn't started yet and will begin next week (phrases like \
  "planning to", "about to start", "will kick off"); "next_action" is the immediate next step for \
  work that's already ONGOING this week (i.e. something you also put in "milestones") - phrases like \
  "next step is", "will continue", "still working on", "need to" describing a next step for existing \
  work. Check every block, not just the first. Only return an EMPTY list for either when truly \
  nothing in any block hints at that kind of forward-looking work.
- "risks_issues": only report a risk/issue that is explicitly described as a blocker, concern, delay, \
  or open question in the notes. If none is mentioned, set this to exactly "None" - never invent a \
  hypothetical risk.
- "project_status_summary": write this LAST, as a short flowing paragraph (not a bullet list) that \
  summarizes the week's overall engagement status, in the tone of a one-paragraph executive summary. \
  Base it ONLY on what you already put in completed_tasks/milestones/upcoming_activities/risks_issues \
  - never introduce a fact here that isn't reflected in those fields.
- NEVER invent, guess, or assign a task number/ticket ID that is not explicitly stated in the notes.
  Do not use a list position, count, or any made-up number as a task number. Most bullets will have
  no task number at all - that is expected and correct, just write the plain sentence.
- Only reuse a task number/ticket ID if the notes explicitly contain one (e.g. "task 82", "#109",
  "ticket 112"), and reuse it exactly as given.
"""


def _status_map():
    return {
        "completed": "\U0001F7E2 Completed",
        "done": "\U0001F7E2 Completed",
        "in progress": "\U0001F7E1 In Progress",
        "ongoing": "\U0001F7E1 In Progress",
        "at risk": "\U0001F534 At Risk",
        "blocked": "\U0001F534 At Risk",
        "not started": "⚪ Not Started",
    }


TASK_NUMBER_PREFIX_RE = re.compile(r"^Task\s*(?:No\.?|#)?\s*[-–:]?\s*(\d+)\s*:?\s*", re.IGNORECASE)


def _strip_unverified_task_numbers(items: list, raw_notes: str) -> list:
    """Belt-and-suspenders check: drops any 'Task No - N:' style prefix the
    model added unless that exact number N is actually present in the raw
    notes, so the model can never invent/hallucinate a task number."""
    cleaned = []
    for item in items:
        text = str(item).strip()
        match = TASK_NUMBER_PREFIX_RE.match(text)
        if match:
            number = match.group(1)
            if not re.search(rf"(?<!\d){re.escape(number)}(?!\d)", raw_notes):
                text = text[match.end():].strip()
                if text:
                    text = text[0].upper() + text[1:]
        cleaned.append(text)
    return cleaned


def _normalize_for_comparison(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]", "", str(text).lower()).strip()


def _is_duplicate_text(a: str, b: str, threshold: float = 0.6) -> bool:
    a_n, b_n = _normalize_for_comparison(a), _normalize_for_comparison(b)
    if not a_n or not b_n:
        return False
    return SequenceMatcher(None, a_n, b_n).ratio() >= threshold


def _milestone_label(milestone: dict) -> str:
    return " - ".join(p for p in (str(milestone.get("area", "")).strip(), str(milestone.get("task", "")).strip()) if p)


def _milestone_summary_text(milestone: dict) -> str:
    label = _milestone_label(milestone)
    remarks = str(milestone.get("remarks", "")).strip()
    if label and remarks and label.lower() != remarks.lower():
        return f"{label}: {remarks}"
    return remarks or label


def _ensure_milestone_fields_filled(report: dict) -> None:
    """Belt-and-suspenders check: every milestone must have BOTH a non-empty
    area/task label and non-empty remarks, since every template renders
    these into a visible cell - only risks_issues is ever allowed to read
    "None". Cross-fills whichever side is empty from the other (never
    invents new information); "Update in progress." is only a last-resort
    fallback if a milestone somehow has neither."""
    for milestone in report.get("milestones", []):
        label = _milestone_label(milestone)
        remarks = str(milestone.get("remarks", "")).strip()
        if not remarks:
            milestone["remarks"] = label or "Update in progress."
        elif not label:
            milestone["area"] = remarks


def reclassify_milestones_by_status(report: dict) -> None:
    """The Milestone & Supporting Activities table is reserved for work
    that's actively ongoing - belt-and-suspenders check that moves any
    milestone row NOT in the "In Progress" state out of "milestones" instead
    of just dropping it, so every point still reaches the deck somewhere:
    a "Completed" row is folded into completed_tasks (completed work only
    ever appears in Completed Tasks), and an "At Risk"/"Not Started"/other
    row is folded into upcoming_activities (it hasn't progressed yet, so
    it's forward-looking). Skips folding in a duplicate if that same point
    is already captured in the destination list."""
    _ensure_milestone_fields_filled(report)
    completed_tasks = report.setdefault("completed_tasks", [])
    upcoming_activities = report.setdefault("upcoming_activities", [])
    kept_milestones = []
    for milestone in report.get("milestones", []):
        status = str(milestone.get("status", "")).strip().lower()
        if "progress" in status:
            kept_milestones.append(milestone)
            continue
        summary = _milestone_summary_text(milestone)
        if not summary:
            continue
        destination = completed_tasks if "complet" in status else upcoming_activities
        if not any(_is_duplicate_text(summary, existing) for existing in destination):
            destination.append(summary)
    report["milestones"] = kept_milestones


def _extract_json(text: str) -> dict:
    text = text.strip()
    fence_match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1)
    else:
        brace_match = re.search(r"\{.*\}", text, re.DOTALL)
        if brace_match:
            text = brace_match.group(0)
    return json.loads(text)


def consolidate_notes(raw_notes: str) -> dict:
    """Sends raw weekly notes to the Databricks model serving endpoint and
    returns the structured report dict. Raises RuntimeError on failure."""

    host = os.getenv("DATABRICKS_HOST", "").rstrip("/")
    token = os.getenv("DATABRICKS_TOKEN", "")
    model = os.getenv("MODEL_NAME", "")

    if not host or not token or not model:
        raise RuntimeError(
            "Missing Databricks configuration. Ensure DATABRICKS_HOST, "
            "DATABRICKS_TOKEN and MODEL_NAME are set in your .env file."
        )

    url = f"{host}/serving-endpoints/{model}/invocations"
    payload = {
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": raw_notes},
        ],
        "temperature": 0.2,
        "max_tokens": 2000,
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(f"Failed to reach the LLM endpoint: {exc}") from exc

    body = resp.json()
    try:
        content = body["choices"][0]["message"]["content"]
    except (KeyError, IndexError) as exc:
        raise RuntimeError(f"Unexpected LLM response shape: {body}") from exc

    try:
        report = _extract_json(content)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"LLM did not return valid JSON: {exc}\nRaw content: {content}") from exc

    report.setdefault("completed_tasks", [])
    report.setdefault("milestones", [])
    report.setdefault("upcoming_activities", [])
    report.setdefault("next_action", [])
    report.setdefault("risks_issues", "None")
    report.setdefault("project_status_summary", "")

    report["completed_tasks"] = _strip_unverified_task_numbers(report["completed_tasks"], raw_notes)

    status_map = _status_map()
    for item in report["milestones"]:
        raw_status = str(item.get("status", "")).strip().lower()
        item["status"] = status_map.get(raw_status, item.get("status", "In Progress"))

    reclassify_milestones_by_status(report)

    report["upcoming_activities"] = _carry_continuing_milestones_forward(
        report["milestones"], report["upcoming_activities"]
    )

    return report


def _carry_continuing_milestones_forward(milestones: list, upcoming_activities: list) -> list:
    """Matches the reference template's own pattern (e.g. 'ECC Migration -
    Continue to Support.'): every remaining milestone is, by this point,
    "In Progress" (anything else was already moved out by
    reclassify_milestones_by_status), so it's also reflected as a next-week
    upcoming activity, unless the notes already gave an explicit different
    next-week plan for it."""
    for item in milestones:
        area = str(item.get("area", "")).strip()
        if not area:
            continue
        already_covered = any(area.lower() in str(u).lower() for u in upcoming_activities)
        if not already_covered:
            upcoming_activities.append(f"{area} – Continue to Support.")
    return upcoming_activities
