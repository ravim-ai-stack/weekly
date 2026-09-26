"""DataPattern Weekly Report Agent - consolidates raw weekly notes into the
DataPattern weekly status report PPT (slide 9 only) and emails it out."""

import glob
import os
import re
import uuid

from dotenv import load_dotenv
from flask import Flask, jsonify, redirect, render_template, request, send_from_directory

from modules.weekly_report.hmh_builder import build_hmh_report
from modules.weekly_report.llm_consolidate import consolidate_notes, reclassify_milestones_by_status
from modules.weekly_report.mailer import dispatch_report_email
from modules.weekly_report.medtronic_builder import build_medtronic_report
from modules.weekly_report.medtronic_timesheet import (
    add_daily_entry,
    build_range_workbook,
    build_weekly_report_from_timesheet,
    get_day,
    remove_daily_entry,
)
from modules.weekly_report.phibro_status_builder import build_phibro_status_report
from modules.weekly_report.pptx_builder import build_report
from modules.weekly_report.teams_config import DEFAULT_LAYOUT, MEDTRONIC_TEAM, PROJECT_LAYOUTS, TEAMS, get_report_config
from modules.weekly_report.updates_store import add_entry, get_entries, remove_entry
from modules.weekly_report.week_utils import format_date_range

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
# Vercel's deployment filesystem is read-only; only /tmp is writable.
WRITABLE_DIR = "/tmp" if os.getenv("VERCEL") else BASE_DIR
OUTPUT_DIR = os.path.join(WRITABLE_DIR, "output")
CLIENT_NAME = "Nova Biomedical"

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

SAFE_FILENAME_RE = re.compile(r"^[A-Za-z0-9_\-. ]+\.(pptx|docx|xlsx)$")


def _resolve_medtronic_source(template_path: str, client_name: str) -> str:
    """Medtronic's report is one growing document - each generation must
    build on top of whatever was generated last time (so earlier weeks are
    preserved), not restart from the frozen template. Returns the most
    recently generated Medtronic output file, or template_path on the very
    first generation."""
    pattern = os.path.join(OUTPUT_DIR, f"DataPattern_{client_name}_Engagement_*.docx")
    candidates = sorted(glob.glob(pattern))
    return candidates[-1] if candidates else template_path


@app.route("/")
def index():
    return redirect("/step1")


@app.route("/step1")
def step1():
    return render_template("step1.html", active_step=1, client_name=CLIENT_NAME, teams=TEAMS)


@app.route("/step2")
def step2():
    return render_template("step2.html", active_step=2, client_name=CLIENT_NAME, medtronic_team=MEDTRONIC_TEAM)


@app.route("/step3")
def step3():
    return render_template(
        "step3.html", active_step=3, client_name=CLIENT_NAME,
        project_layouts=PROJECT_LAYOUTS, default_layout=DEFAULT_LAYOUT,
    )


@app.route("/step4")
def step4():
    return render_template("step4.html", active_step=4, client_name=CLIENT_NAME)


@app.route("/step5")
def step5():
    return render_template(
        "step5.html", active_step=5, client_name=CLIENT_NAME,
        sender_email=os.getenv("SYSTEM_SENDER_EMAIL", ""),
    )


@app.route("/api/week-entries", methods=["GET"])
def api_week_entries_get():
    team = (request.args.get("team") or "").strip()
    project = (request.args.get("project") or "").strip()
    start_date = (request.args.get("start") or "").strip()
    end_date = (request.args.get("end") or "").strip()

    if not all([team, project, start_date, end_date]):
        return jsonify(success=False, error="Missing team/project/week."), 400

    entries = get_entries(team, project, start_date, end_date)
    return jsonify(success=True, entries=entries)


@app.route("/api/week-entries", methods=["POST"])
def api_week_entries_post():
    data = request.get_json(force=True) or {}
    team = (data.get("team") or "").strip()
    project = (data.get("project") or "").strip()
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()
    name = (data.get("name") or "").strip()
    update_text = (data.get("update") or "").strip()

    if not all([team, project, start_date, end_date]):
        return jsonify(success=False, error="Missing team/project/week."), 400
    if not name:
        return jsonify(success=False, error="Please enter your name."), 400
    if not update_text:
        return jsonify(success=False, error="Please enter your update."), 400

    entries = add_entry(team, project, start_date, end_date, name, update_text)
    return jsonify(success=True, entries=entries)


@app.route("/api/week-entries", methods=["DELETE"])
def api_week_entries_delete():
    data = request.get_json(force=True) or {}
    team = (data.get("team") or "").strip()
    project = (data.get("project") or "").strip()
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()
    index = data.get("index")

    if not all([team, project, start_date, end_date]) or index is None:
        return jsonify(success=False, error="Missing team/project/week/index."), 400

    try:
        index = int(index)
    except (TypeError, ValueError):
        return jsonify(success=False, error="Invalid entry index."), 400

    entries = remove_entry(team, project, start_date, end_date, index)
    return jsonify(success=True, entries=entries)


@app.route("/api/medtronic/daily-entry", methods=["GET"])
def api_medtronic_daily_entry_get():
    date_str = (request.args.get("date") or "").strip()
    if not date_str:
        return jsonify(success=False, error="Missing date."), 400
    try:
        day = get_day(date_str)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400
    return jsonify(success=True, day=day)


@app.route("/api/medtronic/daily-entry", methods=["POST"])
def api_medtronic_daily_entry_post():
    data = request.get_json(force=True) or {}
    date_str = (data.get("date") or "").strip()
    person = (data.get("person") or "").strip()
    description = (data.get("description") or "").strip()

    if not date_str:
        return jsonify(success=False, error="Please pick a date."), 400
    if person not in MEDTRONIC_TEAM:
        return jsonify(success=False, error="Please select your name from the list."), 400
    if not description:
        return jsonify(success=False, error="Please enter a description."), 400
    try:
        hrs = float(data.get("hrs"))
    except (TypeError, ValueError):
        return jsonify(success=False, error="Hrs must be a number."), 400
    if hrs <= 0 or hrs > 24:
        return jsonify(success=False, error="Hrs must be between 0 and 24."), 400

    try:
        day = add_daily_entry(date_str, person, description, hrs)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400

    return jsonify(success=True, day=day, download_url="/download/Medtronic_Time_Sheet.xlsx")


@app.route("/api/medtronic/daily-entry", methods=["DELETE"])
def api_medtronic_daily_entry_delete():
    data = request.get_json(force=True) or {}
    date_str = (data.get("date") or "").strip()
    person = (data.get("person") or "").strip()

    if not date_str or not person:
        return jsonify(success=False, error="Missing date/person."), 400

    try:
        day = remove_daily_entry(date_str, person)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400

    return jsonify(success=True, day=day, download_url="/download/Medtronic_Time_Sheet.xlsx")


@app.route("/api/medtronic/custom-download", methods=["POST"])
def api_medtronic_custom_download():
    data = request.get_json(force=True) or {}
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()

    if not start_date or not end_date:
        return jsonify(success=False, error="Please pick both a start and end date."), 400

    try:
        filename = build_range_workbook(start_date, end_date)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400

    return jsonify(success=True, filename=filename, download_url=f"/download/{filename}")


@app.route("/api/medtronic/weekly-from-timesheet", methods=["POST"])
def api_medtronic_weekly_from_timesheet():
    data = request.get_json(force=True) or {}
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()

    if not start_date or not end_date:
        return jsonify(success=False, error="Please pick both a start and end date."), 400

    try:
        week = format_date_range(start_date, end_date)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400

    try:
        report = build_weekly_report_from_timesheet(start_date, end_date)
    except ValueError as exc:
        return jsonify(success=False, error=str(exc)), 400
    except RuntimeError as exc:
        return jsonify(success=False, error=str(exc)), 502

    return jsonify(success=True, report=report, week=week)


@app.route("/api/week", methods=["POST"])
def api_week():
    data = request.get_json(force=True) or {}
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()
    if not start_date or not end_date:
        return jsonify(success=False, error="Please pick both a start and end date."), 400
    try:
        week = format_date_range(start_date, end_date)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400
    return jsonify(success=True, week=week)


@app.route("/api/consolidate", methods=["POST"])
def api_consolidate():
    data = request.get_json(force=True) or {}
    entries = data.get("entries") or []
    raw_notes = (data.get("raw_notes") or "").strip()
    start_date = (data.get("start_date") or "").strip()
    end_date = (data.get("end_date") or "").strip()

    if not raw_notes and entries:
        raw_notes = "\n\n".join(
            f"{(e.get('name') or '').strip()}: {(e.get('update') or '').strip()}"
            for e in entries
            if (e.get("update") or "").strip()
        )

    if not raw_notes:
        return jsonify(success=False, error="At least one person needs to add their weekly update first."), 400
    if not start_date or not end_date:
        return jsonify(success=False, error="Please pick the start and end date for this report."), 400

    try:
        week = format_date_range(start_date, end_date)
    except ValueError:
        return jsonify(success=False, error="Invalid date."), 400

    try:
        report = consolidate_notes(raw_notes)
    except RuntimeError as exc:
        return jsonify(success=False, error=str(exc)), 502

    return jsonify(success=True, report=report, week=week)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True) or {}
    report = data.get("report") or {}
    week = data.get("week") or {}
    project = (data.get("project") or "").strip()

    if not week.get("range_text") or not week.get("slide1_date_text"):
        return jsonify(success=False, error="Missing week range."), 400

    config = get_report_config(project)
    template_path = os.path.join(TEMPLATE_DIR, config["template_file"])
    if not os.path.exists(template_path):
        return jsonify(success=False, error="Template file not found on the server."), 500

    client_name = config["client_name"]
    file_ext = config.get("file_ext", "pptx")
    filename = f"DataPattern_{client_name}_Engagement_{week.get('start', uuid.uuid4().hex[:8])}.{file_ext}"
    filename = re.sub(r"[^A-Za-z0-9_.\- ]", "", filename)
    output_path = os.path.join(OUTPUT_DIR, filename)

    reclassify_milestones_by_status(report)

    try:
        # "builder" picks the backend renderer and defaults to "layout" when
        # not set separately - most projects use the same value for both,
        # but a project can show a review UI meant for a different-shaped
        # template than the one it renders onto (its "builder"). Not
        # currently used that way by any configured project, but kept
        # separate from "layout" for that case.
        builder = config.get("builder", config.get("layout"))
        if builder == "medtronic":
            source_path = _resolve_medtronic_source(template_path, client_name)
            build_medtronic_report(source_path, output_path, week, report)
        elif builder == "hmh":
            build_hmh_report(template_path, output_path, week, report)
        elif builder == "phibro_status":
            build_phibro_status_report(template_path, output_path, week, report)
        else:
            build_report(template_path, output_path, week, report, slide_index=config["slide_index"])
    except Exception as exc:
        return jsonify(success=False, error=f"Failed to generate the report: {exc}"), 500

    return jsonify(success=True, filename=filename, download_url=f"/download/{filename}")


@app.route("/download/<path:filename>")
def download(filename):
    if not SAFE_FILENAME_RE.match(filename):
        return jsonify(success=False, error="Invalid filename."), 400
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route("/api/send-email", methods=["POST"])
def api_send_email():
    data = request.get_json(force=True) or {}
    filename = (data.get("filename") or "").strip()

    if not filename or not SAFE_FILENAME_RE.match(filename):
        return jsonify(success=False, message="Missing or invalid report file. Please generate the report again."), 400

    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(file_path):
        return jsonify(success=False, message="Generated report file was not found. Please generate it again."), 404

    with open(file_path, "rb") as f:
        file_bytes = f.read()

    result = dispatch_report_email(
        recipient_email=data.get("recipient_email", ""),
        cc_email=data.get("cc_email", ""),
        client_name=CLIENT_NAME,
        date_range_text=data.get("range_text", ""),
        filename=filename,
        file_bytes=file_bytes,
        sender_email=data.get("sender_email", ""),
        subject=data.get("subject", ""),
        custom_message=data.get("message", ""),
    )
    return jsonify(result), 200 if result["success"] else 400


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False, host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
