"""Static team -> project hierarchy shown on step1. Add new teams/projects
here as they get set up; an empty list just shows "no projects yet" in the UI."""

TEAMS = {
    "Data Team": ["Nova Bio", "Phibro", "Medtronic"],
}

# Per-project weekly-status-report template: which file to start from, the
# client name used in the generated filename/email, the output file
# extension, and a "layout":
# - "standard" templates (Nova Bio, Phibro) share the same PPTX slide
#   structure - slide_index (0-indexed) picks which slide holds the
#   title/date + Completed Tasks / Milestone / Upcoming Activities /
#   Risk & Issues tables.
# - "medtronic" is a growing Word (.docx) document where every generation
#   prepends a brand-new dated page ahead of all previous weeks, built by
#   medtronic_builder.build_medtronic_report. Its "template_file" is only
#   ever used for the very first generation - after that, app.py resolves
#   the most recently generated Medtronic output as the source to build on.
# - "hmh" is a differently-structured PPTX (cover-slide date, plus Completed
#   Platforms & Tasks / Ongoing Work (In Progress) / Upcoming Activities /
#   Risk & Issues cells on the content slide - no 4-column milestone table)
#   built by hmh_builder.build_hmh_report instead of pptx_builder.build_report.
# Any project not listed here falls back to DEFAULT_REPORT_CONFIG (the Nova
# Bio template).
PROJECT_REPORT_CONFIG = {
    "Nova Bio": {
        "template_file": "DataPattern_Nova Biomedical_Engagement 04- Sep- 2026.pptx",
        "slide_index": 8,  # slide 9
        "client_name": "Nova Biomedical",
        "layout": "standard",
        "file_ext": "pptx",
    },
    "Phibro": {
        "template_file": "Phibro_Weekly_Status_Sep11_1.pptx",
        "client_name": "Phibro Animal Healthcare",
        # "layout" drives the step3 review UI: Phibro gets its own
        # "phibro_status" layout - one unified Milestone/Status/Value table,
        # matching the exact column headers of the template's own Milestone
        # table, instead of the separate Completed Tasks/Milestones/
        # Upcoming Activities/Risk & Issues fields every other "standard"
        # project shows. "builder" picks the backend renderer, which uses
        # that same reviewed table directly (see phibro_status_builder.py's
        # _status_rows_from_review()) to fill the new template's Milestone/
        # Status/Value table plus its Reporting Period/Date and Project
        # Status narrative. Slide 3 (Workstream) is untouched by that
        # builder.
        "layout": "phibro_status",
        "builder": "phibro_status",
        "file_ext": "pptx",
    },
    "Medtronic": {
        "template_file": "Medtronic_Weekly_Status_Report.docx",
        "client_name": "Medtronic",
        "layout": "medtronic",
        "file_ext": "docx",
    },
    "HMH": {
        "template_file": "HMH - DataPatten_Leadership_SyncUp (1).pptx",
        "client_name": "HMH",
        "layout": "hmh",
        "file_ext": "pptx",
    },
}

DEFAULT_REPORT_CONFIG = PROJECT_REPORT_CONFIG["Nova Bio"]

# Names offered in the "Responsible Person" dropdown for Medtronic's daily
# timesheet (step2's "Daily Update" mode) - see medtronic_timesheet.py.
MEDTRONIC_TEAM = ["Sivaprathish Sivamoorthy", "Manikandan", "Praveen Kumar Dharmalingam"]

# project -> layout, exposed to the frontend so step3 (review) can show only
# the fields that actually exist in that project's template.
PROJECT_LAYOUTS = {name: cfg["layout"] for name, cfg in PROJECT_REPORT_CONFIG.items()}
DEFAULT_LAYOUT = DEFAULT_REPORT_CONFIG["layout"]


def get_report_config(project: str) -> dict:
    return PROJECT_REPORT_CONFIG.get(project, DEFAULT_REPORT_CONFIG)
