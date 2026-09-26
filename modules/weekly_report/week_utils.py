"""Turns the user-picked start/end dates into the date strings used on the
PPT: slide 9's 'Week of ...' range and slide 1's single engagement date."""

from datetime import datetime

_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sept", 10: "Oct", 11: "Nov", 12: "Dec",
}


def format_date_range(start_date_str: str, end_date_str: str) -> dict:
    """start_date_str/end_date_str: 'YYYY-MM-DD' from HTML date inputs."""
    start = datetime.strptime(start_date_str, "%Y-%m-%d")
    end = datetime.strptime(end_date_str, "%Y-%m-%d")

    if end < start:
        start, end = end, start

    range_text = (
        f"{_MONTH_ABBR[start.month]} {start.day:02d} - "
        f"{_MONTH_ABBR[end.month]} {end.day:02d}, {end.year}"
    )
    slide1_date_text = f"{end.day:02d}-{end.strftime('%B')}-{end.year}"

    return {
        "start": start.strftime("%Y-%m-%d"),
        "end": end.strftime("%Y-%m-%d"),
        "range_text": range_text,
        "display_text": range_text,
        "slide1_date_text": slide1_date_text,
    }
