"""Builds the HTML email body used when dispatching the weekly status report."""

from html import escape


def build_email_html(client_name: str, date_range_text: str, custom_message: str = None) -> str:
    body = (custom_message or "").strip() or (
        f"Please find attached the weekly status report for {client_name} "
        f"covering {date_range_text}.\n\n"
        "Let us know if you have any questions."
    )
    paragraphs = "\n".join(
        f"<p>{escape(p).replace(chr(10), '<br>')}</p>" for p in body.split("\n\n") if p.strip()
    )

    return f"""\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Weekly Status Report</title>
</head>
<body style="margin:0;padding:0;background:#ffffff;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.7;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;padding:32px 24px;">
    <tr>
      <td align="left" style="max-width:680px;">
        {paragraphs}
      </td>
    </tr>
  </table>
</body>
</html>"""
