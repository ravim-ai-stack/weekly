"""Dispatches the generated weekly status report PPT using Microsoft Graph API."""

import base64
import os
import re

from core.graph_api import graph_post
from .email_template import build_email_html


def is_valid_email(email: str) -> bool:
    pattern = r"^[a-zA-Z0-9_.+\-]+@[a-zA-Z0-9\-]+\.[a-zA-Z0-9\-.]+$"
    return bool(re.match(pattern, email.strip()))


def dispatch_report_email(recipient_email: str, cc_email: str, client_name: str,
                           date_range_text: str, filename: str, file_bytes: bytes,
                           sender_email: str = None, subject: str = None,
                           custom_message: str = None) -> dict:
    if not sender_email or not sender_email.strip():
        sender_email = os.getenv("SYSTEM_SENDER_EMAIL", "")

    if not sender_email.strip() or not is_valid_email(sender_email.strip()):
        return {"success": False, "message": "Missing or invalid sender email. Please check your .env file."}

    if not is_valid_email(recipient_email):
        return {"success": False, "message": "Invalid recipient email address format."}

    cc_recipients = []
    if cc_email and cc_email.strip():
        raw_ccs = [e.strip() for e in cc_email.replace(";", ",").split(",") if e.strip()]
        for cc in raw_ccs:
            if is_valid_email(cc):
                cc_recipients.append({"emailAddress": {"address": cc}})
            else:
                return {"success": False, "message": f"Invalid CC email format detected: {cc}"}

    attachment_content_types = {
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    _, ext = os.path.splitext(filename.lower())

    attachment = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": filename,
        "contentType": attachment_content_types.get(ext, "application/octet-stream"),
        "contentBytes": base64.b64encode(file_bytes).decode("utf-8"),
    }

    email_payload = {
        "message": {
            "subject": (subject or f"Weekly Status Report - {client_name} ({date_range_text})").strip(),
            "body": {
                "contentType": "HTML",
                "content": build_email_html(client_name, date_range_text, custom_message=custom_message),
            },
            "toRecipients": [{"emailAddress": {"address": recipient_email.strip()}}],
            "ccRecipients": cc_recipients,
            "attachments": [attachment],
        },
        "saveToSentItems": "true",
    }

    try:
        endpoint = f"users/{sender_email.strip()}/sendMail"
        graph_post(endpoint, json_data=email_payload)
        return {"success": True, "message": f"Weekly report emailed successfully to {recipient_email}"}
    except Exception as e:
        return {"success": False, "message": f"Microsoft Graph API Error: {str(e)}"}
