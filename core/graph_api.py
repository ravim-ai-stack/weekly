"""
core/graph_api.py
Handles OAuth2 token generation and unified Microsoft Graph requests.
Adapted from the HR agent's core/graph_api.py.
"""

import os
import time
import requests
from requests.exceptions import HTTPError

# Global cache variables for the Graph token
_TOKEN_CACHE = None
_TOKEN_EXPIRY = 0


def get_graph_token():
    global _TOKEN_CACHE, _TOKEN_EXPIRY

    if _TOKEN_CACHE and time.time() < (_TOKEN_EXPIRY - 300):
        return _TOKEN_CACHE

    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }

    resp = requests.post(url, data=data)
    resp.raise_for_status()

    token_data = resp.json()

    _TOKEN_CACHE = token_data["access_token"]
    _TOKEN_EXPIRY = time.time() + token_data.get("expires_in", 3599)

    return _TOKEN_CACHE


def graph_post(endpoint, json_data=None, data=None, content_type="application/json"):
    """A robust POST wrapper for Microsoft Graph."""
    token = get_graph_token()

    url = f"https://graph.microsoft.com/v1.0/{endpoint}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }

    try:
        if json_data is not None:
            resp = requests.post(url, headers=headers, json=json_data)
        else:
            resp = requests.post(url, headers=headers, data=data)

        resp.raise_for_status()
        return resp.json() if resp.text.strip() else {}

    except HTTPError:
        error_details = resp.text if "resp" in locals() else "Unknown error"
        raise RuntimeError(
            f"\nGRAPH POST FAILED\n"
            f"URL: {url}\n"
            f"Status: {resp.status_code if 'resp' in locals() else 'N/A'}\n"
            f"Response: {error_details}\n"
        )
