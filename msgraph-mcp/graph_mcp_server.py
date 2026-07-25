#!/usr/bin/env python3
"""
Microsoft Graph MCP server for PERSONAL Microsoft accounts
(@hotmail.com / @outlook.com / @live.com).

Why this exists: Anthropic's built-in Microsoft 365 connector registers its
OAuth app for work/school tenants only, so personal accounts are rejected.
Graph itself has no such limit -- you just need your own app registration
whose signInAudience includes personal Microsoft accounts.

Auth: MSAL device-code flow. You authenticate once in a browser; the refresh
token is cached on disk and renewed silently thereafter.

Transport: stdio (for Claude Desktop).

Scope of access: read and search mail, create drafts, read and write calendar,
read OneDrive. It CANNOT send mail -- the Mail.Send scope is not requested.
"""

import os
import sys
import json
import atexit
import pathlib
from typing import Any, Optional

import msal
import httpx
from mcp.server.fastmcp import FastMCP

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

CLIENT_ID = os.environ.get("MSGRAPH_CLIENT_ID")
if not CLIENT_ID:
    sys.exit(
        "MSGRAPH_CLIENT_ID is not set.\n"
        "Register an app in Azure (see README) and export its Application "
        "(client) ID as MSGRAPH_CLIENT_ID."
    )

# "common" accepts both personal and work/school accounts.
# Use "consumers" to restrict to personal Microsoft accounts only.
AUTHORITY = os.environ.get(
    "MSGRAPH_AUTHORITY", "https://login.microsoftonline.com/common"
)

# Delegated scopes. MSAL adds openid/profile/offline_access automatically --
# do not list them here or MSAL will raise.
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",   # implies Mail.Read; needed to create drafts
    "Calendars.ReadWrite",
    "Files.Read.All",
]

# Deliberately NOT requested: Mail.Send.
# Every mail tool in this server creates drafts only. Without this scope the
# server is incapable of sending mail even if a future tool tried to.

GRAPH = "https://graph.microsoft.com/v1.0"

CACHE_PATH = pathlib.Path(
    os.environ.get(
        "MSGRAPH_TOKEN_CACHE",
        pathlib.Path.home() / ".msgraph-mcp" / "token_cache.json",
    )
)
CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------
# Token handling
# --------------------------------------------------------------------------

_cache = msal.SerializableTokenCache()
if CACHE_PATH.exists():
    _cache.deserialize(CACHE_PATH.read_text())


def _persist_cache() -> None:
    if _cache.has_state_changed:
        CACHE_PATH.write_text(_cache.serialize())
        try:
            CACHE_PATH.chmod(0o600)
        except OSError:
            pass


atexit.register(_persist_cache)

_app = msal.PublicClientApplication(
    CLIENT_ID, authority=AUTHORITY, token_cache=_cache
)


def _token() -> str:
    """Return a valid access token, refreshing or prompting as needed."""
    accounts = _app.get_accounts()
    if accounts:
        result = _app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _persist_cache()
            return result["access_token"]

    # No cached account, or refresh failed -> device code flow.
    flow = _app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            f"Failed to start device flow: {json.dumps(flow, indent=2)}"
        )

    # Device-code prompts go to stderr so they never corrupt the stdio
    # JSON-RPC stream that Claude Desktop reads on stdout.
    print(flow["message"], file=sys.stderr, flush=True)

    result = _app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(
            f"Authentication failed: {result.get('error_description', result)}"
        )
    _persist_cache()
    return result["access_token"]


def _get(path: str, params: Optional[dict] = None) -> dict:
    r = httpx.get(
        f"{GRAPH}{path}",
        headers={"Authorization": f"Bearer {_token()}"},
        params=params,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def _post(path: str, payload: dict) -> dict:
    r = httpx.post(
        f"{GRAPH}{path}",
        headers={
            "Authorization": f"Bearer {_token()}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )
    r.raise_for_status()
    return r.json() if r.content else {}


def _recipients(addresses: Optional[list[str]]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in (addresses or [])]


mcp = FastMCP("microsoft-graph")

# --------------------------------------------------------------------------
# Identity
# --------------------------------------------------------------------------


@mcp.tool()
def whoami() -> str:
    """Return the signed-in Microsoft account's name and email address."""
    me = _get("/me")
    return json.dumps(
        {
            "displayName": me.get("displayName"),
            "mail": me.get("mail") or me.get("userPrincipalName"),
            "id": me.get("id"),
        },
        indent=2,
    )


# --------------------------------------------------------------------------
# Mail
# --------------------------------------------------------------------------


@mcp.tool()
def search_mail(query: str, limit: int = 15) -> str:
    """
    Full-text search across the mailbox.

    query: words to match in subject, body, sender, or recipients.
    limit: maximum messages to return (1-50).
    """
    limit = max(1, min(limit, 50))
    data = _get(
        "/me/messages",
        params={
            "$search": f'"{query}"',
            "$top": limit,
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments,isDraft",
        },
    )
    out = []
    for m in data.get("value", []):
        out.append(
            {
                "id": m.get("id"),
                "subject": m.get("subject"),
                "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
                "to": [
                    r["emailAddress"]["address"] for r in m.get("toRecipients", [])
                ],
                "received": m.get("receivedDateTime"),
                "preview": (m.get("bodyPreview") or "")[:300],
                "hasAttachments": m.get("hasAttachments"),
                "isDraft": m.get("isDraft"),
            }
        )
    return json.dumps(out, indent=2)


@mcp.tool()
def list_recent_mail(limit: int = 15, folder: str = "inbox") -> str:
    """
    List the most recent messages in a folder.

    folder: well-known folder name such as inbox, sentitems, drafts, archive.
    limit: maximum messages to return (1-50).
    """
    limit = max(1, min(limit, 50))
    data = _get(
        f"/me/mailFolders/{folder}/messages",
        params={
            "$top": limit,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview",
        },
    )
    out = [
        {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
            "received": m.get("receivedDateTime"),
            "preview": (m.get("bodyPreview") or "")[:300],
        }
        for m in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def read_message(message_id: str) -> str:
    """Return the full body and metadata of one message, by its id."""
    m = _get(
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,ccRecipients,"
            "receivedDateTime,body,hasAttachments,conversationId"
        },
    )
    return json.dumps(
        {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
            "to": [r["emailAddress"]["address"] for r in m.get("toRecipients", [])],
            "cc": [r["emailAddress"]["address"] for r in m.get("ccRecipients", [])],
            "received": m.get("receivedDateTime"),
            "contentType": (m.get("body") or {}).get("contentType"),
            "body": (m.get("body") or {}).get("content"),
            "conversationId": m.get("conversationId"),
        },
        indent=2,
    )


@mcp.tool()
def create_draft(
    subject: str,
    body: str,
    to: list[str],
    cc: Optional[list[str]] = None,
    html: bool = False,
) -> str:
    """
    Create a draft email in the mailbox. Does NOT send it.

    to / cc: lists of email addresses.
    html: set true if body contains HTML, otherwise it is treated as plain text.
    """
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML" if html else "Text", "content": body},
        "toRecipients": _recipients(to),
        "ccRecipients": _recipients(cc),
    }
    created = _post("/me/messages", payload)
    return json.dumps(
        {
            "status": "draft created",
            "id": created.get("id"),
            "webLink": created.get("webLink"),
        },
        indent=2,
    )


@mcp.tool()
def reply_draft(message_id: str, comment: str) -> str:
    """Create a draft reply to an existing message, without sending it."""
    r = httpx.post(
        f"{GRAPH}/me/messages/{message_id}/createReply",
        headers={
            "Authorization": f"Bearer {_token()}",
            "Content-Type": "application/json",
        },
        json={"comment": comment},
        timeout=30,
    )
    r.raise_for_status()
    created = r.json()
    return json.dumps(
        {"status": "reply draft created", "id": created.get("id")}, indent=2
    )


# --------------------------------------------------------------------------
# Calendar
# --------------------------------------------------------------------------


@mcp.tool()
def list_events(start: str, end: str, limit: int = 25) -> str:
    """
    List calendar events in a date range.

    start / end: ISO 8601, e.g. 2026-07-24T00:00:00Z
    """
    limit = max(1, min(limit, 50))
    data = _get(
        "/me/calendarView",
        params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": limit,
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay",
        },
    )
    out = [
        {
            "id": e.get("id"),
            "subject": e.get("subject"),
            "start": (e.get("start") or {}).get("dateTime"),
            "end": (e.get("end") or {}).get("dateTime"),
            "timeZone": (e.get("start") or {}).get("timeZone"),
            "location": (e.get("location") or {}).get("displayName"),
            "organizer": (e.get("organizer") or {})
            .get("emailAddress", {})
            .get("address"),
            "isAllDay": e.get("isAllDay"),
        }
        for e in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def create_event(
    subject: str,
    start: str,
    end: str,
    time_zone: str = "America/Toronto",
    attendees: Optional[list[str]] = None,
    location: Optional[str] = None,
    body: Optional[str] = None,
) -> str:
    """
    Create a calendar event.

    start / end: ISO 8601 local wall-clock time, e.g. 2026-07-25T14:00:00
    time_zone: IANA time zone name for the times above.
    """
    payload: dict[str, Any] = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": time_zone},
        "end": {"dateTime": end, "timeZone": time_zone},
    }
    if attendees:
        payload["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees
        ]
    if location:
        payload["location"] = {"displayName": location}
    if body:
        payload["body"] = {"contentType": "Text", "content": body}

    created = _post("/me/events", payload)
    return json.dumps(
        {"status": "event created", "id": created.get("id")}, indent=2
    )


# --------------------------------------------------------------------------
# OneDrive
# --------------------------------------------------------------------------


@mcp.tool()
def search_onedrive(query: str, limit: int = 20) -> str:
    """Search OneDrive for files and folders matching a query string."""
    limit = max(1, min(limit, 50))
    data = _get(f"/me/drive/root/search(q='{query}')", params={"$top": limit})
    out = [
        {
            "id": i.get("id"),
            "name": i.get("name"),
            "type": "folder" if "folder" in i else "file",
            "size": i.get("size"),
            "modified": i.get("lastModifiedDateTime"),
            "webUrl": i.get("webUrl"),
        }
        for i in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def list_onedrive_folder(item_id: str = "root") -> str:
    """List the contents of a OneDrive folder. Use 'root' for the top level."""
    path = (
        "/me/drive/root/children"
        if item_id == "root"
        else f"/me/drive/items/{item_id}/children"
    )
    data = _get(path, params={"$top": 100})
    out = [
        {
            "id": i.get("id"),
            "name": i.get("name"),
            "type": "folder" if "folder" in i else "file",
            "size": i.get("size"),
            "modified": i.get("lastModifiedDateTime"),
        }
        for i in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def read_onedrive_file(item_id: str, max_chars: int = 20000) -> str:
    """
    Download a OneDrive file and return its text content.

    Works for text-based files (txt, md, csv, json, code). Binary formats such
    as PDF or DOCX return a notice rather than usable text.
    """
    r = httpx.get(
        f"{GRAPH}/me/drive/items/{item_id}/content",
        headers={"Authorization": f"Bearer {_token()}"},
        follow_redirects=True,
        timeout=60,
    )
    r.raise_for_status()
    try:
        text = r.content.decode("utf-8")
    except UnicodeDecodeError:
        return (
            "This file is not UTF-8 text (likely PDF, DOCX, or an image). "
            "Download it via its webUrl and upload it to the chat instead."
        )
    if len(text) > max_chars:
        text = text[:max_chars] + f"\n\n[truncated at {max_chars} characters]"
    return text


if __name__ == "__main__":
    # `login` forces the device-code flow, caches the refresh token, and exits.
    # Run this once before wiring the server into Claude Desktop -- Claude Desktop
    # cannot display the device prompt (it goes to stderr), so first sign-in must
    # happen here in a terminal.
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        _token()  # prints device code to stderr, blocks until you finish, caches
        print(f"Authenticated. Token cached at {CACHE_PATH}", file=sys.stderr)
        sys.exit(0)
    mcp.run()