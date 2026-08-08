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

Scope of access: read and search mail, create drafts, organize mail (move
messages, create folders, flag, categorize, mark read, and create inbox
rules that move mail), read and write calendar, read OneDrive, read and
write Microsoft To Do tasks, and read and write contacts.

It CANNOT send mail -- the Mail.Send scope is not requested.

It CANNOT delete anything. No tool in this server issues a DELETE against
a message, event, folder, task, contact, or Drive item, and no tool moves
anything to Deleted Items. The one DELETE call in this file is
delete_mail_rule, which removes an inbox-rule automation, not any mail --
see that tool's docstring for the distinction. OneDrive stays read-only
(Files.Read.All, not Files.ReadWrite.All), so nothing on Drive can be
changed or removed either.
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
    "Mail.ReadWrite",          # implies Mail.Read; drafts, folders, move/flag/categorize
    "Calendars.ReadWrite",
    "Files.Read.All",
    "MailboxSettings.ReadWrite",  # inbox rules (message rules live under mailbox settings)
    "Tasks.ReadWrite",            # Microsoft To Do
    "Contacts.ReadWrite",         # contacts
]

# Deliberately NOT requested:
#
#   Mail.Send -- every mail tool in this server creates drafts or organizes
#   existing mail (move, flag, categorize, mark read, rules). None of them
#   send anything, and without this scope the server is incapable of
#   sending mail even if a future tool tried to.
#
#   Files.ReadWrite.All -- would permit deleting or overwriting Drive items.
#   OneDrive access stays on Files.Read.All, which is read-only.
#
#   No scope grants delete rights anywhere. There is nothing this server
#   requests, and nothing it implements, that can delete a message, event,
#   folder, task, contact, or Drive item, or move one to Deleted Items.

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


def _acquire_token_silent() -> Optional[str]:
    """Return a cached/renewed token when one is available."""
    accounts = _app.get_accounts()
    if not accounts:
        return None
    result = _app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        _persist_cache()
        return result["access_token"]
    return None


def _token() -> str:
    """Return a valid token, allowing prompts only for the `login` command."""
    token = _acquire_token_silent()
    if token:
        return token

    # The explicit `login` command may update the cache while Claude Desktop
    # already has this module running. Reload it once before deciding that an
    # interactive sign-in is required.
    if CACHE_PATH.exists():
        _cache.deserialize(CACHE_PATH.read_text())
        token = _acquire_token_silent()
        if token:
            return token

    # A device prompt is invisible in a normal stdio MCP session. Starting the
    # flow here would make the tool appear frozen until the client times out.
    # Only the explicit terminal command is allowed to start interactive auth.
    if not (len(sys.argv) > 1 and sys.argv[1] == "login"):
        raise RuntimeError(
            "Microsoft authentication is required. Run this command in a "
            "visible terminal, complete the device login, then retry the MCP "
            f"tool: {sys.executable} {pathlib.Path(__file__).resolve()} login"
        )

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


def _patch(path: str, payload: dict) -> dict:
    r = httpx.patch(
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


def _delete_rule(rule_id: str) -> None:
    """
    DELETE an inbox rule (the automation itself, not any email).

    Scoped narrowly to the messageRules endpoint on purpose -- this helper
    only ever talks to /me/mailFolders/inbox/messageRules/{id}, so it can't
    be reused to delete a message, event, folder, task, contact, or Drive
    item by accident. This is the only DELETE call in this file.
    """
    r = httpx.delete(
        f"{GRAPH}/me/mailFolders/inbox/messageRules/{rule_id}",
        headers={"Authorization": f"Bearer {_token()}"},
        timeout=30,
    )
    r.raise_for_status()


def _recipients(addresses: Optional[list[str]]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in (addresses or [])]


def _shape_messages(messages: list[dict]) -> list[dict]:
    """Shared shaping for message-list results (id, subject, from, received, preview)."""
    return [
        {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
            "received": m.get("receivedDateTime"),
            "preview": (m.get("bodyPreview") or "")[:300],
        }
        for m in messages
    ]


_WELL_KNOWN_MAIL_FOLDERS = {
    "inbox",
    "drafts",
    "sentitems",
    "deleteditems",
    "archive",
    "outbox",
    "junkemail",
}


def _resolve_folder_path(path: str) -> dict:
    """
    Walk a '/'-separated mail folder path from the mailbox root, e.g.
    'Inbox/Bank/CIBC_Canada', and return {"id": ..., "displayName": ...}
    for the folder at the end of it.

    Segment names are matched case-insensitively. The first segment may
    also be a Graph well-known folder name (inbox, drafts, sentitems,
    deleteditems, archive, outbox, junkemail).

    Raises ValueError naming the exact segment that failed to resolve,
    instead of surfacing a generic 404.
    """
    segments = [s for s in path.split("/") if s]
    if not segments:
        raise ValueError("Folder path is empty.")

    first = segments[0]
    remaining = segments[1:]

    if first.lower() in _WELL_KNOWN_MAIL_FOLDERS:
        folder = _get(f"/me/mailFolders/{first.lower()}")
        current_id = folder.get("id")
        current_name = folder.get("displayName")
    else:
        top = _get("/me/mailFolders", params={"$top": 250})
        match = next(
            (
                f
                for f in top.get("value", [])
                if (f.get("displayName") or "").lower() == first.lower()
            ),
            None,
        )
        if not match:
            raise ValueError(
                f"Could not resolve folder path segment '{first}' "
                "(no top-level folder with that name)."
            )
        current_id = match["id"]
        current_name = match["displayName"]

    for seg in remaining:
        children = _get(
            f"/me/mailFolders/{current_id}/childFolders", params={"$top": 250}
        )
        match = next(
            (
                f
                for f in children.get("value", [])
                if (f.get("displayName") or "").lower() == seg.lower()
            ),
            None,
        )
        if not match:
            raise ValueError(
                f"Could not resolve folder path segment '{seg}' "
                f"(no child folder with that name under '{current_name}')."
            )
        current_id = match["id"]
        current_name = match["displayName"]

    return {"id": current_id, "displayName": current_name}


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
    return json.dumps(_shape_messages(data.get("value", [])), indent=2)


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
# Mail folders
# --------------------------------------------------------------------------


def _list_folder_tree(parent_id: str, depth: int, max_depth: int = 6) -> list[dict]:
    if depth > max_depth:
        return []
    path = (
        "/me/mailFolders"
        if parent_id == "root"
        else f"/me/mailFolders/{parent_id}/childFolders"
    )
    data = _get(path, params={"$top": 250, "$select": "id,displayName,childFolderCount"})
    tree = []
    for f in data.get("value", []):
        node = {
            "id": f.get("id"),
            "displayName": f.get("displayName"),
            "childFolderCount": f.get("childFolderCount"),
        }
        if f.get("childFolderCount") and depth < max_depth:
            node["children"] = _list_folder_tree(f["id"], depth + 1, max_depth)
        tree.append(node)
    return tree


@mcp.tool()
def list_mail_folders(parent: str = "root") -> str:
    """
    Return the mail folder tree starting from `parent`.

    parent: 'root' for the top level, or a folder id to start from a subfolder.
    Each node has id, displayName, and childFolderCount, with a nested
    "children" list when the folder has subfolders. Recursion stops at
    depth 6 to guard against pathological folder nesting.
    """
    tree = _list_folder_tree(parent, depth=0)
    return json.dumps(tree, indent=2)


@mcp.tool()
def resolve_folder_path(path: str) -> str:
    """
    Resolve a '/'-separated mail folder path, e.g. 'Inbox/Bank/CIBC_Canada',
    to its folder id and displayName. Segment names are matched
    case-insensitively.

    Read-only: this only looks up folders. It does not create, move, or
    change anything. On a bad path, the error names the exact segment that
    failed to resolve.
    """
    try:
        folder = _resolve_folder_path(path)
    except ValueError as e:
        return json.dumps({"error": str(e)}, indent=2)
    return json.dumps(folder, indent=2)


@mcp.tool()
def list_messages_in_folder(path: str, limit: int = 15) -> str:
    """
    List the most recent messages in the folder found at `path`, e.g.
    'Inbox/Bank/CIBC_Canada'. Uses the same message shape as list_recent_mail.

    limit: maximum messages to return (1-50).
    """
    limit = max(1, min(limit, 50))
    try:
        folder = _resolve_folder_path(path)
    except ValueError as e:
        return json.dumps({"error": str(e)}, indent=2)
    data = _get(
        f"/me/mailFolders/{folder['id']}/messages",
        params={
            "$top": limit,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview",
        },
    )
    return json.dumps(_shape_messages(data.get("value", [])), indent=2)


@mcp.tool()
def create_mail_folder(path: str) -> str:
    """
    Create a mail folder at `path`, e.g. 'Inbox/Bank/CIBC_Canada'. Only the
    final segment is created -- every segment before it must already exist,
    resolved the same way resolve_folder_path does.

    If a folder with that name already exists under the parent, this
    returns the existing folder's id instead of creating a duplicate.
    """
    segments = [s for s in path.split("/") if s]
    if not segments:
        return json.dumps({"error": "Folder path is empty."}, indent=2)

    parent_path = "/".join(segments[:-1])
    new_name = segments[-1]

    if parent_path:
        try:
            parent = _resolve_folder_path(parent_path)
        except ValueError as e:
            return json.dumps({"error": str(e)}, indent=2)
        list_path = f"/me/mailFolders/{parent['id']}/childFolders"
    else:
        list_path = "/me/mailFolders"

    existing = _get(list_path, params={"$top": 250})
    match = next(
        (
            f
            for f in existing.get("value", [])
            if (f.get("displayName") or "").lower() == new_name.lower()
        ),
        None,
    )
    if match:
        return json.dumps(
            {
                "status": "already exists",
                "id": match["id"],
                "displayName": match["displayName"],
            },
            indent=2,
        )

    created = _post(list_path, {"displayName": new_name})
    return json.dumps(
        {
            "status": "created",
            "id": created.get("id"),
            "displayName": created.get("displayName"),
        },
        indent=2,
    )


# --------------------------------------------------------------------------
# Message organization
# --------------------------------------------------------------------------


@mcp.tool()
def move_message(message_id: str, destination_path: str) -> str:
    """
    Move a message to another folder, e.g. 'Inbox/Bank/CIBC_Canada'.

    This relocates the message within the mailbox. It never deletes
    anything, and it never moves a message to Deleted Items on this
    server's behalf -- only wherever `destination_path` resolves to.
    """
    try:
        folder = _resolve_folder_path(destination_path)
    except ValueError as e:
        return json.dumps({"error": str(e)}, indent=2)
    moved = _post(f"/me/messages/{message_id}/move", {"destinationId": folder["id"]})
    return json.dumps(
        {"status": "moved", "id": moved.get("id"), "destination": folder["displayName"]},
        indent=2,
    )


@mcp.tool()
def flag_message(message_id: str, flagged: bool = True) -> str:
    """Set (flagged=True) or clear (flagged=False) the follow-up flag on a message."""
    status = "flagged" if flagged else "notFlagged"
    _patch(f"/me/messages/{message_id}", {"flag": {"flagStatus": status}})
    return json.dumps({"status": "updated", "id": message_id, "flagStatus": status}, indent=2)


@mcp.tool()
def categorize_message(message_id: str, categories: list[str]) -> str:
    """Set the category labels on a message, replacing whatever categories it had."""
    _patch(f"/me/messages/{message_id}", {"categories": categories})
    return json.dumps(
        {"status": "updated", "id": message_id, "categories": categories}, indent=2
    )


@mcp.tool()
def mark_message_read(message_id: str, read: bool = True) -> str:
    """Mark a message as read (read=True) or unread (read=False)."""
    _patch(f"/me/messages/{message_id}", {"isRead": read})
    return json.dumps({"status": "updated", "id": message_id, "isRead": read}, indent=2)


# --------------------------------------------------------------------------
# Mail rules
# --------------------------------------------------------------------------

# Every action key an inbox rule created by this server is allowed to carry.
# There is no "delete" entry here, and create_mail_rule below builds its
# actions dict from a fixed, hard-coded shape -- there is no parameter path
# that lets a caller inject anything outside this set.
_ALLOWED_RULE_ACTION_KEYS = {
    "moveToFolder",
    "copyToFolder",
    "markAsRead",
    "assignCategories",
    "stopProcessingRules",
}


@mcp.tool()
def list_mail_rules() -> str:
    """List the inbox rules configured on this mailbox."""
    data = _get("/me/mailFolders/inbox/messageRules")
    out = [
        {
            "id": r.get("id"),
            "displayName": r.get("displayName"),
            "isEnabled": r.get("isEnabled"),
            "sequence": r.get("sequence"),
            "conditions": r.get("conditions"),
            "actions": r.get("actions"),
        }
        for r in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def create_mail_rule(
    name: str,
    sender_contains: list[str],
    destination_path: str,
    stop_processing: bool = True,
) -> str:
    """
    Create an inbox rule that moves matching mail into a folder.

    sender_contains: strings matched against the sender's address or display
        name (Graph's senderContains condition).
    destination_path: folder path the matching mail is moved to, e.g.
        'Inbox/Bank/CIBC_Canada'.
    stop_processing: if true, later rules are skipped once this one matches.

    This tool can only move mail. It builds a fixed action set --
    moveToFolder and stopProcessingRules only -- and has no parameter that
    can add a delete action. Even so, the action dict is checked against an
    allow-list before the API call as defense in depth.
    """
    try:
        folder = _resolve_folder_path(destination_path)
    except ValueError as e:
        return json.dumps({"error": str(e)}, indent=2)

    actions = {"moveToFolder": folder["id"], "stopProcessingRules": stop_processing}

    if not set(actions).issubset(_ALLOWED_RULE_ACTION_KEYS):
        raise ValueError(
            "Refusing to create a rule with an action outside the allowed set: "
            f"{sorted(set(actions) - _ALLOWED_RULE_ACTION_KEYS)}"
        )

    payload = {
        "displayName": name,
        "sequence": 1,
        "isEnabled": True,
        "conditions": {"senderContains": sender_contains},
        "actions": actions,
    }
    created = _post("/me/mailFolders/inbox/messageRules", payload)
    return json.dumps(
        {
            "status": "rule created",
            "id": created.get("id"),
            "displayName": created.get("displayName"),
        },
        indent=2,
    )


@mcp.tool()
def delete_mail_rule(rule_id: str) -> str:
    """
    Delete an inbox rule.

    This deletes the RULE -- the automation that watches for and moves
    matching mail -- not any email message. No message is touched, moved,
    or deleted by this tool; only the rule definition is removed.
    """
    _delete_rule(rule_id)
    return json.dumps({"status": "rule deleted", "id": rule_id}, indent=2)


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


# --------------------------------------------------------------------------
# To Do
# --------------------------------------------------------------------------


@mcp.tool()
def list_task_lists() -> str:
    """List the Microsoft To Do task lists in this account."""
    data = _get("/me/todo/lists")
    out = [
        {"id": lst.get("id"), "displayName": lst.get("displayName")}
        for lst in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def list_tasks(list_id: str, limit: int = 25) -> str:
    """
    List tasks in a To Do list.

    limit: maximum tasks to return (1-100).
    """
    limit = max(1, min(limit, 100))
    data = _get(
        f"/me/todo/lists/{list_id}/tasks",
        params={"$top": limit, "$select": "id,title,status,dueDateTime"},
    )
    out = [
        {
            "id": t.get("id"),
            "title": t.get("title"),
            "status": t.get("status"),
            "due": (t.get("dueDateTime") or {}).get("dateTime"),
        }
        for t in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def create_task(
    list_id: str, title: str, due: Optional[str] = None, body: Optional[str] = None
) -> str:
    """
    Create a task in a To Do list.

    due: ISO 8601 date/time, e.g. 2026-08-15T00:00:00 (treated as UTC).
    """
    payload: dict[str, Any] = {"title": title}
    if due:
        payload["dueDateTime"] = {"dateTime": due, "timeZone": "UTC"}
    if body:
        payload["body"] = {"contentType": "text", "content": body}
    created = _post(f"/me/todo/lists/{list_id}/tasks", payload)
    return json.dumps({"status": "task created", "id": created.get("id")}, indent=2)


@mcp.tool()
def complete_task(list_id: str, task_id: str) -> str:
    """Mark a task as completed. This does not delete the task."""
    _patch(f"/me/todo/lists/{list_id}/tasks/{task_id}", {"status": "completed"})
    return json.dumps({"status": "task completed", "id": task_id}, indent=2)


# --------------------------------------------------------------------------
# Contacts
# --------------------------------------------------------------------------


@mcp.tool()
def search_contacts(query: str, limit: int = 15) -> str:
    """Search contacts by display name or email address."""
    limit = max(1, min(limit, 50))
    data = _get(
        "/me/contacts",
        params={
            "$search": f'"{query}"',
            "$top": limit,
            "$select": "id,displayName,emailAddresses,mobilePhone",
        },
    )
    out = [
        {
            "id": c.get("id"),
            "displayName": c.get("displayName"),
            "emails": [e.get("address") for e in c.get("emailAddresses", [])],
            "mobilePhone": c.get("mobilePhone"),
        }
        for c in data.get("value", [])
    ]
    return json.dumps(out, indent=2)


@mcp.tool()
def create_contact(display_name: str, email: str, phone: Optional[str] = None) -> str:
    """Create a contact with a display name, one email address, and an optional phone number."""
    payload: dict[str, Any] = {
        "displayName": display_name,
        "emailAddresses": [{"address": email, "name": display_name}],
    }
    if phone:
        payload["mobilePhone"] = phone
    created = _post("/me/contacts", payload)
    return json.dumps({"status": "contact created", "id": created.get("id")}, indent=2)


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
