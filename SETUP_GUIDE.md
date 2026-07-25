# Connecting Claude to a personal Outlook / Hotmail account

A Microsoft Graph MCP server for personal Microsoft accounts: read and search mail, create
drafts, manage calendar, and browse OneDrive.

---

## Background

Anthropic's Microsoft 365 connector registers its OAuth application for **work and school
tenants only**. Personal Microsoft accounts (`@hotmail.com`, `@outlook.com`, `@live.com`) are
rejected at authentication.

Microsoft Graph itself has no such restriction. The limit is purely in how that application was
registered. This guide registers your own application, with an account-type setting that accepts
personal accounts, and points a small local MCP server at it.

## What this can and cannot do

| Capability | Granted? |
|---|---|
| Read and search mail | Yes |
| Create drafts and draft replies | Yes |
| **Send mail** | **No, scope deliberately not requested** |
| Read and create calendar events | Yes |
| Search and read OneDrive files | Yes (read only) |

> **On sending:** the `Mail.Send` scope is not requested, so the server is *incapable* of sending
> mail even if asked. Claude composes a draft; you review and send it yourself from Outlook.

---

## Architecture

Three pieces do the work, and each has exactly one job.

**The Entra app registration** is the server's identity with Microsoft. It's a public client, so
there's no secret to store, since a script running on your own machine has nowhere safe to keep
one. It's registered to accept personal Microsoft accounts, and it's granted a fixed list of
delegated permissions. Those permissions are a ceiling, not a suggestion: the server can never do
more than what's granted here, no matter what a tool function tries to call. `Mail.Send` is
deliberately absent, so sending mail is impossible at the Graph level, not merely unimplemented in
the code.

**MSAL and the token cache** handle authentication. First sign-in uses a device-code flow: the
server prints a code, you enter it at microsoft.com/devicelogin in a browser, and Microsoft issues
a token pair, an access token and a refresh token. The refresh token is written to
`~/.msgraph-mcp/token_cache.json`. Every later call to Graph asks MSAL for a token first; MSAL
uses the cached refresh token to get a fresh access token silently, with no browser involved,
until the refresh token itself expires or is revoked.

**The MCP server** (`graph_mcp_server.py`) is a FastMCP application speaking stdio JSON-RPC,
which is how Claude Desktop expects to talk to it. Claude Desktop launches the script as a
subprocess on startup and keeps it running for the session. Each tool, `search_mail`,
`create_draft`, `list_events`, and the rest, is a plain Python function decorated with
`@mcp.tool()`. FastMCP reads the function's type hints and docstring to build the tool schema
Claude sees. Calling a tool means Claude sends a JSON-RPC request over stdin, the function runs,
and its return value goes back over stdout.

A single request, end to end: Claude decides to call `search_mail`, FastMCP dispatches to the
Python function, the function asks MSAL for a token (cached and silent, in the normal case), then
makes an authenticated GET against `https://graph.microsoft.com/v1.0/me/messages`, and the JSON
response is reshaped into the compact structure the tool returns. Nothing here talks to any
Anthropic service directly. Claude Desktop only ever sees stdio, and Microsoft Graph only ever
sees an OAuth bearer token.

```text
Claude Desktop  <-- stdio JSON-RPC -->  graph_mcp_server.py  <-- MSAL token -->  token_cache.json
                                                |
                                                v
                                     Microsoft Graph (v1.0 REST API)
```

---

# Part A: Register the application in Entra

Free. No Azure subscription required. Roughly eight minutes.

## A1. Find App registrations

Go to **entra.microsoft.com**. In the left navigation, expand **Entra ID**, then choose
**App registrations**.



Shortcut: paste this straight into the address bar.

```
https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
```

## A2. Create the registration

Click **New registration** and fill it in exactly as shown:

- **Name:** `claude-graph-mcp`
- **Supported account types:** Accounts in any organizational directory and **personal Microsoft accounts**
- **Redirect URI:** leave blank

> **This is the setting that matters.** The account-type choice is the entire reason the built-in
> connector fails. Any other option here and your personal account will be rejected exactly as
> before.

![The registration form, correctly filled](images/02-register-app.png)

## A3. Copy the client ID

After registering you land on the overview page. Copy the **Application (client) ID**. You'll need
it twice, once for first sign-in and once in the Claude Desktop config.

Confirm **Supported account types** reads *All Microsoft account users*. If it says anything else,
the registration was made with the wrong option and should be redone.

![Overview page showing client ID and supported account types](images/03-overview-client-id.png)

> You may see a banner warning that end users cannot consent to newly registered multitenant apps
> without a verified publisher. That restriction applies to users consenting *inside Entra tenants*;
> consenting as a personal Microsoft account is normally unaffected. If you do hit a consent error
> later, note the exact wording before changing anything.

## A4. Allow public client flows

Left menu → **Authentication** → **Settings** tab → set **Allow public client flows** to
**Enabled** → **Save**.

> Device-code authentication will not work without this. It's the step most commonly missed, and
> the usual cause if sign-in fails in Part C. **The toggle isn't committed until you press Save.**

![Allow public client flows set to Enabled](images/04-public-client-flows.png)

## A5. Add the Graph permissions

Left menu → **API permissions** → **Add a permission** → **Microsoft Graph** →
**Delegated permissions**.

![The Request API permissions pane](images/05-add-permission.png)

Add these five. `User.Read` is usually present by default:

```
User.Read              (usually already there)
Mail.ReadWrite         read mail and create drafts
Calendars.ReadWrite    read and create events
Files.Read.All         read OneDrive
offline_access         keep the session alive
```

> **Do not add `Mail.Send`.** Its absence is what guarantees the server cannot send mail on your
> behalf.
>
> `offline_access` sits under an OpenId grouping rather than with the resource scopes, so search
> for it by name.

## A6. Confirm the final list

Your configured permissions should match this exactly: five delegated Graph permissions, no
`Mail.Send`.

![Final permissions list](images/06-final-permissions.png)

> Ignore the **Grant admin consent** button. It consents on behalf of the Entra tenant you
> registered the app in, which does nothing for a personal account. The consent that matters
> happens in the browser during Part C.

---

# Part B: Install locally

Save `graph_mcp_server.py` to your home folder, then:

```bash
mkdir -p ~/msgraph-mcp
mv ~/Downloads/graph_mcp_server.py ~/msgraph-mcp/
cd ~/msgraph-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" msal httpx
```

Your project folder may end up looking like this, the server and its virtual environment in
`msgraph-mcp` (ignore the images folder, it's just for this guide):

![Project folder layout: images, msgraph-mcp with its venv, and this guide](images/07-folder-structure.png)

---

# Part C: Sign in once

Claude Desktop can't show you an interactive login prompt, so do the first sign-in yourself in a
terminal. The `login` subcommand forces the device-code flow and exits once the token is cached:

```bash
export MSGRAPH_CLIENT_ID="<your client ID from A3>"
python graph_mcp_server.py login
```

It prints a code and a URL. Go to **microsoft.com/devicelogin**, enter the code, and sign in
**with your personal Microsoft account**, not the account you used to register the app. Approve
the permissions.

When it prints `Authenticated. Token cached at …` and exits on its own, you're done. No Ctrl-C
needed. The refresh token is now cached at `~/.msgraph-mcp/token_cache.json` and renews silently
from here on.

> **Why a subcommand?** Running the server with no arguments starts the stdio JSON-RPC loop and
> blocks. It never triggers auth, because sign-in only happens lazily when a *tool* is called.
> `login` calls the token path directly so the device prompt actually appears.

> **The two accounts do different jobs.** The account you registered the app with owns the
> registration. Your personal account is the mailbox being accessed. Signing in with the wrong one
> here is the second most common mistake after skipping A4.

---

# Part D: Connect Claude Desktop

Open the config file:

```bash
mkdir -p ~/Library/Application\ Support/Claude
open -e ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

Paste this, replacing the username in the paths and the client ID with your own:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "/Users/YOUR_USERNAME/msgraph-mcp/.venv/bin/python",
      "args": ["/Users/YOUR_USERNAME/msgraph-mcp/graph_mcp_server.py"],
      "env": {
        "MSGRAPH_CLIENT_ID": "<your client ID from A3>"
      }
    }
  }
}
```

> Use absolute paths. Claude Desktop doesn't expand `~` or read your shell PATH. Then quit Claude
> Desktop completely with **Cmd-Q** (closing the window isn't enough) and reopen it.

---

# Part E: Test

In a new chat in Claude Desktop:

```
Use whoami to check which Microsoft account is connected
```

If it returns your own address, everything is working. Then try a real search against your own
mailbox to confirm the whole path works end to end.

## Tools available

| Tool | What it does |
|---|---|
| `whoami` | Confirm which account is connected |
| `search_mail` | Full-text search across the mailbox |
| `list_recent_mail` | Recent messages in a folder (inbox, sentitems, drafts) |
| `read_message` | Full body of one message |
| `create_draft` | Compose a draft, never sends |
| `reply_draft` | Draft a reply to an existing thread |
| `list_events` | Calendar events in a date range |
| `create_event` | Create a calendar event |
| `search_onedrive` | Search OneDrive files |
| `list_onedrive_folder` | Browse a OneDrive folder |
| `read_onedrive_file` | Read a text-based file's contents |

## Troubleshooting

**Device flow fails to start.** Almost always A4: Allow public client flows still disabled, or
enabled but not saved.

**Signed in but `whoami` returns the wrong account.** You authenticated with the account you
registered the app with instead of your personal one. Delete `~/.msgraph-mcp/token_cache.json`
and repeat Part C, signing in with the personal account this time.

**Server doesn't appear in Claude Desktop.** Check the paths in the config are absolute and correct,
and that you fully quit with Cmd-Q rather than closing the window.

**You later change the scopes.** Delete the token cache and re-authenticate, or the cached token
won't carry the new permissions.

**Basic auth / app passwords.** Any guide suggesting an app password with `imap-mail.outlook.com` is
out of date. Microsoft disabled basic authentication for Outlook.com IMAP and POP in September
2024. OAuth via Graph, as used here, is the supported path.
