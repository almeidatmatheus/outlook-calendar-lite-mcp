# outlook-calendar-lite-mcp

A lightweight [MCP](https://modelcontextprotocol.io) server for Microsoft Outlook Calendar (Microsoft Graph).
Works with work/school accounts (Microsoft 365) and personal accounts (Outlook.com, Hotmail).

- **Zero dependencies.** Node 22.18+ runs the TypeScript directly: no `npm install`, no build.
- **You run it, you own it.** It runs on your machine with your own Azure app. No third-party server sees your calendar.
- **Token-light.** Tabular output.
- **Permanent ids.** Every event comes with its real Microsoft Graph id, requested as an *immutable id*, so you can store it in a database and use it later.
- **Your time zone, always.** Times are read and written in the mailbox's time zone; every listing shows today's date and the current time, so the AI never guesses "tomorrow".
- **Safe by default.** Anything that notifies other people (invitations, updates, cancellations, replies) shows a preview and waits for the user's approval. Events only on your own calendar don't ask.
- **Login on connect.** The first time the client connects, your browser opens on the Microsoft account picker.

> Português: [README.pt-BR.md](README.pt-BR.md). Tool descriptions are in Portuguese; any LLM handles them fine.

## Prerequisites

- **Node.js 22.18 or newer** (`node --version`). Nothing else to install: no `npm install`, no build.
- **git**, to clone the repository (or download the ZIP from GitHub).
- **A Microsoft account** with a calendar: work/school (Microsoft 365) or personal (Outlook.com, Hotmail).
- **Permission to register an app in Microsoft Entra** (free):
  - work account: allowed by default in most companies; if blocked, ask your IT admin to register it or to grant consent;
  - personal account: sign in to <https://portal.azure.com> once with a free Azure account, which creates your own directory.
- **An MCP client**: Claude Desktop, Claude Code, Cursor, VS Code or any other client that supports stdio servers.

## Quick start

### 1. Get the code

```bash
git clone https://github.com/almeidatmatheus/outlook-calendar-lite-mcp.git
cd outlook-calendar-lite-mcp
node --version   # needs 22.18 or newer
```

### 2. Register your own app in Microsoft Entra (once, ~5 min)

> Already using [outlook-lite-mcp](https://github.com/almeidatmatheus/outlook-lite-mcp)? Reuse the same app: just do step 4 (add `Calendars.ReadWrite`) and use the same client ID.

1. Go to <https://entra.microsoft.com> → **Applications** → **App registrations** → **New registration**.
   - **Name:** `outlook-calendar-lite-mcp`
   - **Supported account types:**
     - personal + any company: **Accounts in any organizational directory and personal Microsoft accounts**
     - only your company: **Accounts in this organizational directory only**
   - **Register**.
2. On **Overview**, copy the **Application (client) ID** (and the **Directory (tenant) ID** if you chose "this directory only").
3. **Authentication** → **Add a platform** → **Mobile and desktop applications** → custom redirect URI:
   ```
   http://localhost/cb
   ```
   Then, still in **Authentication**, set **Allow public client flows** to **Yes** → **Save**.
   No client secret is needed for local use.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   `Calendars.ReadWrite`, `MailboxSettings.Read`, `User.Read`, `offline_access`.
   (Never *Application permissions*.)
5. Work accounts: if your company requires it, an admin clicks **Grant admin consent for <company>**.

### 3. Add it to your MCP client

**Claude Desktop** (`claude_desktop_config.json`, see [examples/](examples/claude_desktop_config.json)):

```json
{
  "mcpServers": {
    "calendar": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-calendar-lite-mcp/src/main.ts"],
      "env": { "OUTLOOK_MCP_CLIENT_ID": "<your client id>" }
    }
  }
}
```

**Claude Code:**

```bash
claude mcp add calendar -e OUTLOOK_MCP_CLIENT_ID=<your client id> -- node /absolute/path/to/outlook-calendar-lite-mcp/src/main.ts
```

**Any other MCP client** (Cursor, VS Code, local models…): command `node /absolute/path/to/outlook-calendar-lite-mcp/src/main.ts`, stdio transport, same env var.

Restart the client. The browser opens on the Microsoft account picker; sign in and you're done.
The token is stored in `~/.outlook-calendar-lite-mcp/default.json` (mode 600) and refreshes itself.
To use another account, ask to "switch account" (tool `cx`, action `trocar`).

On macOS, keep the folder outside `Downloads`/`Desktop`/`Documents` (e.g. `~/outlook-calendar-lite-mcp`), or the client may not be allowed to read it.

### Try it

> "what do I have today?" · "find a free hour with ana@acme.com this week" · "book it and invite her, with Teams" · "move my dentist to 11am" · "accept the ACME meeting"

## Tools

| tool | does |
|---|---|
| `ls` | agenda for a period (today, tomorrow, this week, next week, 7/30 days, or dates) |
| `rd` | read an event: time, place, organizer, attendees with each response, Teams link, recurrence, description |
| `find` | search events by title, place, organizer or attendee |
| `free` | free slots for a meeting, crossing your calendar (and attendees', on work accounts) within working hours |
| `cr` | create an event: all-day, recurring, online (Teams), reminder, attendees. **Asks for confirmation when it invites people**; warns about conflicts |
| `ed` | change an event you organize (time, title, place, attendees…). **Asks for confirmation when attendees are notified** |
| `rm` | cancel (attendees get the cancellation, with an optional message) or remove from your calendar. **Asks for confirmation** |
| `rsp` | reply to an invitation: accept, decline or tentative, with a comment. **Asks for confirmation when the organizer is notified** |
| `cx` | account, time zone, working hours and calendars; `trocar` switches account |
| `ntf` | change notifications for your app: `info` shows the subscription, `ligar` creates or renews it, `desligar` removes it (see below) |

Event ids: the `id` column of `ls`/`find` (and the id returned by `cr`) is the event's permanent Graph id (immutable id). Pass it as `alvo` to the other tools, now or in a later session.
Recurring events: `ed`, `rm` and `rsp` act on one occurrence; pass `serie: true` for the whole series.
Dates are written as `YYYY-MM-DDTHH:mm` in your time zone; `YYYY-MM-DD` alone means all day.

**Confirmation.** If the client supports MCP *elicitation*, the server asks the user directly.
Otherwise the first call returns a preview and a code `cf`; only the identical call with that `cf` executes.
Clients should show the preview to the user and send `cf` only after explicit approval.

**Personal accounts.** Checking other people's availability (`free` with attendees) and Teams meetings are Microsoft 365 features; on Outlook.com, `free` uses only your own calendar and says so.

## Options

| env var | default |
|---|---|
| `OUTLOOK_MCP_CLIENT_ID` | required |
| `OUTLOOK_MCP_TENANT` | `common` (set your tenant id to lock to one company) |
| `OUTLOOK_MCP_PROFILE` | `default` (one profile per account) |
| `OUTLOOK_MCP_LOGIN` | browser; `device` = device code (SSH, no browser) |
| `OUTLOOK_MCP_TZ` | the mailbox's time zone; set an IANA name (e.g. `America/Sao_Paulo`) to override date math |
| `OUTLOOK_MCP_HOME` | `~/.outlook-calendar-lite-mcp` |
| `OUTLOOK_MCP_NOTIFY_URL` | change notifications: public HTTPS URL of your app that receives them (unset = `ntf` is off) |
| `OUTLOOK_MCP_NOTIFY_SECRET` | change notifications: secret sent as `clientState` (max. 128 characters) |

Command line: `node src/main.ts login | whoami | logout`.
Behind a corporate proxy: Node 22.21+ with `NODE_USE_ENV_PROXY=1` and `HTTPS_PROXY`.

## Change notifications (optional)

To have another app of yours told when events change, set `OUTLOOK_MCP_NOTIFY_URL` and `OUTLOOK_MCP_NOTIFY_SECRET`.
The MCP only creates and keeps the Microsoft Graph subscription; your app receives the notifications at that URL.

- `ntf acao=ligar` (optionally `calendario`, `dias` 1–3, default 2) creates the subscription or renews it. It's idempotent, so your app can call it on a schedule, before it expires. Returns `ligado:<id>|expira:<ISO UTC>`.
- `ntf` (`acao=info`) returns `assinatura:<id>|expira:<ISO UTC>|calendario:<name>` or `sem assinatura`. `ntf acao=desligar` returns `desligado`.
- Subscriptions are created with the immutable id, so notifications carry the same permanent event id that `ls` and `cr` return.
- On creation, Microsoft Graph calls the URL with a `validationToken` and expects it back as plain text within 10 seconds; if that fails, `ntf` says the URL didn't answer the validation.
- Your app should check that `clientState` equals the secret before trusting a notification.
- The URL and secret come only from the environment, never from tool arguments, so a model can't point notifications elsewhere. The secret never appears in tool output. `ntf` doesn't ask for confirmation: nobody else is notified.
- No new permission: `Calendars.ReadWrite` is enough.
- The URL and secret belong to the server, not to an account: in remote mode with several accounts, each account's subscription points to the same URL. Your app can tell them apart by the `resource` in each notification.

## Remote mode (optional, self-hosted)

To use it from claude.ai as a custom connector, run your own HTTPS instance:

```bash
export OUTLOOK_MCP_CLIENT_ID=<client id>
export OUTLOOK_MCP_URL=https://your-domain.example      # public HTTPS URL
export OUTLOOK_MCP_ALLOW=you@company.com,company.com    # who may connect (emails or domains)
node src/main.ts http                                   # listens on PORT (8787)
```

In Entra add the platform **Web** with redirect `https://your-domain.example/ms/cb` (and a client secret in
`OUTLOOK_MCP_CLIENT_SECRET` if Entra asks for one). In claude.ai add a custom connector with `https://your-domain.example/mcp`.
For a quick test, `cloudflared tunnel --url http://localhost:8787` gives you a temporary HTTPS URL.

**Always set `OUTLOOK_MCP_ALLOW`.** Without it anyone with the URL can connect their own account to your server.

## Security

- Delegated permissions only: the server can do what the signed-in user can do, nothing more.
- Tokens live on your machine (`~/.outlook-calendar-lite-mcp`, mode 600). Remote mode keeps them in `http-store.json` on the server.
- **Invitations are untrusted.** An event description can contain text trying to instruct the AI.
  The mandatory confirmation before anything that notifies other people is the main defense; don't remove it.
- See [SECURITY.md](SECURITY.md) to report a vulnerability.

## Project layout

```
src/
  main.ts     entry: stdio (default) | http | login | logout | whoami
  mcp.ts      MCP protocol: JSON-RPC, schema validation, confirmation (elicitation or cf code)
  http.ts     remote mode: Streamable HTTP + OAuth server that delegates login to Microsoft
  auth.ts     Microsoft login: browser (PKCE), device code, refresh, token file
  ctx.ts      per-user state (token, caches)
  g.ts        minimal Graph client: get/post/patch/delete, retry on 429/401, immutable ids
  ids.ts      validates the event ids received by the tools
  fmt.ts      tabular output, text cleanup
  lib.ts      time zone and date math, periods, events table, recurrence, shared helpers
  tools/      one file per tool + index.ts (ev.ts: event fields shared by create and update)
test/
  mock.ts     fake Microsoft (login + Graph calendar)
  run.ts      real MCP clients against the mock, local and remote
```

## Development

```bash
node test/run.ts   # 88 checks, no network or Microsoft account needed
npx tsc -p .       # type check (optional)
```

## License

[MIT](LICENSE)
