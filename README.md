# mailhub-mcp

A local MCP server that puts many Gmail and Outlook mailboxes behind one connector for Claude.

- Connect as many mailboxes as you like (Gmail, Google Workspace, Microsoft 365, Outlook.com).
- Sign in to each mailbox **once**. Refresh tokens are kept on your computer and renewed automatically.
- Search old and new mail in one mailbox or across all of them in a single question.
- **Read-only by default**: it requests `gmail.readonly` and `Mail.Read` only.
- **Drafts are opt-in** per mailbox (`--drafts`). Claude can then save a draft for you to review and send.
- It has **no way to send, delete or change mail**. Those tools do not exist in the code.

```
Claude desktop app  --stdio-->  mailhub-mcp  --+--> Gmail API            (personal, work-nl, ...)
                                               +--> Microsoft Graph API  (work-uk, outlook-personal, ...)
```

## Tools Claude gets

| Tool | What it does |
|---|---|
| `list_accounts` | Lists the connected mailboxes (alias, provider, address). |
| `search_emails` | Searches one mailbox, several, or `all`. Filters: `query`, `from`, `to`, `subject`, `after`, `before`, `unread_only`, `has_attachment`. Paged. |
| `read_email` | Full message: headers, plain-text body, attachment list. |
| `read_thread` | Every message of a conversation, oldest first. |
| `download_attachment` | Saves an attachment to the local download folder and returns its path. |
| `create_draft` | Saves a draft (new message, reply or reply-all) in a mailbox added with `--drafts`. Never sends. |

Example prompts once it is connected:

- "Search all my mailboxes for the Acme invoice from October."
- "What unread mail came into work-uk this week?"
- "Read the latest thread from the accountant in work-in and summarise it."
- "Draft a reply to that thread saying the documents will follow on Friday."

## Requirements

- Node.js 20 or newer
- The Claude desktop app (local MCP servers run through it)

## 1. Install

```bash
git clone https://github.com/ashwajce/mailhub-mcp.git
cd mailhub-mcp
npm install          # also builds dist/
npm link             # optional: makes the `mailhub-mcp` command available everywhere
```

Without `npm link`, replace `mailhub-mcp` below with `node /full/path/to/mailhub-mcp/dist/cli.js`.

## 2. One-time setup: Google (for all Gmail accounts)

One Google Cloud OAuth client covers every Gmail and Workspace account you connect.

1. Open <https://console.cloud.google.com>, create a project (for example `mailhub`).
2. **APIs & Services > Library**: enable the **Gmail API**.
3. **APIs & Services > OAuth consent screen**: user type **External**, fill in the app name and your email.
   Add the scope `https://www.googleapis.com/auth/gmail.readonly`, and
   `https://www.googleapis.com/auth/gmail.compose` if you plan to use drafts.
4. Set the publishing status to **In production**. This matters: while the app is in "Testing",
   Google expires refresh tokens after 7 days and you would have to sign in again every week.
   For personal use you do not need to submit the app for verification; you will see a
   "Google hasn't verified this app" screen once per account (Advanced > Go to app).
5. **APIs & Services > Credentials > Create credentials > OAuth client ID**, application type **Desktop app**.
6. Save the client ID and secret:

```bash
mailhub-mcp setup google --client-id <CLIENT_ID> --client-secret <CLIENT_SECRET>
```

Google Workspace note: if a Workspace admin restricts third-party apps, allow this client ID under
Admin console > Security > API controls, or create the project inside that Workspace and mark the app **Internal**.

## 3. One-time setup: Microsoft (for all Outlook accounts)

One Entra app registration covers work tenants and personal Outlook.com / Hotmail accounts.

1. Open <https://entra.microsoft.com> > **App registrations > New registration**.
2. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
3. Redirect URI: platform **Public client/native (mobile & desktop)**, value `http://localhost`.
4. **API permissions > Add > Microsoft Graph > Delegated**: `Mail.Read`, `User.Read`, `offline_access`.
   Add `Mail.Read.Shared` too if you want to read shared or delegated mailboxes, and `Mail.ReadWrite`
   (or `Mail.ReadWrite.Shared`) if you plan to use drafts.
5. Copy the **Application (client) ID** from the Overview page (no client secret is needed):

```bash
mailhub-mcp setup microsoft --client-id <APPLICATION_ID>
```

If another tenant blocks user consent, an admin of that tenant has to approve the app once.

## 4. Connect your mailboxes (once each)

Give each mailbox a short alias. A browser window opens for sign-in.

```bash
mailhub-mcp add gmail personal
mailhub-mcp add gmail workspace --login-hint you@yourdomain.com
mailhub-mcp add outlook work-nl
mailhub-mcp add outlook work-uk
mailhub-mcp add outlook hotmail

# A shared or delegated Microsoft 365 mailbox, signed in as yourself:
mailhub-mcp add outlook billing --shared-mailbox billing@yourdomain.com

# Allow Claude to save drafts in a mailbox (it still cannot send):
mailhub-mcp add outlook work-in --drafts

mailhub-mcp list        # show what is connected, and which mailboxes allow drafts
mailhub-mcp check       # verify every stored sign-in still works
mailhub-mcp remove hotmail
```

To switch drafts on for a mailbox you already added, sign in again:
`mailhub-mcp add gmail personal --drafts --force`.

Other options: `--force` re-authorises an existing alias, `--tenant <domain>` signs in against one
specific Microsoft tenant, `--no-browser` prints the sign-in URL instead of opening it.

## 5. Add it to Claude

In the Claude desktop app: **Settings > Developer > Edit Config**, then add the server to
`claude_desktop_config.json` (macOS: `~/Library/Application Support/Claude/`, Windows: `%APPDATA%\Claude\`):

```json
{
  "mcpServers": {
    "mailhub": {
      "command": "node",
      "args": ["/full/path/to/mailhub-mcp/dist/cli.js"]
    }
  }
}
```

Restart the app. If `node` is not found (common with nvm), use the full path from `which node` as the command.

Tip: switch off Claude's built-in Gmail and Microsoft 365 connectors, or tell Claude to use mailhub
for email, so it does not pick the single-account connector by mistake.

## Where things are stored

Everything lives in `~/.mailhub-mcp/` (override with `MAILHUB_HOME`), outside this repository:

| File | Contents |
|---|---|
| `config.json` | Google and Microsoft app IDs. |
| `accounts.json` | One refresh token per mailbox. Owner-only permissions (0600). Treat it like a password file. |
| `downloads/<alias>/` | Attachments saved by `download_attachment`. |

Run `mailhub-mcp paths` to print the exact locations. Environment variables override `config.json`:
`MAILHUB_GOOGLE_CLIENT_ID`, `MAILHUB_GOOGLE_CLIENT_SECRET`, `MAILHUB_MS_CLIENT_ID`, `MAILHUB_MS_TENANT`,
`MAILHUB_DOWNLOAD_DIR`.

To revoke access entirely, remove the alias and also remove the app under
<https://myaccount.google.com/permissions> or <https://myaccount.microsoft.com> (or Entra "Enterprise applications").

## How long do sign-ins last?

- **Google**: indefinitely with the app "In production", unless you revoke access, change the password
  on some account types, or leave a token unused for 6 months.
- **Microsoft**: the refresh token renews itself each time it is used. It lapses after about 90 days
  without any use, or when a tenant's conditional access policy requires a fresh sign-in.

When a sign-in does lapse, the tool error tells you the exact command, for example
`mailhub-mcp add gmail personal --force`.

## Drafts

- Off unless the mailbox was added with `--drafts`. `list_accounts` and `mailhub-mcp list` show which ones allow it.
- A draft lands in the mailbox's normal Drafts folder. You open it in Gmail or Outlook, check it and send it yourself.
- Replies stay in the original conversation. Outlook reply drafts include the quoted original; Gmail reply
  drafts contain only the new text.
- Bodies are plain text. Attachments cannot be added to drafts.
- Scopes: Outlook uses `Mail.ReadWrite`, which cannot send. Gmail has no drafts-only permission, so it uses
  `gmail.compose`, which technically also covers sending. mailhub contains no code that sends.

## Search notes

- The structured filters (`from`, `to`, `subject`, `after`, `before`, ...) work on every provider.
- Plain keywords in `query` work everywhere. Provider operators only work on that provider:
  Gmail search operators (`label:`, `in:sent`, `filename:pdf`) or Outlook KQL (`participants:`, `body:`).
- With `account: "all"` results are grouped per mailbox. To get the next page, call again with that
  one alias and its `next_page_token`.
- Outlook thread view returns up to 50 messages of a conversation.

## Limitations

- No sending, deleting, moving or labelling, by design.
- Local only: it works while this computer is on and the Claude desktop app is open. To use it from
  a phone or from scheduled tasks, host the same code behind a remote MCP endpoint.
- Tokens are stored in a permission-restricted file, not the OS keychain.

## Development

```bash
npm run build
npm test        # unit, mocked Gmail/Graph APIs, OAuth loopback flow, stdio end-to-end
```

Layout:

```
src/cli.ts                 command line and stdio entry point
src/server.ts              MCP tools
src/auth.ts                token refresh, caching, authenticated requests
src/oauth.ts               browser sign-in (authorisation code + PKCE on a loopback port)
src/mime.ts                address parsing and message building for Gmail drafts
src/store.ts, config.ts    accounts.json and config.json
src/providers/gmail.ts     Gmail API
src/providers/outlook.ts   Microsoft Graph API
```
