# Privacy policy

Last updated: 24 November 2026

mailhub-mcp is open-source software that you run on your own computer. It connects an MCP client
(for example the Claude desktop app) to mailboxes that you choose to connect. There is no hosted
service behind it, and the authors of mailhub-mcp do not receive, store or see any of your data.

## What mailhub accesses

Only the mailboxes you sign in to with `mailhub-mcp add`, and only with the permissions you grant:

| Provider | Permission | Used for |
|---|---|---|
| Google | `gmail.readonly` | Searching and reading messages, threads and attachments |
| Google | `gmail.compose` (only if you add `--drafts`) | Saving drafts. mailhub contains no code that sends mail |
| Microsoft | `Mail.Read`, `User.Read`, `offline_access` | Searching and reading mail, identifying the signed-in account, staying signed in |
| Microsoft | `Mail.ReadWrite` (only if you add `--drafts`) | Saving drafts |
| Microsoft | `Mail.Read.Shared` / `Mail.ReadWrite.Shared` (only for shared mailboxes) | The same, for a shared or delegated mailbox |

mailhub never sends, deletes, moves or labels mail.

## Where your data goes

- **Sign-in tokens** are stored in a file on your computer (`~/.mailhub-mcp/accounts.json`), readable
  only by your user account. They are sent only to Google or Microsoft, to obtain access.
- **Email content** is fetched directly from Google or Microsoft to your computer when your MCP client
  asks for it, and is passed to that client. What the client does with it (for example, sending it to
  an AI model so it can answer your question) is governed by that client's own terms and privacy policy.
- **Attachments** you ask to download are saved to a folder on your computer (`~/.mailhub-mcp/downloads`).
- mailhub has no analytics, telemetry or logging to any server, and shares nothing with third parties.

## Google user data

mailhub's use of information received from Google APIs is limited to providing the features described
above to you, on your own computer. Google user data is not transferred to the authors or to anyone
else and is not used for advertising. mailhub itself does not use it to train AI models. No person
other than you can read it through mailhub.

## Keeping and deleting data

mailhub keeps only the sign-in tokens and any attachments you downloaded. To remove them:

- `mailhub-mcp remove <alias>` deletes the stored token for one mailbox.
- Deleting the `~/.mailhub-mcp` folder removes everything mailhub has stored.
- You can also revoke access at any time at <https://myaccount.google.com/permissions> (Google) or
  <https://myaccount.microsoft.com> (Microsoft).

## Contact

Questions or concerns: open an issue at <https://github.com/ashwajce/mailhub-mcp/issues>.
