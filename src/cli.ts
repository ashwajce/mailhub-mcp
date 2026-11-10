#!/usr/bin/env node
import fs from "node:fs";
import { parseArgs } from "node:util";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { GMAIL_SCOPES, OUTLOOK_SCOPES, OUTLOOK_SHARED_SCOPE, oauthSettings } from "./auth.js";
import { accountsPath, configPath, homeDir, loadConfig, saveConfigPatch } from "./config.js";
import { authorizeInBrowser } from "./oauth.js";
import { providerFor } from "./providers/index.js";
import { createServer } from "./server.js";
import { getAccount, hasAccount, listAccounts, removeAccount, saveAccount, validateAlias } from "./store.js";
import { Account, MailhubError } from "./types.js";

const HELP = `mailhub-mcp - many Gmail and Outlook mailboxes behind one MCP server

Usage:
  mailhub-mcp [serve]                         Run the MCP server on stdio (what Claude starts)

  mailhub-mcp setup google --client-id <id> --client-secret <secret>
  mailhub-mcp setup microsoft --client-id <application-id> [--tenant common]

  mailhub-mcp add gmail <alias>               Sign in to a Gmail / Google Workspace mailbox
  mailhub-mcp add outlook <alias>             Sign in to an Outlook / Microsoft 365 mailbox
      --login-hint <email>                    Pre-select the account on the sign-in page
      --shared-mailbox <email>                Outlook: read a shared or delegated mailbox
      --tenant <id-or-domain>                 Outlook: sign in against one specific tenant
      --force                                 Replace an existing alias (re-authorise)
      --no-browser                            Print the sign-in URL instead of opening it

  mailhub-mcp list                            Show connected mailboxes
  mailhub-mcp check [alias]                   Verify that stored sign-ins still work
  mailhub-mcp remove <alias>                  Forget a mailbox
  mailhub-mcp paths                           Show where config and tokens are stored
`;

function version(): string {
  try {
    const pkg = JSON.parse(fs.readFileSync(new URL("../package.json", import.meta.url), "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

async function getJson<T>(url: string, token: string): Promise<T> {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json" } });
  if (!res.ok) throw new MailhubError(`Request to ${new URL(url).host} failed (${res.status}): ${(await res.text()).slice(0, 300)}`);
  return (await res.json()) as T;
}

function setup(provider: string | undefined, opts: Record<string, string | boolean | undefined>): void {
  const clientId = opts["client-id"] as string | undefined;
  if (provider === "google") {
    const clientSecret = opts["client-secret"] as string | undefined;
    if (!clientId || !clientSecret) throw new MailhubError("setup google needs --client-id and --client-secret.");
    saveConfigPatch({ google: { clientId, clientSecret } });
  } else if (provider === "microsoft") {
    if (!clientId) throw new MailhubError("setup microsoft needs --client-id.");
    const tenant = (opts.tenant as string | undefined) ?? "common";
    saveConfigPatch({ microsoft: { clientId, tenant } });
  } else {
    throw new MailhubError('Use "setup google" or "setup microsoft".');
  }
  console.log(`Saved ${provider} settings to ${configPath()}`);
}

async function add(
  provider: string | undefined,
  alias: string | undefined,
  opts: Record<string, string | boolean | undefined>,
): Promise<void> {
  if (provider !== "gmail" && provider !== "outlook") throw new MailhubError('Use "add gmail <alias>" or "add outlook <alias>".');
  if (!alias) throw new MailhubError(`Give the mailbox a short alias, for example: mailhub-mcp add ${provider} work`);
  validateAlias(alias);
  if (hasAccount(alias) && !opts.force) {
    throw new MailhubError(`Alias "${alias}" already exists. Use --force to authorise it again, or pick another alias.`);
  }

  const loginHint = opts["login-hint"] as string | undefined;
  const sharedMailbox = (opts["shared-mailbox"] as string | undefined)?.toLowerCase();
  const tenant = opts.tenant as string | undefined;
  const openBrowser = !opts["no-browser"];
  if (provider === "gmail" && (sharedMailbox || tenant)) {
    throw new MailhubError("--shared-mailbox and --tenant only apply to Outlook accounts.");
  }

  const settings = oauthSettings(provider, tenant);
  let account: Account;

  if (provider === "gmail") {
    const tokens = await authorizeInBrowser({
      ...settings,
      scopes: GMAIL_SCOPES,
      redirectHost: "127.0.0.1",
      // offline + consent makes Google issue a refresh token every time.
      extraParams: { access_type: "offline", prompt: "consent select_account", ...(loginHint ? { login_hint: loginHint } : {}) },
      openBrowser,
    });
    if (!tokens.refresh_token) throw new MailhubError("Google did not return a refresh token. Try again with --force.");
    if (tokens.scope && !tokens.scope.split(" ").includes(GMAIL_SCOPES[0])) {
      throw new MailhubError("Gmail access was not granted. Run the command again and tick the Gmail permission on the consent page.");
    }
    const profile = await getJson<{ emailAddress: string }>(
      "https://gmail.googleapis.com/gmail/v1/users/me/profile",
      tokens.access_token,
    );
    account = {
      alias,
      provider,
      email: profile.emailAddress,
      login: profile.emailAddress,
      refreshToken: tokens.refresh_token,
      scopes: GMAIL_SCOPES,
      addedAt: new Date().toISOString(),
    };
  } else {
    const scopes = sharedMailbox ? [...OUTLOOK_SCOPES, OUTLOOK_SHARED_SCOPE] : OUTLOOK_SCOPES;
    const tokens = await authorizeInBrowser({
      ...settings,
      scopes,
      redirectHost: "localhost",
      extraParams: { prompt: "select_account", ...(loginHint ? { login_hint: loginHint } : {}) },
      openBrowser,
    });
    if (!tokens.refresh_token) throw new MailhubError("Microsoft did not return a refresh token (offline_access was not granted).");
    const me = await getJson<{ mail?: string | null; userPrincipalName?: string }>(
      "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName",
      tokens.access_token,
    );
    const login = (me.mail || me.userPrincipalName || "").toLowerCase();
    if (sharedMailbox) {
      // Fail now, with a clear message, rather than on the first search.
      await getJson(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sharedMailbox)}/messages?$top=1&$select=id`,
        tokens.access_token,
      ).catch((err: Error) => {
        throw new MailhubError(`${login} cannot read the mailbox ${sharedMailbox}. Check the delegation in Microsoft 365. (${err.message})`);
      });
    }
    account = {
      alias,
      provider,
      email: sharedMailbox ?? login,
      login,
      refreshToken: tokens.refresh_token,
      scopes,
      ...(tenant ? { tenant } : {}),
      ...(sharedMailbox ? { sharedMailbox } : {}),
      addedAt: new Date().toISOString(),
    };
  }

  const duplicate = listAccounts().find(
    (a) => a.alias !== alias && a.provider === account.provider && a.email.toLowerCase() === account.email.toLowerCase(),
  );
  saveAccount(account);
  console.log(`Connected ${account.email} as "${alias}" (${provider}).`);
  if (duplicate) console.log(`Note: the same mailbox is also connected as "${duplicate.alias}".`);
}

function list(): void {
  const accounts = listAccounts();
  if (!accounts.length) {
    console.log("No mailboxes connected yet. Start with: mailhub-mcp add gmail <alias>");
    return;
  }
  const width = Math.max(...accounts.map((a) => a.alias.length), 5);
  for (const a of accounts) {
    const via = a.sharedMailbox ? `  (via ${a.login})` : "";
    console.log(`${a.alias.padEnd(width)}  ${a.provider.padEnd(7)}  ${a.email}${via}`);
  }
}

async function check(alias: string | undefined): Promise<void> {
  const accounts = alias ? [getAccount(alias)] : listAccounts();
  if (!accounts.length) {
    console.log("No mailboxes connected yet.");
    return;
  }
  let failed = 0;
  for (const account of accounts) {
    try {
      const result = await providerFor(account).search(account, { maxResults: 1 });
      const latest = result.messages[0]?.date.slice(0, 10);
      console.log(`ok    ${account.alias}  ${account.email}  ${latest ? `latest mail ${latest}` : "mailbox is empty"}`);
    } catch (err) {
      failed++;
      console.log(`FAIL  ${account.alias}  ${account.email}  ${(err as Error).message}`);
    }
  }
  if (failed) process.exitCode = 1;
}

async function serve(): Promise<void> {
  const server = createServer(version());
  await server.connect(new StdioServerTransport());
  // stdout carries the protocol, so diagnostics go to stderr.
  process.stderr.write(`mailhub-mcp ${version()} ready, ${listAccountsSafe()} mailbox(es) connected\n`);
}

function listAccountsSafe(): number {
  try {
    return listAccounts().length;
  } catch {
    return 0;
  }
}

async function main(): Promise<void> {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      "client-id": { type: "string" },
      "client-secret": { type: "string" },
      tenant: { type: "string" },
      "login-hint": { type: "string" },
      "shared-mailbox": { type: "string" },
      force: { type: "boolean" },
      "no-browser": { type: "boolean" },
      help: { type: "boolean", short: "h" },
      version: { type: "boolean", short: "v" },
    },
  });
  const [command = "serve", arg1, arg2] = positionals;

  if (values.help || command === "help") return void console.log(HELP);
  if (values.version) return void console.log(version());

  switch (command) {
    case "serve":
      return serve();
    case "setup":
      return setup(arg1, values);
    case "add":
      return add(arg1, arg2?.toLowerCase(), values);
    case "list":
      return list();
    case "check":
      return check(arg1?.toLowerCase());
    case "remove": {
      if (!arg1) throw new MailhubError("Which alias? Usage: mailhub-mcp remove <alias>");
      const removed = removeAccount(arg1.toLowerCase());
      console.log(removed ? `Removed "${arg1}". Its stored token is deleted.` : `No account with alias "${arg1}".`);
      return;
    }
    case "paths": {
      console.log(`Data folder:  ${homeDir()}`);
      console.log(`Settings:     ${configPath()}`);
      console.log(`Tokens:       ${accountsPath()}`);
      console.log(`Downloads:    ${loadConfig().downloadDir}`);
      return;
    }
    default:
      throw new MailhubError(`Unknown command "${command}". Run: mailhub-mcp --help`);
  }
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  process.stderr.write(`mailhub-mcp: ${message}\n`);
  process.exit(1);
});
