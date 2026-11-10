import { accountsPath, readJson, writeJsonPrivate } from "./config.js";
import { Account, MailhubError } from "./types.js";

interface StoreFile {
  version: 1;
  accounts: Record<string, Account>;
}

const ALIAS_RE = /^[a-z0-9][a-z0-9_-]{0,31}$/;

export function validateAlias(alias: string): void {
  if (!ALIAS_RE.test(alias) || alias === "all") {
    throw new MailhubError(
      `Invalid alias "${alias}". Use 1-32 lowercase letters, digits, "-" or "_" (and not "all").`,
    );
  }
}

function load(): StoreFile {
  const file = readJson<StoreFile>(accountsPath(), { version: 1, accounts: {} });
  return { version: 1, accounts: file.accounts ?? {} };
}

/**
 * Every write re-reads the file first, so the long-running server and the CLI
 * can both update it without clobbering each other's accounts.
 */
function mutate(fn: (file: StoreFile) => void): void {
  const file = load();
  fn(file);
  writeJsonPrivate(accountsPath(), file);
}

export function listAccounts(): Account[] {
  return Object.values(load().accounts).sort((a, b) => a.alias.localeCompare(b.alias));
}

export function getAccount(alias: string): Account {
  const account = load().accounts[alias];
  if (!account) {
    const known = listAccounts().map((a) => a.alias);
    throw new MailhubError(
      `No account with alias "${alias}". Known aliases: ${known.length ? known.join(", ") : "(none yet)"}.`,
    );
  }
  return account;
}

export function hasAccount(alias: string): boolean {
  return alias in load().accounts;
}

export function saveAccount(account: Account): void {
  validateAlias(account.alias);
  mutate((file) => {
    file.accounts[account.alias] = account;
  });
}

export function removeAccount(alias: string): boolean {
  let removed = false;
  mutate((file) => {
    removed = alias in file.accounts;
    delete file.accounts[alias];
  });
  return removed;
}

/** Microsoft rotates refresh tokens, so the newest one is persisted after each refresh. */
export function updateRefreshToken(alias: string, refreshToken: string): void {
  mutate((file) => {
    const account = file.accounts[alias];
    if (account) account.refreshToken = refreshToken;
  });
}
