import { loadConfig } from "./config.js";
import { OAuthError, postToken } from "./oauth.js";
import { updateRefreshToken } from "./store.js";
import { Account, MailhubError, ProviderName } from "./types.js";
import { sleep } from "./util.js";

export const GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];
export const OUTLOOK_SCOPES = ["offline_access", "User.Read", "Mail.Read"];
export const OUTLOOK_SHARED_SCOPE = "Mail.Read.Shared";

// Optional extra scopes, requested only when a mailbox is added with --drafts.
// Gmail has no drafts-only scope: gmail.compose is the narrowest one that can create drafts.
export const GMAIL_DRAFT_SCOPE = "https://www.googleapis.com/auth/gmail.compose";
export const OUTLOOK_DRAFT_SCOPE = "Mail.ReadWrite";
export const OUTLOOK_SHARED_DRAFT_SCOPE = "Mail.ReadWrite.Shared";

/** Scopes to request when signing in to a mailbox. */
export function scopesFor(provider: ProviderName, opts: { drafts?: boolean; shared?: boolean } = {}): string[] {
  if (provider === "gmail") return opts.drafts ? [...GMAIL_SCOPES, GMAIL_DRAFT_SCOPE] : GMAIL_SCOPES;
  const scopes = [...OUTLOOK_SCOPES];
  if (opts.shared) scopes.push(OUTLOOK_SHARED_SCOPE);
  if (opts.drafts) scopes.push(opts.shared ? OUTLOOK_SHARED_DRAFT_SCOPE : OUTLOOK_DRAFT_SCOPE);
  return scopes;
}

/** Whether this mailbox was authorised for creating drafts. */
export function canDraft(account: Account): boolean {
  if (account.provider === "gmail") return account.scopes.includes(GMAIL_DRAFT_SCOPE);
  return account.scopes.includes(account.sharedMailbox ? OUTLOOK_SHARED_DRAFT_SCOPE : OUTLOOK_DRAFT_SCOPE);
}

export function requireDraftAccess(account: Account): void {
  if (canDraft(account)) return;
  const shared = account.sharedMailbox ? ` --shared-mailbox ${account.sharedMailbox}` : "";
  throw new MailhubError(
    `Account "${account.alias}" is connected read-only. To allow drafts, run: ` +
      `mailhub-mcp add ${account.provider} ${account.alias} --drafts --force${shared}`,
  );
}

export interface OAuthSettings {
  authUrl: string;
  tokenUrl: string;
  clientId: string;
  clientSecret?: string;
}

/** Resolve the OAuth app settings for a provider, with a clear error when setup is missing. */
export function oauthSettings(provider: ProviderName, tenant?: string): OAuthSettings {
  const config = loadConfig();
  if (provider === "gmail") {
    const { clientId, clientSecret } = config.google ?? {};
    if (!clientId || !clientSecret) {
      throw new MailhubError(
        "Google OAuth client is not configured. Run: mailhub-mcp setup google --client-id <id> --client-secret <secret>",
      );
    }
    return {
      authUrl: "https://accounts.google.com/o/oauth2/v2/auth",
      tokenUrl: "https://oauth2.googleapis.com/token",
      clientId,
      clientSecret,
    };
  }
  const clientId = config.microsoft?.clientId;
  if (!clientId) {
    throw new MailhubError(
      "Microsoft app registration is not configured. Run: mailhub-mcp setup microsoft --client-id <application-id>",
    );
  }
  const authority = `https://login.microsoftonline.com/${encodeURIComponent(tenant || config.microsoft?.tenant || "common")}/oauth2/v2.0`;
  return { authUrl: `${authority}/authorize`, tokenUrl: `${authority}/token`, clientId };
}

interface CachedToken {
  token: string;
  expiresAt: number;
  /**
   * Refresh tokens this access token belongs to (the one used and, when the
   * provider rotated it, the replacement). A different token on the account
   * means it was authorised again, so the cached access token is dropped.
   */
  refreshTokens: string[];
}

const cache = new Map<string, CachedToken>();
const inflight = new Map<string, Promise<string>>();

export function resetTokenCache(): void {
  cache.clear();
  inflight.clear();
}

async function refresh(account: Account): Promise<string> {
  const settings = oauthSettings(account.provider, account.tenant);
  let tokens;
  try {
    tokens = await postToken(settings.tokenUrl, {
      grant_type: "refresh_token",
      refresh_token: account.refreshToken,
      client_id: settings.clientId,
      ...(settings.clientSecret ? { client_secret: settings.clientSecret } : {}),
    });
  } catch (err) {
    if (err instanceof OAuthError && (err.code === "invalid_grant" || err.code === "interaction_required")) {
      const flags =
        (canDraft(account) ? " --drafts" : "") +
        (account.sharedMailbox ? ` --shared-mailbox ${account.sharedMailbox}` : "");
      throw new MailhubError(
        `Account "${account.alias}" (${account.login}) needs to be authorised again. ` +
          `Run: mailhub-mcp add ${account.provider} ${account.alias} --force${flags}`,
      );
    }
    throw new MailhubError(`Could not refresh the token for "${account.alias}": ${(err as Error).message}`);
  }
  let refreshToken = account.refreshToken;
  if (tokens.refresh_token && tokens.refresh_token !== account.refreshToken) {
    refreshToken = tokens.refresh_token;
    updateRefreshToken(account.alias, refreshToken);
  }
  cache.set(account.alias, {
    token: tokens.access_token,
    expiresAt: Date.now() + (tokens.expires_in - 60) * 1000,
    refreshTokens: [account.refreshToken, refreshToken],
  });
  return tokens.access_token;
}

/** Access token for an account, refreshed when missing, expired or forced. */
export async function accessToken(account: Account, force = false): Promise<string> {
  const cached = cache.get(account.alias);
  if (!force && cached && cached.expiresAt > Date.now() && cached.refreshTokens.includes(account.refreshToken)) {
    return cached.token;
  }
  const pending = inflight.get(account.alias);
  if (pending) return pending;
  const job = refresh(account).finally(() => inflight.delete(account.alias));
  inflight.set(account.alias, job);
  return job;
}

function errorDetail(text: string): string {
  try {
    const body = JSON.parse(text) as { error?: { message?: string } | string; error_description?: string };
    if (typeof body.error === "object" && body.error?.message) return body.error.message;
    if (typeof body.error === "string") return body.error_description ?? body.error;
  } catch {
    // Not JSON.
  }
  return text.slice(0, 300);
}

/**
 * Authenticated JSON request against a mail API (GET, or POST/PATCH when a body is given).
 * Retries once on 401 with a fresh token and backs off on throttling. Transient
 * server errors are retried for reads only, so a write is never applied twice.
 */
export async function apiJson<T>(
  account: Account,
  url: string,
  headers: Record<string, string> = {},
  body?: unknown,
  method: "POST" | "PATCH" = "POST",
): Promise<T> {
  const isWrite = body !== undefined;
  let forced = false;
  for (let attempt = 0; ; attempt++) {
    const token = await accessToken(account);
    const res = await fetch(url, {
      method: isWrite ? method : "GET",
      headers: {
        Accept: "application/json",
        ...(isWrite ? { "Content-Type": "application/json" } : {}),
        ...headers,
        Authorization: `Bearer ${token}`,
      },
      ...(isWrite ? { body: JSON.stringify(body) } : {}),
    });
    if (res.ok) return (await res.json()) as T;

    if (res.status === 401 && !forced) {
      forced = true;
      cache.delete(account.alias);
      continue;
    }
    const transient = res.status === 502 || res.status === 503 || res.status === 504;
    if ((res.status === 429 || (transient && !isWrite)) && attempt < 3) {
      const retryAfter = Number(res.headers.get("retry-after"));
      const waitMs = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 500 * 2 ** attempt;
      await sleep(Math.min(waitMs, 20_000));
      continue;
    }
    const detail = errorDetail(await res.text());
    throw new MailhubError(`${account.provider} API error ${res.status} for "${account.alias}": ${detail}`);
  }
}
