import { loadConfig } from "./config.js";
import { OAuthError, postToken } from "./oauth.js";
import { updateRefreshToken } from "./store.js";
import { Account, MailhubError, ProviderName } from "./types.js";
import { sleep } from "./util.js";

export const GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];
export const OUTLOOK_SCOPES = ["offline_access", "User.Read", "Mail.Read"];
export const OUTLOOK_SHARED_SCOPE = "Mail.Read.Shared";

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
      throw new MailhubError(
        `Account "${account.alias}" (${account.login}) needs to be authorised again. ` +
          `Run: mailhub-mcp add ${account.provider} ${account.alias} --force`,
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
 * Authenticated JSON GET against a mail API. Retries once on 401 with a fresh
 * token and backs off on throttling or transient server errors.
 */
export async function apiJson<T>(account: Account, url: string, headers: Record<string, string> = {}): Promise<T> {
  let forced = false;
  for (let attempt = 0; ; attempt++) {
    const token = await accessToken(account);
    const res = await fetch(url, {
      headers: { Accept: "application/json", ...headers, Authorization: `Bearer ${token}` },
    });
    if (res.ok) return (await res.json()) as T;

    if (res.status === 401 && !forced) {
      forced = true;
      cache.delete(account.alias);
      continue;
    }
    if ((res.status === 429 || res.status === 502 || res.status === 503 || res.status === 504) && attempt < 3) {
      const retryAfter = Number(res.headers.get("retry-after"));
      const waitMs = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 500 * 2 ** attempt;
      await sleep(Math.min(waitMs, 20_000));
      continue;
    }
    const detail = errorDetail(await res.text());
    throw new MailhubError(`${account.provider} API error ${res.status} for "${account.alias}": ${detail}`);
  }
}
