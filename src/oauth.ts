import { spawn } from "node:child_process";
import crypto from "node:crypto";
import http from "node:http";
import type { AddressInfo } from "node:net";

export interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
}

export class OAuthError extends Error {
  constructor(
    public readonly code: string,
    description: string,
  ) {
    super(description ? `${code}: ${description}` : code);
    this.name = "OAuthError";
  }
}

/** POST to an OAuth token endpoint and return the parsed tokens. */
export async function postToken(tokenUrl: string, params: Record<string, string>): Promise<TokenResponse> {
  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: new URLSearchParams(params).toString(),
  });
  const text = await res.text();
  let body: Record<string, unknown> = {};
  try {
    body = JSON.parse(text) as Record<string, unknown>;
  } catch {
    // Non-JSON error page: reported below with the status code.
  }
  if (!res.ok || typeof body.access_token !== "string") {
    const code = typeof body.error === "string" ? body.error : `http_${res.status}`;
    const description = typeof body.error_description === "string" ? body.error_description : "";
    throw new OAuthError(code, description.split("\n")[0]);
  }
  return body as unknown as TokenResponse;
}

function base64url(buf: Buffer): string {
  return buf.toString("base64url");
}

function openBrowser(url: string): void {
  const [cmd, args] =
    process.platform === "darwin"
      ? ["open", [url]]
      : process.platform === "win32"
        ? ["rundll32", ["url.dll,FileProtocolHandler", url]]
        : ["xdg-open", [url]];
  try {
    const child = spawn(cmd as string, args as string[], { stdio: "ignore", detached: true });
    child.on("error", () => {});
    child.unref();
  } catch {
    // The URL is also printed, so the user can open it by hand.
  }
}

function listen(server: http.Server, port: number, host: string): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => {
      server.removeListener("error", reject);
      resolve();
    });
  });
}

const DONE_PAGE = (title: string, detail: string) =>
  `<!doctype html><meta charset="utf-8"><title>mailhub-mcp</title>` +
  `<body style="font-family:system-ui;margin:4rem auto;max-width:32rem">` +
  `<h2>${title}</h2><p>${detail}</p></body>`;

export interface AuthorizeOptions {
  authUrl: string;
  tokenUrl: string;
  clientId: string;
  clientSecret?: string;
  scopes: string[];
  /** Google accepts 127.0.0.1 loopback redirects; Entra desktop apps register http://localhost. */
  redirectHost: "127.0.0.1" | "localhost";
  extraParams?: Record<string, string>;
  openBrowser?: boolean;
  timeoutMs?: number;
  log?: (line: string) => void;
}

/**
 * OAuth 2.0 authorisation code flow with PKCE, using a one-shot web server on
 * the loopback interface to receive the redirect. Returns the token response.
 */
export async function authorizeInBrowser(opts: AuthorizeOptions): Promise<TokenResponse> {
  const log = opts.log ?? ((line: string) => process.stderr.write(line + "\n"));
  const verifier = base64url(crypto.randomBytes(32));
  const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
  const state = base64url(crypto.randomBytes(16));

  let settle: (value: { code?: string; error?: string }) => void = () => {};
  const callback = new Promise<{ code?: string; error?: string }>((resolve) => {
    settle = resolve;
  });

  const handler = (req: http.IncomingMessage, res: http.ServerResponse) => {
    const url = new URL(req.url ?? "/", "http://localhost");
    const params = url.searchParams;
    if (!params.has("code") && !params.has("error")) {
      res.writeHead(404).end();
      return;
    }
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    if (params.get("state") !== state) {
      res.writeHead(400).end(DONE_PAGE("Sign-in rejected", "State mismatch. Start again from the terminal."));
      return;
    }
    const error = params.get("error");
    if (error) {
      res.writeHead(400).end(DONE_PAGE("Sign-in failed", "Return to the terminal for details."));
      settle({ error: `${error}: ${params.get("error_description") ?? ""}`.trim() });
      return;
    }
    res.writeHead(200).end(DONE_PAGE("Account connected", "You can close this tab and return to the terminal."));
    settle({ code: params.get("code") ?? undefined });
  };

  const servers: http.Server[] = [];
  const primary = http.createServer(handler);
  await listen(primary, 0, "127.0.0.1");
  servers.push(primary);
  const port = (primary.address() as AddressInfo).port;
  if (opts.redirectHost === "localhost") {
    // "localhost" may resolve to IPv6 first, so answer on ::1 as well when possible.
    const v6 = http.createServer(handler);
    try {
      await listen(v6, port, "::1");
      servers.push(v6);
    } catch {
      // IPv6 loopback unavailable: browsers fall back to 127.0.0.1.
    }
  }

  const redirectUri = `http://${opts.redirectHost}:${port}`;
  const authUrl = new URL(opts.authUrl);
  authUrl.search = new URLSearchParams({
    client_id: opts.clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    scope: opts.scopes.join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    ...opts.extraParams,
  }).toString();

  log("Open this URL to sign in (it should open in your browser automatically):");
  log("");
  log(authUrl.toString());
  log("");
  if (opts.openBrowser !== false) openBrowser(authUrl.toString());

  let timer: NodeJS.Timeout | undefined;
  const timeout = new Promise<{ code?: string; error?: string }>((resolve) => {
    timer = setTimeout(() => resolve({ error: "Timed out waiting for the browser sign-in." }), opts.timeoutMs ?? 300_000);
  });

  try {
    const result = await Promise.race([callback, timeout]);
    if (result.error || !result.code) {
      throw new OAuthError("authorization_failed", result.error ?? "No authorisation code received.");
    }
    return await postToken(opts.tokenUrl, {
      grant_type: "authorization_code",
      code: result.code,
      redirect_uri: redirectUri,
      client_id: opts.clientId,
      code_verifier: verifier,
      ...(opts.clientSecret ? { client_secret: opts.clientSecret } : {}),
    });
  } finally {
    clearTimeout(timer);
    for (const server of servers) {
      server.close();
      server.closeAllConnections();
    }
  }
}
