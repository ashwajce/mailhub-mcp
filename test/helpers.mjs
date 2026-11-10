import fs from "node:fs";
import os from "node:os";
import path from "node:path";

/** Point mailhub at a fresh temp data folder with two fake accounts and fake OAuth apps. */
export function makeHome() {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), "mailhub-test-"));
  process.env.MAILHUB_HOME = home;
  process.env.MAILHUB_GOOGLE_CLIENT_ID = "g-client";
  process.env.MAILHUB_GOOGLE_CLIENT_SECRET = "g-secret";
  process.env.MAILHUB_MS_CLIENT_ID = "ms-client";
  delete process.env.MAILHUB_MS_TENANT;
  delete process.env.MAILHUB_DOWNLOAD_DIR;
  return home;
}

export const gmailAccount = {
  alias: "personal",
  provider: "gmail",
  email: "me@example.com",
  login: "me@example.com",
  refreshToken: "g-refresh",
  scopes: ["https://www.googleapis.com/auth/gmail.readonly"],
  addedAt: "2026-01-01T00:00:00.000Z",
};

/** Same mailbox types, but signed in with --drafts. */
export const gmailDraftAccount = {
  ...gmailAccount,
  alias: "writer",
  email: "writer@example.com",
  login: "writer@example.com",
  scopes: ["https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/gmail.compose"],
};

export const outlookDraftAccount = {
  alias: "office",
  provider: "outlook",
  email: "me@contoso.example",
  login: "me@contoso.example",
  refreshToken: "ms-refresh-1",
  scopes: ["offline_access", "User.Read", "Mail.Read", "Mail.ReadWrite"],
  addedAt: "2026-01-01T00:00:00.000Z",
};

export const outlookAccount = {
  alias: "work",
  provider: "outlook",
  email: "me@contoso.example",
  login: "me@contoso.example",
  refreshToken: "ms-refresh-1",
  scopes: ["offline_access", "User.Read", "Mail.Read"],
  addedAt: "2026-01-01T00:00:00.000Z",
};

export function json(body, status = 200, headers = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json", ...headers },
  });
}

/**
 * Replace global fetch with a router. `routes` is a list of [matcher, handler];
 * matcher is a substring or RegExp tested against the full URL. Returns the call log.
 */
export function mockFetch(routes) {
  const calls = [];
  globalThis.fetch = async (input, init = {}) => {
    const url = typeof input === "string" ? input : input.url;
    const call = { url, init, body: typeof init.body === "string" ? init.body : undefined };
    calls.push(call);
    for (const [matcher, handler] of routes) {
      const hit = typeof matcher === "string" ? url.includes(matcher) : matcher.test(url);
      if (hit) return handler(call);
    }
    throw new Error(`Unexpected fetch in test: ${url}`);
  };
  return calls;
}

export const b64url = (text, encoding = "utf8") => Buffer.from(text, encoding).toString("base64url");
