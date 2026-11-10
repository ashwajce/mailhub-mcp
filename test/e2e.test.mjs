import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { after, before, test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { authorizeInBrowser } from "../dist/oauth.js";
import { saveAccount } from "../dist/store.js";
import { gmailAccount, gmailDraftAccount, json, makeHome, mockFetch, outlookAccount } from "./helpers.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const cli = path.join(here, "..", "dist", "cli.js");
const preload = pathToFileURL(path.join(here, "fixtures", "mock-apis.mjs")).href;

let home;
let client;

before(async () => {
  home = makeHome();
  saveAccount(gmailAccount);
  saveAccount(outlookAccount);
  saveAccount(gmailDraftAccount);
  client = new Client({ name: "mailhub-test", version: "0.0.0" });
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: ["--import", preload, cli],
      env: { ...process.env, MAILHUB_HOME: home },
      stderr: "ignore",
    }),
  );
});

after(async () => {
  await client?.close();
});

const call = async (name, args = {}) => {
  const result = await client.callTool({ name, arguments: args });
  const text = result.content[0].text;
  return { isError: result.isError === true, text, data: result.isError ? undefined : JSON.parse(text) };
};

test("server advertises its tools and usage instructions over stdio", async () => {
  const { tools } = await client.listTools();
  assert.deepEqual(
    tools.map((t) => t.name).sort(),
    ["create_draft", "download_attachment", "list_accounts", "read_email", "read_thread", "search_emails"],
  );
  // There must be no way to send, delete or modify mail.
  assert.equal(tools.some((t) => /send|delete|trash|move|label|update/i.test(t.name)), false);
  const search = tools.find((t) => t.name === "search_emails");
  assert.ok(search.inputSchema.properties.account);
  assert.equal(search.annotations.readOnlyHint, true);
  assert.match(client.getInstructions(), /untrusted/);
});

test("list_accounts never exposes tokens", async () => {
  const { data, text } = await call("list_accounts");
  assert.deepEqual(data.accounts, [
    { alias: "personal", provider: "gmail", email: "me@example.com", can_draft: false },
    { alias: "work", provider: "outlook", email: "me@contoso.example", can_draft: false },
    { alias: "writer", provider: "gmail", email: "writer@example.com", can_draft: true },
  ]);
  assert.doesNotMatch(text, /refresh/i);
});

test("search_emails fans out across every mailbox by default", async () => {
  const { data } = await call("search_emails", { query: "invoice" });
  assert.equal(data.errors, undefined);
  assert.deepEqual(
    data.results.map((r) => [r.account, r.provider, r.count, r.messages[0].subject, r.messages[0].thread_id]),
    [
      ["personal", "gmail", 1, "Gmail invoice", "gt1"],
      ["work", "outlook", 1, "Outlook invoice", "oc1"],
      ["writer", "gmail", 1, "Gmail invoice", "gt1"],
    ],
  );
});

test("one failing mailbox is reported without losing the others", async () => {
  const { data } = await call("search_emails", { query: "blocked" });
  assert.deepEqual(data.results.map((r) => r.account), ["personal", "writer"]);
  assert.equal(data.errors.length, 1);
  assert.equal(data.errors[0].account, "work");
  assert.match(data.errors[0].error, /403/);
});

test("read_email and download_attachment work end to end", async () => {
  const { data: mail } = await call("read_email", { account: "Personal", id: "g1" });
  assert.equal(mail.body, "Body from Gmail");
  assert.deepEqual(mail.attachments, [{ attachment_id: "1", filename: "../data.csv", mime_type: "text/csv", size: 14 }]);

  const { data: saved } = await call("download_attachment", { account: "personal", message_id: "g1", attachment_id: "1" });
  // The hostile "../" in the attachment name must not escape the download folder.
  assert.equal(saved.path, path.join(home, "downloads", "personal", "data.csv"));
  assert.equal(fs.readFileSync(saved.path, "utf8"), "col1,col2\n1,2\n");
  assert.equal(saved.text, "col1,col2\n1,2\n");

  const { data: again } = await call("download_attachment", { account: "personal", message_id: "g1", attachment_id: "1" });
  assert.equal(path.basename(again.path), "data (1).csv");
});

test("bad input comes back as a tool error, not a crash", async () => {
  const unknown = await call("read_email", { account: "nope", id: "x" });
  assert.equal(unknown.isError, true);
  assert.match(unknown.text, /No account with alias "nope"\. Known aliases: personal, work, writer/);

  const paging = await call("search_emails", { page_token: "abc" });
  assert.equal(paging.isError, true);
  assert.match(paging.text, /single account alias/);

  const date = await call("search_emails", { after: "yesterday" });
  assert.equal(date.isError, true);

  // The server is still alive afterwards.
  assert.equal((await call("list_accounts")).data.accounts.length, 3);
});

test("create_draft saves a draft only where drafting was enabled", async () => {
  const refused = await call("create_draft", { account: "personal", to: ["jane@example.org"], subject: "Hi", body: "Hello" });
  assert.equal(refused.isError, true);
  assert.match(refused.text, /connected read-only.*mailhub-mcp add gmail personal --drafts --force/);

  const { data } = await call("create_draft", {
    account: "writer",
    to: ["Jane Doe <jane@example.org>"],
    subject: "Quarterly figures",
    body: "Hello Jane,\n\nFigures attached.\n",
  });
  assert.match(data.status, /not been sent/);
  assert.deepEqual(
    [data.account, data.draft_id, data.message_id, data.thread_id, data.subject, data.to],
    ["writer", "d1", "dm1", "dt1", "Quarterly figures", "Jane Doe <jane@example.org>"],
  );

  const noSubject = await call("create_draft", { account: "writer", to: ["jane@example.org"], body: "Hello" });
  assert.equal(noSubject.isError, true);
  assert.match(noSubject.text, /needs a subject/);

  const badAddress = await call("create_draft", { account: "writer", to: ["not-an-address"], subject: "x", body: "y" });
  assert.equal(badAddress.isError, true);
  assert.match(badAddress.text, /not a valid email address/);
});

test("cli lists accounts and reports health", () => {
  const env = { ...process.env, MAILHUB_HOME: home };
  const list = execFileSync(process.execPath, [cli, "list"], { env, encoding: "utf8" });
  assert.match(list, /personal\s+gmail\s+read-only\s+me@example\.com/);
  assert.match(list, /work\s+outlook\s+read-only\s+me@contoso\.example/);
  assert.match(list, /writer\s+gmail\s+read\+drafts\s+writer@example\.com/);
  const check = execFileSync(process.execPath, ["--import", preload, cli, "check"], { env, encoding: "utf8" });
  assert.match(check, /ok\s+personal\s+me@example\.com\s+latest mail 2026-01-01/);
  assert.match(check, /ok\s+work/);
});

function httpGet(url) {
  return new Promise((resolve, reject) => {
    http
      .get(url, (res) => {
        res.resume();
        res.on("end", () => resolve(res.statusCode));
      })
      .on("error", reject);
  });
}

test("browser sign-in uses PKCE, checks state and exchanges the code", async () => {
  const lines = [];
  const calls = mockFetch([
    ["token.example/token", () => json({ access_token: "at", refresh_token: "rt", expires_in: 3600 })],
  ]);
  const flow = authorizeInBrowser({
    authUrl: "https://auth.example/authorize",
    tokenUrl: "https://token.example/token",
    clientId: "client-1",
    scopes: ["scope.a", "scope.b"],
    redirectHost: "127.0.0.1",
    extraParams: { prompt: "select_account" },
    openBrowser: false,
    log: (line) => lines.push(line),
  });
  while (!lines.some((l) => l.startsWith("https://"))) await new Promise((r) => setTimeout(r, 10));
  const authUrl = new URL(lines.find((l) => l.startsWith("https://")));
  const redirect = authUrl.searchParams.get("redirect_uri");
  assert.match(redirect, /^http:\/\/127\.0\.0\.1:\d+$/);
  assert.equal(authUrl.searchParams.get("scope"), "scope.a scope.b");
  assert.equal(authUrl.searchParams.get("code_challenge_method"), "S256");
  assert.equal(authUrl.searchParams.get("prompt"), "select_account");

  // A forged callback with the wrong state is refused and does not complete the flow.
  assert.equal(await httpGet(`${redirect}/?code=evil&state=wrong`), 400);
  assert.equal(await httpGet(`${redirect}/?code=good-code&state=${authUrl.searchParams.get("state")}`), 200);

  const tokens = await flow;
  assert.equal(tokens.refresh_token, "rt");
  const form = new URLSearchParams(calls[0].body);
  assert.equal(form.get("code"), "good-code");
  assert.equal(form.get("redirect_uri"), redirect);
  const challenge = crypto.createHash("sha256").update(form.get("code_verifier")).digest("base64url");
  assert.equal(challenge, authUrl.searchParams.get("code_challenge"));
});

test("a denied sign-in is reported", async () => {
  const lines = [];
  mockFetch([]);
  const flow = authorizeInBrowser({
    authUrl: "https://auth.example/authorize",
    tokenUrl: "https://token.example/token",
    clientId: "client-1",
    scopes: ["s"],
    redirectHost: "localhost",
    openBrowser: false,
    log: (line) => lines.push(line),
  });
  const outcome = assert.rejects(flow, /access_denied/);
  while (!lines.some((l) => l.startsWith("https://"))) await new Promise((r) => setTimeout(r, 10));
  const authUrl = new URL(lines.find((l) => l.startsWith("https://")));
  const redirect = authUrl.searchParams.get("redirect_uri");
  assert.match(redirect, /^http:\/\/localhost:\d+$/);
  const port = new URL(redirect).port;
  await httpGet(`http://127.0.0.1:${port}/?error=access_denied&state=${authUrl.searchParams.get("state")}`);
  await outcome;
});
