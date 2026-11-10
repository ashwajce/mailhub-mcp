import assert from "node:assert/strict";
import { beforeEach, test } from "node:test";
import { resetTokenCache } from "../dist/auth.js";
import { gmail } from "../dist/providers/gmail.js";
import { outlook } from "../dist/providers/outlook.js";
import { getAccount, saveAccount } from "../dist/store.js";
import { b64url, gmailAccount, json, makeHome, mockFetch, outlookAccount } from "./helpers.mjs";

const googleToken = ["oauth2.googleapis.com/token", () => json({ access_token: "g-access", expires_in: 3600 })];
const msToken = [
  "login.microsoftonline.com/common/oauth2/v2.0/token",
  () => json({ access_token: "ms-access", refresh_token: "ms-refresh-2", expires_in: 3600 }),
];

beforeEach(() => {
  makeHome();
  resetTokenCache();
  saveAccount(gmailAccount);
  saveAccount(outlookAccount);
});

const gmailMeta = (id, subject, extra = {}) => ({
  id,
  threadId: `t-${id}`,
  labelIds: ["INBOX", "UNREAD"],
  snippet: "It&#39;s here",
  internalDate: "1767225600000",
  payload: {
    headers: [
      { name: "From", value: "Acme <billing@acme.com>" },
      { name: "To", value: "me@example.com" },
      { name: "subject", value: subject },
    ],
  },
  ...extra,
});

test("gmail search lists ids then loads summaries, with one token refresh", async () => {
  const calls = mockFetch([
    googleToken,
    [/\/messages\?/, () => json({ messages: [{ id: "m1" }, { id: "m2" }], nextPageToken: "next-1" })],
    [/\/messages\/m1\?format=metadata/, () => json(gmailMeta("m1", "Invoice 1"))],
    [/\/messages\/m2\?format=metadata/, () => json(gmailMeta("m2", "Invoice 2", { labelIds: ["INBOX"] }))],
  ]);
  const result = await gmail.search(getAccount("personal"), { query: "invoice", unreadOnly: true, maxResults: 2 });

  assert.equal(result.nextPageToken, "next-1");
  assert.deepEqual(
    result.messages.map((m) => [m.id, m.threadId, m.subject, m.unread, m.account, m.mailbox]),
    [
      ["m1", "t-m1", "Invoice 1", true, "personal", "me@example.com"],
      ["m2", "t-m2", "Invoice 2", false, "personal", "me@example.com"],
    ],
  );
  assert.equal(result.messages[0].snippet, "It's here");
  assert.equal(result.messages[0].date, "2026-01-01T00:00:00.000Z");

  const tokenCalls = calls.filter((c) => c.url.includes("oauth2.googleapis.com"));
  assert.equal(tokenCalls.length, 1);
  const form = new URLSearchParams(tokenCalls[0].body);
  assert.equal(form.get("grant_type"), "refresh_token");
  assert.equal(form.get("refresh_token"), "g-refresh");
  assert.equal(form.get("client_secret"), "g-secret");

  const listCall = calls.find((c) => /\/messages\?/.test(c.url));
  assert.equal(new URL(listCall.url).searchParams.get("q"), "invoice is:unread");
  assert.equal(listCall.init.headers.Authorization, "Bearer g-access");
});

test("gmail read prefers plain text, honours charset and lists attachments by part id", async () => {
  mockFetch([
    googleToken,
    [
      /\/messages\/m1\?format=full/,
      () =>
        json(
          gmailMeta("m1", "Café invoice", {
            payload: {
              mimeType: "multipart/mixed",
              headers: [
                { name: "From", value: "Acme <billing@acme.com>" },
                { name: "Subject", value: "Café invoice" },
                { name: "Cc", value: "boss@example.com" },
              ],
              parts: [
                {
                  partId: "0",
                  mimeType: "text/plain",
                  filename: "",
                  headers: [{ name: "Content-Type", value: "text/plain; charset=iso-8859-1" }],
                  body: { data: Buffer.from([0x63, 0x61, 0x66, 0xe9]).toString("base64url") },
                },
                { partId: "1", mimeType: "text/html", filename: "", body: { data: b64url("<b>ignored</b>") } },
                { partId: "2", mimeType: "application/pdf", filename: "inv.pdf", body: { attachmentId: "volatile", size: 3 } },
              ],
            },
          }),
        ),
    ],
  ]);
  const m = await gmail.read(getAccount("personal"), "m1", 20000);
  assert.equal(m.body, "café");
  assert.equal(m.cc, "boss@example.com");
  assert.equal(m.bodyTruncated, false);
  assert.deepEqual(m.attachments, [{ id: "2", filename: "inv.pdf", mimeType: "application/pdf", size: 3 }]);
});

test("gmail read falls back to html, fetches out-of-line bodies and truncates", async () => {
  mockFetch([
    googleToken,
    [/\/messages\/m9\/attachments\/big-body/, () => json({ data: b64url("<p>Hello <b>world</b></p><p>Second</p>") })],
    [
      /\/messages\/m9\?format=full/,
      () => json(gmailMeta("m9", "Html only", { payload: { mimeType: "text/html", body: { attachmentId: "big-body" } } })),
    ],
  ]);
  const account = getAccount("personal");
  assert.equal((await gmail.read(account, "m9", 20000)).body, "Hello world\nSecond");
  const cut = await gmail.read(account, "m9", 5);
  assert.equal(cut.body, "Hello");
  assert.equal(cut.bodyTruncated, true);
});

test("gmail attachment is resolved by stable part id", async () => {
  mockFetch([
    googleToken,
    [/\/messages\/m1\/attachments\/fresh-id/, () => json({ data: b64url("PDF") })],
    [
      /\/messages\/m1\?format=full/,
      () =>
        json(
          gmailMeta("m1", "x", {
            payload: {
              mimeType: "multipart/mixed",
              parts: [{ partId: "2", mimeType: "application/pdf", filename: "inv.pdf", body: { attachmentId: "fresh-id" } }],
            },
          }),
        ),
    ],
  ]);
  const account = getAccount("personal");
  const att = await gmail.attachment(account, "m1", "2");
  assert.equal(att.filename, "inv.pdf");
  assert.equal(att.data.toString(), "PDF");
  await assert.rejects(gmail.attachment(account, "m1", "7"), /not found/);
});

const graphMsg = (id, received, extra = {}) => ({
  id,
  conversationId: "conv-1",
  subject: `Subject ${id}`,
  from: { emailAddress: { name: "Jane Doe", address: "jane@contoso.example" } },
  toRecipients: [{ emailAddress: { name: "me@contoso.example", address: "me@contoso.example" } }],
  receivedDateTime: received,
  bodyPreview: "Preview\r\n text",
  isRead: false,
  hasAttachments: false,
  ...extra,
});

test("outlook search applies exact date and unread filters and saves the rotated refresh token", async () => {
  const calls = mockFetch([
    msToken,
    [
      "graph.microsoft.com/v1.0/me/messages?$search=",
      () =>
        json({
          value: [
            graphMsg("a", "2026-01-15T10:00:00Z"),
            graphMsg("b", "2026-01-09T23:00:00Z"), // inside the widened KQL window, before the exact bound
            graphMsg("c", "2026-01-16T10:00:00Z", { isRead: true }),
          ],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc",
        }),
    ],
  ]);
  const result = await outlook.search(getAccount("work"), {
    query: "invoice",
    after: new Date("2026-01-10T00:00:00Z"),
    unreadOnly: true,
    maxResults: 10,
  });
  assert.deepEqual(result.messages.map((m) => m.id), ["a"]);
  const withFiles = await outlook.search(getAccount("work"), { query: "invoice", hasAttachment: true, maxResults: 10 });
  assert.deepEqual(withFiles.messages, []);
  assert.equal(result.messages[0].from, "Jane Doe <jane@contoso.example>");
  assert.equal(result.messages[0].to, "me@contoso.example");
  assert.equal(result.messages[0].snippet, "Preview text");
  assert.equal(result.messages[0].threadId, "conv-1");
  assert.equal(result.nextPageToken, "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc");

  const form = new URLSearchParams(calls[0].body);
  assert.equal(form.get("client_id"), "ms-client");
  assert.equal(form.get("client_secret"), null);
  assert.equal(getAccount("work").refreshToken, "ms-refresh-2");
  assert.match(calls[1].init.headers.Prefer, /ImmutableId/);

  // A second call with the stale in-memory account object must reuse the cached access token.
  await outlook.search({ ...outlookAccount }, { query: "invoice", maxResults: 1 });
  assert.equal(calls.filter((c) => c.url.includes("login.microsoftonline.com")).length, 1);
});

test("outlook paging only follows Graph links", async () => {
  mockFetch([msToken, ["$skiptoken=abc", () => json({ value: [graphMsg("z", "2026-01-01T00:00:00Z")] })]]);
  const account = getAccount("work");
  const page = await outlook.search(account, {
    maxResults: 10,
    pageToken: "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc",
  });
  assert.deepEqual(page.messages.map((m) => m.id), ["z"]);
  await assert.rejects(
    outlook.search(account, { maxResults: 10, pageToken: "https://evil.example/steal" }),
    /Invalid page_token/,
  );
});

test("outlook read returns the text body and attachments; thread is sorted oldest first", async () => {
  mockFetch([
    msToken,
    [
      "/messages/a/attachments",
      () => json({ value: [{ id: "att-1", name: "contract.docx", contentType: "application/msword", size: 2048 }] }),
    ],
    [
      "/messages/a?$select=",
      () =>
        json(
          graphMsg("a", "2026-01-15T10:00:00Z", {
            hasAttachments: true,
            ccRecipients: [{ emailAddress: { name: "Bob", address: "bob@contoso.example" } }],
            body: { contentType: "text", content: "Hello\r\nWorld\r\n" },
          }),
        ),
    ],
    [
      "conversationId",
      () =>
        json({
          value: [
            graphMsg("late", "2026-01-16T10:00:00Z", { body: { contentType: "html", content: "<p>Second</p>" } }),
            graphMsg("early", "2026-01-15T10:00:00Z", { body: { contentType: "text", content: "First" } }),
          ],
        }),
    ],
  ]);
  const account = getAccount("work");
  const m = await outlook.read(account, "a", 20000);
  assert.equal(m.body, "Hello\nWorld");
  assert.equal(m.cc, "Bob <bob@contoso.example>");
  assert.deepEqual(m.attachments, [{ id: "att-1", filename: "contract.docx", mimeType: "application/msword", size: 2048 }]);

  const thread = await outlook.thread(account, "conv'1", 20000);
  assert.deepEqual(thread.map((t) => [t.id, t.body]), [["early", "First"], ["late", "Second"]]);
});

test("outlook attachment download decodes file attachments and rejects other kinds", async () => {
  mockFetch([
    msToken,
    ["/attachments/file", () => json({ "@odata.type": "#microsoft.graph.fileAttachment", name: "a.txt", contentType: "text/plain", contentBytes: Buffer.from("hi").toString("base64") })],
    ["/attachments/item", () => json({ "@odata.type": "#microsoft.graph.itemAttachment", name: "Meeting" })],
  ]);
  const account = getAccount("work");
  assert.equal((await outlook.attachment(account, "m", "file")).data.toString(), "hi");
  await assert.rejects(outlook.attachment(account, "m", "item"), /not a file attachment/);
});

test("a 401 triggers one token refresh and retry; throttling is retried", async () => {
  let tokens = 0;
  let apiCalls = 0;
  mockFetch([
    ["oauth2.googleapis.com/token", () => json({ access_token: `g-access-${++tokens}`, expires_in: 3600 })],
    [
      /\/messages\?/,
      ({ init }) => {
        apiCalls++;
        if (init.headers.Authorization === "Bearer g-access-1") return json({ error: { message: "expired" } }, 401);
        if (apiCalls === 2) return json({ error: { message: "slow down" } }, 429, { "retry-after": "0.01" });
        return json({ messages: [] });
      },
    ],
  ]);
  const result = await gmail.search(getAccount("personal"), { maxResults: 1 });
  assert.deepEqual(result.messages, []);
  assert.equal(tokens, 2);
  assert.equal(apiCalls, 3);
});

test("a revoked sign-in produces a re-authorise instruction; API errors carry the detail", async () => {
  mockFetch([
    ["oauth2.googleapis.com/token", () => json({ error: "invalid_grant", error_description: "Token has been expired or revoked." }, 400)],
  ]);
  await assert.rejects(
    gmail.search(getAccount("personal"), { maxResults: 1 }),
    /needs to be authorised again\. Run: mailhub-mcp add gmail personal --force/,
  );

  resetTokenCache();
  mockFetch([msToken, ["graph.microsoft.com", () => json({ error: { code: "ErrorAccessDenied", message: "Access is denied." } }, 403)]]);
  await assert.rejects(outlook.read(getAccount("work"), "x", 1000), /outlook API error 403 for "work": Access is denied\./);
});
