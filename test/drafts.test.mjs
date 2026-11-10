import assert from "node:assert/strict";
import { beforeEach, test } from "node:test";
import { canDraft, resetTokenCache, scopesFor } from "../dist/auth.js";
import { buildMime, encodeHeaderText, formatAddress, parseAddress, parseAddresses, splitAddressList } from "../dist/mime.js";
import { gmail } from "../dist/providers/gmail.js";
import { composeReplyBody, outlook } from "../dist/providers/outlook.js";
import { getAccount, saveAccount } from "../dist/store.js";
import { gmailAccount, gmailDraftAccount, json, makeHome, mockFetch, outlookAccount, outlookDraftAccount } from "./helpers.mjs";

const googleToken = ["oauth2.googleapis.com/token", () => json({ access_token: "g-access", expires_in: 3600 })];
const msToken = ["login.microsoftonline.com", () => json({ access_token: "ms-access", expires_in: 3600 })];

beforeEach(() => {
  makeHome();
  resetTokenCache();
  for (const a of [gmailAccount, gmailDraftAccount, outlookAccount, outlookDraftAccount]) saveAccount(a);
});

/** Split a raw MIME message into unfolded headers and the decoded body. */
function parseMime(raw) {
  const [head, body] = raw.split("\r\n\r\n");
  const headers = {};
  for (const line of head.replace(/\r\n[ \t]+/g, " ").split("\r\n")) {
    const i = line.indexOf(":");
    headers[line.slice(0, i).toLowerCase()] = line.slice(i + 1).trim();
  }
  return { headers, body: Buffer.from(body.replace(/\r\n/g, ""), "base64").toString("utf8"), head };
}

test("address lists are split and parsed safely", () => {
  assert.deepEqual(splitAddressList('"Doe, Jane" <jane@x.com>, bob@y.com ,  Al <al@z.org>'), [
    '"Doe, Jane" <jane@x.com>',
    "bob@y.com",
    "Al <al@z.org>",
  ]);
  assert.deepEqual(parseAddress('"Doe, Jane" <jane@x.com>'), { name: "Doe, Jane", email: "jane@x.com" });
  assert.deepEqual(parseAddress(" bob@y.com "), { email: "bob@y.com" });
  assert.deepEqual(parseAddresses(["a@x.com, B <b@x.com>", "c@x.com"]).map((a) => a.email), ["a@x.com", "b@x.com", "c@x.com"]);
  assert.throws(() => parseAddress("not-an-address"), /not a valid email address/);
  assert.throws(() => parseAddress("a@x.com\r\nBcc: evil@x.com"), /line breaks/);
  assert.throws(() => parseAddress("Evil <a@x.com> <b@y.com>"), /not a valid email address/);
});

test("headers are encoded when needed and cannot be injected", () => {
  assert.equal(encodeHeaderText("Plain subject"), "Plain subject");
  const encoded = encodeHeaderText("Offerte voor café €500");
  assert.match(encoded, /^=\?UTF-8\?B\?[A-Za-z0-9+/=]+\?=$/);
  const long = encodeHeaderText("é".repeat(120));
  for (const word of long.split("\r\n ")) assert.ok(word.length <= 75, `encoded word too long: ${word.length}`);
  assert.throws(() => encodeHeaderText("Hi\r\nBcc: evil@x.com"), /line breaks/);
  assert.equal(formatAddress({ name: "Doe, Jane", email: "j@x.com" }), '"Doe, Jane" <j@x.com>');
  assert.equal(formatAddress({ name: "Jane Doe", email: "j@x.com" }), "Jane Doe <j@x.com>");
  assert.match(formatAddress({ name: "José", email: "j@x.com" }), /^=\?UTF-8\?B\?.+\?= <j@x\.com>$/);
});

test("a MIME message round-trips subject, recipients and a unicode body", () => {
  const body = "Hallo,\n\nDe prijs is €500.\n" + "x".repeat(200);
  const raw = buildMime({
    to: [{ name: "Jane Doe", email: "jane@x.com" }, { email: "bob@y.com" }],
    cc: [{ email: "cc@y.com" }],
    subject: "Offerte café",
    body,
    inReplyTo: "<abc@mail.example>",
    references: "<first@mail.example> <abc@mail.example>",
  });
  const { headers, body: decoded, head } = parseMime(raw);
  assert.equal(headers.to, "Jane Doe <jane@x.com>, bob@y.com");
  assert.equal(headers.cc, "cc@y.com");
  assert.equal(headers.bcc, undefined);
  assert.equal(headers.from, undefined);
  assert.equal(headers["in-reply-to"], "<abc@mail.example>");
  assert.equal(headers["content-type"], "text/plain; charset=UTF-8");
  assert.equal(Buffer.from(/=\?UTF-8\?B\?(.+)\?=/.exec(headers.subject)[1], "base64").toString("utf8"), "Offerte café");
  assert.equal(decoded, body.replace(/\n/g, "\r\n"));
  for (const line of raw.split("\r\n")) assert.ok(line.length <= 78, "line too long");
  assert.doesNotMatch(head, /[^\x00-\x7f]/);
});

test("draft scopes are opt-in", () => {
  assert.deepEqual(scopesFor("gmail"), ["https://www.googleapis.com/auth/gmail.readonly"]);
  assert.ok(scopesFor("gmail", { drafts: true }).includes("https://www.googleapis.com/auth/gmail.compose"));
  assert.deepEqual(scopesFor("outlook"), ["offline_access", "User.Read", "Mail.Read"]);
  assert.deepEqual(scopesFor("outlook", { drafts: true }), ["offline_access", "User.Read", "Mail.Read", "Mail.ReadWrite"]);
  assert.deepEqual(scopesFor("outlook", { drafts: true, shared: true }), [
    "offline_access",
    "User.Read",
    "Mail.Read",
    "Mail.Read.Shared",
    "Mail.ReadWrite.Shared",
  ]);
  // No scope in any combination allows sending from Outlook.
  assert.equal(scopesFor("outlook", { drafts: true, shared: true }).some((s) => /send/i.test(s)), false);
  assert.equal(canDraft(gmailAccount), false);
  assert.equal(canDraft(gmailDraftAccount), true);
  assert.equal(canDraft(outlookDraftAccount), true);
  // A shared mailbox needs the shared write scope, not the personal one.
  assert.equal(canDraft({ ...outlookDraftAccount, sharedMailbox: "billing@contoso.example" }), false);
});

test("read-only accounts refuse drafts before any network call", async () => {
  const calls = mockFetch([]);
  const input = { to: ["jane@x.com"], cc: [], bcc: [], subject: "Hi", body: "Hello" };
  await assert.rejects(gmail.createDraft(getAccount("personal"), input), /connected read-only.*add gmail personal --drafts --force/);
  await assert.rejects(outlook.createDraft(getAccount("work"), input), /connected read-only.*add outlook work --drafts --force/);
  assert.equal(calls.length, 0);
});

test("gmail creates a new draft through the drafts endpoint only", async () => {
  const calls = mockFetch([googleToken, [/\/drafts$/, () => json({ id: "d1", message: { id: "m1", threadId: "t1" } })]]);
  const result = await gmail.createDraft(getAccount("writer"), {
    to: ["Jane Doe <jane@x.com>"],
    cc: [],
    bcc: ["audit@x.com"],
    subject: "Hello",
    body: "Line 1\nLine 2",
  });
  assert.deepEqual(result, { draftId: "d1", messageId: "m1", threadId: "t1", subject: "Hello", to: "Jane Doe <jane@x.com>", cc: "" });

  const post = calls.find((c) => /\/drafts$/.test(c.url));
  assert.equal(post.init.method, "POST");
  assert.equal(calls.some((c) => /\/send/.test(c.url)), false);
  const payload = JSON.parse(post.body);
  assert.equal(payload.message.threadId, undefined);
  const { headers, body } = parseMime(Buffer.from(payload.message.raw, "base64url").toString("utf8"));
  assert.equal(headers.to, "Jane Doe <jane@x.com>");
  assert.equal(headers.bcc, "audit@x.com");
  assert.equal(headers.subject, "Hello");
  assert.equal(body, "Line 1\r\nLine 2");
});

test("gmail reply-all draft threads correctly and addresses everyone except me", async () => {
  const calls = mockFetch([
    googleToken,
    [/\/drafts$/, () => json({ id: "d2", message: { id: "m2", threadId: "thread-9" } })],
    [
      /\/messages\/orig\?format=metadata/,
      () =>
        json({
          id: "orig",
          threadId: "thread-9",
          payload: {
            headers: [
              { name: "Message-Id", value: "<orig@mail.example>" },
              { name: "References", value: "<root@mail.example>" },
              { name: "Subject", value: "Budget 2027" },
              { name: "From", value: '"Doe, Jane" <jane@x.com>' },
              { name: "To", value: "writer@example.com, Bob <bob@y.com>" },
              { name: "Cc", value: "carol@z.org, JANE@x.com, undisclosed-recipients:;" },
            ],
          },
        }),
    ],
  ]);
  const result = await gmail.createDraft(getAccount("writer"), {
    to: [],
    cc: [],
    bcc: [],
    subject: "ignored for replies",
    body: "Agreed.",
    replyToMessageId: "orig",
    replyAll: true,
  });
  assert.equal(result.subject, "Re: Budget 2027");
  assert.equal(result.to, '"Doe, Jane" <jane@x.com>');
  assert.equal(result.cc, "Bob <bob@y.com>, carol@z.org");
  assert.equal(result.threadId, "thread-9");

  const payload = JSON.parse(calls.find((c) => /\/drafts$/.test(c.url)).body);
  assert.equal(payload.message.threadId, "thread-9");
  const { headers } = parseMime(Buffer.from(payload.message.raw, "base64url").toString("utf8"));
  assert.equal(headers["in-reply-to"], "<orig@mail.example>");
  assert.equal(headers.references, "<root@mail.example> <orig@mail.example>");
  assert.equal(headers.subject, "Re: Budget 2027");
});

test("gmail plain reply answers only the sender, and a reply to my own mail goes to its recipients", async () => {
  const original = (from, to) => () =>
    json({
      id: "o",
      threadId: "t",
      payload: {
        headers: [
          { name: "Message-ID", value: "<o@mail.example>" },
          { name: "Subject", value: "RE: Status" },
          { name: "From", value: from },
          { name: "To", value: to },
          { name: "Cc", value: "carol@z.org" },
        ],
      },
    });
  const input = { to: [], cc: [], bcc: [], body: "Ok", replyToMessageId: "o", replyAll: false };

  mockFetch([googleToken, [/\/drafts$/, () => json({ id: "d" })], [/\/messages\/o\?/, original("jane@x.com", "writer@example.com")]]);
  const plain = await gmail.createDraft(getAccount("writer"), input);
  assert.deepEqual([plain.to, plain.cc, plain.subject, plain.threadId], ["jane@x.com", "", "RE: Status", "t"]);

  mockFetch([googleToken, [/\/drafts$/, () => json({ id: "d" })], [/\/messages\/o\?/, original("Writer <writer@example.com>", "bob@y.com")]]);
  const followUp = await gmail.createDraft(getAccount("writer"), input);
  assert.equal(followUp.to, "bob@y.com");
});

test("outlook creates a new draft as a message in Drafts, without sending", async () => {
  const calls = mockFetch([
    msToken,
    [
      /\/me\/messages$/,
      ({ body }) => {
        const sent = JSON.parse(body);
        return json({ id: "od1", conversationId: "oc1", subject: sent.subject, toRecipients: sent.toRecipients, webLink: "https://outlook.office.com/mail/od1" });
      },
    ],
  ]);
  const result = await outlook.createDraft(getAccount("office"), {
    to: ["Jane Doe <jane@x.com>", "bob@y.com"],
    cc: [],
    bcc: [],
    subject: "Hello",
    body: "Text body",
  });
  assert.deepEqual(result, {
    draftId: "od1",
    messageId: "od1",
    threadId: "oc1",
    subject: "Hello",
    to: "Jane Doe <jane@x.com>, bob@y.com",
    cc: "",
    webLink: "https://outlook.office.com/mail/od1",
  });
  const post = calls[1];
  assert.equal(post.init.method, "POST");
  assert.deepEqual(JSON.parse(post.body), {
    subject: "Hello",
    body: { contentType: "text", content: "Text body" },
    toRecipients: [{ emailAddress: { address: "jane@x.com", name: "Jane Doe" } }, { emailAddress: { address: "bob@y.com" } }],
    ccRecipients: [],
    bccRecipients: [],
  });
  assert.equal(calls.some((c) => /\/send/i.test(c.url)), false);
});

test("outlook reply draft keeps the quoted original and puts my text on top", async () => {
  const calls = mockFetch([
    msToken,
    [
      /\/messages\/orig\/createReplyAll$/,
      () =>
        json({
          id: "reply-1",
          conversationId: "conv-7",
          subject: "RE: Budget",
          toRecipients: [{ emailAddress: { name: "Jane", address: "jane@x.com" } }],
          ccRecipients: [{ emailAddress: { address: "bob@y.com" } }],
          body: { contentType: "html", content: '<html><body lang="en"><div>quoted original</div></body></html>' },
        }),
    ],
    [/\/messages\/reply-1$/, ({ body }) => json({ id: "reply-1", body: JSON.parse(body).body })],
  ]);
  const result = await outlook.createDraft(getAccount("office"), {
    to: [],
    cc: [],
    bcc: [],
    body: "Agreed <ok> & done.\nSecond line",
    replyToMessageId: "orig",
    replyAll: true,
  });
  assert.deepEqual(
    [result.draftId, result.threadId, result.subject, result.to, result.cc],
    ["reply-1", "conv-7", "RE: Budget", "Jane <jane@x.com>", "bob@y.com"],
  );
  assert.deepEqual(JSON.parse(calls[1].body), {});
  const patch = calls[2];
  assert.equal(patch.init.method, "PATCH");
  assert.equal(
    JSON.parse(patch.body).body.content,
    '<html><body lang="en"><div>Agreed &lt;ok&gt; &amp; done.<br>Second line</div><br><div>quoted original</div></body></html>',
  );
});

test("outlook reply recipients can be overridden; plain-text originals stay plain text", async () => {
  const calls = mockFetch([
    msToken,
    [/\/createReply$/, () => json({ id: "r2", body: { contentType: "text", content: "> old" } })],
    [/\/messages\/r2$/, () => json({ id: "r2" })],
  ]);
  await outlook.createDraft(getAccount("office"), { to: ["new@x.com"], cc: [], bcc: [], body: "Hi", replyToMessageId: "m" });
  assert.deepEqual(JSON.parse(calls[1].body), { message: { toRecipients: [{ emailAddress: { address: "new@x.com" } }] } });
  assert.deepEqual(JSON.parse(calls[2].body).body, { contentType: "text", content: "Hi\n\n> old" });
  assert.deepEqual(composeReplyBody("Hi", undefined), { contentType: "text", content: "Hi" });
  assert.deepEqual(composeReplyBody("Hi", { contentType: "html", content: "<p>q</p>" }), { contentType: "html", content: "<div>Hi</div><br><p>q</p>" });
});

test("a failed write is not repeated, so a draft is never created twice", async () => {
  let posts = 0;
  mockFetch([
    googleToken,
    [/\/drafts$/, () => (posts++, json({ error: { message: "backend error" } }, 503))],
  ]);
  await assert.rejects(
    gmail.createDraft(getAccount("writer"), { to: ["a@x.com"], cc: [], bcc: [], subject: "s", body: "b" }),
    /gmail API error 503/,
  );
  assert.equal(posts, 1);
});
