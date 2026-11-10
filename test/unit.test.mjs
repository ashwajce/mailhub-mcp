import assert from "node:assert/strict";
import fs from "node:fs";
import { test } from "node:test";
import { accountsPath } from "../dist/config.js";
import { buildGmailQuery, collectParts } from "../dist/providers/gmail.js";
import { buildFilter, buildKql, buildSearchUrl } from "../dist/providers/outlook.js";
import { getAccount, listAccounts, removeAccount, saveAccount, updateRefreshToken, validateAlias } from "../dist/store.js";
import { decodeBytes, htmlToText, parseDate, safeFilename } from "../dist/util.js";
import { gmailAccount, makeHome, outlookAccount } from "./helpers.mjs";

test("gmail query combines structured fields", () => {
  const q = buildGmailQuery({
    query: "invoice",
    from: "billing@acme.com",
    subject: "October statement",
    after: new Date("2026-01-01T00:00:00Z"),
    before: new Date("2026-02-01T00:00:00Z"),
    unreadOnly: true,
    hasAttachment: true,
    maxResults: 10,
  });
  assert.equal(
    q,
    "invoice from:(billing@acme.com) subject:(October statement) after:1767225600 before:1769904000 is:unread has:attachment",
  );
  assert.equal(buildGmailQuery({ maxResults: 5 }), "");
});

test("outlook uses KQL search when there are text terms", () => {
  const params = {
    query: "invoice",
    from: "Jane Doe",
    after: new Date("2026-01-10T00:00:00Z"),
    before: new Date("2026-01-20T00:00:00Z"),
    hasAttachment: true,
    maxResults: 10,
  };
  // Date window is widened by a day on each side; exact bounds are applied after the fetch.
  assert.equal(buildKql(params), 'invoice from:"Jane Doe" received>=2026-01-09 received<=2026-01-21');
  const url = new URL(buildSearchUrl(outlookAccount, params));
  assert.equal(url.pathname, "/v1.0/me/messages");
  assert.equal(url.searchParams.get("$search"), '"invoice from:\\"Jane Doe\\" received>=2026-01-09 received<=2026-01-21"');
  assert.equal(url.searchParams.get("$filter"), null);
  assert.equal(url.searchParams.get("$top"), "10");
});

test("outlook uses a date-ordered filter when there are no text terms", () => {
  const params = { after: new Date("2026-01-10T00:00:00Z"), unreadOnly: true, maxResults: 3 };
  assert.equal(buildKql(params), "");
  assert.equal(buildFilter(params), "receivedDateTime ge 2026-01-10T00:00:00.000Z and isRead eq false");
  const url = new URL(buildSearchUrl({ ...outlookAccount, sharedMailbox: "billing@contoso.example" }, params));
  assert.equal(url.pathname, "/v1.0/users/billing%40contoso.example/messages");
  assert.equal(url.searchParams.get("$orderby"), "receivedDateTime desc");
  assert.equal(url.searchParams.get("$search"), null);
});

test("html is reduced to readable text", () => {
  const html =
    "<html><head><style>p{color:red}</style></head><body><p>Hello&nbsp;<b>Alex</b> &amp; team</p>" +
    "<div>Line two<br>Line three &#8364;5 &#x41;</div><script>alert(1)</script></body></html>";
  assert.equal(htmlToText(html), "Hello Alex & team\nLine two\nLine three \u20ac5 A");
});

test("bytes are decoded with the declared charset", () => {
  assert.equal(decodeBytes(Buffer.from([0x63, 0x61, 0x66, 0xe9]), 'text/plain; charset="ISO-8859-1"'), "caf\u00e9");
  assert.equal(decodeBytes(Buffer.from("caf\u00e9", "utf8"), "text/plain"), "caf\u00e9");
  assert.equal(decodeBytes(Buffer.from("ok", "utf8"), "text/plain; charset=not-a-charset"), "ok");
});

test("dates accept day or ISO format and reject junk", () => {
  assert.equal(parseDate("2026-03-05", "after").toISOString(), "2026-03-05T00:00:00.000Z");
  assert.equal(parseDate("2026-03-05T10:30:00+01:00", "after").toISOString(), "2026-03-05T09:30:00.000Z");
  assert.equal(parseDate(undefined, "after"), undefined);
  assert.throws(() => parseDate("last week", "after"), /YYYY-MM-DD/);
});

test("attachment names cannot escape the download folder", () => {
  assert.equal(safeFilename("../../etc/passwd"), "passwd");
  assert.equal(safeFilename("C:\\Users\\x\\report Q3 (final).pdf"), "report Q3 (final).pdf");
  assert.equal(safeFilename(".bashrc"), "bashrc");
  assert.equal(safeFilename("a:b*c?.txt"), "a_b_c_.txt");
  assert.equal(safeFilename(""), "attachment");
});

test("gmail MIME tree is split into body parts and attachments", () => {
  const leaves = collectParts({
    mimeType: "multipart/mixed",
    parts: [
      {
        mimeType: "multipart/alternative",
        parts: [
          { partId: "0.0", mimeType: "text/plain", filename: "", body: { data: "aGk" } },
          { partId: "0.1", mimeType: "text/html", filename: "", body: { data: "PGI-aGk8L2I-" } },
        ],
      },
      { partId: "1", mimeType: "application/pdf", filename: "invoice.pdf", body: { attachmentId: "x", size: 1234 } },
    ],
  });
  assert.deepEqual(leaves.plain.map((p) => p.partId), ["0.0"]);
  assert.deepEqual(leaves.html.map((p) => p.partId), ["0.1"]);
  assert.deepEqual(leaves.attachments.map((p) => p.filename), ["invoice.pdf"]);
});

test("account store round-trips, keeps tokens private and validates aliases", () => {
  makeHome();
  assert.deepEqual(listAccounts(), []);
  saveAccount(outlookAccount);
  saveAccount(gmailAccount);
  assert.deepEqual(listAccounts().map((a) => a.alias), ["personal", "work"]);
  if (process.platform !== "win32") {
    assert.equal(fs.statSync(accountsPath()).mode & 0o777, 0o600);
  }
  updateRefreshToken("work", "ms-refresh-2");
  assert.equal(getAccount("work").refreshToken, "ms-refresh-2");
  assert.equal(getAccount("personal").refreshToken, "g-refresh");
  assert.throws(() => getAccount("nope"), /Known aliases: personal, work/);
  assert.equal(removeAccount("work"), true);
  assert.equal(removeAccount("work"), false);
  assert.throws(() => validateAlias("all"), /Invalid alias/);
  assert.throws(() => validateAlias("Has Space"), /Invalid alias/);
  validateAlias("work-nl_2");
});
