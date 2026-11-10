import { apiJson, requireDraftAccess } from "../auth.js";
import { Address, buildMime, formatAddress, parseAddress, parseAddresses, splitAddressList } from "../mime.js";
import {
  Account,
  AttachmentData,
  AttachmentInfo,
  DraftInput,
  DraftResult,
  MailhubError,
  MessageFull,
  MessageSummary,
  Provider,
  SearchParams,
  SearchResult,
} from "../types.js";
import { decodeBytes, htmlToText, mapLimit, truncate } from "../util.js";

const API = "https://gmail.googleapis.com/gmail/v1/users/me";

interface GmailHeader {
  name: string;
  value: string;
}

interface GmailPart {
  partId?: string;
  mimeType?: string;
  filename?: string;
  headers?: GmailHeader[];
  body?: { data?: string; attachmentId?: string; size?: number };
  parts?: GmailPart[];
}

interface GmailMessage {
  id: string;
  threadId: string;
  labelIds?: string[];
  snippet?: string;
  internalDate?: string;
  payload?: GmailPart;
}

/** Translate the structured search fields into Gmail's search syntax. */
export function buildGmailQuery(p: SearchParams): string {
  const parts: string[] = [];
  if (p.query?.trim()) parts.push(p.query.trim());
  if (p.from) parts.push(`from:(${p.from})`);
  if (p.to) parts.push(`to:(${p.to})`);
  if (p.subject) parts.push(`subject:(${p.subject})`);
  // Epoch seconds avoid Gmail's date-only handling, which uses the account's time zone.
  if (p.after) parts.push(`after:${Math.floor(p.after.getTime() / 1000)}`);
  if (p.before) parts.push(`before:${Math.floor(p.before.getTime() / 1000)}`);
  if (p.unreadOnly) parts.push("is:unread");
  if (p.hasAttachment) parts.push("has:attachment");
  return parts.join(" ");
}

function header(part: GmailPart | undefined, name: string): string {
  const lower = name.toLowerCase();
  return part?.headers?.find((h) => h.name.toLowerCase() === lower)?.value ?? "";
}

function toSummary(account: Account, m: GmailMessage): MessageSummary {
  const millis = Number(m.internalDate);
  return {
    id: m.id,
    threadId: m.threadId,
    account: account.alias,
    mailbox: account.email,
    from: header(m.payload, "From"),
    to: header(m.payload, "To"),
    subject: header(m.payload, "Subject"),
    date: Number.isFinite(millis) && millis > 0 ? new Date(millis).toISOString() : header(m.payload, "Date"),
    snippet: htmlToText(m.snippet ?? ""),
    unread: m.labelIds?.includes("UNREAD") ?? false,
  };
}

interface Leaves {
  plain: GmailPart[];
  html: GmailPart[];
  attachments: GmailPart[];
}

/** Walk the MIME tree and sort the leaf parts into body text and attachments. */
export function collectParts(part: GmailPart | undefined, acc: Leaves = { plain: [], html: [], attachments: [] }): Leaves {
  if (!part) return acc;
  if (part.parts?.length) {
    for (const child of part.parts) collectParts(child, acc);
    return acc;
  }
  const mime = (part.mimeType ?? "").toLowerCase();
  if (part.filename) acc.attachments.push(part);
  else if (mime === "text/plain") acc.plain.push(part);
  else if (mime === "text/html") acc.html.push(part);
  return acc;
}

async function partBytes(account: Account, messageId: string, part: GmailPart): Promise<Buffer> {
  let data = part.body?.data;
  if (!data && part.body?.attachmentId) {
    // Large parts are stored out of line and fetched separately.
    const att = await apiJson<{ data?: string }>(
      account,
      `${API}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(part.body.attachmentId)}`,
    );
    data = att.data;
  }
  return Buffer.from(data ?? "", "base64url");
}

async function partText(account: Account, messageId: string, part: GmailPart): Promise<string> {
  return decodeBytes(await partBytes(account, messageId, part), header(part, "Content-Type"));
}

async function toFull(account: Account, m: GmailMessage, maxBodyChars: number): Promise<MessageFull> {
  const leaves = collectParts(m.payload);
  let body = "";
  if (leaves.plain.length) {
    body = (await Promise.all(leaves.plain.map((p) => partText(account, m.id, p)))).join("\n").trim();
  }
  if (!body && leaves.html.length) {
    body = htmlToText((await Promise.all(leaves.html.map((p) => partText(account, m.id, p)))).join("\n"));
  }
  const cut = truncate(body, maxBodyChars);
  const attachments: AttachmentInfo[] = leaves.attachments.map((p) => ({
    // Gmail's own attachment ids change between requests; the MIME part id is stable.
    id: p.partId ?? "",
    filename: p.filename ?? "",
    mimeType: p.mimeType ?? "application/octet-stream",
    size: p.body?.size ?? 0,
  }));
  return {
    ...toSummary(account, m),
    hasAttachments: attachments.length > 0,
    cc: header(m.payload, "Cc"),
    body: cut.text,
    bodyTruncated: cut.truncated,
    attachments,
  };
}

/** Addresses from a received header. Entries that cannot be parsed are skipped rather than failing the reply. */
function headerAddresses(value: string): Address[] {
  return splitAddressList(value).flatMap((item) => {
    try {
      return [parseAddress(item)];
    } catch {
      return [];
    }
  });
}

const REPLY_HEADERS = ["Message-ID", "References", "Subject", "From", "Reply-To", "To", "Cc"]
  .map((h) => `metadataHeaders=${h}`)
  .join("&");

export const gmail: Provider = {
  async search(account: Account, params: SearchParams): Promise<SearchResult> {
    const qs = new URLSearchParams({ maxResults: String(params.maxResults) });
    const q = buildGmailQuery(params);
    if (q) qs.set("q", q);
    if (params.pageToken) qs.set("pageToken", params.pageToken);
    const list = await apiJson<{ messages?: { id: string }[]; nextPageToken?: string }>(
      account,
      `${API}/messages?${qs}`,
    );
    const ids = (list.messages ?? []).map((m) => m.id);
    const meta = "format=metadata&metadataHeaders=From&metadataHeaders=To&metadataHeaders=Subject&metadataHeaders=Date";
    const messages = await mapLimit(ids, 8, async (id) =>
      toSummary(account, await apiJson<GmailMessage>(account, `${API}/messages/${encodeURIComponent(id)}?${meta}`)),
    );
    return { messages, nextPageToken: list.nextPageToken };
  },

  async read(account: Account, id: string, maxBodyChars: number): Promise<MessageFull> {
    const m = await apiJson<GmailMessage>(account, `${API}/messages/${encodeURIComponent(id)}?format=full`);
    return toFull(account, m, maxBodyChars);
  },

  async thread(account: Account, threadId: string, maxBodyChars: number): Promise<MessageFull[]> {
    const t = await apiJson<{ messages?: GmailMessage[] }>(
      account,
      `${API}/threads/${encodeURIComponent(threadId)}?format=full`,
    );
    return Promise.all((t.messages ?? []).map((m) => toFull(account, m, maxBodyChars)));
  },

  async attachment(account: Account, messageId: string, attachmentId: string): Promise<AttachmentData> {
    const m = await apiJson<GmailMessage>(account, `${API}/messages/${encodeURIComponent(messageId)}?format=full`);
    const part = collectParts(m.payload).attachments.find((p) => p.partId === attachmentId);
    if (!part) {
      throw new MailhubError(
        `Attachment "${attachmentId}" not found on message ${messageId}. Use the attachment ids returned by read_email.`,
      );
    }
    return {
      filename: part.filename ?? "attachment",
      mimeType: part.mimeType ?? "application/octet-stream",
      data: await partBytes(account, messageId, part),
    };
  },

  async createDraft(account: Account, draft: DraftInput): Promise<DraftResult> {
    requireDraftAccess(account);
    let to = parseAddresses(draft.to);
    let cc = parseAddresses(draft.cc);
    const bcc = parseAddresses(draft.bcc);
    let subject = draft.subject ?? "";
    let threadId: string | undefined;
    let inReplyTo: string | undefined;
    let references: string | undefined;

    if (draft.replyToMessageId) {
      const original = await apiJson<GmailMessage>(
        account,
        `${API}/messages/${encodeURIComponent(draft.replyToMessageId)}?format=metadata&${REPLY_HEADERS}`,
      );
      const h = (name: string) => header(original.payload, name).replace(/\s+/g, " ").trim();
      threadId = original.threadId;
      // These two headers are what keep the reply inside the conversation for every recipient.
      inReplyTo = h("Message-ID") || undefined;
      references = [h("References"), h("Message-ID")].filter(Boolean).join(" ") || undefined;
      subject = /^re:/i.test(h("Subject")) ? h("Subject") : `Re: ${h("Subject")}`;

      const self = account.email.toLowerCase();
      const from = headerAddresses(h("Reply-To") || h("From"));
      const originalTo = headerAddresses(h("To"));
      const sentByMe = from.some((a) => a.email.toLowerCase() === self);
      if (!to.length) to = sentByMe ? originalTo : from;
      if (draft.replyAll && !cc.length) {
        const seen = new Set([self, ...to.map((a) => a.email.toLowerCase())]);
        cc = [...originalTo, ...headerAddresses(h("Cc"))].filter((a) => {
          const key = a.email.toLowerCase();
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        });
      }
    }

    const raw = buildMime({ to, cc, bcc, subject, body: draft.body, inReplyTo, references });
    const created = await apiJson<{ id: string; message?: { id?: string; threadId?: string } }>(
      account,
      `${API}/drafts`,
      {},
      { message: { raw: Buffer.from(raw, "utf8").toString("base64url"), ...(threadId ? { threadId } : {}) } },
    );
    return {
      draftId: created.id,
      messageId: created.message?.id ?? "",
      threadId: created.message?.threadId ?? threadId ?? "",
      subject,
      to: to.map(formatAddress).join(", "),
      cc: cc.map(formatAddress).join(", "),
    };
  },
};
