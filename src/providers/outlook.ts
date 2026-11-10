import { apiJson } from "../auth.js";
import {
  Account,
  AttachmentData,
  AttachmentInfo,
  MailhubError,
  MessageFull,
  MessageSummary,
  Provider,
  SearchParams,
  SearchResult,
} from "../types.js";
import { htmlToText, truncate } from "../util.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

// Plain-text bodies, and ids that stay valid when a message moves between folders.
const PREFER = { Prefer: 'outlook.body-content-type="text", IdType="ImmutableId"' };

const SUMMARY_FIELDS =
  "id,conversationId,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments";
const FULL_FIELDS = `${SUMMARY_FIELDS},ccRecipients,body`;

interface GraphRecipient {
  emailAddress?: { name?: string; address?: string };
}

interface GraphMessage {
  id: string;
  conversationId?: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  body?: { contentType?: string; content?: string };
}

interface GraphList<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

interface GraphAttachment {
  "@odata.type"?: string;
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  contentBytes?: string;
}

const DAY_MS = 86_400_000;

function mailboxBase(account: Account): string {
  return account.sharedMailbox ? `${GRAPH}/users/${encodeURIComponent(account.sharedMailbox)}` : `${GRAPH}/me`;
}

function ymd(date: Date): string {
  return date.toISOString().slice(0, 10);
}

function kqlValue(value: string): string {
  const cleaned = value.replace(/"/g, "").trim();
  return /\s/.test(cleaned) ? `"${cleaned}"` : cleaned;
}

/** Translate the text search fields into Outlook's KQL search syntax. Empty when there are none. */
export function buildKql(p: SearchParams): string {
  const parts: string[] = [];
  if (p.query?.trim()) parts.push(p.query.trim());
  if (p.from) parts.push(`from:${kqlValue(p.from)}`);
  if (p.to) parts.push(`to:${kqlValue(p.to)}`);
  if (p.subject) parts.push(`subject:${kqlValue(p.subject)}`);
  if (!parts.length) return "";
  // KQL dates are day-granular and follow the mailbox time zone, so the window is
  // widened by a day each side; exact bounds are applied to the results afterwards.
  if (p.after) parts.push(`received>=${ymd(new Date(p.after.getTime() - DAY_MS))}`);
  if (p.before) parts.push(`received<=${ymd(new Date(p.before.getTime() + DAY_MS))}`);
  return parts.join(" ");
}

/** OData filter used when there are no text terms, so results come back newest first. */
export function buildFilter(p: SearchParams): string {
  // receivedDateTime must lead the filter because it is also the sort key.
  const parts = [`receivedDateTime ge ${(p.after ?? new Date(0)).toISOString()}`];
  if (p.before) parts.push(`receivedDateTime lt ${p.before.toISOString()}`);
  if (p.unreadOnly) parts.push("isRead eq false");
  if (p.hasAttachment) parts.push("hasAttachments eq true");
  return parts.join(" and ");
}

/** First-page URL for a search. Exported for tests. */
export function buildSearchUrl(account: Account, p: SearchParams): string {
  const common = `$top=${p.maxResults}&$select=${SUMMARY_FIELDS}`;
  const kql = buildKql(p);
  if (kql) {
    const quoted = `"${kql.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
    return `${mailboxBase(account)}/messages?$search=${encodeURIComponent(quoted)}&${common}`;
  }
  return (
    `${mailboxBase(account)}/messages?$filter=${encodeURIComponent(buildFilter(p))}` +
    `&$orderby=${encodeURIComponent("receivedDateTime desc")}&${common}`
  );
}

function formatRecipient(r: GraphRecipient | undefined): string {
  const { name, address } = r?.emailAddress ?? {};
  if (name && address && name !== address) return `${name} <${address}>`;
  return address ?? name ?? "";
}

function toSummary(account: Account, m: GraphMessage): MessageSummary {
  return {
    id: m.id,
    threadId: m.conversationId ?? "",
    account: account.alias,
    mailbox: account.email,
    from: formatRecipient(m.from),
    to: (m.toRecipients ?? []).map(formatRecipient).join(", "),
    subject: m.subject ?? "",
    date: m.receivedDateTime ?? "",
    snippet: (m.bodyPreview ?? "").replace(/\s+/g, " ").trim(),
    unread: m.isRead === false,
    hasAttachments: m.hasAttachments ?? false,
  };
}

function toFull(account: Account, m: GraphMessage, maxBodyChars: number, attachments: AttachmentInfo[]): MessageFull {
  const raw = m.body?.content ?? "";
  const body = m.body?.contentType?.toLowerCase() === "html" ? htmlToText(raw) : raw.replace(/\r\n/g, "\n").trim();
  const cut = truncate(body, maxBodyChars);
  return {
    ...toSummary(account, m),
    cc: (m.ccRecipients ?? []).map(formatRecipient).join(", "),
    body: cut.text,
    bodyTruncated: cut.truncated,
    attachments,
  };
}

async function listAttachments(account: Account, messageId: string): Promise<AttachmentInfo[]> {
  const list = await apiJson<GraphList<GraphAttachment>>(
    account,
    `${mailboxBase(account)}/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size`,
    PREFER,
  );
  return (list.value ?? []).map((a) => ({
    id: a.id,
    filename: a.name ?? "",
    mimeType: a.contentType ?? "application/octet-stream",
    size: a.size ?? 0,
  }));
}

export const outlook: Provider = {
  async search(account: Account, params: SearchParams): Promise<SearchResult> {
    let url: string;
    if (params.pageToken) {
      // The page token is Graph's next-page link. Only ever send the access token to Graph.
      if (!params.pageToken.startsWith(`${GRAPH}/`)) {
        throw new MailhubError("Invalid page_token for an Outlook account.");
      }
      url = params.pageToken;
    } else {
      url = buildSearchUrl(account, params);
    }
    const list = await apiJson<GraphList<GraphMessage>>(account, url, PREFER);
    const after = params.after?.getTime();
    const before = params.before?.getTime();
    const messages = (list.value ?? [])
      .filter((m) => {
        const received = m.receivedDateTime ? Date.parse(m.receivedDateTime) : NaN;
        if (after !== undefined && received < after) return false;
        if (before !== undefined && received >= before) return false;
        // Read state and attachments are not reliable search keywords, so they are checked here.
        if (params.unreadOnly && m.isRead !== false) return false;
        if (params.hasAttachment && !m.hasAttachments) return false;
        return true;
      })
      .map((m) => toSummary(account, m));
    return { messages, nextPageToken: list["@odata.nextLink"] };
  },

  async read(account: Account, id: string, maxBodyChars: number): Promise<MessageFull> {
    const m = await apiJson<GraphMessage>(
      account,
      `${mailboxBase(account)}/messages/${encodeURIComponent(id)}?$select=${FULL_FIELDS}`,
      PREFER,
    );
    const attachments = m.hasAttachments ? await listAttachments(account, id) : [];
    return toFull(account, m, maxBodyChars, attachments);
  },

  async thread(account: Account, threadId: string, maxBodyChars: number): Promise<MessageFull[]> {
    const filter = `conversationId eq '${threadId.replace(/'/g, "''")}'`;
    const list = await apiJson<GraphList<GraphMessage>>(
      account,
      `${mailboxBase(account)}/messages?$filter=${encodeURIComponent(filter)}&$top=50&$select=${FULL_FIELDS}`,
      PREFER,
    );
    // Graph rejects sorting together with a conversation filter, so sort here.
    return (list.value ?? [])
      .sort((a, b) => (a.receivedDateTime ?? "").localeCompare(b.receivedDateTime ?? ""))
      .map((m) => toFull(account, m, maxBodyChars, []));
  },

  async attachment(account: Account, messageId: string, attachmentId: string): Promise<AttachmentData> {
    const a = await apiJson<GraphAttachment>(
      account,
      `${mailboxBase(account)}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
      PREFER,
    );
    if (typeof a.contentBytes !== "string") {
      throw new MailhubError(
        `Attachment "${a.name ?? attachmentId}" is not a file attachment (${a["@odata.type"] ?? "unknown type"}) and cannot be downloaded.`,
      );
    }
    return {
      filename: a.name ?? "attachment",
      mimeType: a.contentType ?? "application/octet-stream",
      data: Buffer.from(a.contentBytes, "base64"),
    };
  },
};
