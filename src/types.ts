export type ProviderName = "gmail" | "outlook";

/** One authorised mailbox. Stored in accounts.json (never committed). */
export interface Account {
  alias: string;
  provider: ProviderName;
  /** Address of the mailbox that is read (the shared mailbox when one is set). */
  email: string;
  /** Address of the identity that signed in. */
  login: string;
  refreshToken: string;
  scopes: string[];
  /** Outlook only: tenant used for sign-in ("common" unless overridden). */
  tenant?: string;
  /** Outlook only: read this shared/delegated mailbox instead of the signed-in user's own. */
  sharedMailbox?: string;
  addedAt: string;
}

export interface SearchParams {
  query?: string;
  from?: string;
  to?: string;
  subject?: string;
  after?: Date;
  before?: Date;
  unreadOnly?: boolean;
  hasAttachment?: boolean;
  maxResults: number;
  pageToken?: string;
}

export interface MessageSummary {
  id: string;
  threadId: string;
  account: string;
  mailbox: string;
  from: string;
  to: string;
  subject: string;
  date: string;
  snippet: string;
  unread: boolean;
  hasAttachments?: boolean;
}

export interface AttachmentInfo {
  id: string;
  filename: string;
  mimeType: string;
  size: number;
}

export interface MessageFull extends MessageSummary {
  cc: string;
  body: string;
  bodyTruncated: boolean;
  attachments: AttachmentInfo[];
}

export interface SearchResult {
  messages: MessageSummary[];
  nextPageToken?: string;
}

export interface AttachmentData {
  filename: string;
  mimeType: string;
  data: Buffer;
}

export interface DraftInput {
  /** Addresses as "jane@x.com" or "Jane Doe <jane@x.com>". */
  to: string[];
  cc: string[];
  bcc: string[];
  subject?: string;
  body: string;
  /** When set, the draft is a reply to this message and stays in its thread. */
  replyToMessageId?: string;
  replyAll?: boolean;
}

export interface DraftResult {
  draftId: string;
  messageId: string;
  threadId: string;
  subject: string;
  to: string;
  cc: string;
  webLink?: string;
}

export interface Provider {
  search(account: Account, params: SearchParams): Promise<SearchResult>;
  read(account: Account, id: string, maxBodyChars: number): Promise<MessageFull>;
  thread(account: Account, threadId: string, maxBodyChars: number): Promise<MessageFull[]>;
  attachment(account: Account, messageId: string, attachmentId: string): Promise<AttachmentData>;
  /** Creates a draft in the mailbox. Nothing is sent. */
  createDraft(account: Account, draft: DraftInput): Promise<DraftResult>;
}

export class MailhubError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MailhubError";
  }
}
