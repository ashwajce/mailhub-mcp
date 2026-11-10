import { MailhubError } from "./types.js";

export interface Address {
  name?: string;
  email: string;
}

const EMAIL_RE = /^[^\s@<>"(),;:\\]+@[^\s@<>"(),;:\\]+\.[^\s@<>"(),;:\\]+$/;

/** Split a header such as `"Doe, Jane" <jane@x.com>, bob@y.com` on the commas that separate addresses. */
export function splitAddressList(value: string): string[] {
  const out: string[] = [];
  let current = "";
  let inQuotes = false;
  let depth = 0;
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch === "\\" && inQuotes) {
      current += ch + (value[++i] ?? "");
      continue;
    }
    if (ch === '"') inQuotes = !inQuotes;
    else if (!inQuotes && ch === "<") depth++;
    else if (!inQuotes && ch === ">") depth = Math.max(0, depth - 1);
    if (ch === "," && !inQuotes && depth === 0) {
      if (current.trim()) out.push(current.trim());
      current = "";
    } else {
      current += ch;
    }
  }
  if (current.trim()) out.push(current.trim());
  return out;
}

/** Parse `jane@x.com` or `Jane Doe <jane@x.com>`. Rejects anything that could inject extra headers. */
export function parseAddress(value: string): Address {
  if (/[\r\n]/.test(value)) throw new MailhubError("Addresses must not contain line breaks.");
  const text = value.trim();
  const angle = /^(.*)<([^<>]+)>$/.exec(text);
  const email = (angle ? angle[2] : text).trim();
  // A display name holding a second <address> is ambiguous about who receives the mail.
  if (!EMAIL_RE.test(email) || (angle && /[<>]/.test(angle[1]))) {
    throw new MailhubError(`"${value}" is not a valid email address.`);
  }
  const name = angle ? angle[1].trim().replace(/^"(.*)"$/, "$1").replace(/\\(.)/g, "$1").trim() : "";
  return name ? { name, email } : { email };
}

export function parseAddresses(values: string[] | undefined): Address[] {
  return (values ?? []).flatMap(splitAddressList).map(parseAddress);
}

/** RFC 2047 encoded-words for header text that is not plain ASCII. */
export function encodeHeaderText(text: string): string {
  if (/[\r\n]/.test(text)) throw new MailhubError("Header values must not contain line breaks.");
  if (/^[\x20-\x7e]*$/.test(text)) return text;
  // Each encoded-word may be at most 75 characters, so the text is split on character boundaries.
  const words: string[] = [];
  let chunk = "";
  for (const ch of text) {
    if (Buffer.byteLength(chunk + ch, "utf8") > 45) {
      words.push(chunk);
      chunk = "";
    }
    chunk += ch;
  }
  if (chunk) words.push(chunk);
  return words.map((w) => `=?UTF-8?B?${Buffer.from(w, "utf8").toString("base64")}?=`).join("\r\n ");
}

export function formatAddress(address: Address): string {
  if (!address.name) return address.email;
  const name = /^[\x20-\x7e]*$/.test(address.name)
    ? /^[A-Za-z0-9 !#$%&'*+\-/=?^_`{|}~]*$/.test(address.name)
      ? address.name
      : `"${address.name.replace(/(["\\])/g, "\\$1")}"`
    : encodeHeaderText(address.name);
  return `${name} <${address.email}>`;
}

export interface MimeInput {
  to: Address[];
  cc?: Address[];
  bcc?: Address[];
  subject: string;
  body: string;
  inReplyTo?: string;
  references?: string;
}

/** Build a plain-text RFC 5322 message. The From header is left for the mail service to fill in. */
export function buildMime(input: MimeInput): string {
  const headers: string[] = [];
  const addressHeader = (name: string, list: Address[] | undefined) => {
    if (list?.length) headers.push(`${name}: ${list.map(formatAddress).join(",\r\n ")}`);
  };
  addressHeader("To", input.to);
  addressHeader("Cc", input.cc);
  addressHeader("Bcc", input.bcc);
  headers.push(`Subject: ${encodeHeaderText(input.subject)}`);
  for (const [name, value] of [
    ["In-Reply-To", input.inReplyTo],
    ["References", input.references],
  ] as const) {
    if (!value) continue;
    if (/[\r\n]/.test(value)) throw new MailhubError(`${name} must not contain line breaks.`);
    headers.push(`${name}: ${value}`);
  }
  headers.push("MIME-Version: 1.0", "Content-Type: text/plain; charset=UTF-8", "Content-Transfer-Encoding: base64");
  const normalised = input.body.replace(/\r\n?/g, "\n").replace(/\n/g, "\r\n");
  const encoded = Buffer.from(normalised, "utf8").toString("base64").replace(/(.{76})/g, "$1\r\n");
  return `${headers.join("\r\n")}\r\n\r\n${encoded}`;
}
