import { MailhubError } from "./types.js";

const ENTITIES: Record<string, string> = {
  nbsp: " ",
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  "#39": "'",
};

/** Good-enough HTML to text conversion for reading mail bodies. */
export function htmlToText(html: string): string {
  return html
    .replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6]|table|blockquote)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&(#x?[0-9a-f]+|[a-z]+[0-9]*);/gi, (match, name: string) => {
      const key = name.toLowerCase();
      if (key in ENTITIES) return ENTITIES[key];
      if (key.startsWith("#")) {
        const code = key[1] === "x" ? parseInt(key.slice(2), 16) : parseInt(key.slice(1), 10);
        if (Number.isFinite(code) && code > 0 && code <= 0x10ffff) return String.fromCodePoint(code);
      }
      return match;
    })
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function truncate(text: string, maxChars: number): { text: string; truncated: boolean } {
  if (text.length <= maxChars) return { text, truncated: false };
  return { text: text.slice(0, maxChars), truncated: true };
}

/** Decode bytes using the charset named in a Content-Type header, falling back to UTF-8. */
export function decodeBytes(data: Buffer, contentType: string | undefined): string {
  const charset = /charset="?([^";\s]+)"?/i.exec(contentType ?? "")?.[1];
  if (charset) {
    try {
      return new TextDecoder(charset).decode(data);
    } catch {
      // Unknown charset label: fall through to UTF-8.
    }
  }
  return data.toString("utf8");
}

/** Accepts YYYY-MM-DD or a full ISO 8601 timestamp. */
export function parseDate(value: string | undefined, field: string): Date | undefined {
  if (value === undefined || value === "") return undefined;
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T00:00:00Z` : value;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    throw new MailhubError(`"${field}" must be YYYY-MM-DD or an ISO 8601 timestamp, got "${value}".`);
  }
  return date;
}

/** Run an async mapper over items with a concurrency cap, preserving order. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const index = next++;
      results[index] = await fn(items[index]);
    }
  });
  await Promise.all(workers);
  return results;
}

export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Make an attachment name safe to use as a file name inside the download folder. */
export function safeFilename(name: string): string {
  const base = name.split(/[\\/]/).pop() ?? "";
  const cleaned = base.replace(/[^\w.\- ()\[\]@+,]/g, "_").replace(/^\.+/, "").trim();
  return cleaned.slice(0, 180) || "attachment";
}
