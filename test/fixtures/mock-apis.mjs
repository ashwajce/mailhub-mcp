// Preloaded into the server process by the end-to-end test (node --import):
// stands in for Google and Microsoft so the real stdio server can be exercised offline.
const json = (body, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

const b64url = (text) => Buffer.from(text).toString("base64url");

globalThis.fetch = async (input) => {
  const url = typeof input === "string" ? input : input.url;

  if (url.includes("oauth2.googleapis.com/token")) return json({ access_token: "g", expires_in: 3600 });
  if (url.includes("login.microsoftonline.com")) return json({ access_token: "m", expires_in: 3600 });

  if (url.includes("gmail.googleapis.com")) {
    if (/\/drafts$/.test(url)) return json({ id: "d1", message: { id: "dm1", threadId: "dt1" } });
    if (/\/messages\?/.test(url)) return json({ messages: [{ id: "g1" }] });
    if (/\/messages\/g1\/attachments\/att/.test(url)) return json({ data: b64url("col1,col2\n1,2\n") });
    if (/\/messages\/g1\?/.test(url)) {
      return json({
        id: "g1",
        threadId: "gt1",
        labelIds: ["INBOX"],
        snippet: "Gmail snippet",
        internalDate: "1767225600000",
        payload: {
          mimeType: "multipart/mixed",
          headers: [
            { name: "From", value: "Acme <billing@acme.com>" },
            { name: "To", value: "me@example.com" },
            { name: "Subject", value: "Gmail invoice" },
          ],
          parts: [
            { partId: "0", mimeType: "text/plain", filename: "", body: { data: b64url("Body from Gmail") } },
            { partId: "1", mimeType: "text/csv", filename: "../data.csv", body: { attachmentId: "att", size: 14 } },
          ],
        },
      });
    }
  }

  if (url.includes("graph.microsoft.com")) {
    if (decodeURIComponent(url).includes("blocked")) {
      return json({ error: { code: "ErrorAccessDenied", message: "Access is denied." } }, 403);
    }
    return json({
      value: [
        {
          id: "o1",
          conversationId: "oc1",
          subject: "Outlook invoice",
          from: { emailAddress: { name: "Jane", address: "jane@contoso.example" } },
          toRecipients: [{ emailAddress: { address: "me@contoso.example" } }],
          receivedDateTime: "2026-01-02T00:00:00Z",
          bodyPreview: "Outlook snippet",
          isRead: true,
          hasAttachments: false,
        },
      ],
    });
  }

  return json({ error: { message: `unmocked ${url}` } }, 500);
};
