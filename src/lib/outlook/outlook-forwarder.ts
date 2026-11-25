/**
 * OUTLOOK EMAIL FORWARDER SERVICE
 * Forwards newly received emails (within the configured lookback window, e.g. 24h)
 * to `ip.kdpmanagement@gmail.com` with full KDP-<number> identification in both Subject and HTML Body.
 */

import * as tls from "tls";
import { OutlookEmailMessage } from "./outlook-imap";
import { OutlookMailboxRecord } from "./outlook-db";

export interface ForwardEmailOptions {
  mailbox: OutlookMailboxRecord;
  message: OutlookEmailMessage;
  targetEmail: string;
  accessToken: string;
  fallbackAccessTokens?: string[];
  gmailSmtpUser?: string;
  gmailSmtpAppPassword?: string;
}

/**
 * Normalize a KDP code into `KDP-<số thứ tự>` format (with hyphen `-`).
 * Examples:
 * - "KDP-050" -> "KDP-050"
 * - "KDP_050" -> "KDP-050"
 * - "050"     -> "KDP-050"
 * - "50"      -> "KDP-50"
 */
export function normalizeKdpCode(rawCode?: string): string {
  const trimmed = (rawCode || "").trim();
  if (!trimmed) return "";
  if (/^\d+$/.test(trimmed)) {
    return `KDP-${trimmed}`;
  }
  const kdpMatch = trimmed.match(/^KDP[-_\s]*(\d+.*)$/i);
  if (kdpMatch) {
    return `KDP-${kdpMatch[1].replace(/[_\s]+/g, "-")}`.toUpperCase();
  }
  return trimmed.replace(/[_\s]+/g, "-").toUpperCase();
}

/**
 * Build target forwarding email in the format `ip.kdpmanagement+KDP-<số thứ tự>@gmail.com`.
 * If the mailbox has no `accountCode`, returns the clean base email `ip.kdpmanagement@gmail.com`.
 */
export function buildKdpScheduleForwardAddress(
  baseEmail: string,
  accountCode?: string
): string {
  const cleanBase = (baseEmail || "ip.kdpmanagement@gmail.com").trim();
  const atIdx = cleanBase.indexOf("@");
  if (atIdx <= 0) return cleanBase;
  const localPart = cleanBase.slice(0, atIdx).split("+")[0];
  const domain = cleanBase.slice(atIdx + 1);
  const normCode = normalizeKdpCode(accountCode);
  if (!normCode) {
    return `${localPart}@${domain}`;
  }
  return `${localPart}+${normCode}@${domain}`;
}

function escapeHtml(str: string): string {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Build original Subject & HTML Body for the forwarded email (100% original content, no custom wrapper).
 */
export function buildForwardedEmailPayload(
  _mailbox: OutlookMailboxRecord,
  message: OutlookEmailMessage
): { subject: string; htmlContent: string } {
  const subject = message.subject || "(No Subject)";

  const htmlContent = message.htmlBody
    ? message.htmlBody
    : `<pre style="white-space: pre-wrap; word-wrap: break-word; font-family: inherit; font-size: 14px;">${escapeHtml(
        message.textBody || message.snippet || ""
      )}</pre>`;

  return { subject, htmlContent };
}

/**
 * Extract raw Microsoft Graph message ID from composite ID `${email}_${folder}_${graphId}`
 */
function extractRawGraphMessageId(compositeId: string, email: string): string | null {
  if (!compositeId) return null;
  const prefixInbox = `${email}_INBOX_`;
  const prefixJunk = `${email}_Junk_`;
  if (compositeId.startsWith(prefixInbox)) {
    return compositeId.slice(prefixInbox.length);
  }
  if (compositeId.startsWith(prefixJunk)) {
    return compositeId.slice(prefixJunk.length);
  }
  const match = compositeId.match(/^[^_]+_(?:INBOX|Junk)_(.+)$/i);
  if (match?.[1]) {
    return match[1];
  }
  return null;
}

/**
 * Native Outlook Forward via Microsoft Graph API `POST /v1.0/me/messages/{id}/forward`
 * Preserves 100% original email HTML, headers, inline images, and attachments.
 */
async function forwardNativeViaMicrosoftGraph(
  accessToken: string,
  rawGraphMessageId: string,
  targetEmail: string
): Promise<void> {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(rawGraphMessageId)}/forward`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        comment: "",
        toRecipients: [
          {
            emailAddress: {
              address: targetEmail.trim(),
            },
          },
        ],
      }),
    }
  );

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`Graph native forward HTTP ${res.status}: ${errText.slice(0, 220)}`);
  }
}

/**
 * Send email via Microsoft Graph API `POST /v1.0/me/sendMail`
 */
async function sendViaMicrosoftGraph(
  accessToken: string,
  targetEmail: string,
  subject: string,
  htmlContent: string
): Promise<void> {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject,
        body: {
          contentType: "HTML",
          content: htmlContent,
        },
        toRecipients: [
          {
            emailAddress: {
              address: targetEmail.trim(),
            },
          },
        ],
      },
      saveToSentItems: false,
    }),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`Graph sendMail HTTP ${res.status}: ${errText.slice(0, 220)}`);
  }
}

/**
 * Minimal zero-dependency Gmail SMTP over TLS (port 465) sender
 * for users who optionally configure a Gmail App Password.
 */
async function sendViaGmailSmtpTls(params: {
  smtpUser: string;
  smtpPass: string;
  toEmail: string;
  subject: string;
  htmlContent: string;
}): Promise<void> {
  return new Promise((resolve, reject) => {
    const socket = tls.connect(
      {
        host: "smtp.gmail.com",
        port: 465,
        servername: "smtp.gmail.com",
      },
      () => {
        // Connected
      }
    );

    let buffer = "";
    let step = 0;
    const timeout = setTimeout(() => {
      socket.destroy();
      reject(new Error("Gmail SMTP connection timeout (20s)"));
    }, 20000);

    const finish = (err?: Error) => {
      clearTimeout(timeout);
      try {
        socket.end();
      } catch {
        // Ignore
      }
      if (err) reject(err);
      else resolve();
    };

    const sendLine = (line: string) => {
      socket.write(line + "\r\n");
    };

    socket.on("error", (err) => finish(err));

    socket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      const lines = buffer.split("\r\n");
      buffer = lines.pop() || "";

      for (const line of lines) {
        if (!line) continue;
        // Multi-line SMTP response has '-' at index 3; wait for final line with ' '
        if (line.length >= 4 && line[3] === "-") continue;
        const code = parseInt(line.slice(0, 3), 10);

        if (code >= 400) {
          return finish(new Error(`SMTP Error (${code}): ${line}`));
        }

        if (step === 0 && code === 220) {
          step = 1;
          sendLine("EHLO kdp-management.local");
        } else if (step === 1 && code === 250) {
          step = 2;
          const authPlain = Buffer.from(
            `\u0000${params.smtpUser.trim()}\u0000${params.smtpPass.replace(/\s+/g, "")}`,
            "utf8"
          ).toString("base64");
          sendLine(`AUTH PLAIN ${authPlain}`);
        } else if (step === 2 && code === 235) {
          step = 3;
          sendLine(`MAIL FROM:<${params.smtpUser.trim()}>`);
        } else if (step === 3 && code === 250) {
          step = 4;
          sendLine(`RCPT TO:<${params.toEmail.trim()}>`);
        } else if (step === 4 && code === 250) {
          step = 5;
          sendLine("DATA");
        } else if (step === 5 && code === 354) {
          step = 6;
          const encodedSubject = `=?UTF-8?B?${Buffer.from(params.subject, "utf8").toString(
            "base64"
          )}?=`;
          const bodyBase64 =
            Buffer.from(params.htmlContent, "utf8")
              .toString("base64")
              .match(/.{1,76}/g)
              ?.join("\r\n") || "";

          const mime = [
            `From: "KDP Outlook Monitor" <${params.smtpUser.trim()}>`,
            `To: <${params.toEmail.trim()}>`,
            `Subject: ${encodedSubject}`,
            `Date: ${new Date().toUTCString()}`,
            "MIME-Version: 1.0",
            'Content-Type: text/html; charset="UTF-8"',
            "Content-Transfer-Encoding: base64",
            "",
            bodyBase64,
            ".",
          ].join("\r\n");
          sendLine(mime);
        } else if (step === 6 && code === 250) {
          step = 7;
          sendLine("QUIT");
          return finish();
        }
      }
    });
  });
}

/**
 * Forward an email message to `targetEmail` (e.g. ip.kdpmanagement@gmail.com).
 * Uses Gmail SMTP if configured, otherwise Microsoft Graph `/me/sendMail` (with fallback tokens if needed).
 */
export async function forwardOutlookMessage(options: ForwardEmailOptions): Promise<{
  method: "GRAPH_NATIVE_FORWARD" | "GMAIL_SMTP" | "GRAPH_API";
}> {
  const { subject, htmlContent } = buildForwardedEmailPayload(
    options.mailbox,
    options.message
  );

  // 1. Try Native Outlook Forward via Microsoft Graph (`POST /v1.0/me/messages/{id}/forward`)
  // This forwards the exact original email (original HTML, inline images, attachments) directly from Outlook!
  const rawGraphId = extractRawGraphMessageId(
    options.message.id,
    options.mailbox.email
  );
  if (rawGraphId && options.accessToken) {
    try {
      await forwardNativeViaMicrosoftGraph(
        options.accessToken,
        rawGraphId,
        options.targetEmail
      );
      return { method: "GRAPH_NATIVE_FORWARD" };
    } catch {
      // Fall through to sendViaMicrosoftGraph if native forward fails
    }
  }

  // 2. If Gmail SMTP credentials are configured, try Gmail SMTP
  if (options.gmailSmtpUser && options.gmailSmtpAppPassword) {
    await sendViaGmailSmtpTls({
      smtpUser: options.gmailSmtpUser,
      smtpPass: options.gmailSmtpAppPassword,
      toEmail: options.targetEmail,
      subject,
      htmlContent,
    });
    return { method: "GMAIL_SMTP" };
  }

  // 3. Otherwise use Microsoft Graph API `POST /v1.0/me/sendMail` with original subject & body
  const tokensToTry = [
    options.accessToken,
    ...(options.fallbackAccessTokens || []),
  ].filter(Boolean);

  let lastErr: Error | null = null;
  for (const token of tokensToTry) {
    try {
      await sendViaMicrosoftGraph(token, options.targetEmail, subject, htmlContent);
      return { method: "GRAPH_API" };
    } catch (err: any) {
      lastErr = err;
    }
  }

  throw lastErr || new Error("Không thể gửi mail forward qua Microsoft Graph API.");
}
