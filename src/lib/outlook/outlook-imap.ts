/**
 * OUTLOOK MAIL READER (IMAP XOAUTH2 + Graph REST API Fallback)
 * Connects to outlook.office365.com:993 using OAuth2 accessToken via ImapFlow,
 * fetches recent emails from both INBOX and Junk (Spam), parses HTML/Text bodies,
 * and extracts OTP codes automatically.
 */

import { ImapFlow } from "imapflow";
import { simpleParser } from "mailparser";
import { fetchViaOutlookProxy, parseOutlookProxy } from "./outlook-proxy";

export interface OutlookEmailMessage {
  id: string;
  folder: "INBOX" | "Junk";
  subject: string;
  fromName: string;
  fromAddress: string;
  toAddress: string;
  date: string;
  snippet: string;
  textBody: string;
  htmlBody: string;
  otpCode: string | null;
  category: "OTP" | "KDP_ALERT" | "GENERAL";
  isUnread: boolean;
}

/**
 * Extract 4-8 digit OTP / verification code from subject + text body.
 */
export function extractOtpFromEmail(subject: string, textBody: string): string | null {
  const combined = `${subject || ""}\n${textBody || ""}`;
  const lower = combined.toLowerCase();

  // Check if email looks like an OTP / verification / security code email
  const isOtpContext =
    lower.includes("verification code") ||
    lower.includes("security code") ||
    lower.includes("one-time") ||
    lower.includes("otp") ||
    lower.includes("passcode") ||
    lower.includes("mã bảo mật") ||
    lower.includes("mã xác minh") ||
    lower.includes("xác thực") ||
    lower.includes("verify your") ||
    lower.includes("sign-in") ||
    lower.includes("single-use code");

  if (!isOtpContext) {
    return null;
  }

  // Match explicit patterns like "code is: 123456" or standalone 6-digit numbers
  const explicitPatterns = [
    /(?:code|otp|passcode|mã)[\s:islà]*([0-9]{6})\b/i,
    /\b([0-9]{6})\b/,
    /\b([0-9]{4,8})\b/,
  ];

  for (const regex of explicitPatterns) {
    const matches = combined.match(new RegExp(regex.source, "gi"));
    if (matches) {
      for (const m of matches) {
        const digits = m.replace(/\D/g, "");
        // Exclude years like 2024, 2025, 2026
        if (digits.length >= 6 && !/^(19|20)\d{2}$/.test(digits)) {
          return digits;
        }
      }
    }
  }

  return null;
}

/**
 * Classify email for KDP / OTP highlighting.
 */
export function classifyEmail(
  subject: string,
  fromAddress: string,
  otpCode: string | null
): "OTP" | "KDP_ALERT" | "GENERAL" {
  if (otpCode) return "OTP";
  const s = (subject || "").toLowerCase();
  const f = (fromAddress || "").toLowerCase();
  if (
    f.includes("amazon") ||
    f.includes("kdp") ||
    f.includes("draft2digital") ||
    s.includes("blocked") ||
    s.includes("terminated") ||
    s.includes("action required") ||
    s.includes("kindle direct publishing")
  ) {
    return "KDP_ALERT";
  }
  return "GENERAL";
}

/**
 * Fetch latest emails via IMAP XOAUTH2 (outlook.office365.com:993)
 */
export async function fetchEmailsViaImap(
  email: string,
  accessToken: string,
  limitPerFolder: number = 12,
  proxy?: string
): Promise<OutlookEmailMessage[]> {
  const parsedProxy = parseOutlookProxy(proxy);
  let client = new ImapFlow({
    host: "outlook.office365.com",
    port: 993,
    secure: true,
    proxy: parsedProxy?.proxyUrl,
    auth: {
      user: email.trim(),
      accessToken: accessToken.trim(),
    },
    logger: false,
  });

  const messages: OutlookEmailMessage[] = [];

  try {
    try {
      await client.connect();
    } catch (connectErr) {
      if (parsedProxy) {
        // Fallback to direct connection if proxy fails
        client = new ImapFlow({
          host: "outlook.office365.com",
          port: 993,
          secure: true,
          auth: {
            user: email.trim(),
            accessToken: accessToken.trim(),
          },
          logger: false,
        });
        await client.connect();
      } else {
        throw connectErr;
      }
    }

    const foldersToCheck: Array<{ path: string; label: "INBOX" | "Junk" }> = [
      { path: "INBOX", label: "INBOX" },
      { path: "Junk", label: "Junk" },
    ];

    for (const folderInfo of foldersToCheck) {
      let lock;
      try {
        lock = await client.getMailboxLock(folderInfo.path);
        const total = client.mailbox ? client.mailbox.exists : 0;
        if (total > 0) {
          const startSeq = Math.max(1, total - limitPerFolder + 1);
          const range = `${startSeq}:*`;

          for await (const msg of client.fetch(range, {
            uid: true,
            flags: true,
            envelope: true,
            source: true,
            internalDate: true,
          })) {
            try {
              const parsed = await simpleParser(msg.source as Buffer);
              const subject = parsed.subject || msg.envelope?.subject || "(Không có tiêu đề)";
              const fromObj = parsed.from?.value?.[0] || msg.envelope?.from?.[0];
              const fromName = fromObj?.name || fromObj?.address || "Unknown";
              const fromAddress = fromObj?.address || "";
              const toObj = parsed.to
                ? Array.isArray(parsed.to)
                  ? parsed.to[0]?.value?.[0]
                  : parsed.to.value?.[0]
                : msg.envelope?.to?.[0];
              const toAddress = toObj?.address || email;
              const rawText = (parsed.text || "").trim();
              const htmlBody = typeof parsed.html === "string" ? parsed.html : "";
              const plainFromHtml =
                !rawText && htmlBody
                  ? htmlBody
                      .replace(/<style[\s\S]*?<\/style>/gi, " ")
                      .replace(/<script[\s\S]*?<\/script>/gi, " ")
                      .replace(/<[^>]+>/g, " ")
                      .replace(/&nbsp;/gi, " ")
                      .replace(/\s+/g, " ")
                      .trim()
                  : "";
              const textBody = rawText || plainFromHtml;
              const snippet = textBody.replace(/\s+/g, " ").slice(0, 180);
              const dateObj =
                (msg.internalDate ? new Date(msg.internalDate) : null) ||
                parsed.date ||
                new Date();
              const otpCode = extractOtpFromEmail(subject, textBody);
              const isUnread = !msg.flags?.has("\\Seen");

              messages.push({
                id: `${email}_${folderInfo.label}_${msg.uid}`,
                folder: folderInfo.label,
                subject,
                fromName,
                fromAddress,
                toAddress,
                date: dateObj.toISOString(),
                snippet,
                textBody,
                htmlBody,
                otpCode,
                category: classifyEmail(subject, fromAddress, otpCode),
                isUnread,
              });
            } catch {
              // Skip malformed message
            }
          }
        }
      } catch {
        // Folder might not exist or be inaccessible; continue
      } finally {
        if (lock) lock.release();
      }
    }
  } finally {
    try {
      await client.logout();
    } catch {
      // Ignore logout errors
    }
  }

  // Sort newest first
  messages.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
  return messages;
}

/**
 * Fetch emails from both INBOX and Junk Email via Microsoft Graph REST API v1.0.
 * Works even when POP/IMAP is disabled in Outlook settings!
 */
export async function fetchEmailsViaGraph(
  email: string,
  accessToken: string,
  limitPerFolder: number = 15,
  proxy?: string
): Promise<OutlookEmailMessage[]> {
  const folders: Array<{ graphId: string; label: "INBOX" | "Junk" }> = [
    { graphId: "inbox", label: "INBOX" },
    { graphId: "junkemail", label: "Junk" },
  ];

  const folderResults = await Promise.all(
    folders.map(async (folderInfo) => {
      const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderInfo.graphId}/messages?$top=${limitPerFolder}&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,body,isRead&$orderby=receivedDateTime desc`;
      const res = await fetchViaOutlookProxy(
        url,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        },
        proxy
      );

      if (!res.ok) {
        const errText = await res.text().catch(() => "");
        throw new Error(`Graph API Error (${res.status}) [${folderInfo.graphId}]: ${errText}`);
      }

      const data = await res.json();
      const items: any[] = data.value || [];

      return items.map((item): OutlookEmailMessage => {
        const subject = item.subject || "(Không có tiêu đề)";
        const fromName = item.from?.emailAddress?.name || "";
        const fromAddress = item.from?.emailAddress?.address || "";
        const toAddress = item.toRecipients?.[0]?.emailAddress?.address || email;
        const content = item.body?.content || "";
        const isHtml = item.body?.contentType?.toLowerCase() === "html";
        const plainFromHtml = isHtml
          ? content
              .replace(/<style[\s\S]*?<\/style>/gi, " ")
              .replace(/<script[\s\S]*?<\/script>/gi, " ")
              .replace(/<[^>]+>/g, " ")
              .replace(/\s+/g, " ")
              .trim()
          : content;
        const textBody = plainFromHtml || item.bodyPreview || "";
        const htmlBody = isHtml ? content : "";
        const otpCode = extractOtpFromEmail(subject, `${item.bodyPreview || ""}\n${textBody}`);

        return {
          id: `${email}_${folderInfo.label}_${item.id}`,
          folder: folderInfo.label,
          subject,
          fromName: fromName || fromAddress,
          fromAddress,
          toAddress,
          date: item.receivedDateTime || new Date().toISOString(),
          snippet: item.bodyPreview || textBody.slice(0, 180),
          textBody,
          htmlBody,
          otpCode,
          category: classifyEmail(subject, fromAddress, otpCode),
          isUnread: !item.isRead,
        };
      });
    })
  );

  const allMessages = folderResults.flat();
  allMessages.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
  return allMessages;
}

/**
 * Unified Mail Reader:
 * 1. Uses Microsoft Graph API v1.0 (reads INBOX + Junk in ~0.4s via HTTP and works even when IMAP is disabled on newer Hotmail accounts).
 * 2. Automatically falls back to IMAP XOAUTH2 (outlook.office365.com:993) if the token only has IMAP scope.
 */
export async function fetchOutlookEmails(
  email: string,
  accessToken: string,
  scope: string = "",
  limitPerFolder: number = 25,
  proxy?: string
): Promise<OutlookEmailMessage[]> {
  const safeLimit = Math.max(5, Math.min(200, Number(limitPerFolder) || 25));

  // Prefer Graph API if scope includes Mail.Read or Mail.ReadWrite
  if (/\bMail\.Read(Write)?\b/i.test(scope)) {
    try {
      return await fetchEmailsViaGraph(email, accessToken, safeLimit, proxy);
    } catch (graphErr: any) {
      if (proxy) {
        // Retry Graph API directly if the proxy connection timed out
        return await fetchEmailsViaGraph(email, accessToken, safeLimit, undefined);
      }
      throw new Error(graphErr?.message || "Lỗi kết nối Microsoft Graph API");
    }
  }

  try {
    return await fetchEmailsViaImap(email, accessToken, safeLimit, proxy);
  } catch (imapErr: any) {
    // Try Graph API as a backup
    try {
      return await fetchEmailsViaGraph(email, accessToken, safeLimit, proxy);
    } catch {
      const detail = imapErr?.responseText
        ? `${imapErr.message} (${imapErr.responseText})`
        : imapErr?.message || "Lỗi kết nối đọc mail";
      throw new Error(detail);
    }
  }
}
