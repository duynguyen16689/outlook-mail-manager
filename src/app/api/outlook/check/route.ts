import { NextRequest, NextResponse } from "next/server";
import { ensureFreshDatabase } from "@/lib/db/prisma";
import {
  getForwardedRecordsForEmail,
  getOutlookMailboxByIdOrEmail,
  hasEmailBeenForwarded,
  recordForwardedEmail,
  updateMailboxTokens,
} from "@/lib/outlook/outlook-db";
import { refreshOutlookAccessToken } from "@/lib/outlook/outlook-oauth";
import { fetchOutlookEmails } from "@/lib/outlook/outlook-imap";
import {
  buildKdpScheduleForwardAddress,
  forwardOutlookMessage,
} from "@/lib/outlook/outlook-forwarder";
import { loadOutlookScheduleConfig } from "@/lib/outlook/outlook-scheduler";

export const dynamic = "force-dynamic";

function nowStr() {
  return new Date().toLocaleTimeString("vi-VN", { hour12: false });
}

/**
 * POST /api/outlook/check
 * Body: `{ id: string }` or `{ email: string }`
 * Uses saved `refreshToken` + `clientId` to obtain a fresh `accessToken`,
 * reads INBOX + Junk via IMAP XOAUTH2 (or Graph API fallback), and extracts OTPs.
 */
export async function POST(request: NextRequest) {
  const logs: string[] = [];
  const addLog = (level: "INFO" | "OK" | "WARN" | "ERROR", email: string, msg: string) => {
    const line = `[${nowStr()}] [${level}] [${email}] ${msg}`;
    logs.push(line);
    if (level === "ERROR") console.error(line);
    else console.log(line);
  };

  try {
    await ensureFreshDatabase(false);
    const body = await request.json();
    const idOrEmail = body.id || body.email;
    const limitPerFolder = Math.max(5, Math.min(200, Number(body.limit) || 25));

    if (!idOrEmail) {
      return NextResponse.json(
        { success: false, error: "Thiếu ID hoặc địa chỉ email cần check.", logs },
        { status: 400 }
      );
    }

    const mailbox = await getOutlookMailboxByIdOrEmail(idOrEmail);
    if (!mailbox) {
      return NextResponse.json(
        { success: false, error: "Không tìm thấy tài khoản Outlook.", logs },
        { status: 404 }
      );
    }

    const logTag = mailbox.accountCode
      ? `${mailbox.accountCode} | ${mailbox.email}`
      : mailbox.email;

    if (!mailbox.refreshToken) {
      addLog("WARN", logTag, "Chưa có Refresh Token. Cần lấy Token trước khi đọc mail.");
      return NextResponse.json(
        {
          success: false,
          error:
            "Tài khoản này mới có Email/Pass, chưa có Refresh Token. Vui lòng bấm 'Lấy Token' (1 lần duy nhất) trước khi đọc mail.",
          needsToken: true,
          logs,
        },
        { status: 400 }
      );
    }

    try {
      const proxyLabel = mailbox.proxy
        ? ` qua Proxy ${mailbox.proxy.split(":").slice(0, 2).join(":")}`
        : "";
      // 1. Refresh access_token from Microsoft OAuth2
      addLog("INFO", logTag, `Đang làm mới Access Token từ Refresh Token${proxyLabel}...`);
      const tokenData = await refreshOutlookAccessToken(
        mailbox.refreshToken,
        mailbox.clientId,
        mailbox.proxy
      );
      addLog("OK", logTag, "Đã nhận Access Token hợp lệ từ Microsoft.");

      // 2. Fetch emails from INBOX + Junk
      addLog(
        "INFO",
        logTag,
        `Đang tải tối đa ${limitPerFolder} thư/mục từ INBOX & Junk/Spam${proxyLabel}...`
      );
      const messages = await fetchOutlookEmails(
        mailbox.email,
        tokenData.accessToken,
        tokenData.scope,
        limitPerFolder,
        mailbox.proxy
      );

      const inboxCount = messages.filter((m) => m.folder === "INBOX").length;
      const junkCount = messages.filter((m) => m.folder === "Junk").length;
      const latestMsg = messages[0] || null;
      const latestOtpMsg = messages.find((m) => !!m.otpCode) || null;

      addLog(
        "OK",
        logTag,
        `Đọc hòm thư thành công: Tổng ${messages.length} thư (INBOX: ${inboxCount}, Junk/Spam: ${junkCount})${
          latestOtpMsg?.otpCode ? ` • Phát hiện OTP: ${latestOtpMsg.otpCode}` : ""
        }`
      );

      const nowIso = new Date().toISOString();
      const expiresAt = new Date(
        Date.now() + (tokenData.expiresIn - 60) * 1000
      ).toISOString();

      await updateMailboxTokens(mailbox.id, {
        refreshToken: tokenData.refreshToken || mailbox.refreshToken,
        accessToken: tokenData.accessToken,
        tokenExpiresAt: expiresAt,
        status: "ACTIVE",
        errorMessage: "",
        lastOtp: latestOtpMsg?.otpCode || mailbox.lastOtp || "",
        lastSubject: latestMsg?.subject || "",
        lastSender: latestMsg ? `${latestMsg.fromName} <${latestMsg.fromAddress}>` : "",
        lastEmailAt: latestMsg?.date || null,
        lastCheckedAt: nowIso,
      });

      // 3. Auto-forward any unforwarded emails within lookbackHours (e.g. 24h) immediately on Check Mail
      try {
        const schedConfig = await loadOutlookScheduleConfig();
        if (schedConfig.autoForwardEnabled && schedConfig.forwardToEmail) {
          const lookbackMs = (schedConfig.lookbackHours || 24) * 3600 * 1000;
          const nowMs = Date.now();
          const targetForwardEmail = buildKdpScheduleForwardAddress(
            schedConfig.forwardToEmail,
            mailbox.accountCode
          );

          const recentMessages = messages.filter((m) => {
            const msgMs = new Date(m.date).getTime();
            return !isNaN(msgMs) && nowMs - msgMs <= lookbackMs && nowMs - msgMs >= -3600000;
          });

          for (const msg of recentMessages) {
            const messageKey = `${mailbox.email.toLowerCase()}_${msg.id}`;
            const legacyKey = `${mailbox.accountCode || mailbox.email}_${msg.id}`;
            const alreadySent =
              (await hasEmailBeenForwarded(messageKey, mailbox.email, msg.subject, msg.date)) ||
              (await hasEmailBeenForwarded(legacyKey, mailbox.email, msg.subject, msg.date));

            if (!alreadySent) {
              try {
                await forwardOutlookMessage({
                  mailbox,
                  message: msg,
                  targetEmail: targetForwardEmail,
                  accessToken: tokenData.accessToken,
                  gmailSmtpUser: schedConfig.gmailSmtpUser,
                  gmailSmtpAppPassword: schedConfig.gmailSmtpAppPassword,
                });

                await recordForwardedEmail({
                  messageKey,
                  accountCode: mailbox.accountCode || "",
                  email: mailbox.email,
                  folder: msg.folder,
                  subject: msg.subject,
                  fromAddress: msg.fromAddress,
                  receivedAt: msg.date,
                  forwardedTo: targetForwardEmail,
                });

                addLog(
                  "OK",
                  logTag,
                  `Đã tự động Forward thư mới (24h): "${msg.subject}" -> ${targetForwardEmail}`
                );
              } catch (fwdErr: any) {
                addLog("WARN", logTag, `Lỗi Forward thư "${msg.subject}": ${fwdErr.message}`);
              }
            }
          }
        }
      } catch {
        // Ignore auto-forward check error
      }

      const updatedMailbox = await getOutlookMailboxByIdOrEmail(mailbox.id);
      const fwdRecords = await getForwardedRecordsForEmail(mailbox.email);
      const enrichedMessages = messages.map((m) => {
        const key1 = `${mailbox.email.toLowerCase()}_${m.id}`;
        const key2 = `${mailbox.accountCode || mailbox.email}_${m.id}`;
        const matchedFwd = fwdRecords.find(
          (f) =>
            f.messageKey === key1 ||
            f.messageKey === key2 ||
            (f.subject === m.subject && f.receivedAt === m.date)
        );
        return {
          ...m,
          isForwarded: Boolean(matchedFwd),
          forwardedTo: matchedFwd?.forwardedTo || undefined,
          forwardedAt: matchedFwd?.forwardedAt || undefined,
        };
      });

      return NextResponse.json({
        success: true,
        mailbox: updatedMailbox,
        messages: enrichedMessages,
        latestOtp: latestOtpMsg?.otpCode || null,
        logs,
      });
    } catch (checkErr: any) {
      const errMsg = checkErr.message || "Lỗi khi kết nối đọc mail Outlook.";
      addLog("ERROR", logTag, `Lỗi đọc mail: ${errMsg}`);
      const updatedMailbox = await updateMailboxTokens(mailbox.id, {
        status: "ERROR",
        errorMessage: errMsg,
        lastCheckedAt: new Date().toISOString(),
      });

      return NextResponse.json(
        {
          success: false,
          error: errMsg,
          mailbox: updatedMailbox,
          logs,
        },
        { status: 400 }
      );
    }
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Lỗi hệ thống khi check mail.", logs },
      { status: 500 }
    );
  }
}
