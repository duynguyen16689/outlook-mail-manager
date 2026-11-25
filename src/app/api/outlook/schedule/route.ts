import { NextRequest, NextResponse } from "next/server";
import { ensureFreshDatabase } from "@/lib/db/prisma";
import {
  ensureOutlookSchedulerStarted,
  executeOutlookScheduledScan,
  getOutlookSchedulerStatus,
  requestStopOutlookScan,
  saveOutlookScheduleConfig,
  tickOutlookScheduler,
} from "@/lib/outlook/outlook-scheduler";
import {
  clearAllOutlookMailboxes,
  getOutlookMailboxByIdOrEmail,
  listOutlookMailboxes,
  parseBulkOutlookAccountLine,
  recordForwardedEmail,
  upsertOutlookMailbox,
} from "@/lib/outlook/outlook-db";
import { refreshOutlookAccessToken } from "@/lib/outlook/outlook-oauth";
import {
  buildKdpScheduleForwardAddress,
  forwardOutlookMessage,
} from "@/lib/outlook/outlook-forwarder";

export const dynamic = "force-dynamic";

export async function GET() {
  try {
    await ensureFreshDatabase(false);
    ensureOutlookSchedulerStarted();
    await tickOutlookScheduler();
    const status = await getOutlookSchedulerStatus();
    return NextResponse.json({
      success: true,
      ...status,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Không thể tải trạng thái Schedule." },
      { status: 500 }
    );
  }
}

export async function PUT(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    ensureOutlookSchedulerStarted();
    const body = await request.json();

    let importedCount = 0;
    if (typeof body.rawLines === "string" && body.rawLines.trim()) {
      if (body.replaceExisting) {
        await clearAllOutlookMailboxes();
      }
      const lines = body.rawLines
        .split(/\r?\n/)
        .map((l: string) => l.trim())
        .filter(Boolean);
      for (const line of lines) {
        const parsed = parseBulkOutlookAccountLine(line);
        if (!parsed) continue;
        await upsertOutlookMailbox(parsed);
        importedCount++;
      }
      if (importedCount > 0) {
        body.enabled = true;
      }
    }

    await saveOutlookScheduleConfig(body);
    await tickOutlookScheduler();
    const status = await getOutlookSchedulerStatus();
    const mailboxes = importedCount > 0 || body.replaceExisting ? await listOutlookMailboxes() : undefined;
    return NextResponse.json({
      success: true,
      importedCount,
      mailboxes,
      ...status,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Không thể lưu cấu hình Schedule." },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    ensureOutlookSchedulerStarted();
    const body = await request.json();

    if (body.action === "stop") {
      requestStopOutlookScan();
      const status = await getOutlookSchedulerStatus();
      return NextResponse.json({
        success: true,
        message: "Đã gửi yêu cầu dừng phiên quét.",
        ...status,
      });
    }

    if (body.action === "forward_single") {
      const mailbox = await getOutlookMailboxByIdOrEmail(body.mailboxId || body.email || "");
      if (!mailbox || !mailbox.refreshToken) {
        return NextResponse.json(
          { success: false, error: "Tài khoản chưa có Refresh Token để gửi forward." },
          { status: 400 }
        );
      }

      const status = await getOutlookSchedulerStatus();
      const baseEmail =
        body.targetEmail || status.config.forwardToEmail || "ip.kdpmanagement@gmail.com";
      const targetEmail = buildKdpScheduleForwardAddress(baseEmail, mailbox.accountCode);

      const tokenData = await refreshOutlookAccessToken(
        mailbox.refreshToken,
        mailbox.clientId,
        mailbox.proxy
      );

      const res = await forwardOutlookMessage({
        mailbox,
        message: body.message,
        targetEmail,
        accessToken: tokenData.accessToken,
        gmailSmtpUser: status.config.gmailSmtpUser,
        gmailSmtpAppPassword: status.config.gmailSmtpAppPassword,
      });

      const messageKey = `${mailbox.email.toLowerCase()}_${body.message?.id || Date.now()}`;
      await recordForwardedEmail({
        messageKey,
        accountCode: mailbox.accountCode || "",
        email: mailbox.email,
        folder: body.message?.folder || "INBOX",
        subject: body.message?.subject || "",
        fromAddress: body.message?.fromAddress || "",
        receivedAt: body.message?.date || new Date().toISOString(),
        forwardedTo: targetEmail,
      });

      const updatedStatus = await getOutlookSchedulerStatus();
      return NextResponse.json({
        success: true,
        method: res.method,
        forwardedTo: targetEmail,
        message: `Đã Forward thư [${mailbox.accountCode || mailbox.email}] -> ${targetEmail} thành công!`,
        ...updatedStatus,
      });
    }

    // Default action: "run_now" (launch asynchronously so UI can poll progress)
    executeOutlookScheduledScan("MANUAL").catch((err) => {
      console.error("[OutlookScheduleRoute] Manual scan error:", err);
    });

    // Brief wait so `isRunning = true` is reflected immediately
    await new Promise((r) => setTimeout(r, 150));
    const status = await getOutlookSchedulerStatus();
    return NextResponse.json({
      success: true,
      message: "Đã khởi chạy phiên quét & forward toàn bộ tài khoản LIVE Hotmail!",
      ...status,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Lỗi thực thi Schedule." },
      { status: 500 }
    );
  }
}
