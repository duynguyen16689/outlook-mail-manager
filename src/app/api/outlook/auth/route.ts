import { NextRequest, NextResponse } from "next/server";
import { ensureFreshDatabase } from "@/lib/db/prisma";
import {
  getOutlookMailboxByIdOrEmail,
  updateMailboxTokens,
} from "@/lib/outlook/outlook-db";
import {
  buildOutlookAuthUrl,
  DEFAULT_OUTLOOK_CLIENT_ID,
  exchangeCodeForTokens,
} from "@/lib/outlook/outlook-oauth";
import { captureOutlookTokenViaBrowser } from "@/lib/outlook/outlook-auto-browser";

export const dynamic = "force-dynamic";

function nowStr() {
  return new Date().toLocaleTimeString("vi-VN", { hour12: false });
}

/**
 * GET /api/outlook/auth?id=...
 * Returns the Microsoft OAuth2 Authorization URL for a mailbox.
 */
export async function GET(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    const { searchParams } = new URL(request.url);
    const idOrEmail = searchParams.get("id") || searchParams.get("email") || "";
    const mailbox = idOrEmail ? await getOutlookMailboxByIdOrEmail(idOrEmail) : null;

    const email = mailbox?.email || searchParams.get("email") || "";
    const clientId = mailbox?.clientId || DEFAULT_OUTLOOK_CLIENT_ID;
    const authUrl = buildOutlookAuthUrl(email, clientId);

    return NextResponse.json({
      success: true,
      email,
      clientId,
      authUrl,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Không thể tạo link OAuth2." },
      { status: 500 }
    );
  }
}

/**
 * POST /api/outlook/auth
 * Modes:
 * - `mode: "manual"`: Exchange `redirectUrlOrCode` (`https://localhost/?code=...`) for tokens
 * - `mode: "auto_browser"`: Launch visible Chrome/Edge via Puppeteer, auto-fill email + pass, intercept code & save tokens
 */
export async function POST(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    const body = await request.json();
    const idOrEmail = body.id || body.email;

    if (!idOrEmail) {
      return NextResponse.json(
        { success: false, error: "Thiếu ID hoặc địa chỉ email tài khoản." },
        { status: 400 }
      );
    }

    const mailbox = await getOutlookMailboxByIdOrEmail(idOrEmail);
    if (!mailbox) {
      return NextResponse.json(
        { success: false, error: "Không tìm thấy tài khoản Outlook trong hệ thống." },
        { status: 404 }
      );
    }

    const clientId = mailbox.clientId || DEFAULT_OUTLOOK_CLIENT_ID;
    const logTag = mailbox.accountCode
      ? `${mailbox.accountCode} | ${mailbox.email}`
      : mailbox.email;
    let tokenResult;
    let logs: string[] = [];

    try {
      if (body.mode === "auto_browser") {
        const res = await captureOutlookTokenViaBrowser(
          mailbox.email,
          mailbox.password,
          clientId,
          mailbox.proxy,
          mailbox.accountCode
        );
        tokenResult = res;
        logs = res.logs || [];
      } else {
        logs.push(
          `[${nowStr()}] [INFO] [${logTag}] Đang đổi mã code thủ công sang Refresh Token${
            mailbox.proxy ? ` (Proxy: ${mailbox.proxy.split(":").slice(0, 2).join(":")})` : ""
          }...`
        );
        const redirectUrlOrCode = body.redirectUrlOrCode || body.code || "";
        tokenResult = await exchangeCodeForTokens(redirectUrlOrCode, clientId, mailbox.proxy);
        logs.push(
          `[${nowStr()}] [OK] [${logTag}] Đổi Token thành công (Refresh Token: ${tokenResult.refreshToken.length} ký tự).`
        );
      }
    } catch (authErr: any) {
      const errMsg = authErr.message || "Lỗi khi lấy Token OAuth2.";
      if (Array.isArray(authErr?.logs)) {
        logs = authErr.logs;
      } else {
        logs.push(`[${nowStr()}] [ERROR] [${mailbox.email}] ${errMsg}`);
      }
      const failedMailbox = await updateMailboxTokens(mailbox.id, {
        errorMessage: errMsg,
      });
      return NextResponse.json(
        {
          success: false,
          error: errMsg,
          mailbox: failedMailbox,
          logs,
        },
        { status: 400 }
      );
    }

    const expiresAt = new Date(
      Date.now() + (tokenResult.expiresIn - 60) * 1000
    ).toISOString();

     const updated = await updateMailboxTokens(mailbox.id, {
       proxy: (tokenResult as any).resolvedProxy || mailbox.proxy,
       refreshToken: tokenResult.refreshToken,
       accessToken: tokenResult.accessToken,
       tokenExpiresAt: expiresAt,
       status: "ACTIVE",
       errorMessage: "",
     });

    return NextResponse.json({
      success: true,
      mailbox: updated,
      logs,
      message: "Đã lấy và lưu Refresh Token thành công!",
    });
  } catch (err: any) {
    return NextResponse.json(
      {
        success: false,
        error: err.message || "Lỗi khi xác thực OAuth2 với Microsoft.",
      },
      { status: 400 }
    );
  }
}
