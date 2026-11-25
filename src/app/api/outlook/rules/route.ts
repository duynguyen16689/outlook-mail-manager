import { NextRequest, NextResponse } from "next/server";
import { ensureFreshDatabase } from "@/lib/db/prisma";
import {
  getOutlookMailboxByIdOrEmail,
  updateMailboxRuleStatus,
} from "@/lib/outlook/outlook-db";
import { setupOutlookRuleAndForwardingViaBrowser } from "@/lib/outlook/outlook-rule-setup";

export const dynamic = "force-dynamic";

/**
 * POST /api/outlook/rules
 * Body:
 * {
 *   id: string,
 *   baseForwardEmail?: string, // default: "ip.kdpmanagement@gmail.com"
 *   ruleName?: string, // default: "KDP"
 *   keywords?: string[], // default: ["Amazon", "Kindle"]
 *   keepForwardedCopy?: boolean, // default: false
 *   setupRulesTab?: boolean, // default: true
 *   setupForwardingTab?: boolean // default: true
 * }
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
        { success: false, error: "Không tìm thấy tài khoản trong danh sách." },
        { status: 404 }
      );
    }

    const keywords: string[] = Array.isArray(body.keywords)
      ? body.keywords
      : typeof body.keywords === "string"
      ? body.keywords
          .split(",")
          .map((s: string) => s.trim())
          .filter(Boolean)
      : ["Amazon", "Kindle"];

    const result = await setupOutlookRuleAndForwardingViaBrowser(mailbox, {
      baseForwardEmail: body.baseForwardEmail || "ip.kdpmanagement@gmail.com",
      ruleName: body.ruleName || "KDP",
      keywords: keywords.length > 0 ? keywords : ["Amazon", "Kindle"],
      keepForwardedCopy: !!body.keepForwardedCopy,
      setupRulesTab: body.setupRulesTab !== false,
      setupForwardingTab: body.setupForwardingTab !== false,
    });

    const updated = await updateMailboxRuleStatus(mailbox.id, {
      ruleStatus: result.ruleStatus,
      ruleErrorMessage: result.ruleStatus === "SUCCESS" ? "" : result.message,
      proxy: result.resolvedProxy || mailbox.proxy,
    });

    return NextResponse.json(
      {
        success: result.success,
        ruleStatus: result.ruleStatus,
        message: result.message,
        error: result.success ? undefined : result.message,
        ruleForwardAddress: result.ruleForwardAddress,
        forwardingAddress: result.forwardingAddress,
        mailbox: updated,
        logs: result.logs,
      },
      { status: result.success ? 200 : 400 }
    );
  } catch (err: any) {
    return NextResponse.json(
      {
        success: false,
        error: err.message || "Lỗi hệ thống khi cài đặt Rule & Forwarding.",
      },
      { status: 500 }
    );
  }
}
