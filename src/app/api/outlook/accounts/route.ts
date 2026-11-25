import { NextRequest, NextResponse } from "next/server";
import { ensureFreshDatabase } from "@/lib/db/prisma";
import {
  clearAllOutlookMailboxes,
  deleteOutlookMailbox,
  listOutlookMailboxes,
  parseBulkOutlookAccountLine,
  syncLiveKdpHotmailAccounts,
  upsertOutlookMailbox,
} from "@/lib/outlook/outlook-db";
import { ensureOutlookSchedulerStarted } from "@/lib/outlook/outlook-scheduler";

export const dynamic = "force-dynamic";

export async function GET() {
  try {
    await ensureFreshDatabase(false);
    ensureOutlookSchedulerStarted();
    const mailboxes = await listOutlookMailboxes();
    return NextResponse.json({
      success: true,
      mailboxes,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Không thể tải danh sách Outlook." },
      { status: 500 }
    );
  }
}

/**
 * POST /api/outlook/accounts
 * Supports:
 * 1) `{ action: "sync_kdp", fromNum?: number, toNum?: number, removeNonMatching?: boolean }`
 * 2) Bulk text lines: `rawLines` with format `KDP-034|email|password|refresh_token|client_id|...` or `email|password|refresh_token|client_id`
 * 3) Single mailbox object: `{ accountCode, email, password, refreshToken, clientId, note }`
 */
export async function POST(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    ensureOutlookSchedulerStarted();
    const body = await request.json();

    if (body.action === "clear_all") {
      const deletedCount = await clearAllOutlookMailboxes();
      return NextResponse.json({
        success: true,
        deletedCount,
        mailboxes: [],
      });
    }

    if (body.action === "sync_kdp") {
      const fromNum = Number(body.fromNum) || 34;
      const toNum = Number(body.toNum) || 126;
      const removeNonMatching = body.removeNonMatching !== undefined ? !!body.removeNonMatching : true;
      const result = await syncLiveKdpHotmailAccounts(fromNum, toNum, removeNonMatching);
      return NextResponse.json({
        success: true,
        syncedCount: result.syncedCount,
        removedCount: result.removedCount,
        mailboxes: result.mailboxes,
      });
    }

    if (typeof body.rawLines === "string" && body.rawLines.trim()) {
      if (body.replaceExisting) {
        await clearAllOutlookMailboxes();
      }

      const lines = body.rawLines
        .split(/\r?\n/)
        .map((l: string) => l.trim())
        .filter(Boolean);

      const imported = [];
      for (const line of lines) {
        const parsed = parseBulkOutlookAccountLine(line);
        if (!parsed) continue;

        const record = await upsertOutlookMailbox({
          ...parsed,
          note: body.note || parsed.note || "",
        });
        imported.push(record);
      }

      const mailboxes = await listOutlookMailboxes();
      return NextResponse.json({
        success: true,
        importedCount: imported.length,
        mailboxes,
      });
    }

    if (!body.email || !String(body.email).includes("@")) {
      return NextResponse.json(
        { success: false, error: "Địa chỉ email không hợp lệ." },
        { status: 400 }
      );
    }

    const mailbox = await upsertOutlookMailbox({
      accountCode: body.accountCode,
      email: body.email,
      password: body.password,
      refreshToken: body.refreshToken,
      clientId: body.clientId,
      note: body.note,
    });

    const mailboxes = await listOutlookMailboxes();
    return NextResponse.json({
      success: true,
      mailbox,
      mailboxes,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Lỗi khi lưu tài khoản Outlook." },
      { status: 500 }
    );
  }
}

export async function DELETE(request: NextRequest) {
  try {
    await ensureFreshDatabase(false);
    const { searchParams } = new URL(request.url);
    const id = searchParams.get("id");
    if (!id) {
      return NextResponse.json(
        { success: false, error: "Thiếu ID tài khoản cần xóa." },
        { status: 400 }
      );
    }

    await deleteOutlookMailbox(id);
    const mailboxes = await listOutlookMailboxes();
    return NextResponse.json({
      success: true,
      mailboxes,
    });
  } catch (err: any) {
    return NextResponse.json(
      { success: false, error: err.message || "Lỗi khi xóa tài khoản Outlook." },
      { status: 500 }
    );
  }
}
