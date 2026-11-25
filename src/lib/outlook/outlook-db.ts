/**
 * STANDALONE OUTLOOK MAILBOX REPOSITORY (SQLite via Prisma Raw SQL)
 * Self-initializes the `OutlookMailbox` and `OutlookForwardedEmail` tables
 * and synchronizes LIVE Hotmail/Outlook accounts from KDP `Account` (KDP-034 .. KDP-126).
 */

import { prisma } from "@/lib/db/prisma";
import { AccountRegistry, RevenueRegistry } from "@/lib/registries";
import { BitBrowserClient } from "@/lib/bitbrowser/bitbrowser-client";
import { DEFAULT_OUTLOOK_CLIENT_ID } from "./outlook-oauth";
import { normalizeKdpCode } from "./outlook-forwarder";

export interface OutlookMailboxRecord {
  id: string;
  accountCode: string;
  email: string;
  password: string;
  proxy: string;
  note: string;
  clientId: string;
  refreshToken: string;
  accessToken: string;
  tokenExpiresAt: string | null;
  status: "UNLINKED" | "ACTIVE" | "ERROR";
  errorMessage: string;
  ruleStatus: "NONE" | "SUCCESS" | "VERIFY_REQUIRED" | "ERROR";
  ruleErrorMessage: string;
  ruleUpdatedAt: string | null;
  lastOtp: string;
  lastSubject: string;
  lastSender: string;
  lastEmailAt: string | null;
  lastCheckedAt: string | null;
  forwardedCount?: number;
  lastForwardedAt?: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface OutlookForwardedEmailRecord {
  messageKey: string;
  accountCode: string;
  email: string;
  folder: string;
  subject: string;
  fromAddress: string;
  receivedAt: string;
  forwardedTo: string;
  forwardedAt: string;
}

let tableInitialized = false;

function generateId(): string {
  return `om_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}

/**
 * Ensure `OutlookMailbox` and `OutlookForwardedEmail` tables exist,
 * migrate `accountCode` column if needed, and auto-sync LIVE KDP Hotmail accounts if test accounts exist.
 */
export async function ensureOutlookTable(): Promise<void> {
  if (tableInitialized) return;

  await prisma.$executeRawUnsafe(`
    CREATE TABLE IF NOT EXISTS "OutlookMailbox" (
      "id" TEXT NOT NULL PRIMARY KEY,
      "accountCode" TEXT DEFAULT '',
      "email" TEXT NOT NULL UNIQUE,
      "password" TEXT DEFAULT '',
      "note" TEXT DEFAULT '',
      "clientId" TEXT NOT NULL DEFAULT '9e5f94bc-e8a4-4e73-b8be-63364c29d753',
      "refreshToken" TEXT DEFAULT '',
      "accessToken" TEXT DEFAULT '',
      "tokenExpiresAt" TEXT,
      "status" TEXT NOT NULL DEFAULT 'UNLINKED',
      "errorMessage" TEXT DEFAULT '',
      "lastOtp" TEXT DEFAULT '',
      "lastSubject" TEXT DEFAULT '',
      "lastSender" TEXT DEFAULT '',
      "lastEmailAt" TEXT,
      "lastCheckedAt" TEXT,
      "createdAt" TEXT NOT NULL,
      "updatedAt" TEXT NOT NULL
    )
  `);

  // Ensure accountCode, proxy, and ruleStatus columns exist on older SQLite tables
  const cols = await prisma.$queryRawUnsafe<Array<{ name: string }>>(
    `PRAGMA table_info("OutlookMailbox")`
  );
  if (!cols.some((c) => c.name === "accountCode")) {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "OutlookMailbox" ADD COLUMN "accountCode" TEXT DEFAULT ''`
    );
  }
  if (!cols.some((c) => c.name === "proxy")) {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "OutlookMailbox" ADD COLUMN "proxy" TEXT DEFAULT ''`
    );
  }
  if (!cols.some((c) => c.name === "ruleStatus")) {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "OutlookMailbox" ADD COLUMN "ruleStatus" TEXT DEFAULT 'NONE'`
    );
  }
  if (!cols.some((c) => c.name === "ruleErrorMessage")) {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "OutlookMailbox" ADD COLUMN "ruleErrorMessage" TEXT DEFAULT ''`
    );
  }
  if (!cols.some((c) => c.name === "ruleUpdatedAt")) {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "OutlookMailbox" ADD COLUMN "ruleUpdatedAt" TEXT`
    );
  }

  await prisma.$executeRawUnsafe(`
    CREATE INDEX IF NOT EXISTS "OutlookMailbox_email_idx" ON "OutlookMailbox"("email")
  `);

  await prisma.$executeRawUnsafe(`
    CREATE TABLE IF NOT EXISTS "OutlookForwardedEmail" (
      "messageKey" TEXT NOT NULL PRIMARY KEY,
      "accountCode" TEXT DEFAULT '',
      "email" TEXT NOT NULL,
      "folder" TEXT DEFAULT 'INBOX',
      "subject" TEXT DEFAULT '',
      "fromAddress" TEXT DEFAULT '',
      "receivedAt" TEXT NOT NULL,
      "forwardedTo" TEXT NOT NULL,
      "forwardedAt" TEXT NOT NULL
    )
  `);

  tableInitialized = true;
}

/**
 * Synchronize LIVE (ACTIVE) Hotmail/Outlook accounts from the main KDP `Account` table
 * within the range `KDP-034` .. `KDP-126` (default: 34 to 126).
 */
export async function syncLiveKdpHotmailAccounts(
  fromNum: number = 50,
  toNum: number = 125,
  removeNonMatching: boolean = false
): Promise<{
  syncedCount: number;
  removedCount: number;
  mailboxes: OutlookMailboxRecord[];
}> {
  await ensureOutlookTable();

  const accounts = await prisma.account.findMany({
    orderBy: { code: "asc" },
  });

  // Also pull latest active proxies from BitBrowser Local API (if online)
  const bbProxyMap = new Map<string, string>();
  try {
    const bbClient = new BitBrowserClient();
    const health = await bbClient.checkHealth();
    if (health.online) {
      const profiles = await bbClient.listAllProfiles();
      for (const p of profiles) {
        if (p.accountCode && p.proxy?.raw) {
          bbProxyMap.set(p.accountCode.toUpperCase(), p.proxy.raw.trim());
        }
      }
    }
  } catch {
    // Ignore if BitBrowser is offline
  }

  const liveHotmailEntries: Array<{
    accountCode: string;
    email: string;
    password: string;
    proxy: string;
    note: string;
  }> = [];

  for (const acc of accounts) {
    const num = parseInt(acc.code.replace(/\D/g, ""), 10);
    if (isNaN(num) || num < fromNum || num > toNum) continue;

    const isExcluded = RevenueRegistry.isExcludedFromRevenue(acc);
    const normalizedAcc = {
      ...acc,
      status: isExcluded ? (acc.status === "BLOCKED" ? "BLOCKED" : "CLOSED") : acc.status,
    };
    if (!AccountRegistry.matchesFilter(normalizedAcc, "ACTIVE")) continue;

    const summary = acc.credentialsSummary || "";
    const mailMatch = summary.match(/Mail:\s*([^\s|]+)/i);
    const passMatch = summary.match(/Pass Mail:\s*([^|\n\r]+)/i);
    const email = (mailMatch?.[1] || "").trim().toLowerCase();
    const password = (passMatch?.[1] || "").trim();

    if (!email || !/(hotmail|outlook|live|msn)\./i.test(email)) continue;

    const accountCode = acc.code.toUpperCase();
    const bbProxy = bbProxyMap.get(accountCode);
    const effectiveProxy = bbProxy || (acc.proxyInfo || "").trim();

    // Keep Account.proxyInfo in sync with BitBrowser if BitBrowser has an updated proxy
    if (bbProxy && bbProxy !== (acc.proxyInfo || "").trim()) {
      await prisma.account
        .update({
          where: { id: acc.id },
          data: { proxyInfo: bbProxy },
        })
        .catch(() => {});
    }

    liveHotmailEntries.push({
      accountCode,
      email,
      password,
      proxy: effectiveProxy,
      note: acc.nickname || "",
    });
  }

  for (const item of liveHotmailEntries) {
    await upsertOutlookMailbox({
      accountCode: item.accountCode,
      email: item.email,
      password: item.password,
      proxy: item.proxy,
      note: item.note,
    });
  }

  let removedCount = 0;
  if (removeNonMatching && liveHotmailEntries.length > 0) {
    const allRows = await prisma.$queryRawUnsafe<OutlookMailboxRecord[]>(
      `SELECT * FROM "OutlookMailbox"`
    );
    const validEmails = new Set(liveHotmailEntries.map((x) => x.email.toLowerCase()));
    for (const row of allRows) {
      if (!validEmails.has(row.email.toLowerCase())) {
        await prisma.$executeRawUnsafe(`DELETE FROM "OutlookMailbox" WHERE "id" = ?`, row.id);
        removedCount++;
      }
    }
  }

  const mailboxes = await listOutlookMailboxes();
  return {
    syncedCount: liveHotmailEntries.length,
    removedCount,
    mailboxes,
  };
}

export async function listOutlookMailboxes(): Promise<OutlookMailboxRecord[]> {
  await ensureOutlookTable();
  const rows = await prisma.$queryRawUnsafe<OutlookMailboxRecord[]>(
    `SELECT * FROM "OutlookMailbox"`
  );
  const fwdStats = await prisma.$queryRawUnsafe<
    Array<{ emailLower: string; cnt: number; lastAt: string | null }>
  >(
    `SELECT lower("email") as emailLower, COUNT(*) as cnt, MAX("forwardedAt") as lastAt
     FROM "OutlookForwardedEmail"
     GROUP BY lower("email")`
  );
  const statsMap = new Map<string, { cnt: number; lastAt: string | null }>();
  for (const s of fwdStats || []) {
    if (s.emailLower) {
      statsMap.set(s.emailLower, { cnt: Number(s.cnt || 0), lastAt: s.lastAt || null });
    }
  }

  const list = (rows || []).map((r) => {
    const st = statsMap.get((r.email || "").toLowerCase());
    return {
      ...r,
      forwardedCount: st?.cnt || 0,
      lastForwardedAt: st?.lastAt || null,
    };
  });
  list.sort((a, b) => {
    const codeA = a.accountCode || "";
    const codeB = b.accountCode || "";
    if (codeA && codeB) {
      const numA = parseInt(codeA.replace(/\D/g, ""), 10) || 999999;
      const numB = parseInt(codeB.replace(/\D/g, ""), 10) || 999999;
      if (numA !== numB) return numA - numB;
      return codeA.localeCompare(codeB);
    }
    if (codeA && !codeB) return -1;
    if (!codeA && codeB) return 1;
    return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
  });
  return list;
}

export async function getOutlookMailboxByIdOrEmail(
  idOrEmail: string
): Promise<OutlookMailboxRecord | null> {
  await ensureOutlookTable();
  const key = (idOrEmail || "").trim();
  const rows = await prisma.$queryRawUnsafe<OutlookMailboxRecord[]>(
    `SELECT * FROM "OutlookMailbox" WHERE "id" = ? OR lower("email") = lower(?) OR upper("accountCode") = upper(?) LIMIT 1`,
    key,
    key,
    key
  );
  const row = rows?.[0] || null;
  if (!row) return null;
  const fwdStats = await prisma.$queryRawUnsafe<
    Array<{ cnt: number; lastAt: string | null }>
  >(
    `SELECT COUNT(*) as cnt, MAX("forwardedAt") as lastAt
     FROM "OutlookForwardedEmail"
     WHERE lower("email") = lower(?)`,
    row.email
  );
  return {
    ...row,
    forwardedCount: Number(fwdStats?.[0]?.cnt || 0),
    lastForwardedAt: fwdStats?.[0]?.lastAt || null,
  };
}

export interface UpsertMailboxInput {
  accountCode?: string;
  email: string;
  password?: string;
  proxy?: string;
  refreshToken?: string;
  clientId?: string;
  note?: string;
}

export async function upsertOutlookMailbox(
  input: UpsertMailboxInput
): Promise<OutlookMailboxRecord> {
  await ensureOutlookTable();
  const email = input.email.trim().toLowerCase();
  const existing = await getOutlookMailboxByIdOrEmail(email);
  const now = new Date().toISOString();

  let accountCode =
    input.accountCode && input.accountCode.trim()
      ? normalizeKdpCode(input.accountCode)
      : existing?.accountCode || "";
  let password =
    input.password && input.password.trim()
      ? input.password.trim()
      : existing?.password || "";
  let proxy =
    input.proxy && input.proxy.trim()
      ? input.proxy.trim()
      : existing?.proxy || "";
  const refreshToken =
    input.refreshToken !== undefined ? input.refreshToken.trim() : existing?.refreshToken || "";
  const clientId =
    input.clientId?.trim() || existing?.clientId || DEFAULT_OUTLOOK_CLIENT_ID;
  let note = input.note !== undefined && input.note.trim() ? input.note.trim() : existing?.note || "";
  const status = refreshToken ? "ACTIVE" : existing?.status === "ACTIVE" ? "ACTIVE" : "UNLINKED";

  // Auto-match KDP accountCode, password, and proxy from main KDP `Account` table + BitBrowser if missing
  if (!accountCode || !proxy || !password) {
    try {
      const matchedAcc = await prisma.account.findFirst({
        where: {
          OR: [
            ...(accountCode ? [{ code: accountCode }] : []),
            { credentialsSummary: { contains: email } },
          ],
        },
      });
      if (matchedAcc) {
        if (!accountCode) accountCode = matchedAcc.code.toUpperCase();
        if (!proxy && matchedAcc.proxyInfo) proxy = matchedAcc.proxyInfo.trim();
        if (!note && matchedAcc.nickname) note = matchedAcc.nickname.trim();
        if (!password && matchedAcc.credentialsSummary) {
          const passMatch = matchedAcc.credentialsSummary.match(/Pass Mail:\s*([^|\n\r]+)/i);
          if (passMatch?.[1]) password = passMatch[1].trim();
        }
      }
      if (accountCode) {
        const bbClient = new BitBrowserClient();
        const bbProfile = await bbClient.findProfileByCode(accountCode).catch(() => null);
        if (bbProfile?.proxy?.raw) {
          proxy = bbProfile.proxy.raw.trim();
        }
      }
    } catch {
      // Ignore auto-match error
    }
  }

  if (existing) {
    await prisma.$executeRawUnsafe(
      `UPDATE "OutlookMailbox"
       SET "accountCode" = ?, "password" = ?, "proxy" = ?, "refreshToken" = ?, "clientId" = ?, "note" = ?, "status" = ?, "errorMessage" = '', "updatedAt" = ?
       WHERE "id" = ?`,
      accountCode,
      password,
      proxy,
      refreshToken,
      clientId,
      note,
      status,
      now,
      existing.id
    );
    return (await getOutlookMailboxByIdOrEmail(existing.id))!;
  } else {
    const id = generateId();
    await prisma.$executeRawUnsafe(
      `INSERT INTO "OutlookMailbox" (
        "id", "accountCode", "email", "password", "proxy", "note", "clientId", "refreshToken", "status", "ruleStatus", "ruleErrorMessage", "createdAt", "updatedAt"
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'NONE', '', ?, ?)`,
      id,
      accountCode,
      email,
      password,
      proxy,
      note,
      clientId,
      refreshToken,
      status,
      now,
      now
    );
    return (await getOutlookMailboxByIdOrEmail(id))!;
  }
}

export async function updateMailboxTokens(
  id: string,
  data: {
    proxy?: string;
    refreshToken?: string;
    accessToken?: string;
    tokenExpiresAt?: string | null;
    status?: "UNLINKED" | "ACTIVE" | "ERROR";
    errorMessage?: string;
    lastOtp?: string;
    lastSubject?: string;
    lastSender?: string;
    lastEmailAt?: string | null;
    lastCheckedAt?: string | null;
  }
): Promise<OutlookMailboxRecord | null> {
  await ensureOutlookTable();
  const current = await getOutlookMailboxByIdOrEmail(id);
  if (!current) return null;

  const now = new Date().toISOString();
  const proxy = data.proxy !== undefined ? data.proxy : current.proxy;
  const refreshToken = data.refreshToken !== undefined ? data.refreshToken : current.refreshToken;
  const accessToken = data.accessToken !== undefined ? data.accessToken : current.accessToken;
  const tokenExpiresAt =
    data.tokenExpiresAt !== undefined ? data.tokenExpiresAt : current.tokenExpiresAt;
  const status = data.status !== undefined ? data.status : current.status;
  const errorMessage = data.errorMessage !== undefined ? data.errorMessage : current.errorMessage;
  const lastOtp = data.lastOtp !== undefined ? data.lastOtp : current.lastOtp;
  const lastSubject = data.lastSubject !== undefined ? data.lastSubject : current.lastSubject;
  const lastSender = data.lastSender !== undefined ? data.lastSender : current.lastSender;
  const lastEmailAt = data.lastEmailAt !== undefined ? data.lastEmailAt : current.lastEmailAt;
  const lastCheckedAt =
    data.lastCheckedAt !== undefined ? data.lastCheckedAt : current.lastCheckedAt;

  await prisma.$executeRawUnsafe(
    `UPDATE "OutlookMailbox"
     SET "proxy" = ?,
         "refreshToken" = ?,
         "accessToken" = ?,
         "tokenExpiresAt" = ?,
         "status" = ?,
         "errorMessage" = ?,
         "lastOtp" = ?,
         "lastSubject" = ?,
         "lastSender" = ?,
         "lastEmailAt" = ?,
         "lastCheckedAt" = ?,
         "updatedAt" = ?
     WHERE "id" = ?`,
    proxy,
    refreshToken,
    accessToken,
    tokenExpiresAt,
    status,
    errorMessage,
    lastOtp,
    lastSubject,
    lastSender,
    lastEmailAt,
    lastCheckedAt,
    now,
    current.id
  );

  return getOutlookMailboxByIdOrEmail(current.id);
}

export async function updateMailboxRuleStatus(
  id: string,
  data: {
    ruleStatus: "NONE" | "SUCCESS" | "VERIFY_REQUIRED" | "ERROR";
    ruleErrorMessage?: string;
    proxy?: string;
  }
): Promise<OutlookMailboxRecord | null> {
  await ensureOutlookTable();
  const current = await getOutlookMailboxByIdOrEmail(id);
  if (!current) return null;

  const now = new Date().toISOString();
  const proxy = data.proxy !== undefined ? data.proxy : current.proxy;
  const ruleErrorMessage = data.ruleErrorMessage !== undefined ? data.ruleErrorMessage : "";

  await prisma.$executeRawUnsafe(
    `UPDATE "OutlookMailbox"
     SET "proxy" = ?,
         "ruleStatus" = ?,
         "ruleErrorMessage" = ?,
         "ruleUpdatedAt" = ?,
         "updatedAt" = ?
     WHERE "id" = ?`,
    proxy,
    data.ruleStatus,
    ruleErrorMessage,
    now,
    now,
    current.id
  );

  return getOutlookMailboxByIdOrEmail(current.id);
}

export async function clearAllOutlookMailboxes(): Promise<number> {
  await ensureOutlookTable();
  const rows = await prisma.$queryRawUnsafe<Array<{ count: number }>>(
    `SELECT COUNT(*) as count FROM "OutlookMailbox"`
  );
  const count = Number(rows?.[0]?.count || 0);
  await prisma.$executeRawUnsafe(`DELETE FROM "OutlookMailbox"`);
  return count;
}

export async function deleteOutlookMailbox(id: string): Promise<boolean> {
  await ensureOutlookTable();
  await prisma.$executeRawUnsafe(`DELETE FROM "OutlookMailbox" WHERE "id" = ?`, id);
  return true;
}

export async function hasEmailBeenForwarded(
  messageKey: string,
  email?: string,
  subject?: string,
  receivedAt?: string
): Promise<boolean> {
  await ensureOutlookTable();
  const rows = await prisma.$queryRawUnsafe<Array<{ messageKey: string }>>(
    `SELECT "messageKey" FROM "OutlookForwardedEmail" WHERE "messageKey" = ? LIMIT 1`,
    messageKey
  );
  if (rows && rows.length > 0) return true;

  if (email && subject !== undefined && receivedAt) {
    const byContent = await prisma.$queryRawUnsafe<Array<{ messageKey: string }>>(
      `SELECT "messageKey" FROM "OutlookForwardedEmail"
       WHERE lower("email") = lower(?) AND "subject" = ? AND "receivedAt" = ?
       LIMIT 1`,
      email.trim(),
      subject,
      receivedAt
    );
    if (byContent && byContent.length > 0) return true;
  }

  return false;
}

export async function getForwardedRecordsForEmail(
  email: string
): Promise<OutlookForwardedEmailRecord[]> {
  await ensureOutlookTable();
  const rows = await prisma.$queryRawUnsafe<OutlookForwardedEmailRecord[]>(
    `SELECT * FROM "OutlookForwardedEmail" WHERE lower("email") = lower(?) ORDER BY "forwardedAt" DESC`,
    (email || "").trim()
  );
  return rows || [];
}

/**
 * Parse a single bulk import line of any of these formats:
 * 1) `KDP-050|email|password|refresh_token|client_id|recovery_email`
 * 2) `050|email|password|refresh_token|client_id`
 * 3) `email|password|refresh_token|client_id|recovery_email|KDP-050`
 * 4) `email|password|refresh_token|client_id|KDP-050`
 * 5) `KDP-050|email|password` or `email|password`
 */
export function parseBulkOutlookAccountLine(line: string): UpsertMailboxInput | null {
  const parts = line
    .split("|")
    .map((p) => p.trim())
    .filter((_, idx, arr) => idx < arr.length);

  if (parts.length === 0) return null;

  let accountCode = "";
  let email = "";
  let password = "";
  let proxy = "";
  let refreshToken = "";
  let clientId = DEFAULT_OUTLOOK_CLIENT_ID;
  let extraNote = "";

  // Case A: First column does NOT contain '@' and second column contains '@' -> First column is KDP Name/Code
  if (parts[0] && !parts[0].includes("@") && parts[1] && parts[1].includes("@")) {
    accountCode = normalizeKdpCode(parts[0]);
    email = parts[1];
    password = parts[2] || "";
    const p3 = parts[3] || "";
    if (p3.includes(":") && !p3.startsWith("M.")) {
      proxy = p3;
      refreshToken = parts[4] || "";
      clientId = parts[5] || DEFAULT_OUTLOOK_CLIENT_ID;
      extraNote = parts[6] || "";
    } else {
      refreshToken = p3;
      clientId = parts[4] || DEFAULT_OUTLOOK_CLIENT_ID;
      extraNote = parts[5] || "";
    }
  } else if (parts[0] && parts[0].includes("@")) {
    // Case B: First column is email
    email = parts[0];
    password = parts[1] || "";
    const p2 = parts[2] || "";
    if (p2.includes(":") && !p2.startsWith("M.")) {
      proxy = p2;
      refreshToken = parts[3] || "";
      clientId = parts[4] || DEFAULT_OUTLOOK_CLIENT_ID;
      extraNote = parts[5] || "";
    } else {
      refreshToken = p2;
      clientId = parts[3] || DEFAULT_OUTLOOK_CLIENT_ID;
      const p4 = parts[4] || "";
      const p5 = parts[5] || "";
      if (p5 && /^KDP[-_\s]?\d+|^\d+$/i.test(p5) && !p5.includes("@")) {
        extraNote = p4;
        accountCode = normalizeKdpCode(p5);
      } else if (p4 && (/^KDP[-_\s]?\d+|^\d+$/i.test(p4) || !p4.includes("@"))) {
        accountCode = normalizeKdpCode(p4);
        extraNote = p5;
      } else {
        extraNote = p4;
      }
    }
  }

  if (!email || !email.includes("@")) return null;

  return {
    accountCode,
    email,
    password,
    proxy: proxy || undefined,
    refreshToken,
    clientId,
    note: extraNote,
  };
}

export async function recordForwardedEmail(entry: {
  messageKey: string;
  accountCode: string;
  email: string;
  folder: string;
  subject: string;
  fromAddress: string;
  receivedAt: string;
  forwardedTo: string;
}): Promise<void> {
  await ensureOutlookTable();
  const now = new Date().toISOString();
  await prisma.$executeRawUnsafe(
    `INSERT OR REPLACE INTO "OutlookForwardedEmail" (
      "messageKey", "accountCode", "email", "folder", "subject", "fromAddress", "receivedAt", "forwardedTo", "forwardedAt"
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    entry.messageKey,
    entry.accountCode || "",
    entry.email,
    entry.folder || "INBOX",
    entry.subject || "",
    entry.fromAddress || "",
    entry.receivedAt,
    entry.forwardedTo,
    now
  );
}

export async function listRecentForwardedEmails(
  limit: number = 50
): Promise<OutlookForwardedEmailRecord[]> {
  await ensureOutlookTable();
  const rows = await prisma.$queryRawUnsafe<OutlookForwardedEmailRecord[]>(
    `SELECT * FROM "OutlookForwardedEmail" ORDER BY "forwardedAt" DESC LIMIT ?`,
    limit
  );
  return rows || [];
}
