/**
 * OUTLOOK RECURRING SCHEDULER & AUTO-FORWARD ENGINE
 * Modeled after KDP Standalone Report Scheduler:
 * - Supports INTERVAL mode (default: every 2 hours) and CUSTOM slots mode (e.g. "08:00, 14:00, 20:00").
 * - Automatically scans all LIVE KDP Hotmail accounts in range (default: KDP-034 to KDP-126).
 * - Checks for emails received within `lookbackHours` (default: 24h) and forwards new emails
 *   to `ip.kdpmanagerment@gmail.com` with full `KDP-<number>` identification.
 */

import { prisma } from "@/lib/db/prisma";
import {
  hasEmailBeenForwarded,
  listOutlookMailboxes,
  listRecentForwardedEmails,
  recordForwardedEmail,
  syncLiveKdpHotmailAccounts,
  updateMailboxTokens,
} from "./outlook-db";
import { refreshOutlookAccessToken } from "./outlook-oauth";
import { fetchOutlookEmails } from "./outlook-imap";
import {
  buildKdpScheduleForwardAddress,
  forwardOutlookMessage,
} from "./outlook-forwarder";

export interface OutlookScheduleConfig {
  enabled: boolean;
  scheduleMode: "INTERVAL" | "CUSTOM";
  dailyTime: string; // e.g. "00:00"
  intervalHours: number; // default: 2
  customScheduleSlots: string; // e.g. "08:00, 14:00, 20:00"
  timezone: string; // default: "Asia/Ho_Chi_Minh"
  fromAccountNum: number; // default: 1
  toAccountNum: number; // default: 9999
  lookbackHours: number; // default: 24
  autoForwardEnabled: boolean; // default: true
  forwardToEmail: string; // default: "ip.kdpmanagement@gmail.com"
  gmailSmtpUser: string; // optional Gmail SMTP sender
  gmailSmtpAppPassword: string; // optional 16-char Gmail App Password
}

export interface OutlookScanRunState {
  isRunning: boolean;
  stopRequested: boolean;
  triggerSource: "SCHEDULE" | "MANUAL" | null;
  startedAt: string | null;
  finishedAt: string | null;
  currentAccountCode: string;
  currentEmail: string;
  processedCount: number;
  totalToScan: number;
  skippedNoTokenCount: number;
  newEmails24hCount: number;
  alreadyForwardedCount: number;
  forwardedCount: number;
  errorCount: number;
  lastScheduledSlotKey: string;
  lastRunSummary: string;
  logs: string[];
}

const SETTING_KEY = "outlook_schedule_config_v1";
const RUN_STATE_KEY = "outlook_schedule_run_state_v1";

export const DEFAULT_OUTLOOK_SCHEDULE_CONFIG: OutlookScheduleConfig = {
  enabled: false,
  scheduleMode: "INTERVAL",
  dailyTime: "00:00",
  intervalHours: 2,
  customScheduleSlots: "00:00, 02:00, 04:00, 06:00, 08:00, 10:00, 12:00, 14:00, 16:00, 18:00, 20:00, 22:00",
  timezone: "Asia/Ho_Chi_Minh",
  fromAccountNum: 1,
  toAccountNum: 9999,
  lookbackHours: 24,
  autoForwardEnabled: true,
  forwardToEmail: "ip.kdpmanagement@gmail.com",
  gmailSmtpUser: "ip.kdpmanagement@gmail.com",
  gmailSmtpAppPassword: "",
};

const g = globalThis as unknown as {
  __outlookSchedulerTimer?: ReturnType<typeof setInterval>;
  __outlookSchedulerVersion?: number;
  __outlookScheduleConfigCache?: OutlookScheduleConfig;
  __outlookScanState?: OutlookScanRunState;
};

const CURRENT_SCHEDULER_VERSION = 4;

function getScanState(): OutlookScanRunState {
  if (!g.__outlookScanState) {
    g.__outlookScanState = {
      isRunning: false,
      stopRequested: false,
      triggerSource: null,
      startedAt: null,
      finishedAt: null,
      currentAccountCode: "",
      currentEmail: "",
      processedCount: 0,
      totalToScan: 0,
      skippedNoTokenCount: 0,
      newEmails24hCount: 0,
      alreadyForwardedCount: 0,
      forwardedCount: 0,
      errorCount: 0,
      lastScheduledSlotKey: "",
      lastRunSummary: "Chưa chạy phiên quét tự động nào.",
      logs: [],
    };
  }
  return g.__outlookScanState;
}

function nowTimeStr(): string {
  return new Date().toLocaleTimeString("vi-VN", {
    timeZone: "Asia/Ho_Chi_Minh",
    hour12: false,
  });
}

function appendScanLog(
  level: "INFO" | "OK" | "WARN" | "ERROR",
  tag: string,
  msg: string
): void {
  const state = getScanState();
  const line = `[${nowTimeStr()}] [${level}] [${tag}] ${msg}`;
  state.logs.push(line);
  if (state.logs.length > 300) {
    state.logs = state.logs.slice(-300);
  }
}

export async function loadOutlookScheduleConfig(): Promise<OutlookScheduleConfig> {
  try {
    const row = await prisma.appSetting.findUnique({
      where: { key: SETTING_KEY },
    });
    if (row?.value) {
      const parsed = JSON.parse(row.value);
      delete parsed.rawLines;
      delete parsed.replaceExisting;
      if (parsed.forwardToEmail === "ip.kdpmanagerment@gmail.com") {
        parsed.forwardToEmail = "ip.kdpmanagement@gmail.com";
      }
      if (parsed.gmailSmtpUser === "ip.kdpmanagerment@gmail.com") {
        parsed.gmailSmtpUser = "ip.kdpmanagement@gmail.com";
      }
      const merged: OutlookScheduleConfig = {
        ...DEFAULT_OUTLOOK_SCHEDULE_CONFIG,
        ...parsed,
        fromAccountNum: 1,
        toAccountNum: 99999,
      };
      g.__outlookScheduleConfigCache = merged;
      return merged;
    }
  } catch {
    if (g.__outlookScheduleConfigCache) {
      return g.__outlookScheduleConfigCache;
    }
  }
  g.__outlookScheduleConfigCache = { ...DEFAULT_OUTLOOK_SCHEDULE_CONFIG };
  return g.__outlookScheduleConfigCache;
}

export async function saveOutlookScheduleConfig(
  partial: Partial<OutlookScheduleConfig>
): Promise<OutlookScheduleConfig> {
  const current = await loadOutlookScheduleConfig();
  const cleanPartial: any = { ...partial };
  delete cleanPartial.rawLines;
  delete cleanPartial.replaceExisting;

  const updated: OutlookScheduleConfig = {
    ...current,
    ...cleanPartial,
    enabled: cleanPartial.enabled !== undefined ? Boolean(cleanPartial.enabled) : true,
    intervalHours: Math.max(0.5, Math.min(24, Number(cleanPartial.intervalHours ?? current.intervalHours) || 2)),
    fromAccountNum: Math.max(1, Number(cleanPartial.fromAccountNum ?? current.fromAccountNum) || 1),
    toAccountNum: Math.max(1, Number(cleanPartial.toAccountNum ?? current.toAccountNum) || 9999),
    lookbackHours: Math.max(1, Math.min(720, Number(cleanPartial.lookbackHours ?? current.lookbackHours) || 24)),
    forwardToEmail: (cleanPartial.forwardToEmail ?? current.forwardToEmail ?? "ip.kdpmanagement@gmail.com").trim(),
  };

  g.__outlookScheduleConfigCache = updated;

  try {
    await prisma.appSetting.upsert({
      where: { key: SETTING_KEY },
      create: {
        key: SETTING_KEY,
        value: JSON.stringify(updated),
      },
      update: {
        value: JSON.stringify(updated),
      },
    });
  } catch {
    // Ignore DB write error if table not ready
  }

  return updated;
}

export function getZonedDateAndTime(
  now = new Date(),
  timeZone = "Asia/Ho_Chi_Minh"
): { dateKey: string; timeHm: string } {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  const parts = formatter.formatToParts(now);
  const get = (type: string) => parts.find((p) => p.type === type)?.value || "00";
  const dateKey = `${get("year")}-${get("month")}-${get("day")}`;
  const hour = get("hour") === "24" ? "00" : get("hour");
  const timeHm = `${hour}:${get("minute")}`;
  return { dateKey, timeHm };
}

export function computeOutlookScheduleSlots(
  config: OutlookScheduleConfig
): string[] {
  if (config.scheduleMode === "CUSTOM" && config.customScheduleSlots?.trim()) {
    const rawParts = config.customScheduleSlots.split(/[,;\s]+/);
    const parsed: string[] = [];
    for (const part of rawParts) {
      const match = part.trim().match(/^(\d{1,2}):(\d{2})$/);
      if (match) {
        const h = Math.max(0, Math.min(23, parseInt(match[1], 10)));
        const m = Math.max(0, Math.min(59, parseInt(match[2], 10)));
        const formatted = `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
        if (!parsed.includes(formatted)) parsed.push(formatted);
      }
    }
    if (parsed.length > 0) {
      parsed.sort();
      return parsed;
    }
  }

  const [hStr, mStr] = (config.dailyTime || "00:00").split(":");
  const startH = Math.max(0, Math.min(23, parseInt(hStr || "0", 10) || 0));
  const startM = Math.max(0, Math.min(59, parseInt(mStr || "0", 10) || 0));
  const startTotalMinutes = startH * 60 + startM;

  const validInterval =
    typeof config.intervalHours === "number" &&
    Number.isFinite(config.intervalHours) &&
    config.intervalHours > 0 &&
    config.intervalHours < 24
      ? config.intervalHours
      : 2;

  const intervalMinutes = Math.max(15, Math.round(validInterval * 60));
  const slotsMinutes: number[] = [];
  let curr = startTotalMinutes;

  for (let i = 0; i < 96; i++) {
    const normalized = ((curr % 1440) + 1440) % 1440;
    if (slotsMinutes.includes(normalized)) break;
    slotsMinutes.push(normalized);
    curr += intervalMinutes;
    if (curr >= startTotalMinutes + 1440) break;
  }

  slotsMinutes.sort((a, b) => a - b);

  return slotsMinutes.map((min) => {
    const hh = String(Math.floor(min / 60)).padStart(2, "0");
    const mm = String(min % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  });
}

async function syncPersistedRunState(state: OutlookScanRunState): Promise<void> {
  try {
    const row = await prisma.appSetting.findUnique({ where: { key: RUN_STATE_KEY } });
    if (row?.value) {
      const saved = JSON.parse(row.value);
      if (!state.lastScheduledSlotKey && saved.lastScheduledSlotKey) {
        state.lastScheduledSlotKey = saved.lastScheduledSlotKey;
      }
      if (
        (!state.lastRunSummary || state.lastRunSummary === "Chưa chạy phiên quét tự động nào.") &&
        saved.lastRunSummary
      ) {
        state.lastRunSummary = saved.lastRunSummary;
      }
      if (!state.finishedAt && saved.finishedAt) {
        state.finishedAt = saved.finishedAt;
      }
    }
  } catch {
    // Ignore DB read error
  }
}

async function savePersistedRunState(state: OutlookScanRunState): Promise<void> {
  try {
    const payload = JSON.stringify({
      lastScheduledSlotKey: state.lastScheduledSlotKey,
      lastRunSummary: state.lastRunSummary,
      finishedAt: state.finishedAt,
    });
    await prisma.appSetting.upsert({
      where: { key: RUN_STATE_KEY },
      create: { key: RUN_STATE_KEY, value: payload },
      update: { value: payload },
    });
  } catch {
    // Ignore DB write error
  }
}

export async function getOutlookSchedulerStatus(now = new Date()) {
  const config = await loadOutlookScheduleConfig();
  const state = getScanState();
  await syncPersistedRunState(state);
  const slots = computeOutlookScheduleSlots(config);
  const { dateKey, timeHm } = getZonedDateAndTime(now, config.timezone);

  let nextSlot = slots[0] || "00:00";
  let dayLabel = "Ngày mai";
  for (const s of slots) {
    if (s > timeHm) {
      nextSlot = s;
      dayLabel = "Hôm nay";
      break;
    }
  }

  const nextRunDisplay = !config.enabled
    ? "Đã tắt (Disabled)"
    : `${dayLabel} lúc ${nextSlot} (${
        config.scheduleMode === "CUSTOM"
          ? `Mốc cố định`
          : `Lặp mỗi ${config.intervalHours} giờ`
      })`;

  const recentForwarded = await listRecentForwardedEmails(20);

  return {
    config,
    slots,
    nextSlot,
    dayLabel,
    nextRunDisplay,
    currentTimeHm: timeHm,
    state,
    recentForwarded,
  };
}

export function requestStopOutlookScan(): void {
  const state = getScanState();
  if (state.isRunning) {
    state.stopRequested = true;
    appendScanLog("WARN", "SCHEDULER", "Đã nhận lệnh dừng phiên quét từ người dùng...");
  }
}

/**
 * Execute a full scan across all LIVE KDP Hotmail accounts in range (e.g. KDP-034 .. KDP-126)
 * and forward any newly received emails within `lookbackHours` (e.g. 24h) to `forwardToEmail`.
 */
export async function executeOutlookScheduledScan(
  triggerSource: "SCHEDULE" | "MANUAL" = "MANUAL"
): Promise<OutlookScanRunState> {
  const state = getScanState();
  if (state.isRunning) {
    return state;
  }

  const config = await loadOutlookScheduleConfig();
  state.isRunning = true;
  state.stopRequested = false;
  state.triggerSource = triggerSource;
  state.startedAt = new Date().toISOString();
  state.finishedAt = null;
  state.currentAccountCode = "";
  state.currentEmail = "";
  state.processedCount = 0;
  state.skippedNoTokenCount = 0;
  state.newEmails24hCount = 0;
  state.alreadyForwardedCount = 0;
  state.forwardedCount = 0;
  state.errorCount = 0;

  appendScanLog(
    "INFO",
    "SCHEDULER",
    `Bắt đầu phiên quét Hotmail [${triggerSource}] (Lọc thư mới nhận trong ${config.lookbackHours}h -> Forward về ${config.forwardToEmail.replace("@", "+KDP-<số>@")})...`
  );

  try {
    const allMailboxes = await listOutlookMailboxes();

    const linkedMailboxes = allMailboxes.filter((mb) => !!mb.refreshToken);
    state.skippedNoTokenCount = allMailboxes.length - linkedMailboxes.length;
    state.totalToScan = linkedMailboxes.length;

    if (linkedMailboxes.length === 0) {
      state.lastRunSummary = `Có ${allMailboxes.length} tài khoản trong danh sách, nhưng chưa tài khoản nào có Refresh Token để quét.`;
      appendScanLog("WARN", "SCHEDULER", state.lastRunSummary);
      return state;
    }

    // Collect valid Graph access tokens as fallback relays for forwarding
    const graphRelayTokens: string[] = [];
    const lookbackMs = config.lookbackHours * 3600 * 1000;
    const nowMs = Date.now();

    for (const mb of linkedMailboxes) {
      if (state.stopRequested) {
        appendScanLog("WARN", "SCHEDULER", "Đã dừng phiên quét theo yêu cầu.");
        break;
      }

      const tag = mb.accountCode ? `${mb.accountCode} | ${mb.email}` : mb.email;
      const targetForwardEmail = buildKdpScheduleForwardAddress(
        config.forwardToEmail,
        mb.accountCode
      );
      state.currentAccountCode = mb.accountCode || "";
      state.currentEmail = mb.email;

      try {
        const tokenData = await refreshOutlookAccessToken(
          mb.refreshToken,
          mb.clientId,
          mb.proxy
        );
        if (tokenData.scope.includes("Mail.Send") || tokenData.scope.includes("graph.microsoft.com")) {
          if (!graphRelayTokens.includes(tokenData.accessToken)) {
            graphRelayTokens.push(tokenData.accessToken);
          }
        }

        const messages = await fetchOutlookEmails(
          mb.email,
          tokenData.accessToken,
          tokenData.scope,
          25,
          mb.proxy
        );

        const latestMsg = messages[0] || null;
        const latestOtpMsg = messages.find((m) => !!m.otpCode) || null;
        const expiresAt = new Date(Date.now() + (tokenData.expiresIn - 60) * 1000).toISOString();

        await updateMailboxTokens(mb.id, {
          refreshToken: tokenData.refreshToken || mb.refreshToken,
          accessToken: tokenData.accessToken,
          tokenExpiresAt: expiresAt,
          status: "ACTIVE",
          errorMessage: "",
          lastOtp: latestOtpMsg?.otpCode || mb.lastOtp || "",
          lastSubject: latestMsg?.subject || "",
          lastSender: latestMsg ? `${latestMsg.fromName} <${latestMsg.fromAddress}>` : "",
          lastEmailAt: latestMsg?.date || null,
          lastCheckedAt: new Date().toISOString(),
        });

        // Filter messages received within lookbackHours (e.g. 24h)
        const recentMessages = messages.filter((m) => {
          const msgMs = new Date(m.date).getTime();
          return !isNaN(msgMs) && nowMs - msgMs <= lookbackMs && nowMs - msgMs >= -3600000;
        });

        if (recentMessages.length > 0) {
          state.newEmails24hCount += recentMessages.length;
        } else {
          appendScanLog(
            "OK",
            tag,
            `Đã quét ${messages.length} thư (Không có thư mới trong ${config.lookbackHours}h qua).`
          );
        }

        // Forward each new message if not already forwarded
        if (config.autoForwardEnabled && config.forwardToEmail && recentMessages.length > 0) {
          let sentForThisMb = 0;
          let skippedForThisMb = 0;

          for (const msg of recentMessages) {
            const messageKey = `${mb.email.toLowerCase()}_${msg.id}`;
            const legacyKey = `${mb.accountCode || mb.email}_${msg.id}`;
            const alreadySent =
              (await hasEmailBeenForwarded(messageKey, mb.email, msg.subject, msg.date)) ||
              (await hasEmailBeenForwarded(legacyKey, mb.email, msg.subject, msg.date));

            if (alreadySent) {
              skippedForThisMb++;
              state.alreadyForwardedCount++;
              continue;
            }

            try {
              const res = await forwardOutlookMessage({
                mailbox: mb,
                message: msg,
                targetEmail: targetForwardEmail,
                accessToken: tokenData.accessToken,
                fallbackAccessTokens: graphRelayTokens,
                gmailSmtpUser: config.gmailSmtpUser,
                gmailSmtpAppPassword: config.gmailSmtpAppPassword,
              });

              await recordForwardedEmail({
                messageKey,
                accountCode: mb.accountCode || "",
                email: mb.email,
                folder: msg.folder,
                subject: msg.subject,
                fromAddress: msg.fromAddress,
                receivedAt: msg.date,
                forwardedTo: targetForwardEmail,
              });

              sentForThisMb++;
              state.forwardedCount++;
              appendScanLog(
                "OK",
                tag,
                `Đã Forward thư [${msg.folder}] "${msg.subject}" -> ${targetForwardEmail} (${res.method}).`
              );
            } catch (fwdErr: any) {
              appendScanLog(
                "ERROR",
                tag,
                `Lỗi Forward thư "${msg.subject}" -> ${targetForwardEmail}: ${fwdErr.message}`
              );
            }
          }

          if (skippedForThisMb > 0 && sentForThisMb === 0) {
            appendScanLog(
              "INFO",
              tag,
              `Có ${recentMessages.length} thư trong ${config.lookbackHours}h qua nhưng tất cả (${skippedForThisMb} thư) đều đã được Forward trước đó -> Bỏ qua không gửi lại.`
            );
          } else if (skippedForThisMb > 0 && sentForThisMb > 0) {
            appendScanLog(
              "INFO",
              tag,
              `Đã Forward ${sentForThisMb} thư mới -> ${targetForwardEmail} (Bỏ qua ${skippedForThisMb} thư đã Forward trước đó).`
            );
          }
        }
      } catch (err: any) {
        state.errorCount++;
        await updateMailboxTokens(mb.id, {
          status: "ERROR",
          errorMessage: err.message || "Lỗi quét hộp thư",
          lastCheckedAt: new Date().toISOString(),
        });
        appendScanLog("ERROR", tag, `Lỗi khi quét hộp thư: ${err.message}`);
      } finally {
        state.processedCount++;
      }
    }

    state.lastRunSummary = `Hoàn tất quét ${state.processedCount}/${state.totalToScan} tài khoản • Thư trong ${config.lookbackHours}h: ${state.newEmails24hCount} • Đã Forward mới: ${state.forwardedCount} • Đã bỏ qua (trùng): ${state.alreadyForwardedCount}.`;
    appendScanLog("OK", "SCHEDULER", state.lastRunSummary);
  } catch (err: any) {
    state.lastRunSummary = `Lỗi phiên quét: ${err.message}`;
    appendScanLog("ERROR", "SCHEDULER", state.lastRunSummary);
  } finally {
    state.isRunning = false;
    state.currentAccountCode = "";
    state.currentEmail = "";
    state.finishedAt = new Date().toISOString();
    await savePersistedRunState(state);
  }

  return state;
}

function hmToMinutes(hm: string): number {
  const [h, m] = (hm || "00:00").split(":");
  return (parseInt(h || "0", 10) || 0) * 60 + (parseInt(m || "0", 10) || 0);
}

/**
 * Tick function called every 25 seconds by the background timer (and on status poll).
 * Uses a 15-minute grace window after each slot so a slot is never missed even if enabled/polled a few minutes past the hour.
 */
export async function tickOutlookScheduler(now = new Date()): Promise<boolean> {
  const config = await loadOutlookScheduleConfig();
  if (!config.enabled) return false;

  const state = getScanState();
  await syncPersistedRunState(state);
  if (state.isRunning) return false;

  const slots = computeOutlookScheduleSlots(config);
  const { dateKey, timeHm } = getZonedDateAndTime(now, config.timezone);
  const nowMin = hmToMinutes(timeHm);

  let matchedSlot: string | null = null;
  for (const s of slots) {
    const slotMin = hmToMinutes(s);
    const diff = nowMin - slotMin;
    if (diff >= 0 && diff <= 15) {
      matchedSlot = s;
    }
  }

  if (!matchedSlot) {
    return false;
  }

  const slotKey = `${dateKey}@${matchedSlot}`;
  if (state.lastScheduledSlotKey === slotKey) {
    return false;
  }

  state.lastScheduledSlotKey = slotKey;
  executeOutlookScheduledScan("SCHEDULE").catch((err) => {
    console.error("[OutlookScheduler] Error in scheduled scan:", err);
  });
  return true;
}

/**
 * Ensure background scheduler timer is started and always runs the latest code version after HMR.
 */
export function ensureOutlookSchedulerStarted(): void {
  if (g.__outlookSchedulerTimer && g.__outlookSchedulerVersion === CURRENT_SCHEDULER_VERSION) {
    return;
  }
  if (g.__outlookSchedulerTimer) {
    clearInterval(g.__outlookSchedulerTimer);
  }
  g.__outlookSchedulerVersion = CURRENT_SCHEDULER_VERSION;
  g.__outlookSchedulerTimer = setInterval(() => {
    tickOutlookScheduler().catch(() => {
      // Ignore background tick errors
    });
  }, 25000);
}
