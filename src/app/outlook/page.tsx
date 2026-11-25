"use client";

import React, { useEffect, useState, useCallback } from "react";
import {
  Mail,
  KeyRound,
  RefreshCw,
  Plus,
  Trash2,
  Copy,
  Check,
  ExternalLink,
  ShieldAlert,
  CheckCircle2,
  AlertTriangle,
  Inbox,
  Sparkles,
  Eye,
  EyeOff,
  Bot,
  Link2,
  Search,
  Clock,
  FolderOpen,
  Terminal,
  CalendarClock,
  Send,
  Play,
  Square,
  Settings2,
} from "lucide-react";

interface OutlookMailbox {
  id: string;
  accountCode: string;
  email: string;
  password: string;
  proxy: string;
  note: string;
  clientId: string;
  refreshToken: string;
  status: "UNLINKED" | "ACTIVE" | "ERROR";
  errorMessage: string;
  ruleStatus?: "NONE" | "SUCCESS" | "VERIFY_REQUIRED" | "ERROR";
  ruleErrorMessage?: string;
  ruleUpdatedAt?: string | null;
  lastOtp: string;
  lastSubject: string;
  lastSender: string;
  lastEmailAt: string | null;
  lastCheckedAt: string | null;
  forwardedCount?: number;
  lastForwardedAt?: string | null;
}

interface OutlookMessage {
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
  isForwarded?: boolean;
  forwardedTo?: string;
  forwardedAt?: string;
}

interface ScheduleConfig {
  enabled: boolean;
  scheduleMode: "INTERVAL" | "CUSTOM";
  dailyTime: string;
  intervalHours: number;
  customScheduleSlots: string;
  timezone: string;
  fromAccountNum: number;
  toAccountNum: number;
  lookbackHours: number;
  autoForwardEnabled: boolean;
  forwardToEmail: string;
  gmailSmtpUser: string;
  gmailSmtpAppPassword: string;
}

export default function OutlookModulePage() {
  const [mailboxes, setMailboxes] = useState<OutlookMailbox[]>([]);
  const [loadingList, setLoadingList] = useState(true);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [syncingKdp, setSyncingKdp] = useState(false);

  // Mail reading state
  const [messages, setMessages] = useState<OutlookMessage[]>([]);
  const [messagesCache, setMessagesCache] = useState<Record<string, OutlookMessage[]>>({});
  const [selectedMessage, setSelectedMessage] = useState<OutlookMessage | null>(null);
  const [checkingId, setCheckingId] = useState<string | null>(null);
  const [folderFilter, setFolderFilter] = useState<"ALL" | "INBOX" | "Junk" | "OTP">("ALL");
  const [mailLimit, setMailLimit] = useState<number>(25);
  const [forwardingMsgId, setForwardingMsgId] = useState<string | null>(null);
  const initialLoadedRef = React.useRef(false);

  // Token linking & Batch Auto-Token state
  const [linkingMailbox, setLinkingMailbox] = useState<OutlookMailbox | null>(null);
  const [authUrl, setAuthUrl] = useState("");
  const [redirectUrlInput, setRedirectUrlInput] = useState("");
  const [isExchanging, setIsExchanging] = useState(false);
  const [autoBrowserId, setAutoBrowserId] = useState<string | null>(null);
  const [batchTokenRunning, setBatchTokenRunning] = useState(false);
  const batchTokenStopRef = React.useRef(false);

  // Bulk Setup Rule & Forwarding inside Hotmail Web (Puppeteer)
  const [ruleSetupId, setRuleSetupId] = useState<string | null>(null);
  const [batchRuleRunning, setBatchRuleRunning] = useState(false);
  const batchRuleStopRef = React.useRef(false);
  const [isRuleConfigOpen, setIsRuleConfigOpen] = useState(false);
  const [ruleBaseForwardEmail, setRuleBaseForwardEmail] = useState("ip.kdpmanagement@gmail.com");
  const [ruleNameInput, setRuleNameInput] = useState("KDP");
  const [ruleKeywordsInput, setRuleKeywordsInput] = useState("Amazon, Kindle");
  const [clearingAll, setClearingAll] = useState(false);

  // Schedule & Auto-Forward state
  const [scheduleStatus, setScheduleStatus] = useState<any>(null);
  const [isScheduleModalOpen, setIsScheduleModalOpen] = useState(false);
  const [scheduleForm, setScheduleForm] = useState<ScheduleConfig | null>(null);
  const [scheduleBulkLines, setScheduleBulkLines] = useState("");
  const [scheduleReplaceExisting, setScheduleReplaceExisting] = useState(false);
  const [savingSchedule, setSavingSchedule] = useState(false);

  // Import modal state
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [rawImportText, setRawImportText] = useState("");
  const [importReplaceExisting, setImportReplaceExisting] = useState(false);
  const [importNote, setImportNote] = useState("");
  const [isImporting, setIsImporting] = useState(false);

  // UI feedback
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [showPassMap, setShowPassMap] = useState<Record<string, boolean>>({});
  const [toast, setToast] = useState<{ type: "success" | "error" | "info"; text: string } | null>(
    null
  );

  // Activity & Error Logs state (chronological: oldest at top -> newest at bottom)
  const [activityLogs, setActivityLogs] = useState<string[]>([]);
  const [logFilter, setLogFilter] = useState<"ALL" | "ERROR">("ALL");
  const logContainerRef = React.useRef<HTMLDivElement | null>(null);
  const lastSyncedScanLogLenRef = React.useRef(0);

  const appendLogs = useCallback((lines: string | string[]) => {
    const arr = Array.isArray(lines) ? lines : [lines];
    if (arr.length === 0) return;
    setActivityLogs((prev) => [...prev, ...arr].slice(-300));
  }, []);

  useEffect(() => {
    if (logContainerRef.current) {
      logContainerRef.current.scrollTop = logContainerRef.current.scrollHeight;
    }
  }, [activityLogs]);

  const addClientLog = useCallback(
    (level: "INFO" | "OK" | "WARN" | "ERROR", tag: string, msg: string) => {
      const time = new Date().toLocaleTimeString("vi-VN", { hour12: false });
      appendLogs(`[${time}] [${level}] [${tag}] ${msg}`);
    },
    [appendLogs]
  );

  const showNotice = (type: "success" | "error" | "info", text: string) => {
    setToast({ type, text });
    setTimeout(() => setToast(null), 5000);
  };

  const copyText = (text: string, key: string) => {
    if (!text) return;
    navigator.clipboard.writeText(text);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey(null), 1800);
  };

  const formatMbTag = (mb: OutlookMailbox) =>
    mb.accountCode ? `${mb.accountCode} | ${mb.email}` : mb.email;

  const fetchMailboxes = useCallback(async () => {
    try {
      setLoadingList(true);
      const res = await fetch("/api/outlook/accounts", { cache: "no-store" });
      const data = await res.json();
      if (data.success && Array.isArray(data.mailboxes)) {
        setMailboxes(data.mailboxes);
        if (!selectedId && data.mailboxes.length > 0) {
          setSelectedId(data.mailboxes[0].id);
        }
      }
    } catch (err: any) {
      showNotice("error", err.message || "Không thể tải danh sách mail");
    } finally {
      setLoadingList(false);
    }
  }, [selectedId]);

  const fetchScheduleStatus = useCallback(async () => {
    try {
      const res = await fetch("/api/outlook/schedule", { cache: "no-store" });
      const data = await res.json();
      if (data.success) {
        setScheduleStatus(data);
        setScheduleForm((prev) => prev || data.config);
        const serverLogs: string[] = data.state?.logs || [];
        if (serverLogs.length > lastSyncedScanLogLenRef.current) {
          const newLines = serverLogs.slice(lastSyncedScanLogLenRef.current);
          lastSyncedScanLogLenRef.current = serverLogs.length;
          appendLogs(newLines);
        } else if (serverLogs.length < lastSyncedScanLogLenRef.current) {
          lastSyncedScanLogLenRef.current = serverLogs.length;
        }
      }
    } catch {
      // Ignore background poll error
    }
  }, [appendLogs]);

  useEffect(() => {
    fetchMailboxes();
    fetchScheduleStatus();
  }, [fetchMailboxes, fetchScheduleStatus]);

  // Poll schedule status every 3s (and refresh mailbox list when scan finishes)
  const prevScanRunningRef = React.useRef(false);
  useEffect(() => {
    const timer = setInterval(() => {
      fetchScheduleStatus();
    }, 3000);
    return () => clearInterval(timer);
  }, [fetchScheduleStatus]);

  useEffect(() => {
    const isNowRunning = !!scheduleStatus?.state?.isRunning;
    if (prevScanRunningRef.current && !isNowRunning) {
      fetchMailboxes();
    }
    prevScanRunningRef.current = isNowRunning;
  }, [scheduleStatus?.state?.isRunning, fetchMailboxes]);

  const selectedMailbox = mailboxes.find((m) => m.id === selectedId) || null;

  // Sync 31 LIVE KDP Hotmail accounts (KDP-034 .. KDP-126)
  const handleSyncLiveKdp = async () => {
    try {
      setSyncingKdp(true);
      const fromNum = scheduleStatus?.config?.fromAccountNum || 34;
      const toNum = scheduleStatus?.config?.toAccountNum || 126;
      addClientLog(
        "INFO",
        "SYNC-KDP",
        `Đang đồng bộ các tài khoản LIVE có Hotmail từ KDP-${String(fromNum).padStart(
          3,
          "0"
        )} đến KDP-${String(toNum).padStart(3, "0")}...`
      );
      const res = await fetch("/api/outlook/accounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "sync_kdp",
          fromNum,
          toNum,
          removeNonMatching: true,
        }),
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || "Đồng bộ KDP thất bại");
      }
      setMailboxes(data.mailboxes || []);
      if (data.mailboxes?.length > 0) {
        setSelectedId(data.mailboxes[0].id);
      }
      addClientLog(
        "OK",
        "SYNC-KDP",
        `Đã đồng bộ ${data.syncedCount} tài khoản LIVE Hotmail (KDP-${String(fromNum).padStart(
          3,
          "0"
        )}..KDP-${String(toNum).padStart(3, "0")})!`
      );
      showNotice(
        "success",
        `Đã đồng bộ ${data.syncedCount} tài khoản LIVE có Hotmail (KDP-${String(fromNum).padStart(
          3,
          "0"
        )} → KDP-${String(toNum).padStart(3, "0")})!`
      );
    } catch (err: any) {
      addClientLog("ERROR", "SYNC-KDP", err.message);
      showNotice("error", err.message);
    } finally {
      setSyncingKdp(false);
    }
  };

  // Open manual OAuth2 link modal
  const handleOpenManualLink = async (mb: OutlookMailbox) => {
    setLinkingMailbox(mb);
    setRedirectUrlInput("");
    setAuthUrl("");
    try {
      const res = await fetch(`/api/outlook/auth?id=${encodeURIComponent(mb.id)}`);
      const data = await res.json();
      if (data.success && data.authUrl) {
        setAuthUrl(data.authUrl);
      }
    } catch (err: any) {
      showNotice("error", err.message || "Lỗi tạo link OAuth2");
    }
  };

  // Exchange pasted localhost/?code=... URL for refresh_token
  const handleConfirmManualCode = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!linkingMailbox || !redirectUrlInput.trim()) return;

    try {
      setIsExchanging(true);
      const res = await fetch("/api/outlook/auth", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id: linkingMailbox.id,
          mode: "manual",
          redirectUrlOrCode: redirectUrlInput.trim(),
        }),
      });
      const data = await res.json();
      if (Array.isArray(data.logs)) {
        appendLogs(data.logs);
      }
      if (!res.ok || !data.success) {
        if (data.mailbox) {
          setMailboxes((prev) =>
            prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
          );
        }
        throw new Error(data.error || "Không thể đổi mã code lấy Token.");
      }

      if (data.mailbox) {
        setMailboxes((prev) =>
          prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
        );
      }
      showNotice("success", `Đã kết nối Token thành công cho ${formatMbTag(linkingMailbox)}!`);
      setLinkingMailbox(null);
      setRedirectUrlInput("");
      await fetchMailboxes();
      if (data.mailbox) {
        handleCheckMail(data.mailbox);
      }
    } catch (err: any) {
      addClientLog("ERROR", formatMbTag(linkingMailbox), err.message);
      showNotice("error", err.message);
    } finally {
      setIsExchanging(false);
    }
  };

  // Auto Browser Token Capture via Puppeteer (Single Account)
  const handleAutoBrowserToken = async (mb: OutlookMailbox): Promise<boolean> => {
    try {
      setAutoBrowserId(mb.id);
      setSelectedId(mb.id);
      addClientLog(
        "INFO",
        formatMbTag(mb),
        "Bắt đầu tiến trình Lấy Token Tự Động (Mở trình duyệt Chrome/Edge)..."
      );
      showNotice(
        "info",
        `Đang mở trình duyệt cho [${mb.accountCode || "MAIL"}] ${mb.email}...`
      );
      const res = await fetch("/api/outlook/auth", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id: mb.id,
          mode: "auto_browser",
        }),
      });
      const data = await res.json();
      if (Array.isArray(data.logs)) {
        appendLogs(data.logs);
      }
      if (!res.ok || !data.success) {
        if (data.mailbox) {
          setMailboxes((prev) =>
            prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
          );
        }
        throw new Error(data.error || "Lấy Token tự động thất bại.");
      }

      if (data.mailbox) {
        setMailboxes((prev) =>
          prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
        );
      }
      showNotice("success", `Đã tự động lấy Refresh Token thành công cho ${formatMbTag(mb)}!`);
      await fetchMailboxes();
      if (data.mailbox && !batchTokenRunning) {
        handleCheckMail(data.mailbox);
      }
      return true;
    } catch (err: any) {
      showNotice("error", err.message);
      return false;
    } finally {
      setAutoBrowserId(null);
    }
  };

  // Batch Auto-Token for all UNLINKED accounts sequentially
  const handleBatchAutoToken = async () => {
    const unlinked = mailboxes.filter((m) => !m.refreshToken);
    if (unlinked.length === 0) {
      showNotice("info", "Tất cả tài khoản trong danh sách đều đã có Refresh Token!");
      return;
    }

    setBatchTokenRunning(true);
    batchTokenStopRef.current = false;
    addClientLog(
      "INFO",
      "BATCH-TOKEN",
      `Bắt đầu lấy Token tự động hàng loạt cho ${unlinked.length} tài khoản chưa có Token...`
    );

    let okCount = 0;
    for (let i = 0; i < unlinked.length; i++) {
      if (batchTokenStopRef.current) {
        addClientLog("WARN", "BATCH-TOKEN", "Đã dừng tiến trình lấy Token hàng loạt theo yêu cầu.");
        break;
      }
      const mb = unlinked[i];
      addClientLog(
        "INFO",
        "BATCH-TOKEN",
        `[${i + 1}/${unlinked.length}] Đang lấy Token cho ${formatMbTag(mb)}...`
      );
      const ok = await handleAutoBrowserToken(mb);
      if (ok) okCount++;
      await new Promise((r) => setTimeout(r, 1200));
    }

    setBatchTokenRunning(false);
    addClientLog(
      "OK",
      "BATCH-TOKEN",
      `Hoàn tất tiến trình lấy Token hàng loạt: Thành công ${okCount}/${unlinked.length} tài khoản.`
    );
    showNotice("success", `Đã lấy Token xong cho ${okCount}/${unlinked.length} tài khoản!`);
  };

  // Setup Rule & Forwarding inside Hotmail Web (Single Account)
  const handleSetupRuleSingle = async (
    mb: OutlookMailbox
  ): Promise<"SUCCESS" | "VERIFY_REQUIRED" | "ERROR"> => {
    try {
      setRuleSetupId(mb.id);
      setSelectedId(mb.id);
      const keywords = ruleKeywordsInput
        .split(",")
        .map((k) => k.trim())
        .filter(Boolean);

      addClientLog(
        "INFO",
        formatMbTag(mb),
        `Bắt đầu cài đặt Rule (${ruleNameInput}: [${keywords.join(", ")}]) & Forwarding (${ruleBaseForwardEmail}) trong Hotmail Web...`
      );
      showNotice(
        "info",
        `Đang mở trình duyệt cài Rule & Forwarding cho [${mb.accountCode || "MAIL"}] ${mb.email}...`
      );

      const res = await fetch("/api/outlook/rules", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id: mb.id,
          baseForwardEmail: ruleBaseForwardEmail.trim() || "ip.kdpmanagement@gmail.com",
          ruleName: ruleNameInput.trim() || "KDP",
          keywords: keywords.length > 0 ? keywords : ["Amazon", "Kindle"],
        }),
      });
      const data = await res.json();
      if (Array.isArray(data.logs)) {
        appendLogs(data.logs);
      }
      if (data.mailbox) {
        setMailboxes((prev) =>
          prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
        );
      }

      if (data.ruleStatus === "VERIFY_REQUIRED") {
        showNotice(
          "error",
          `[${mb.accountCode || mb.email}] Bắt nhập mã từ mail khôi phục — Đã tự động bỏ qua và ghi nhận trạng thái!`
        );
        return "VERIFY_REQUIRED";
      }

      if (!res.ok || !data.success) {
        throw new Error(data.error || "Cài đặt Rule & Forwarding thất bại.");
      }

      showNotice(
        "success",
        `Đã cài Rule (${data.ruleForwardAddress}) & Forwarding (${data.forwardingAddress}) thành công cho ${formatMbTag(mb)}!`
      );
      return "SUCCESS";
    } catch (err: any) {
      showNotice("error", err.message);
      return "ERROR";
    } finally {
      setRuleSetupId(null);
    }
  };

  // Batch Setup Rule & Forwarding for all accounts not yet SUCCESS
  const handleBatchSetupRule = async (retryAll = false) => {
    const targets = retryAll
      ? mailboxes
      : mailboxes.filter((m) => m.ruleStatus !== "SUCCESS");

    if (targets.length === 0) {
      showNotice(
        "info",
        "Tất cả tài khoản trong danh sách đều đã cài đặt Rule & Forwarding thành công!"
      );
      return;
    }

    setBatchRuleRunning(true);
    batchRuleStopRef.current = false;
    addClientLog(
      "INFO",
      "BATCH-RULE",
      `Bắt đầu chạy Setup Rule & Forwarding hàng loạt cho ${targets.length} tài khoản Hotmail (Gặp tài khoản bắt mã mail khôi phục sẽ tự động bỏ qua)...`
    );

    let okCount = 0;
    let verifySkipCount = 0;
    let errCount = 0;

    for (let i = 0; i < targets.length; i++) {
      if (batchRuleStopRef.current) {
        addClientLog("WARN", "BATCH-RULE", "Đã dừng tiến trình Setup Rule & Forwarding hàng loạt theo yêu cầu.");
        break;
      }
      const mb = targets[i];
      addClientLog(
        "INFO",
        "BATCH-RULE",
        `[${i + 1}/${targets.length}] Đang xử lý ${formatMbTag(mb)}...`
      );
      const status = await handleSetupRuleSingle(mb);
      if (status === "SUCCESS") okCount++;
      else if (status === "VERIFY_REQUIRED") verifySkipCount++;
      else errCount++;

      await new Promise((r) => setTimeout(r, 1200));
    }

    setBatchRuleRunning(false);
    await fetchMailboxes();
    const summaryMsg = `Hoàn tất Setup Rule & Forward hàng loạt: Thành công ${okCount}/${targets.length} • Bắt mã khôi phục (Đã bỏ qua): ${verifySkipCount} • Lỗi khác: ${errCount}.`;
    addClientLog("OK", "BATCH-RULE", summaryMsg);
    showNotice("success", summaryMsg);
  };

  // Clear all mailboxes in the list
  const handleClearAllMailboxes = async () => {
    if (mailboxes.length === 0) return;
    if (
      !confirm(
        `Bạn có chắc chắn muốn xóa toàn bộ ${mailboxes.length} tài khoản khỏi danh sách hiện tại không?`
      )
    ) {
      return;
    }
    try {
      setClearingAll(true);
      const res = await fetch("/api/outlook/accounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "clear_all" }),
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || "Không thể xóa danh sách.");
      }
      setMailboxes([]);
      setSelectedId(null);
      setMessages([]);
      setSelectedMessage(null);
      addClientLog("OK", "CLEAR-LIST", `Đã xóa toàn bộ ${data.deletedCount || 0} tài khoản khỏi danh sách.`);
      showNotice("info", `Đã xóa sạch danh sách (${data.deletedCount || 0} tài khoản).`);
    } catch (err: any) {
      showNotice("error", err.message);
    } finally {
      setClearingAll(false);
    }
  };

  // Check Mail via OAuth2 API
  const handleCheckMail = async (mb: OutlookMailbox, customLimit?: number) => {
    setSelectedId(mb.id);
    if (!mb.refreshToken) {
      handleOpenManualLink(mb);
      return;
    }

    const limitToUse = customLimit ?? mailLimit;

    try {
      setCheckingId(mb.id);
      addClientLog(
        "INFO",
        formatMbTag(mb),
        `Gửi yêu cầu Check Mail qua OAuth2 API (Số lượng: ${limitToUse} thư/mục)...`
      );
      const res = await fetch("/api/outlook/check", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: mb.id, limit: limitToUse }),
      });
      const data = await res.json();
      if (Array.isArray(data.logs)) {
        appendLogs(data.logs);
      }
      if (!res.ok || !data.success) {
        if (data.mailbox) {
          setMailboxes((prev) =>
            prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
          );
        }
        throw new Error(data.error || "Lỗi khi đọc hòm thư Outlook.");
      }

      const fetchedMessages: OutlookMessage[] = data.messages || [];
      setMessages(fetchedMessages);
      setMessagesCache((prev) => ({ ...prev, [mb.id]: fetchedMessages }));
      setSelectedMessage(fetchedMessages[0] || null);
      if (data.mailbox) {
        setMailboxes((prev) =>
          prev.map((item) => (item.id === data.mailbox.id ? data.mailbox : item))
        );
      }
      showNotice(
        "success",
        `[${mb.accountCode || "MAIL"}] Đã tải ${fetchedMessages.length} thư từ ${mb.email}${
          data.latestOtp ? ` • OTP: ${data.latestOtp}` : ""
        }`
      );
    } catch (err: any) {
      showNotice("error", err.message);
    } finally {
      setCheckingId(null);
    }
  };

  // Forward a single selected message immediately to ip.kdpmanagement+KDP-<số thứ tự>@gmail.com
  const handleForwardSingleMessage = async (mb: OutlookMailbox, msg: OutlookMessage) => {
    try {
      setForwardingMsgId(msg.id);
      const baseEmail =
        scheduleStatus?.config?.forwardToEmail || "ip.kdpmanagement@gmail.com";
      const previewTarget = mb.accountCode
        ? baseEmail.replace("@", `+${mb.accountCode}@`)
        : baseEmail;
      addClientLog(
        "INFO",
        formatMbTag(mb),
        `Đang Forward thư "${msg.subject}" -> ${previewTarget}...`
      );
      const res = await fetch("/api/outlook/schedule", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "forward_single",
          mailboxId: mb.id,
          message: msg,
          targetEmail: baseEmail,
        }),
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || "Forward email thất bại");
      }
      setScheduleStatus(data);
      const actualForwardedTo = data.forwardedTo || previewTarget;
      setMessages((prev) =>
        prev.map((m) =>
          m.id === msg.id
            ? {
                ...m,
                isForwarded: true,
                forwardedTo: actualForwardedTo,
                forwardedAt: new Date().toISOString(),
              }
            : m
        )
      );
      if (selectedMessage?.id === msg.id) {
        setSelectedMessage((prev) =>
          prev
            ? {
                ...prev,
                isForwarded: true,
                forwardedTo: actualForwardedTo,
                forwardedAt: new Date().toISOString(),
              }
            : null
        );
      }
      await fetchMailboxes();
      addClientLog(
        "OK",
        formatMbTag(mb),
        `Đã Forward thành công thư "${msg.subject}" -> ${actualForwardedTo} (${data.method})!`
      );
      showNotice("success", data.message);
    } catch (err: any) {
      addClientLog("ERROR", formatMbTag(mb), `Lỗi Forward: ${err.message}`);
      showNotice("error", err.message);
    } finally {
      setForwardingMsgId(null);
    }
  };

  // Trigger manual scheduled scan across all accounts with Refresh Token
  const handleTriggerScheduledScan = async () => {
    try {
      const res = await fetch("/api/outlook/schedule", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "run_now" }),
      });
      const data = await res.json();
      if (data.success) {
        setScheduleStatus(data);
        showNotice("info", data.message);
      }
    } catch (err: any) {
      showNotice("error", err.message);
    }
  };

  const handleStopScheduledScan = async () => {
    try {
      const res = await fetch("/api/outlook/schedule", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "stop" }),
      });
      const data = await res.json();
      if (data.success) {
        setScheduleStatus(data);
        showNotice("info", data.message);
      }
    } catch (err: any) {
      showNotice("error", err.message);
    }
  };

  const handleSaveScheduleConfig = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!scheduleForm) return;
    try {
      setSavingSchedule(true);
      const res = await fetch("/api/outlook/schedule", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...scheduleForm,
          rawLines: scheduleBulkLines.trim() || undefined,
          replaceExisting: scheduleReplaceExisting,
        }),
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || "Lỗi lưu cấu hình Schedule");
      }
      setScheduleStatus(data);
      setScheduleForm(data.config);
      if (Array.isArray(data.mailboxes)) {
        setMailboxes(data.mailboxes);
        if (data.mailboxes.length > 0 && !selectedId) {
          setSelectedId(data.mailboxes[0].id);
        }
      }
      setScheduleBulkLines("");
      setScheduleReplaceExisting(false);
      setIsScheduleModalOpen(false);
      showNotice(
        "success",
        data.importedCount > 0
          ? `Đã lưu cấu hình Schedule & nạp ${data.importedCount} tài khoản mail (KDP + Token)!`
          : "Đã lưu cấu hình Schedule & Auto-Forward thành công!"
      );
    } catch (err: any) {
      showNotice("error", err.message);
    } finally {
      setSavingSchedule(false);
    }
  };

  // Click on a mailbox in the left list -> switch view & auto-load emails if token exists
  const handleSelectMailbox = (mb: OutlookMailbox) => {
    setSelectedId(mb.id);
    const cached = messagesCache[mb.id];
    if (cached) {
      setMessages(cached);
      setSelectedMessage(cached[0] || null);
      return;
    }
    setMessages([]);
    setSelectedMessage(null);
    if (mb.refreshToken && checkingId !== mb.id) {
      handleCheckMail(mb);
    }
  };

  // Auto-load emails for the first selected mailbox with a token on initial page load
  useEffect(() => {
    if (!initialLoadedRef.current && mailboxes.length > 0 && selectedId) {
      const firstMb = mailboxes.find((m) => m.id === selectedId);
      if (firstMb) {
        initialLoadedRef.current = true;
        if (firstMb.refreshToken) {
          handleCheckMail(firstMb);
        }
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mailboxes, selectedId]);

  // Bulk Import Mailboxes
  const handleImportSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!rawImportText.trim()) return;
    try {
      setIsImporting(true);
      const res = await fetch("/api/outlook/accounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rawLines: rawImportText,
          replaceExisting: importReplaceExisting,
          note: importNote,
        }),
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || "Lỗi khi thêm danh sách mail.");
      }
      setMailboxes(data.mailboxes || []);
      if (data.mailboxes?.length > 0) {
        setSelectedId(data.mailboxes[0].id);
      }
      setIsImportOpen(false);
      setRawImportText("");
      setImportReplaceExisting(false);
      setImportNote("");
      showNotice("success", `Đã lưu ${data.importedCount || 1} tài khoản Outlook!`);
    } catch (err: any) {
      showNotice("error", err.message);
    } finally {
      setIsImporting(false);
    }
  };

  const handleDeleteMailbox = async (mb: OutlookMailbox) => {
    if (!confirm(`Xóa tài khoản ${formatMbTag(mb)} khỏi danh sách?`)) return;
    try {
      const res = await fetch(`/api/outlook/accounts?id=${encodeURIComponent(mb.id)}`, {
        method: "DELETE",
      });
      const data = await res.json();
      if (data.success) {
        setMailboxes(data.mailboxes || []);
        if (selectedId === mb.id) {
          setSelectedId(data.mailboxes?.[0]?.id || null);
          setMessages([]);
          setSelectedMessage(null);
        }
        showNotice("info", `Đã xóa ${formatMbTag(mb)}`);
      }
    } catch (err: any) {
      showNotice("error", err.message);
    }
  };

  const filteredMailboxes = mailboxes.filter((m) => {
    if (!searchQuery.trim()) return true;
    const q = searchQuery.toLowerCase();
    return (
      (m.accountCode || "").toLowerCase().includes(q) ||
      m.email.toLowerCase().includes(q) ||
      (m.note || "").toLowerCase().includes(q) ||
      (m.lastSubject || "").toLowerCase().includes(q)
    );
  });

  const filteredMessages = messages.filter((msg) => {
    if (folderFilter === "INBOX") return msg.folder === "INBOX";
    if (folderFilter === "Junk") return msg.folder === "Junk";
    if (folderFilter === "OTP") return Boolean(msg.otpCode);
    return true;
  });

  const linkedCount = mailboxes.filter((m) => !!m.refreshToken).length;
  const unlinkedCount = mailboxes.length - linkedCount;
  const ruleSuccessCount = mailboxes.filter((m) => m.ruleStatus === "SUCCESS").length;
  const ruleVerifySkipCount = mailboxes.filter((m) => m.ruleStatus === "VERIFY_REQUIRED").length;
  const ruleErrorCount = mailboxes.filter((m) => m.ruleStatus === "ERROR").length;
  const rulePendingCount = mailboxes.filter((m) => m.ruleStatus !== "SUCCESS").length;
  const isScanRunning = !!scheduleStatus?.state?.isRunning;

  return (
    <div className="container mx-auto max-w-7xl px-3 sm:px-6 py-6 space-y-5">
      {/* Header Banner */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 rounded-2xl border border-border bg-card p-5 shadow-sm">
        <div className="flex items-start gap-3.5">
          <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-blue-500 to-indigo-600 text-white shadow-md shadow-blue-500/20">
            <Mail className="h-5 w-5" />
          </div>
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <h1 className="text-lg sm:text-xl font-bold tracking-tight text-foreground">
                KDP Outlook / Hotmail Manager &amp; Bulk Rule Setup
              </h1>
              <span className="rounded-full bg-emerald-500/10 px-2.5 py-0.5 text-[11px] font-bold text-emerald-500 border border-emerald-500/20">
                {mailboxes.length} Tài khoản trong danh sách
              </span>
              <span className="rounded-full bg-teal-500/10 px-2.5 py-0.5 text-[11px] font-bold text-teal-500 border border-teal-500/20">
                Đã cài Rule &amp; Forward: {ruleSuccessCount}/{mailboxes.length}
              </span>
              <span className="rounded-full bg-blue-500/10 px-2.5 py-0.5 text-[11px] font-bold text-blue-500 border border-blue-500/20">
                Đã có Token: {linkedCount}/{mailboxes.length}
              </span>
            </div>
            <p className="text-xs text-muted-foreground mt-1">
              Tự động cài đặt <strong>Rule (Amazon, Kindle)</strong> &amp;{" "}
              <strong>Forwarding + POP/IMAP</strong> hàng loạt bên trong Hotmail Web về{" "}
              <code className="text-primary font-semibold">{ruleBaseForwardEmail}</code> (kèm{" "}
              <code>+KDP_XXX</code>). Gặp nick bắt mã mail khôi phục sẽ tự động bỏ qua và ghi log.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {mailboxes.length > 0 && (
            <button
              onClick={handleClearAllMailboxes}
              disabled={clearingAll || batchRuleRunning || batchTokenRunning}
              className="inline-flex items-center gap-1.5 rounded-xl border border-rose-500/30 bg-rose-500/10 hover:bg-rose-500/20 px-3 py-2 text-xs font-bold text-rose-500 shadow-sm transition-all disabled:opacity-50"
              title="Xóa toàn bộ danh sách email hiện tại"
            >
              <Trash2 className="h-3.5 w-3.5" />
              <span>{clearingAll ? "Đang xóa..." : "Xóa Danh Sách"}</span>
            </button>
          )}

          <button
            onClick={handleSyncLiveKdp}
            disabled={syncingKdp}
            className="inline-flex items-center gap-1.5 rounded-xl border border-border bg-card hover:bg-muted px-3 py-2 text-xs font-bold text-foreground shadow-sm transition-all disabled:opacity-50"
            title="Đồng bộ danh sách tài khoản LIVE có Hotmail từ KDP-034 đến KDP-126"
          >
            <RefreshCw className={`h-3.5 w-3.5 text-primary ${syncingKdp ? "animate-spin" : ""}`} />
            <span>{syncingKdp ? "Đang đồng bộ..." : "Đồng Bộ LIVE KDP"}</span>
          </button>

          {batchTokenRunning ? (
            <button
              onClick={() => {
                batchTokenStopRef.current = true;
              }}
              className="inline-flex items-center gap-1.5 rounded-xl bg-rose-500 hover:bg-rose-600 px-3.5 py-2 text-xs font-bold text-white shadow-sm transition-all"
            >
              <Square className="h-3.5 w-3.5" />
              <span>Dừng Lấy Token</span>
            </button>
          ) : (
            unlinkedCount > 0 && (
              <button
                onClick={handleBatchAutoToken}
                className="inline-flex items-center gap-1.5 rounded-xl bg-amber-500 hover:bg-amber-600 px-3.5 py-2 text-xs font-bold text-white shadow-sm transition-all"
                title="Tự động mở trình duyệt lấy Refresh Token lần lượt cho các tài khoản chưa có Token"
              >
                <Bot className="h-3.5 w-3.5" />
                <span>Lấy Token Hàng Loạt ({unlinkedCount})</span>
              </button>
            )
          )}

          <button
            onClick={() => setIsImportOpen(true)}
            className="inline-flex items-center gap-1.5 rounded-xl bg-primary px-3.5 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 transition-all"
          >
            <Plus className="h-3.5 w-3.5" />
            <span>Thêm / Import Mail</span>
          </button>
        </div>
      </div>

      {/* BULK SETUP RULE & FORWARDING INSIDE HOTMAIL WEB CONTROL BAR */}
      <div className="rounded-2xl border border-emerald-500/30 bg-card p-4 shadow-sm space-y-3">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex items-start gap-3">
            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-emerald-500/15 text-emerald-500 border border-emerald-500/30">
              <Sparkles className="h-4 w-4" />
            </div>
            <div className="space-y-1">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs sm:text-sm font-bold text-foreground">
                  Setup Rule &amp; Forwarding Hàng Loạt Bên Trong Hotmail (Browser + Proxy)
                </span>
                <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 px-2 py-0.5 text-[10px] font-bold text-emerald-500">
                  ● Thành công: {ruleSuccessCount}
                </span>
                {ruleVerifySkipCount > 0 && (
                  <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/15 border border-amber-500/30 px-2 py-0.5 text-[10px] font-bold text-amber-500">
                    ● Bắt mã khôi phục (Đã bỏ qua): {ruleVerifySkipCount}
                  </span>
                )}
                {ruleErrorCount > 0 && (
                  <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/15 border border-rose-500/30 px-2 py-0.5 text-[10px] font-bold text-rose-500">
                    ● Lỗi: {ruleErrorCount}
                  </span>
                )}
              </div>
              <p className="text-[11px] text-muted-foreground">
                <strong>1. Tab Rules:</strong> Tên <code>{ruleNameInput}</code> • Điều kiện{" "}
                <code>Message body includes [{ruleKeywordsInput}]</code> → Forward to{" "}
                <code className="text-emerald-500 font-semibold">
                  {ruleBaseForwardEmail.replace("@", "+KDP_XXX@")}
                </code>{" "}
                • <strong>2. Tab Forwarding &amp; IMAP:</strong> Forward to{" "}
                <code className="text-primary font-semibold">{ruleBaseForwardEmail}</code> (Không
                giữ bản sao, Bật POP &amp; IMAP).
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {batchRuleRunning ? (
              <button
                onClick={() => {
                  batchRuleStopRef.current = true;
                }}
                className="inline-flex items-center gap-1.5 rounded-xl bg-rose-500 hover:bg-rose-600 px-3.5 py-2 text-xs font-bold text-white shadow-sm"
              >
                <Square className="h-3.5 w-3.5" />
                <span>Dừng Setup Rule Hàng Loạt</span>
              </button>
            ) : (
              <button
                onClick={() => handleBatchSetupRule(false)}
                disabled={mailboxes.length === 0 || !!ruleSetupId}
                className="inline-flex items-center gap-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 px-3.5 py-2 text-xs font-bold text-white shadow-sm disabled:opacity-50"
              >
                <Play className="h-3.5 w-3.5" />
                <span>
                  Setup Rule &amp; Forward Hàng Loạt ({rulePendingCount} mail)
                </span>
              </button>
            )}

            <button
              onClick={() => setIsRuleConfigOpen((v) => !v)}
              className="inline-flex items-center gap-1.5 rounded-xl border border-border bg-card hover:bg-muted px-3 py-2 text-xs font-bold text-foreground"
            >
              <Settings2 className="h-3.5 w-3.5 text-emerald-500" />
              <span>Cấu Hình Rule</span>
            </button>
          </div>
        </div>

        {isRuleConfigOpen && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-border text-xs">
            <div>
              <label className="font-bold text-foreground block mb-1">
                Email Gmail nhận Forward chính:
              </label>
              <input
                type="email"
                value={ruleBaseForwardEmail}
                onChange={(e) => setRuleBaseForwardEmail(e.target.value)}
                placeholder="ip.kdpmanagement@gmail.com"
                className="w-full rounded-xl border border-border bg-muted/40 px-3 py-1.5 text-xs font-mono text-foreground"
              />
            </div>
            <div>
              <label className="font-bold text-foreground block mb-1">
                Tên Rule (Name your rule):
              </label>
              <input
                type="text"
                value={ruleNameInput}
                onChange={(e) => setRuleNameInput(e.target.value)}
                placeholder="KDP"
                className="w-full rounded-xl border border-border bg-muted/40 px-3 py-1.5 text-xs font-mono text-foreground"
              />
            </div>
            <div>
              <label className="font-bold text-foreground block mb-1">
                Từ khóa (Message body includes, cách bởi dấu phẩy):
              </label>
              <input
                type="text"
                value={ruleKeywordsInput}
                onChange={(e) => setRuleKeywordsInput(e.target.value)}
                placeholder="Amazon, Kindle"
                className="w-full rounded-xl border border-border bg-muted/40 px-3 py-1.5 text-xs font-mono text-foreground"
              />
            </div>
          </div>
        )}
      </div>

      {/* SCHEDULE & AUTO-FORWARD CONTROL BAR (Like KDP Standalone Report) */}
      <div className="rounded-2xl border border-indigo-500/30 bg-card p-4 shadow-sm space-y-3">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex items-start gap-3">
            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-indigo-500/15 text-indigo-500 border border-indigo-500/30">
              <CalendarClock className="h-4 w-4" />
            </div>
            <div className="space-y-0.5">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs sm:text-sm font-bold text-foreground">
                  Lịch Quét Tự Động (OAuth2 Token) &amp; Forward Mail (24h)
                </span>
                {scheduleStatus?.config?.enabled ? (
                  <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 px-2 py-0.5 text-[10px] font-bold text-emerald-500">
                    ● Đang Bật (Mỗi {scheduleStatus?.config?.intervalHours || 2} giờ)
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 rounded-full bg-muted px-2 py-0.5 text-[10px] font-bold text-muted-foreground">
                    Đã Tắt
                  </span>
                )}
                <span className="text-[11px] font-semibold text-primary">
                  Kế tiếp: {scheduleStatus?.nextRunDisplay || "Đang tải..."}
                </span>
                <span className="inline-flex items-center gap-1 rounded-full bg-indigo-500/10 border border-indigo-500/20 px-2 py-0.5 text-[10px] font-bold text-indigo-400">
                  ● Đã Forward chống trùng: {(scheduleStatus?.recentForwarded || []).length} thư gần đây
                </span>
              </div>
              <p className="text-[11px] text-muted-foreground">
                Mốc chạy trong ngày:{" "}
                <code className="text-foreground font-mono">
                  {(scheduleStatus?.slots || []).join(", ")}
                </code>{" "}
                • Nơi nhận Forward:{" "}
                <code className="text-indigo-400 font-bold">
                  {(scheduleStatus?.config?.forwardToEmail || "ip.kdpmanagement@gmail.com").replace(
                    "@",
                    "+KDP-<số thứ tự>@"
                  )}
                </code>{" "}
                •{" "}
                <span className="text-emerald-500 font-semibold">
                  Thư đã Forward được lưu vết vĩnh viễn, lần quét sau tự động bỏ qua.
                </span>
              </p>
              {scheduleStatus?.state?.lastRunSummary && (
                <p className="text-[11px] text-muted-foreground">
                  <span className="font-semibold text-foreground">Phiên gần nhất:</span>{" "}
                  {scheduleStatus.state.lastRunSummary}
                </p>
              )}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {isScanRunning ? (
              <button
                onClick={handleStopScheduledScan}
                className="inline-flex items-center gap-1.5 rounded-xl bg-rose-500 hover:bg-rose-600 px-3.5 py-2 text-xs font-bold text-white shadow-sm"
              >
                <Square className="h-3.5 w-3.5" />
                <span>
                  Dừng Quét ({scheduleStatus?.state?.processedCount || 0}/
                  {scheduleStatus?.state?.totalToScan || 0})
                </span>
              </button>
            ) : (
              <button
                onClick={handleTriggerScheduledScan}
                className="inline-flex items-center gap-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 px-3.5 py-2 text-xs font-bold text-white shadow-sm"
              >
                <Play className="h-3.5 w-3.5" />
                <span>Quét &amp; Forward Ngay (24h)</span>
              </button>
            )}

            <button
              onClick={() => {
                if (scheduleStatus?.config) {
                  setScheduleForm({ ...scheduleStatus.config });
                }
                setIsScheduleModalOpen(true);
              }}
              className="inline-flex items-center gap-1.5 rounded-xl border border-indigo-500/30 bg-indigo-500/10 hover:bg-indigo-500/20 px-3 py-2 text-xs font-bold text-indigo-400"
            >
              <Settings2 className="h-3.5 w-3.5 text-indigo-400" />
              <span>Cài Đặt Schedule &amp; Nạp Danh Sách Mail</span>
            </button>
          </div>
        </div>

        {isScanRunning && (
          <div className="rounded-xl bg-indigo-500/10 border border-indigo-500/20 px-3.5 py-2 text-xs flex items-center justify-between">
            <div className="flex items-center gap-2 text-indigo-400 font-semibold">
              <RefreshCw className="h-3.5 w-3.5 animate-spin" />
              <span>
                Đang quét:{" "}
                <strong>
                  {scheduleStatus?.state?.currentAccountCode}{" "}
                  {scheduleStatus?.state?.currentEmail}
                </strong>{" "}
                ({scheduleStatus?.state?.processedCount}/{scheduleStatus?.state?.totalToScan})
              </span>
            </div>
            <span className="text-[11px] font-bold text-emerald-400">
              Thư mới 24h: {scheduleStatus?.state?.newEmails24hCount || 0} • Đã Forward mới:{" "}
              {scheduleStatus?.state?.forwardedCount || 0} • Đã bỏ qua (trùng):{" "}
              {scheduleStatus?.state?.alreadyForwardedCount || 0}
            </span>
          </div>
        )}
      </div>

      {/* Toast Notification */}
      {toast && (
        <div
          className={`flex items-center justify-between gap-3 rounded-xl border px-4 py-3 text-xs font-semibold shadow-sm transition-all ${
            toast.type === "error"
              ? "border-rose-500/30 bg-rose-500/10 text-rose-500"
              : toast.type === "info"
              ? "border-blue-500/30 bg-blue-500/10 text-blue-500"
              : "border-emerald-500/30 bg-emerald-500/10 text-emerald-500"
          }`}
        >
          <div className="flex items-center gap-2">
            {toast.type === "error" ? (
              <AlertTriangle className="h-4 w-4 shrink-0" />
            ) : (
              <CheckCircle2 className="h-4 w-4 shrink-0" />
            )}
            <span>{toast.text}</span>
          </div>
          <button onClick={() => setToast(null)} className="text-xs underline opacity-75">
            Đóng
          </button>
        </div>
      )}

      {/* Main Grid: Left (Accounts List) + Right (Inbox & Message Reader) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* LEFT COLUMN: Mailboxes List (5 cols) */}
        <div className="lg:col-span-5 flex flex-col rounded-2xl border border-border bg-card shadow-sm overflow-hidden">
          <div className="border-b border-border p-4 space-y-3">
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-bold text-foreground flex items-center gap-2">
                <Inbox className="h-4 w-4 text-primary" />
                <span>Danh Sách Tài Khoản LIVE ({filteredMailboxes.length})</span>
              </h2>
              <button
                onClick={fetchMailboxes}
                className="text-xs text-muted-foreground hover:text-foreground flex items-center gap-1"
              >
                <RefreshCw className={`h-3.5 w-3.5 ${loadingList ? "animate-spin" : ""}`} />
                <span>Làm mới</span>
              </button>
            </div>

            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-muted-foreground" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Tìm theo mã KDP-034, email, tên tác giả..."
                className="w-full rounded-xl border border-border bg-muted/40 pl-9 pr-3 py-2 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:ring-1 focus:ring-primary"
              />
            </div>
          </div>

          <div className="divide-y divide-border max-h-[680px] overflow-y-auto">
            {filteredMailboxes.length === 0 ? (
              <div className="p-8 text-center text-xs text-muted-foreground">
                Không tìm thấy tài khoản nào khớp bộ lọc.
              </div>
            ) : (
              filteredMailboxes.map((mb) => {
                const isSelected = mb.id === selectedId;
                const isChecking = checkingId === mb.id;
                const isAutoRunning = autoBrowserId === mb.id;
                const showPass = !!showPassMap[mb.id];

                return (
                  <div
                    key={mb.id}
                    onClick={() => handleSelectMailbox(mb)}
                    className={`p-4 transition-colors cursor-pointer space-y-3 ${
                      isSelected ? "bg-primary/5" : "hover:bg-muted/40"
                    }`}
                  >
                    {/* Top row: KDP Code + Email + Status Badge */}
                    <div className="flex items-start justify-between gap-2">
                      <div className="min-w-0">
                        <div className="flex flex-wrap items-center gap-1.5">
                          {mb.accountCode && (
                            <span className="shrink-0 rounded-md bg-blue-600 px-2 py-0.5 text-[11px] font-extrabold text-white shadow-sm">
                              {mb.accountCode}
                            </span>
                          )}
                          <span className="font-bold text-xs sm:text-sm text-foreground truncate">
                            {mb.email}
                          </span>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              copyText(mb.email, `email_${mb.id}`);
                            }}
                            title="Copy Email"
                            className="p-1 rounded hover:bg-muted text-muted-foreground hover:text-foreground"
                          >
                            {copiedKey === `email_${mb.id}` ? (
                              <Check className="h-3.5 w-3.5 text-emerald-500" />
                            ) : (
                              <Copy className="h-3.5 w-3.5" />
                            )}
                          </button>
                        </div>

                        {mb.note && (
                          <p className="text-[11px] font-medium text-muted-foreground mt-0.5">
                            Tên KDP: <span className="text-foreground">{mb.note}</span>
                          </p>
                        )}

                        {mb.proxy && (
                          <div className="flex items-center gap-1.5 mt-0.5 text-[11px] text-muted-foreground">
                            <span>Proxy:</span>
                            <code className="rounded bg-indigo-500/10 border border-indigo-500/20 px-1.5 py-0.5 font-mono text-[10px] font-semibold text-indigo-400">
                              {mb.proxy.split(":").slice(0, 2).join(":")}
                            </code>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                copyText(mb.proxy, `proxy_${mb.id}`);
                              }}
                              className="p-0.5 hover:text-foreground"
                              title={`Copy full proxy: ${mb.proxy}`}
                            >
                              {copiedKey === `proxy_${mb.id}` ? (
                                <Check className="h-3 w-3 text-emerald-500" />
                              ) : (
                                <Copy className="h-3 w-3" />
                              )}
                            </button>
                          </div>
                        )}

                        {/* Password row */}
                        {mb.password && (
                          <div className="flex items-center gap-1.5 mt-1 text-[11px] text-muted-foreground">
                            <span>Pass:</span>
                            <code className="rounded bg-muted px-1.5 py-0.5 font-mono text-foreground">
                              {showPass ? mb.password : "••••••••••"}
                            </code>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setShowPassMap((p) => ({ ...p, [mb.id]: !p[mb.id] }));
                              }}
                              className="p-0.5 hover:text-foreground"
                              title={showPass ? "Ẩn mật khẩu" : "Hiện mật khẩu"}
                            >
                              {showPass ? (
                                <EyeOff className="h-3 w-3" />
                              ) : (
                                <Eye className="h-3 w-3" />
                              )}
                            </button>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                copyText(mb.password, `pass_${mb.id}`);
                              }}
                              className="p-0.5 hover:text-foreground"
                              title="Copy mật khẩu"
                            >
                              {copiedKey === `pass_${mb.id}` ? (
                                <Check className="h-3 w-3 text-emerald-500" />
                              ) : (
                                <Copy className="h-3 w-3" />
                              )}
                            </button>
                          </div>
                        )}
                      </div>

                      {/* Status badge */}
                      <div className="flex flex-col items-end gap-1 shrink-0">
                        <div className="flex items-center gap-1.5">
                          {mb.ruleStatus === "SUCCESS" ? (
                            <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/15 border border-emerald-500/30 px-2.5 py-0.5 text-[10px] font-bold text-emerald-500">
                              <CheckCircle2 className="h-3 w-3" />
                              Đã cài Rule &amp; Forward
                            </span>
                          ) : mb.ruleStatus === "VERIFY_REQUIRED" ? (
                            <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/15 border border-amber-500/30 px-2.5 py-0.5 text-[10px] font-bold text-amber-500">
                              <AlertTriangle className="h-3 w-3" />
                              Bắt mã khôi phục (Bỏ qua)
                            </span>
                          ) : mb.ruleStatus === "ERROR" ? (
                            <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/15 border border-rose-500/30 px-2.5 py-0.5 text-[10px] font-bold text-rose-500">
                              <ShieldAlert className="h-3 w-3" />
                              Lỗi cài Rule
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 rounded-full bg-muted border border-border px-2.5 py-0.5 text-[10px] font-bold text-muted-foreground">
                              Chưa cài Rule
                            </span>
                          )}

                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleDeleteMailbox(mb);
                            }}
                            title="Xóa tài khoản"
                            className="p-1 rounded-lg text-muted-foreground hover:text-rose-500 hover:bg-rose-500/10"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        </div>

                        {mb.status === "ACTIVE" && mb.refreshToken ? (
                          <span className="inline-flex items-center gap-1 rounded-full bg-blue-500/10 border border-blue-500/20 px-2 py-0.5 text-[9px] font-bold text-blue-500">
                            <CheckCircle2 className="h-2.5 w-2.5" />
                            Đã có Token
                          </span>
                        ) : mb.status === "ERROR" ? (
                          <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/10 border border-rose-500/20 px-2 py-0.5 text-[9px] font-bold text-rose-500">
                            <ShieldAlert className="h-2.5 w-2.5" />
                            Lỗi Token
                          </span>
                        ) : null}

                        {Number(mb.forwardedCount || 0) > 0 && (
                          <span
                            className="inline-flex items-center gap-1 rounded-full bg-indigo-500/15 border border-indigo-500/30 px-2 py-0.5 text-[9px] font-bold text-indigo-400"
                            title={
                              mb.lastForwardedAt
                                ? `Forward gần nhất: ${new Date(mb.lastForwardedAt).toLocaleString("vi-VN")}`
                                : ""
                            }
                          >
                            <Send className="h-2.5 w-2.5" />
                            Đã Forward: {mb.forwardedCount} thư
                          </span>
                        )}
                      </div>
                    </div>

                    {/* Latest OTP, Rule Status Note, or Subject summary */}
                    {(mb.lastOtp || mb.lastSubject || mb.errorMessage || mb.ruleErrorMessage) && (
                      <div className="rounded-xl bg-muted/50 px-3 py-2 text-[11px] space-y-1">
                        {mb.lastOtp && (
                          <div className="flex items-center justify-between">
                            <span className="text-muted-foreground font-medium">OTP gần nhất:</span>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                copyText(mb.lastOtp, `otp_${mb.id}`);
                              }}
                              className="inline-flex items-center gap-1.5 font-mono font-bold text-emerald-500 bg-emerald-500/10 px-2 py-0.5 rounded border border-emerald-500/20"
                            >
                              <span>{mb.lastOtp}</span>
                              {copiedKey === `otp_${mb.id}` ? (
                                <Check className="h-3 w-3" />
                              ) : (
                                <Copy className="h-3 w-3" />
                              )}
                            </button>
                          </div>
                        )}
                        {mb.lastSubject && (
                          <p className="truncate text-muted-foreground">
                            <span className="font-semibold text-foreground">Thư mới:</span>{" "}
                            {mb.lastSubject}
                          </p>
                        )}
                        {mb.ruleErrorMessage && (
                          <p
                            className={`font-medium break-words ${
                              mb.ruleStatus === "SUCCESS"
                                ? "text-emerald-500"
                                : mb.ruleStatus === "VERIFY_REQUIRED"
                                ? "text-amber-500"
                                : "text-rose-500"
                            }`}
                          >
                            Rule/Forward: {mb.ruleErrorMessage}
                          </p>
                        )}
                        {mb.errorMessage && (
                          <p className="text-rose-500 font-medium break-words">
                            Lỗi Token: {mb.errorMessage}
                          </p>
                        )}
                      </div>
                    )}

                    {/* Action buttons */}
                    <div className="flex flex-wrap items-center gap-2 pt-1">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          handleSetupRuleSingle(mb);
                        }}
                        disabled={ruleSetupId === mb.id || batchRuleRunning}
                        className="inline-flex items-center justify-center gap-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 text-xs font-bold shadow-sm transition-all disabled:opacity-50"
                        title="Mở trình duyệt đăng nhập Hotmail qua Proxy để tự động tạo Rule (Amazon, Kindle) và bật Forwarding + POP/IMAP"
                      >
                        <Sparkles
                          className={`h-3.5 w-3.5 ${ruleSetupId === mb.id ? "animate-spin" : ""}`}
                        />
                        <span>
                          {ruleSetupId === mb.id
                            ? "Đang cài Rule..."
                            : mb.ruleStatus === "SUCCESS"
                            ? "Cài Lại Rule & Forward"
                            : "Cài Rule & Forward (Browser)"}
                        </span>
                      </button>

                      {mb.refreshToken ? (
                        <>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleCheckMail(mb);
                            }}
                            disabled={isChecking}
                            className="flex-1 inline-flex items-center justify-center gap-1.5 rounded-xl bg-primary px-3 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 disabled:opacity-50"
                          >
                            <RefreshCw
                              className={`h-3.5 w-3.5 ${isChecking ? "animate-spin" : ""}`}
                            />
                            <span>{isChecking ? "Đang đọc mail..." : "Check Mail Ngay"}</span>
                          </button>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleOpenManualLink(mb);
                            }}
                            className="inline-flex items-center gap-1 rounded-xl border border-border bg-card px-2.5 py-2 text-[11px] font-semibold text-muted-foreground hover:text-foreground"
                            title="Lấy lại Refresh Token mới"
                          >
                            <KeyRound className="h-3.5 w-3.5" />
                            <span>Đổi Token</span>
                          </button>
                        </>
                      ) : (
                        <>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleAutoBrowserToken(mb);
                            }}
                            disabled={isAutoRunning}
                            className="inline-flex items-center justify-center gap-1.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-white px-2.5 py-2 text-xs font-bold shadow-sm transition-all disabled:opacity-50"
                          >
                            <Bot
                              className={`h-3.5 w-3.5 ${isAutoRunning ? "animate-spin" : ""}`}
                            />
                            <span>
                              {isAutoRunning ? "Đang mở..." : "Lấy Token"}
                            </span>
                          </button>

                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleOpenManualLink(mb);
                            }}
                            className="inline-flex items-center gap-1 rounded-xl border border-border bg-card hover:bg-muted px-2.5 py-2 text-xs font-bold text-foreground transition-all"
                          >
                            <Link2 className="h-3.5 w-3.5 text-primary" />
                            <span>Thủ Công</span>
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>

        {/* RIGHT COLUMN: Inbox & Email Content Viewer (7 cols) */}
        <div className="lg:col-span-7 flex flex-col rounded-2xl border border-border bg-card shadow-sm overflow-hidden min-h-[680px]">
          {/* Top Bar */}
          <div className="border-b border-border p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h2 className="text-sm font-bold text-foreground flex flex-wrap items-center gap-2">
                <FolderOpen className="h-4 w-4 text-primary" />
                {selectedMailbox ? (
                  <>
                    {selectedMailbox.accountCode && (
                      <span className="rounded-md bg-blue-600 px-2 py-0.5 text-xs font-extrabold text-white">
                        {selectedMailbox.accountCode}
                      </span>
                    )}
                    <span>{selectedMailbox.email}</span>
                    {selectedMailbox.note && (
                      <span className="text-xs font-normal text-muted-foreground">
                        ({selectedMailbox.note})
                      </span>
                    )}
                  </>
                ) : (
                  <span>Chọn một tài khoản để xem thư</span>
                )}
              </h2>
              {selectedMailbox?.lastCheckedAt && (
                <p className="text-[11px] text-muted-foreground flex items-center gap-1 mt-0.5">
                  <Clock className="h-3 w-3" />
                  <span>
                    Check lần cuối:{" "}
                    {new Date(selectedMailbox.lastCheckedAt).toLocaleString("vi-VN")}
                  </span>
                </p>
              )}
            </div>

            {/* Limit Selector & Folder Filter Tabs */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center gap-1.5 rounded-xl border border-border bg-muted/40 px-2.5 py-1">
                <span className="text-[11px] font-semibold text-muted-foreground whitespace-nowrap">
                  Số thư/mục:
                </span>
                <select
                  value={mailLimit}
                  onChange={(e) => {
                    const nextLimit = Number(e.target.value) || 25;
                    setMailLimit(nextLimit);
                    if (selectedMailbox && selectedMailbox.refreshToken && !checkingId) {
                      handleCheckMail(selectedMailbox, nextLimit);
                    }
                  }}
                  className="bg-transparent text-[11px] font-bold text-foreground focus:outline-none cursor-pointer"
                  title="Số lượng thư tối đa tải về từ mỗi thư mục (Inbox & Spam/Junk)"
                >
                  <option value={15} className="bg-card text-foreground">
                    15 thư (Tổng tối đa 30)
                  </option>
                  <option value={25} className="bg-card text-foreground">
                    25 thư (Tổng tối đa 50)
                  </option>
                  <option value={50} className="bg-card text-foreground">
                    50 thư (Tổng tối đa 100)
                  </option>
                  <option value={100} className="bg-card text-foreground">
                    100 thư (Tổng tối đa 200)
                  </option>
                  <option value={200} className="bg-card text-foreground">
                    200 thư (Tổng tối đa 400)
                  </option>
                </select>
              </div>

              <div className="flex items-center gap-1 rounded-xl bg-muted/60 p-1">
                {[
                  { id: "ALL", label: `Tất cả (${messages.length})` },
                  {
                    id: "INBOX",
                    label: (() => {
                      const inboxMsgs = messages.filter((m) => m.folder === "INBOX");
                      const unreadCnt = inboxMsgs.filter((m) => m.isUnread).length;
                      return unreadCnt > 0 && unreadCnt !== inboxMsgs.length
                        ? `Inbox (${inboxMsgs.length} • ${unreadCnt} chưa đọc)`
                        : `Inbox (${inboxMsgs.length})`;
                    })(),
                  },
                  {
                    id: "Junk",
                    label: `Spam/Junk (${messages.filter((m) => m.folder === "Junk").length})`,
                  },
                  {
                    id: "OTP",
                    label: `OTP (${messages.filter((m) => !!m.otpCode).length})`,
                  },
                ].map((tab) => (
                  <button
                    key={tab.id}
                    onClick={() => {
                      setFolderFilter(tab.id as any);
                      const nextList = messages.filter((m) => {
                        if (tab.id === "INBOX") return m.folder === "INBOX";
                        if (tab.id === "Junk") return m.folder === "Junk";
                        if (tab.id === "OTP") return !!m.otpCode;
                        return true;
                      });
                      setSelectedMessage(nextList[0] || null);
                    }}
                    className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition-all ${
                      folderFilter === tab.id
                        ? "bg-card text-foreground shadow-sm"
                        : "text-muted-foreground hover:text-foreground"
                    }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Highlighted OTP Banner if found */}
          {messages.some((m) => !!m.otpCode) && (
            <div className="bg-emerald-500/10 border-b border-emerald-500/20 px-4 py-2.5 flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-emerald-600 dark:text-emerald-400">
                <Sparkles className="h-4 w-4" />
                <span>
                  Phát hiện mã OTP mới nhất:{" "}
                  <code className="text-sm font-mono underline">
                    {messages.find((m) => !!m.otpCode)?.otpCode}
                  </code>
                </span>
              </div>
              <button
                onClick={() => {
                  const code = messages.find((m) => !!m.otpCode)?.otpCode;
                  if (code) copyText(code, "banner_otp");
                }}
                className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-500 px-3 py-1 text-xs font-bold text-white shadow-sm hover:bg-emerald-600"
              >
                {copiedKey === "banner_otp" ? (
                  <>
                    <Check className="h-3.5 w-3.5" />
                    <span>Đã Copy</span>
                  </>
                ) : (
                  <>
                    <Copy className="h-3.5 w-3.5" />
                    <span>Copy OTP</span>
                  </>
                )}
              </button>
            </div>
          )}

          {/* Split view inside Right Column: Messages list + Selected Message Body */}
          {filteredMessages.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center p-10 text-center space-y-3">
              <div className="h-12 w-12 rounded-2xl bg-muted flex items-center justify-center text-muted-foreground">
                <Mail className="h-6 w-6" />
              </div>
              <div className="max-w-sm space-y-1">
                <p className="text-sm font-bold text-foreground">
                  {checkingId
                    ? "Đang kết nối OAuth2 tới máy chủ Outlook..."
                    : selectedMailbox && !selectedMailbox.refreshToken
                    ? `Tài khoản ${selectedMailbox.accountCode || ""} (${selectedMailbox.email}) chưa lấy Refresh Token`
                    : "Chưa có dữ liệu thư hiển thị"}
                </p>
                <p className="text-xs text-muted-foreground">
                  {selectedMailbox && !selectedMailbox.refreshToken
                    ? "Hãy bấm nút 'Lấy Token Tự Động (Browser)' ở tài khoản bên trái hoặc 'Lấy Token Hàng Loạt' ở góc trên cùng."
                    : "Bấm nút 'Check Mail Ngay' ở tài khoản bên trái để tải toàn bộ thư từ Inbox và Junk/Spam."}
                </p>
              </div>
            </div>
          ) : (
            <div className="flex-1 grid grid-cols-1 md:grid-cols-12 divide-y md:divide-y-0 md:divide-x divide-border overflow-hidden">
              {/* Message List (5 cols) */}
              <div className="md:col-span-5 divide-y divide-border overflow-y-auto max-h-[580px]">
                {filteredMessages.map((msg) => {
                  const isActive = selectedMessage?.id === msg.id;
                  return (
                    <div
                      key={msg.id}
                      onClick={() => setSelectedMessage(msg)}
                      className={`p-3.5 cursor-pointer transition-colors space-y-1.5 ${
                        isActive ? "bg-primary/10" : "hover:bg-muted/40"
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-[11px] font-bold text-foreground truncate flex items-center gap-1.5">
                          {msg.isUnread && (
                            <span
                              className="h-2 w-2 rounded-full bg-blue-500 shrink-0"
                              title="Thư chưa đọc"
                            />
                          )}
                          <span className="truncate">{msg.fromName || msg.fromAddress}</span>
                        </span>
                        <div className="flex items-center gap-1 shrink-0">
                          {msg.isForwarded && (
                            <span
                              className="rounded bg-emerald-500/15 border border-emerald-500/30 px-1.5 py-0.5 text-[9px] font-bold text-emerald-500"
                              title={`Đã Forward về ${msg.forwardedTo || "Gmail"}${
                                msg.forwardedAt
                                  ? ` lúc ${new Date(msg.forwardedAt).toLocaleString("vi-VN")}`
                                  : ""
                              }`}
                            >
                              ✓ Đã Forward
                            </span>
                          )}
                          {!msg.isUnread && (
                            <span className="rounded bg-muted px-1.5 py-0.5 text-[9px] font-medium text-muted-foreground">
                              Đã đọc
                            </span>
                          )}
                          <span
                            className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                              msg.folder === "Junk"
                                ? "bg-amber-500/15 text-amber-500"
                                : "bg-blue-500/10 text-blue-500"
                            }`}
                          >
                            {msg.folder}
                          </span>
                        </div>
                      </div>

                      <p className="text-xs font-semibold text-foreground line-clamp-2">
                        {msg.subject}
                      </p>

                      {msg.otpCode && (
                        <div className="inline-flex items-center gap-1.5 rounded bg-emerald-500/15 px-2 py-0.5 text-[11px] font-mono font-bold text-emerald-500">
                          <span>OTP: {msg.otpCode}</span>
                        </div>
                      )}

                      <p className="text-[11px] text-muted-foreground line-clamp-2">
                        {msg.snippet}
                      </p>

                      <p className="text-[10px] text-muted-foreground">
                        {new Date(msg.date).toLocaleString("vi-VN")}
                      </p>
                    </div>
                  );
                })}
              </div>

              {/* Message Content Detail (7 cols) */}
              <div className="md:col-span-7 flex flex-col overflow-y-auto max-h-[580px] p-4 space-y-4">
                {selectedMessage ? (
                  <>
                    <div className="border-b border-border pb-3 space-y-2">
                      <div className="flex items-start justify-between gap-2">
                        <h3 className="text-sm sm:text-base font-bold text-foreground">
                          {selectedMessage.subject}
                        </h3>
                        <div className="flex items-center gap-1.5 shrink-0">
                          {selectedMailbox && (
                            <button
                              onClick={() =>
                                handleForwardSingleMessage(selectedMailbox, selectedMessage)
                              }
                              disabled={forwardingMsgId === selectedMessage.id}
                              className="inline-flex items-center gap-1 rounded-lg bg-indigo-600 hover:bg-indigo-700 px-2.5 py-1 text-[10px] font-bold text-white shadow-sm disabled:opacity-50"
                              title={`Forward thư này kèm mã ${
                                selectedMailbox.accountCode || ""
                              } về ${
                                scheduleStatus?.config?.forwardToEmail ||
                                "ip.kdpmanagement@gmail.com"
                              }`}
                            >
                              <Send className="h-3 w-3" />
                              <span>
                                {forwardingMsgId === selectedMessage.id
                                  ? "Đang Forward..."
                                  : "Forward Về Gmail"}
                              </span>
                            </button>
                          )}
                          <span className="rounded-full bg-muted px-2.5 py-0.5 text-[10px] font-bold text-muted-foreground">
                            {selectedMessage.folder}
                          </span>
                        </div>
                      </div>

                      <div className="text-xs text-muted-foreground space-y-0.5">
                        {selectedMailbox?.accountCode && (
                          <p>
                            <span className="font-semibold text-foreground">Tài khoản KDP:</span>{" "}
                            <span className="font-bold text-primary">
                              {selectedMailbox.accountCode}
                            </span>{" "}
                            {selectedMailbox.note ? `(${selectedMailbox.note})` : ""}
                          </p>
                        )}
                        <p>
                          <span className="font-semibold text-foreground">Từ:</span>{" "}
                          {selectedMessage.fromName} &lt;{selectedMessage.fromAddress}&gt;
                        </p>
                        <p>
                          <span className="font-semibold text-foreground">Tới:</span>{" "}
                          {selectedMessage.toAddress}
                        </p>
                        <p>
                          <span className="font-semibold text-foreground">Thời gian nhận:</span>{" "}
                          {new Date(selectedMessage.date).toLocaleString("vi-VN")}
                        </p>
                      </div>

                      {selectedMessage.otpCode && (
                        <div className="flex items-center justify-between rounded-xl bg-emerald-500/10 border border-emerald-500/20 px-3 py-2">
                          <span className="text-xs font-bold text-emerald-500">
                            Mã xác thực (OTP):{" "}
                            <code className="text-sm font-mono">{selectedMessage.otpCode}</code>
                          </span>
                          <button
                            onClick={() =>
                              copyText(selectedMessage.otpCode!, `detail_otp_${selectedMessage.id}`)
                            }
                            className="inline-flex items-center gap-1 rounded-lg bg-emerald-500 px-2.5 py-1 text-[11px] font-bold text-white"
                          >
                            {copiedKey === `detail_otp_${selectedMessage.id}` ? (
                              <>
                                <Check className="h-3 w-3" />
                                <span>Đã Copy</span>
                              </>
                            ) : (
                              <>
                                <Copy className="h-3 w-3" />
                                <span>Copy</span>
                              </>
                            )}
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Email Body */}
                    {selectedMessage.htmlBody ? (
                      <div className="rounded-xl border border-border bg-white text-slate-900 p-3 overflow-x-auto">
                        <iframe
                          key={selectedMessage.id}
                          title="email-body"
                          srcDoc={selectedMessage.htmlBody}
                          className="w-full min-h-[380px] border-0"
                          sandbox="allow-popups allow-popups-to-escape-sandbox"
                        />
                      </div>
                    ) : (
                      <pre className="whitespace-pre-wrap break-words rounded-xl bg-muted/40 p-4 text-xs text-foreground font-sans leading-relaxed">
                        {selectedMessage.textBody || "(Không có nội dung văn bản)"}
                      </pre>
                    )}
                  </>
                ) : (
                  <div className="flex-1 flex items-center justify-center text-xs text-muted-foreground">
                    Chọn một bức thư để xem chi tiết nội dung
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ACTIVITY & ERROR LOG CONSOLE */}
      <div className="rounded-2xl border border-border bg-card shadow-sm overflow-hidden">
        <div className="border-b border-border px-4 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-muted/30">
          <div className="flex items-center gap-2">
            <Terminal className="h-4 w-4 text-primary" />
            <h2 className="text-xs sm:text-sm font-bold text-foreground">
              Nhật Ký Hoạt Động &amp; Báo Lỗi Chi Tiết ({activityLogs.length})
            </h2>
          </div>

          <div className="flex items-center gap-2">
            <div className="flex items-center gap-1 rounded-lg bg-muted p-0.5">
              <button
                type="button"
                onClick={() => setLogFilter("ALL")}
                className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all ${
                  logFilter === "ALL"
                    ? "bg-card text-foreground shadow-sm"
                    : "text-muted-foreground hover:text-foreground"
                }`}
              >
                Tất cả ({activityLogs.length})
              </button>
              <button
                type="button"
                onClick={() => setLogFilter("ERROR")}
                className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all ${
                  logFilter === "ERROR"
                    ? "bg-rose-500 text-white shadow-sm"
                    : "text-muted-foreground hover:text-foreground"
                }`}
              >
                Chỉ hiện Lỗi (
                {
                  activityLogs.filter(
                    (l) => l.includes("[ERROR]") || l.includes("[WARN]")
                  ).length
                }
                )
              </button>
            </div>

            {activityLogs.length > 0 && (
              <>
                <button
                  type="button"
                  onClick={() => copyText(activityLogs.join("\n"), "copy_all_logs")}
                  className="inline-flex items-center gap-1 rounded-lg border border-border bg-card px-2.5 py-1 text-[11px] font-semibold text-foreground hover:bg-muted"
                >
                  {copiedKey === "copy_all_logs" ? (
                    <>
                      <Check className="h-3 w-3 text-emerald-500" />
                      <span>Đã Copy</span>
                    </>
                  ) : (
                    <>
                      <Copy className="h-3 w-3" />
                      <span>Copy Log</span>
                    </>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setActivityLogs([])}
                  className="inline-flex items-center gap-1 rounded-lg border border-border bg-card px-2.5 py-1 text-[11px] font-semibold text-muted-foreground hover:text-rose-500"
                >
                  <Trash2 className="h-3 w-3" />
                  <span>Xóa Log</span>
                </button>
              </>
            )}
          </div>
        </div>

        <div
          ref={logContainerRef}
          className="bg-zinc-950 text-zinc-200 font-mono text-xs p-4 max-h-[260px] overflow-y-auto space-y-1.5"
        >
          {activityLogs.filter((l) =>
            logFilter === "ERROR"
              ? l.includes("[ERROR]") || l.includes("[WARN]")
              : true
          ).length === 0 ? (
            <p className="text-zinc-500 text-center py-4 font-sans">
              Chưa có bản ghi log nào. Khi bạn bấm &ldquo;Lấy Token Tự Động&rdquo;, &ldquo;Check Mail Ngay&rdquo; hoặc chạy Schedule, chi tiết từng bước kèm mã KDP-xxx sẽ hiện tại đây.
            </p>
          ) : (
            activityLogs
              .filter((l) =>
                logFilter === "ERROR"
                  ? l.includes("[ERROR]") || l.includes("[WARN]")
                  : true
              )
              .map((line, idx) => {
                const isErr = line.includes("[ERROR]");
                const isWarn = line.includes("[WARN]");
                const isOk = line.includes("[OK]");
                return (
                  <div
                    key={idx}
                    className={`leading-relaxed break-words rounded px-2 py-1 ${
                      isErr
                        ? "bg-rose-500/15 text-rose-300 border border-rose-500/30"
                        : isWarn
                        ? "bg-amber-500/15 text-amber-300 border border-amber-500/20"
                        : isOk
                        ? "bg-emerald-500/10 text-emerald-300"
                        : "text-zinc-300"
                    }`}
                  >
                    {line}
                  </div>
                );
              })
          )}
        </div>
      </div>

      {/* MODAL 1: Manual OAuth2 Token Activation */}
      {linkingMailbox && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="w-full max-w-lg rounded-2xl border border-border bg-card p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div className="flex items-center gap-2">
                <KeyRound className="h-5 w-5 text-primary" />
                <h3 className="text-base font-bold text-foreground">
                  Lấy Refresh Token OAuth2 ({linkingMailbox.accountCode || "Hotmail"})
                </h3>
              </div>
              <button
                onClick={() => setLinkingMailbox(null)}
                className="text-xs text-muted-foreground hover:text-foreground"
              >
                Đóng
              </button>
            </div>

            <div className="rounded-xl bg-muted/50 p-3.5 text-xs space-y-1.5">
              {linkingMailbox.accountCode && (
                <p>
                  <span className="font-semibold text-foreground">Tài khoản KDP:</span>{" "}
                  <span className="font-bold text-primary">{linkingMailbox.accountCode}</span>{" "}
                  {linkingMailbox.note ? `(${linkingMailbox.note})` : ""}
                </p>
              )}
              <p>
                <span className="font-semibold text-foreground">Email:</span>{" "}
                <code className="font-mono font-bold text-primary">{linkingMailbox.email}</code>
              </p>
              {linkingMailbox.password && (
                <div className="flex items-center gap-2">
                  <span className="font-semibold text-foreground">Mật khẩu:</span>
                  <code className="font-mono bg-card px-2 py-0.5 rounded border border-border">
                    {linkingMailbox.password}
                  </code>
                  <button
                    onClick={() => copyText(linkingMailbox.password, "modal_pass")}
                    className="inline-flex items-center gap-1 text-[11px] text-primary font-semibold"
                  >
                    {copiedKey === "modal_pass" ? "Đã copy!" : "Copy Pass"}
                  </button>
                </div>
              )}
            </div>

            <div className="space-y-3 text-xs">
              <div className="rounded-xl border border-border p-3.5 space-y-2">
                <p className="font-bold text-foreground">
                  Bước 1: Mở trang cấp quyền Microsoft OAuth2
                </p>
                <p className="text-muted-foreground">
                  Bấm nút bên dưới để mở trang đăng nhập Microsoft (hoặc copy link dán vào
                  BitBrowser), đăng nhập tài khoản và bấm <b>Accept (Đồng ý)</b>.
                </p>
                <div className="flex items-center gap-2 pt-1">
                  <a
                    href={authUrl || "#"}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 rounded-xl bg-blue-600 hover:bg-blue-700 px-3.5 py-2 text-xs font-bold text-white shadow-sm"
                  >
                    <ExternalLink className="h-3.5 w-3.5" />
                    <span>Mở Trang Đăng Nhập Microsoft</span>
                  </a>
                  <button
                    type="button"
                    onClick={() => copyText(authUrl, "modal_auth_url")}
                    className="inline-flex items-center gap-1.5 rounded-xl border border-border bg-card px-3 py-2 text-xs font-semibold text-foreground hover:bg-muted"
                  >
                    {copiedKey === "modal_auth_url" ? (
                      <>
                        <Check className="h-3.5 w-3.5 text-emerald-500" />
                        <span>Đã Copy Link</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-3.5 w-3.5" />
                        <span>Copy Link OAuth2</span>
                      </>
                    )}
                  </button>
                </div>
              </div>

              <form
                onSubmit={handleConfirmManualCode}
                className="rounded-xl border border-border p-3.5 space-y-2.5"
              >
                <p className="font-bold text-foreground">
                  Bước 2: Dán đường link `https://localhost/?code=...` vào đây
                </p>
                <textarea
                  rows={3}
                  value={redirectUrlInput}
                  onChange={(e) => setRedirectUrlInput(e.target.value)}
                  placeholder="https://localhost/?code=M.R3_BAY..."
                  className="w-full rounded-xl border border-border bg-muted/40 p-2.5 text-xs font-mono text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
                  required
                />
                <div className="flex justify-end gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => setLinkingMailbox(null)}
                    className="rounded-xl border border-border px-4 py-2 text-xs font-semibold text-muted-foreground hover:text-foreground"
                  >
                    Hủy
                  </button>
                  <button
                    type="submit"
                    disabled={isExchanging || !redirectUrlInput.trim()}
                    className="inline-flex items-center gap-1.5 rounded-xl bg-primary px-4 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 disabled:opacity-50"
                  >
                    {isExchanging ? (
                      <>
                        <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                        <span>Đang lấy Token...</span>
                      </>
                    ) : (
                      <>
                        <CheckCircle2 className="h-3.5 w-3.5" />
                        <span>Lưu Refresh Token &amp; Đọc Mail</span>
                      </>
                    )}
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: Schedule & Auto-Forward Configuration + Bulk Mail Input */}
      {isScheduleModalOpen && scheduleForm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <form
            onSubmit={handleSaveScheduleConfig}
            className="w-full max-w-2xl rounded-2xl border border-border bg-card p-6 shadow-xl space-y-4 max-h-[90vh] overflow-y-auto"
          >
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div className="flex items-center gap-2">
                <CalendarClock className="h-5 w-5 text-primary" />
                <h3 className="text-base font-bold text-foreground">
                  Cài Đặt Schedule Quét Hotmail (24h) &amp; Nạp Danh Sách Mail (KDP + Token)
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsScheduleModalOpen(false)}
                className="text-xs text-muted-foreground hover:text-foreground"
              >
                Đóng
              </button>
            </div>

            <div className="space-y-3 text-xs">
              {/* Bulk Input Section inside Schedule Modal */}
              <div className="rounded-xl border border-indigo-500/30 bg-indigo-500/5 p-3.5 space-y-2">
                <label className="font-bold text-foreground block">
                  1. Nạp danh sách Email chạy Schedule kèm mã KDP (Tùy chọn - Mỗi dòng 1 nick):
                </label>
                <p className="text-muted-foreground">
                  Hỗ trợ định dạng:{" "}
                  <code>KDP-050|email|pass|refresh_token|client_id|mail_khoi_phuc</code> (hoặc để{" "}
                  <code>KDP-050</code> ở cột đầu / cột cuối).
                </p>
                <textarea
                  rows={5}
                  value={scheduleBulkLines}
                  onChange={(e) => setScheduleBulkLines(e.target.value)}
                  placeholder={
                    "KDP-050|alaiyahzohair1991@outlook.com|kgkkyxdxy3|M.C527_BAY...|9e5f94bc-e8a4-4e73-b8be-63364c29d753|alaiyahzohair1991m3ph1h6@fviainboxes.com"
                  }
                  className="w-full rounded-xl border border-border bg-card p-2.5 text-xs font-mono text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
                />
                <label className="flex items-center gap-2 font-semibold text-foreground cursor-pointer">
                  <input
                    type="checkbox"
                    checked={scheduleReplaceExisting}
                    onChange={(e) => setScheduleReplaceExisting(e.target.checked)}
                    className="rounded border-border"
                  />
                  <span>
                    Xóa danh sách tài khoản cũ trên bảng và thay bằng danh sách mới vừa dán
                  </span>
                </label>
              </div>

              <label className="flex items-center gap-2 font-bold text-foreground cursor-pointer pt-1">
                <input
                  type="checkbox"
                  checked={scheduleForm.enabled}
                  onChange={(e) =>
                    setScheduleForm({ ...scheduleForm, enabled: e.target.checked })
                  }
                  className="rounded border-border"
                />
                <span>2. Bật tự động quét Hotmail định kỳ theo lịch (Schedule Enabled)</span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="font-bold text-foreground block mb-1">
                    Chế độ lập lịch:
                  </label>
                  <select
                    value={scheduleForm.scheduleMode}
                    onChange={(e) =>
                      setScheduleForm({
                        ...scheduleForm,
                        scheduleMode: e.target.value as "INTERVAL" | "CUSTOM",
                      })
                    }
                    className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs text-foreground"
                  >
                    <option value="INTERVAL">Lặp theo chu kỳ giờ (Interval)</option>
                    <option value="CUSTOM">Mốc giờ cố định trong ngày (Custom)</option>
                  </select>
                </div>

                {scheduleForm.scheduleMode === "INTERVAL" ? (
                  <>
                    <div>
                      <label className="font-bold text-foreground block mb-1">
                        Giờ bắt đầu:
                      </label>
                      <input
                        type="text"
                        value={scheduleForm.dailyTime}
                        onChange={(e) =>
                          setScheduleForm({ ...scheduleForm, dailyTime: e.target.value })
                        }
                        placeholder="00:00"
                        className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs font-mono text-foreground"
                      />
                    </div>
                    <div>
                      <label className="font-bold text-foreground block mb-1">
                        Chu kỳ lặp (Giờ):
                      </label>
                      <input
                        type="number"
                        step="0.5"
                        min="0.5"
                        max="24"
                        value={scheduleForm.intervalHours}
                        onChange={(e) =>
                          setScheduleForm({
                            ...scheduleForm,
                            intervalHours: Number(e.target.value) || 2,
                          })
                        }
                        className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs font-mono text-foreground"
                      />
                    </div>
                  </>
                ) : (
                  <div className="sm:col-span-2">
                    <label className="font-bold text-foreground block mb-1">
                      Các mốc giờ (cách nhau dấu phẩy):
                    </label>
                    <input
                      type="text"
                      value={scheduleForm.customScheduleSlots}
                      onChange={(e) =>
                        setScheduleForm({
                          ...scheduleForm,
                          customScheduleSlots: e.target.value,
                        })
                      }
                      placeholder="00:00, 02:00, 04:00, 08:00..."
                      className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs font-mono text-foreground"
                    />
                  </div>
                )}
              </div>

              <div className="border-t border-border pt-3 space-y-2.5">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <label className="flex items-center gap-2 font-bold text-foreground cursor-pointer">
                    <input
                      type="checkbox"
                      checked={scheduleForm.autoForwardEnabled}
                      onChange={(e) =>
                        setScheduleForm({
                          ...scheduleForm,
                          autoForwardEnabled: e.target.checked,
                        })
                      }
                      className="rounded border-border"
                    />
                    <span>
                      3. Đọc thấy thư trong vòng {scheduleForm.lookbackHours}h → Tự động Forward về{" "}
                      <code className="text-indigo-400">
                        {(scheduleForm.forwardToEmail || "ip.kdpmanagement@gmail.com").replace(
                          "@",
                          "+KDP-<số thứ tự>@"
                        )}
                      </code>
                    </span>
                  </label>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div className="sm:col-span-2">
                    <label className="font-bold text-foreground block mb-1">
                      Email Gmail gốc nhận Forward (Tự động gắn <code>+KDP-&lt;số thứ tự&gt;</code>):
                    </label>
                    <input
                      type="email"
                      value={scheduleForm.forwardToEmail}
                      onChange={(e) =>
                        setScheduleForm({ ...scheduleForm, forwardToEmail: e.target.value })
                      }
                      placeholder="ip.kdpmanagement@gmail.com"
                      className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs font-mono text-foreground"
                      required
                    />
                  </div>
                  <div>
                    <label className="font-bold text-foreground block mb-1">
                      Lọc thư mới trong (Giờ):
                    </label>
                    <input
                      type="number"
                      value={scheduleForm.lookbackHours}
                      onChange={(e) =>
                        setScheduleForm({
                          ...scheduleForm,
                          lookbackHours: Number(e.target.value) || 24,
                        })
                      }
                      className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs font-mono text-foreground"
                    />
                  </div>
                </div>

                <p className="text-[11px] text-emerald-500 font-semibold">
                  ✓ Cơ chế chống gửi trùng: Mỗi bức thư đã Forward sẽ được lưu khóa định danh vĩnh viễn vào bảng <code>OutlookForwardedEmail</code>. Các lần quét tiếp theo gặp lại thư đó sẽ tự động bỏ qua, tuyệt đối không forward lại lần 2.
                </p>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setIsScheduleModalOpen(false)}
                className="rounded-xl border border-border px-4 py-2 text-xs font-semibold text-muted-foreground hover:text-foreground"
              >
                Hủy
              </button>
              <button
                type="submit"
                disabled={savingSchedule}
                className="inline-flex items-center gap-1.5 rounded-xl bg-primary px-4 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 disabled:opacity-50"
              >
                {savingSchedule ? "Đang lưu..." : "Lưu Cấu Hình & Nạp Danh Sách"}
              </button>
            </div>
          </form>
        </div>
      )}

      {/* MODAL 3: Bulk Import Mailboxes */}
      {isImportOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <form
            onSubmit={handleImportSubmit}
            className="w-full max-w-xl rounded-2xl border border-border bg-card p-6 shadow-xl space-y-4"
          >
            <div className="flex items-center justify-between border-b border-border pb-3">
              <h3 className="text-base font-bold text-foreground">
                Thêm / Import Danh Sách Mail Outlook (KDP + OAuth2 Token)
              </h3>
              <button
                type="button"
                onClick={() => setIsImportOpen(false)}
                className="text-xs text-muted-foreground hover:text-foreground"
              >
                Đóng
              </button>
            </div>

            <div className="space-y-2 text-xs">
              <label className="font-bold text-foreground block">
                Danh sách tài khoản (Mỗi dòng 1 mail):
              </label>
              <p className="text-muted-foreground">
                Hỗ trợ đầy đủ: <code>KDP-050|email|password|refresh_token|client_id|recovery_mail</code>{" "}
                hoặc <code>KDP-050|email|password</code>
              </p>
              <textarea
                rows={6}
                value={rawImportText}
                onChange={(e) => setRawImportText(e.target.value)}
                placeholder={
                  "KDP-050|alaiyahzohair1991@outlook.com|kgkkyxdxy3|M.C527_BAY...|9e5f94bc-e8a4-4e73-b8be-63364c29d753|alaiyahzohair1991m3ph1h6@fviainboxes.com"
                }
                className="w-full rounded-xl border border-border bg-muted/40 p-3 text-xs font-mono text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
                required
              />
              <label className="flex items-center gap-2 font-semibold text-foreground cursor-pointer pt-1">
                <input
                  type="checkbox"
                  checked={importReplaceExisting}
                  onChange={(e) => setImportReplaceExisting(e.target.checked)}
                  className="rounded border-border"
                />
                <span>Xóa danh sách cũ và thay bằng danh sách mới này</span>
              </label>
            </div>

            <div className="space-y-1.5 text-xs">
              <label className="font-bold text-foreground block">Ghi chú / Nhãn (Tùy chọn):</label>
              <input
                type="text"
                value={importNote}
                onChange={(e) => setImportNote(e.target.value)}
                placeholder="VD: Tên tác giả KDP..."
                className="w-full rounded-xl border border-border bg-muted/40 px-3 py-2 text-xs text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setIsImportOpen(false)}
                className="rounded-xl border border-border px-4 py-2 text-xs font-semibold text-muted-foreground hover:text-foreground"
              >
                Hủy
              </button>
              <button
                type="submit"
                disabled={isImporting || !rawImportText.trim()}
                className="inline-flex items-center gap-1.5 rounded-xl bg-primary px-4 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 disabled:opacity-50"
              >
                {isImporting ? "Đang lưu..." : "Lưu Danh Sách Mail"}
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}
