/**
 * AUTOMATED OUTLOOK WEB (OWA) RULE & FORWARDING SETUP VIA PUPPETEER
 * Logs into `https://login.live.com` -> `https://outlook.live.com/mail/0/` through the account's Proxy,
 * opens Settings (⚙️), and configures:
 * 1) `Rules` tab:
 *    - Rule name: `KDP`
 *    - Condition: `Message body includes` -> `Amazon`, `Kindle`
 *    - Action: `Forward to` -> `ip.kdpmanagement+KDP_XXX@gmail.com`
 *    - `Stop processing more rules` = checked -> `Save`
 * 2) `Forwarding and IMAP` tab:
 *    - `Enable forwarding` = ON -> `ip.kdpmanagement@gmail.com`
 *    - `Keep a copy of forwarded messages` = unchecked
 *    - `Let devices and apps use POP` = ON ("Don't allow devices and apps to delete messages...")
 *    - `Let devices and apps use IMAP` = ON -> `Save`
 *
 * If Microsoft prompts for a recovery email verification code ("Verify your identity" / "Help us protect your account")
 * without allowing password sign-in, immediately aborts and returns `ruleStatus: "VERIFY_REQUIRED"`.
 */

import fs from "fs";
import puppeteer, { Page } from "puppeteer-core";
import { BitBrowserClient } from "@/lib/bitbrowser/bitbrowser-client";
import { OutlookMailboxRecord } from "./outlook-db";
import { parseOutlookProxy, verifyOutlookProxy } from "./outlook-proxy";

export interface RuleSetupConfigInput {
  baseForwardEmail?: string; // default: "ip.kdpmanagement@gmail.com"
  ruleName?: string; // default: "KDP"
  keywords?: string[]; // default: ["Amazon", "Kindle"]
  keepForwardedCopy?: boolean; // default: false (as shown in screenshot)
  setupRulesTab?: boolean; // default: true
  setupForwardingTab?: boolean; // default: true
}

export interface RuleSetupResult {
  success: boolean;
  ruleStatus: "SUCCESS" | "VERIFY_REQUIRED" | "ERROR";
  message: string;
  ruleForwardAddress: string;
  forwardingAddress: string;
  resolvedProxy?: string;
  logs: string[];
}

function nowTimeStr(): string {
  return new Date().toLocaleTimeString("vi-VN", { hour12: false });
}

function findLocalBrowserExecutable(): string {
  const candidates = [
    process.env.CHROME_PATH,
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  ].filter(Boolean) as string[];

  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }

  throw new Error("Không tìm thấy trình duyệt Chrome hoặc Edge trên máy để chạy tự động.");
}

/**
 * Format `ip.kdpmanagement+KDP_126@gmail.com` from `ip.kdpmanagement@gmail.com` + `KDP-126`.
 */
export function buildRuleAliasEmail(baseEmail: string, accountCode?: string): string {
  const cleanEmail = (baseEmail || "ip.kdpmanagement@gmail.com").trim();
  const atIdx = cleanEmail.indexOf("@");
  if (atIdx === -1) return cleanEmail;

  const rawCode = (accountCode || "").trim();
  if (!rawCode) return cleanEmail;

  // Convert "KDP-126" -> "KDP_126" to match user's screenshot: ip.kdpmanagement+KDP_126@gmail.com
  const aliasTag = rawCode.toUpperCase().replace(/[^A-Z0-9]+/g, "_");
  const localPart = cleanEmail.slice(0, atIdx).split("+")[0];
  const domainPart = cleanEmail.slice(atIdx + 1);
  return `${localPart}+${aliasTag}@${domainPart}`;
}

/**
 * Helper to check if a Microsoft login/auth page is demanding a code from recovery email
 * without any option to use password.
 */
async function checkAndHandleAuthScreen(
  page: Page,
  password: string,
  addLog: (level: "INFO" | "OK" | "WARN" | "ERROR", msg: string) => void
): Promise<"AUTHENTICATED" | "VERIFY_REQUIRED" | "WRONG_PASSWORD"> {
  const start = Date.now();
  let typedPassword = false;

  while (Date.now() - start < 24000) {
    await new Promise((r) => setTimeout(r, 800));
    if (page.isClosed()) return "AUTHENTICATED";

    const url = page.url();
    if (
      url.includes("outlook.live.com/mail") ||
      url.includes("account.microsoft.com") ||
      url.includes("complete-client-signin-oauth-silent")
    ) {
      return "AUTHENTICATED";
    }

    const screenInfo = await page
      .evaluate(() => {
        const bodyText = (document.body ? document.body.innerText : "").trim();
        const lower = bodyText.toLowerCase();

        // Check skip / accept buttons first
        const skipLink = document.querySelector("#iShowSkip") as HTMLElement | null;
        if (skipLink && skipLink.offsetParent !== null) {
          skipLink.click();
          return { action: "CLICKED_SKIP", lower };
        }

        const primaryBtns = Array.from(
          document.querySelectorAll(
            '#acceptButton, #idBtn_Accept, [data-testid="primaryButton"], [data-testid="secondaryButton"], button, input[type="submit"]'
          )
        ) as HTMLElement[];
        for (const b of primaryBtns) {
          if (b.offsetParent === null) continue;
          const t = (b.innerText || (b as HTMLInputElement).value || "").trim().toLowerCase();
          if (
            t === "yes" ||
            t === "accept" ||
            t === "có" ||
            t.includes("skip for now") ||
            ((lower.includes("passkey") || lower.includes("break free from your passwords")) &&
              (t === "cancel" || t === "hủy" || t.includes("not now")))
          ) {
            b.click();
            return { action: `CLICKED_${t.toUpperCase()}`, lower };
          }
        }

        // Check if password input is visible
        const passEl = document.querySelector('input[type="password"]') as HTMLElement | null;
        if (passEl && passEl.offsetParent !== null) {
          return { action: "NEED_PASSWORD", lower };
        }

        // Check if "Use your password" / "Other ways to sign in" is available
        const clickables = Array.from(
          document.querySelectorAll('span[role="button"], a, button, div[role="button"]')
        ) as HTMLElement[];
        for (const el of clickables) {
          if (el.offsetParent === null) continue;
          const t = (el.innerText || "").trim().toLowerCase();
          if (
            t.includes("use your password") ||
            t.includes("dùng mật khẩu") ||
            t.includes("other ways to sign in") ||
            t.includes("sign in another way")
          ) {
            el.click();
            return { action: "CLICKED_USE_PASS", lower };
          }
        }

        return { action: "NONE", lower };
      })
      .catch(() => ({ action: "NONE", lower: "" }));

    if (screenInfo.action === "CLICKED_USE_PASS") {
      addLog("INFO", `Chuyển sang đăng nhập bằng mật khẩu ("Use your password")...`);
      await new Promise((r) => setTimeout(r, 2200));
      continue;
    }

    if (screenInfo.action === "NEED_PASSWORD" && !typedPassword) {
      await new Promise((r) => setTimeout(r, 900));
      const passInput = await page.$('input[type="password"]');
      if (passInput) {
        addLog("INFO", "Điền mật khẩu Microsoft và bấm Enter...");
        await passInput.click().catch(() => {});
        await passInput.type(password.trim(), { delay: 35 });
        const actualVal = await passInput
          .evaluate((el) => (el as HTMLInputElement).value)
          .catch(() => "");
        if (actualVal !== password.trim()) {
          await passInput.click({ count: 3 }).catch(() => {});
          await page.keyboard.press("Backspace").catch(() => {});
          await passInput.type(password.trim(), { delay: 45 });
        }
        typedPassword = true;
        await new Promise((r) => setTimeout(r, 400));
        await page.keyboard.press("Enter");
        await new Promise((r) => setTimeout(r, 3500));
        continue;
      }
    }

    const lower = screenInfo.lower;
    if (
      lower.includes("that password is incorrect") ||
      lower.includes("your account or password is incorrect") ||
      lower.includes("mật khẩu không đúng")
    ) {
      return "WRONG_PASSWORD";
    }

    if (lower.includes("tried to sign in too many times")) {
      return lower.includes("send a code") ? "VERIFY_REQUIRED" : "WRONG_PASSWORD";
    }

    // Detect recovery email code requirement without password fallback
    if (
      lower.includes("help us protect your account") ||
      lower.includes("detected something unusual about this sign-in") ||
      lower.includes("verify your identity") ||
      (lower.includes("verify your email") &&
        lower.includes("send a code") &&
        screenInfo.action === "NONE") ||
      (lower.includes("enter the code") && lower.includes("sent to"))
    ) {
      return "VERIFY_REQUIRED";
    }
  }

  return "AUTHENTICATED";
}

export async function setupOutlookRuleAndForwardingViaBrowser(
  mailbox: OutlookMailboxRecord,
  config: RuleSetupConfigInput = {}
): Promise<RuleSetupResult> {
  const logs: string[] = [];
  const logTag = mailbox.accountCode
    ? `${mailbox.accountCode} | ${mailbox.email}`
    : mailbox.email;

  const addLog = (level: "INFO" | "OK" | "WARN" | "ERROR", msg: string) => {
    const line = `[${nowTimeStr()}] [${level}] [${logTag}] ${msg}`;
    logs.push(line);
    if (level === "ERROR") console.error(line);
    else console.log(line);
  };

  const baseForwardEmail = (config.baseForwardEmail || "ip.kdpmanagement@gmail.com").trim();
  const ruleName = (config.ruleName || "KDP").trim();
  const keywords =
    Array.isArray(config.keywords) && config.keywords.length > 0
      ? config.keywords.map((k) => k.trim()).filter(Boolean)
      : ["Amazon", "Kindle"];
  const keepForwardedCopy = !!config.keepForwardedCopy;
  const setupRulesTab = config.setupRulesTab !== false;
  const setupForwardingTab = config.setupForwardingTab !== false;

  const ruleForwardAddress = buildRuleAliasEmail(baseForwardEmail, mailbox.accountCode);

  if (!mailbox.password || !mailbox.password.trim()) {
    const msg = `Tài khoản ${mailbox.email} chưa có mật khẩu để đăng nhập cài Rule.`;
    addLog("ERROR", msg);
    return {
      success: false,
      ruleStatus: "ERROR",
      message: msg,
      ruleForwardAddress,
      forwardingAddress: baseForwardEmail,
      logs,
    };
  }

  let browser;
  try {
    const executablePath = findLocalBrowserExecutable();
    let effectiveProxyRaw = (mailbox.proxy || "").trim();
    let parsedProxy = parseOutlookProxy(effectiveProxyRaw);
    let activeProxy = parsedProxy;

    const launchArgs = [
      "--window-size=1280,880",
      "--disable-blink-features=AutomationControlled",
      "--disable-popup-blocking",
      "--no-first-run",
      "--no-default-browser-check",
    ];

    if (parsedProxy) {
      addLog("INFO", `Kiểm tra Proxy riêng (${parsedProxy.displayLabel})...`);
      const check = await verifyOutlookProxy(effectiveProxyRaw, 5500);
      if (check.alive) {
        launchArgs.push(`--proxy-server=${parsedProxy.serverArg}`);
        addLog("OK", `Proxy hoạt động tốt (IP: ${check.ip}).`);
      } else {
        activeProxy = null;
        addLog("WARN", `Proxy lưu trong DB (${parsedProxy.displayLabel}) không phản hồi (${check.error}).`);
      }
    }

    if (!activeProxy) {
      try {
        const bbClient = new BitBrowserClient();
        if (mailbox.accountCode) {
          const bbProfile = await bbClient.findProfileByCode(mailbox.accountCode);
          const bbProxyRaw = bbProfile?.proxy?.raw || "";
          if (bbProxyRaw && bbProxyRaw !== effectiveProxyRaw) {
            const bbParsed = parseOutlookProxy(bbProxyRaw);
            if (bbParsed) {
              const bbCheck = await verifyOutlookProxy(bbProxyRaw, 5500);
              if (bbCheck.alive) {
                effectiveProxyRaw = bbProxyRaw;
                parsedProxy = bbParsed;
                activeProxy = bbParsed;
                launchArgs.push(`--proxy-server=${bbParsed.serverArg}`);
                addLog("OK", `Đã lấy Proxy sống từ BitBrowser (IP: ${bbCheck.ip}).`);
              }
            }
          }
        }

        // Fallback: If account's own proxy is dead (e.g. CLOSED account with expired wslpweyo proxy),
        // pick a working US static residential proxy from BitBrowser so we don't connect from local VN IP.
        if (!activeProxy) {
          const allProfiles = await bbClient.listAllProfiles();
          const staticCandidates = allProfiles
            .filter((p) => (p.proxy?.raw || "").includes("staticresidential"))
            .sort((a, b) => {
              const aFast = (a.proxy?.raw || "").includes("mnl6j6636mfl") ? 0 : 1;
              const bFast = (b.proxy?.raw || "").includes("mnl6j6636mfl") ? 0 : 1;
              if (aFast !== bFast) return aFast - bFast;
              return (a.name || "").localeCompare(b.name || "");
            });
          for (const cand of staticCandidates.slice(0, 5)) {
            const candRaw = cand.proxy?.raw || "";
            const candParsed = parseOutlookProxy(candRaw);
            if (!candParsed) continue;
            const candCheck = await verifyOutlookProxy(candRaw, 5000);
            if (candCheck.alive) {
              effectiveProxyRaw = candRaw;
              parsedProxy = candParsed;
              activeProxy = candParsed;
              launchArgs.push(`--proxy-server=${candParsed.serverArg}`);
              addLog(
                "OK",
                `Proxy gốc đã chết -> Dùng Proxy US dự phòng từ BitBrowser (${cand.name} - IP: ${candCheck.ip}).`
              );
              break;
            }
          }
        }
      } catch {
        // Ignore BitBrowser lookup error
      }
    }

    browser = await puppeteer.launch({
      executablePath,
      headless: false,
      defaultViewport: null,
      args: launchArgs,
    });

    const page = (await browser.pages())[0] || (await browser.newPage());
    if (activeProxy?.username && activeProxy?.password) {
      await page.authenticate({
        username: activeProxy.username,
        password: activeProxy.password,
      });
    }

    // Automatically authenticate any popup windows created by Outlook
    browser.on("targetcreated", async (target) => {
      if (target.type() === "page") {
        const p = await target.page().catch(() => null);
        if (p && activeProxy?.username && activeProxy?.password) {
          await p
            .authenticate({
              username: activeProxy.username,
              password: activeProxy.password,
            })
            .catch(() => {});
        }
      }
    });

    // STEP 1: Login via login.live.com
    addLog("INFO", "Bước 1: Mở trang đăng nhập Microsoft (login.live.com)...");
    await page
      .goto("https://login.live.com/", { waitUntil: "domcontentloaded", timeout: 35000 })
      .catch(() => {});

    await page
      .waitForSelector('input[type="email"], input[name="loginfmt"]', { timeout: 15000 })
      .catch(() => {});
    const emailInput = await page.$('input[type="email"], input[name="loginfmt"]');
    if (emailInput) {
      addLog("INFO", `Điền Email (${mailbox.email}) và bấm Enter...`);
      await emailInput.click();
      await emailInput.type(mailbox.email.trim(), { delay: 25 });
      await new Promise((r) => setTimeout(r, 400));
      await page.keyboard.press("Enter");
    }

    const authStatus = await checkAndHandleAuthScreen(page, mailbox.password, addLog);
    if (authStatus === "WRONG_PASSWORD") {
      const msg = `Sai mật khẩu Microsoft cho tài khoản ${mailbox.email}.`;
      addLog("ERROR", msg);
      return {
        success: false,
        ruleStatus: "ERROR",
        message: msg,
        ruleForwardAddress,
        forwardingAddress: baseForwardEmail,
        resolvedProxy: effectiveProxyRaw,
        logs,
      };
    }

    if (authStatus === "VERIFY_REQUIRED") {
      const msg = `Bị Microsoft yêu cầu nhập mã xác minh từ mail khôi phục khi đăng nhập -> Tự động bỏ qua.`;
      addLog("WARN", msg);
      return {
        success: false,
        ruleStatus: "VERIFY_REQUIRED",
        message: msg,
        ruleForwardAddress,
        forwardingAddress: baseForwardEmail,
        resolvedProxy: effectiveProxyRaw,
        logs,
      };
    }

    // STEP 2: Open Outlook Web Mail & Settings Dialog
    addLog("INFO", "Bước 2: Mở hòm thư Outlook Web (outlook.live.com/mail/0/)...");
    await page
      .goto("https://outlook.live.com/mail/0/", { waitUntil: "domcontentloaded", timeout: 35000 })
      .catch(() => {});

    // Check if redirected to another security challenge or Stay signed in prompt
    for (let i = 0; i < 25; i++) {
      await new Promise((r) => setTimeout(r, 1200));
      const settingsBtn = await page.$('button[aria-label="Settings"]').catch(() => null);
      if (settingsBtn) break;

      const bodyLower = await page
        .evaluate(() => {
          const txt = document.body ? document.body.innerText.toLowerCase() : "";
          // Click "Yes" / "Skip for now" if redirected through login.live.com on the way to Outlook
          const btns = Array.from(
            document.querySelectorAll(
              '#acceptButton, #idBtn_Accept, [data-testid="primaryButton"], [data-testid="secondaryButton"], button, input[type="submit"], #iShowSkip'
            )
          ) as HTMLElement[];
          for (const b of btns) {
            if (b.offsetParent === null) continue;
            const t = (b.innerText || (b as HTMLInputElement).value || "").trim().toLowerCase();
            if (
              t === "yes" ||
              t === "accept" ||
              t.includes("skip for now") ||
              ((txt.includes("passkey") || txt.includes("break free from your passwords")) &&
                (t === "cancel" || t === "hủy" || t.includes("not now")))
            ) {
              b.click();
              break;
            }
          }
          return txt;
        })
        .catch(() => "");
      if (
        bodyLower.includes("help us protect your account") ||
        bodyLower.includes("verify your identity") ||
        (bodyLower.includes("verify your email") && bodyLower.includes("send a code")) ||
        (bodyLower.includes("enter the code") && bodyLower.includes("sent to"))
      ) {
        const msg = `Bị Microsoft yêu cầu mã từ mail khôi phục khi vào Outlook Web -> Tự động bỏ qua.`;
        addLog("WARN", msg);
        return {
          success: false,
          ruleStatus: "VERIFY_REQUIRED",
          message: msg,
          ruleForwardAddress,
          forwardingAddress: baseForwardEmail,
          resolvedProxy: effectiveProxyRaw,
          logs,
        };
      }
    }

    const settingsBtn = await page
      .waitForSelector('button[aria-label="Settings"]', { visible: true, timeout: 20000 })
      .catch(() => null);
    if (!settingsBtn) {
      throw new Error("Không tìm thấy nút Settings (⚙️) trên giao diện Outlook Web.");
    }

    addLog("INFO", "Mở hộp thoại Settings (⚙️)...");
    await settingsBtn.click();

    await page.waitForSelector('button[role="tab"][value="rules"]', {
      visible: true,
      timeout: 20000,
    });

    // =========================================================================
    // STEP 3: SETUP TAB `Rules` (Name: KDP | Body includes: Amazon, Kindle | Forward to: ip.kdpmanagement+KDP_XXX@gmail.com)
    // =========================================================================
    if (setupRulesTab) {
      addLog(
        "INFO",
        `Bước 3: Mở tab Rules -> Cài đặt Rule "${ruleName}" (Body: ${keywords.join(
          ", "
        )} -> Forward to: ${ruleForwardAddress})...`
      );
      await page.click('button[role="tab"][value="rules"]');
      await new Promise((r) => setTimeout(r, 3500));

      // Check if Rules tab asks for "Sign in" verification
      const rulesNeedsSignIn = await page.evaluate(() => {
        const dialog = document.querySelector('[role="dialog"]');
        if (!dialog) return false;
        const txt = (dialog as HTMLElement).innerText || "";
        return txt.includes("Sign in and verify your account");
      });

      if (rulesNeedsSignIn) {
        const verified = await handleSettingsTabSignIn(browser, page, mailbox.password, addLog);
        if (verified === "VERIFY_REQUIRED") {
          const msg = `Mục Settings -> Rules yêu cầu mã xác minh từ mail khôi phục -> Tự động bỏ qua.`;
          addLog("WARN", msg);
          return {
            success: false,
            ruleStatus: "VERIFY_REQUIRED",
            message: msg,
            ruleForwardAddress,
            forwardingAddress: baseForwardEmail,
            resolvedProxy: effectiveProxyRaw,
            logs,
          };
        }
      }

      // Optional cleanup: If a rule with the exact same name (`KDP`) already exists, delete it first
      // so Outlook Web does not block `Save` due to duplicate rule name.
      for (let delRound = 0; delRound < 3; delRound++) {
        const deletedExisting = await page
          .evaluate((targetRuleName) => {
            const dialog = document.querySelector('[role="dialog"]');
            if (!dialog) return false;
            const delBtns = Array.from(
              dialog.querySelectorAll(
                'button[aria-label*="Delete" i], button[title*="Delete" i], button[aria-label*="Xóa" i]'
              )
            ) as HTMLElement[];
            for (const btn of delBtns) {
              const rect = btn.getBoundingClientRect();
              if (rect.width === 0 || rect.height === 0) continue;
              // Check if the rule container row contains targetRuleName
              let row: HTMLElement | null = btn.parentElement;
              let matched = false;
              for (let depth = 0; depth < 5 && row && row !== dialog; depth++) {
                const txt = (row.innerText || "").trim();
                if (
                  txt.toLowerCase().includes(targetRuleName.toLowerCase()) &&
                  txt.length < 400
                ) {
                  matched = true;
                  break;
                }
                row = row.parentElement;
              }
              if (matched) {
                btn.click();
                return true;
              }
            }
            return false;
          }, ruleName)
          .catch(() => false);

        if (!deletedExisting) break;
        addLog("INFO", `Phát hiện Rule "${ruleName}" cũ đã tồn tại -> Đang xóa để cập nhật mới...`);
        await new Promise((r) => setTimeout(r, 1000));
        // Confirm modal "OK" / "Yes" / "Delete" button if shown
        await page
          .evaluate(() => {
            const btns = Array.from(document.querySelectorAll("button")) as HTMLElement[];
            for (const b of btns) {
              const r = b.getBoundingClientRect();
              if (r.width === 0 || r.height === 0) continue;
              const t = (b.innerText || "").trim().toLowerCase();
              if (t === "ok" || t === "yes" || t === "delete" || t === "xóa" || t === "đồng ý") {
                b.click();
                return true;
              }
            }
            return false;
          })
          .catch(() => false);
        await new Promise((r) => setTimeout(r, 1800));
      }

      let ruleSavedDone = false;
      for (let ruleFormAttempt = 0; ruleFormAttempt < 2 && !ruleSavedDone; ruleFormAttempt++) {
        // Ensure Settings -> Rules dialog is open (in case Sign in step-up reloaded the page)
        const rulesTabBtn = await page.$('button[role="tab"][value="rules"]');
        if (!rulesTabBtn) {
          const sBtn = await page
            .waitForSelector('button[aria-label="Settings"]', { visible: true, timeout: 15000 })
            .catch(() => null);
          if (sBtn) await sBtn.click();
          await page
            .waitForSelector('button[role="tab"][value="rules"]', {
              visible: true,
              timeout: 15000,
            })
            .catch(() => {});
          await page.click('button[role="tab"][value="rules"]').catch(() => {});
          await new Promise((r) => setTimeout(r, 3000));
        }

        // Click "Add rule" button if we are on the Rules list view
        let clickedAddRule = false;
        for (let attempt = 0; attempt < 12; attempt++) {
          const hasNameInput = await page.$(
            '[role="dialog"] input[placeholder="Name your rule"]'
          );
          if (hasNameInput) {
            clickedAddRule = true;
            break;
          }
          const handles = await page.$$('[role="dialog"] button');
          for (const h of handles) {
            const txt = await h
              .evaluate((el) => (el.innerText || "").trim().toLowerCase())
              .catch(() => "");
            if (
              txt.includes("add rule") ||
              txt.includes("add new rule") ||
              txt.includes("thêm quy tắc")
            ) {
              await h.click();
              clickedAddRule = true;
              break;
            }
          }
          if (clickedAddRule) break;
          await new Promise((r) => setTimeout(r, 800));
        }

        if (!clickedAddRule) {
          throw new Error("Không tìm thấy nút 'Add rule' trong mục Settings -> Rules.");
        }

        await new Promise((r) => setTimeout(r, 2000));

        // 3.1: Fill Rule Name ("KDP")
        const nameInput = await page.waitForSelector(
          '[role="dialog"] input[placeholder="Name your rule"], [role="dialog"] input[type="text"]:not([placeholder*="Search" i])',
          { visible: true, timeout: 10000 }
        );
        if (nameInput) {
          await nameInput.click({ count: 3 }).catch(() => {});
          await page.keyboard.press("Backspace").catch(() => {});
          await nameInput.type(ruleName, { delay: 35 });
        }

        // 3.2: Select Condition -> "Message body includes"
        const condCombo = await page.$(
          '[role="dialog"] button[role="combobox"][aria-label*="condition" i]'
        );
        if (!condCombo) {
          throw new Error("Không tìm thấy ô chọn Condition ('Select a condition') trong form Rule.");
        }
        await condCombo.click();
        await new Promise((r) => setTimeout(r, 1200));

        const selectedCond = await selectDropdownOption(page, [
          "message body includes",
          "nội dung thư bao gồm",
        ]);
        if (!selectedCond) {
          throw new Error("Không chọn được điều kiện 'Message body includes' trong danh sách Condition.");
        }
        await new Promise((r) => setTimeout(r, 1200));

        // 3.3: Enter keywords ("Amazon", "Kindle") into "Enter words to look for"
        const wordInput = await page.waitForSelector(
          '[role="dialog"] input[placeholder*="Enter words" i], [role="dialog"] input[placeholder*="words to look for" i]',
          { visible: true, timeout: 8000 }
        );
        if (!wordInput) {
          throw new Error("Không tìm thấy ô nhập từ khóa ('Enter words to look for').");
        }
        for (const kw of keywords) {
          await wordInput.click();
          await wordInput.type(kw, { delay: 35 });
          await new Promise((r) => setTimeout(r, 350));
          await page.keyboard.press("Enter");
          await new Promise((r) => setTimeout(r, 500));
        }

        // 3.4: Select Action -> "Forward to"
        const actionCombo = await page.$(
          '[role="dialog"] button[role="combobox"][aria-label*="action" i]'
        );
        if (!actionCombo) {
          throw new Error("Không tìm thấy ô chọn Action ('Select an action') trong form Rule.");
        }
        await actionCombo.click();
        await new Promise((r) => setTimeout(r, 1200));

        const selectedAction = await selectDropdownOption(page, [
          "forward to",
          "chuyển tiếp đến",
        ]);
        if (!selectedAction) {
          throw new Error("Không chọn được hành động 'Forward to' trong danh sách Action.");
        }
        // Wait 2.2s for Outlook to finish mounting the final RoosterJS [contenteditable="true"] PeoplePicker node
        await new Promise((r) => setTimeout(r, 2200));

        // 3.5: Focus & click `[role="dialog"] [contenteditable="true"]` and type `ip.kdpmanagement+KDP_XXX@gmail.com`
        await page
          .evaluate(() => {
            const el = document.querySelector(
              '[role="dialog"] [contenteditable="true"]'
            ) as HTMLElement | null;
            if (el) {
              el.focus();
              el.click();
            }
          })
          .catch(() => {});

        await new Promise((r) => setTimeout(r, 400));
        await page.keyboard.type(ruleForwardAddress, { delay: 30 });

        // Wait for Outlook PeoplePicker's "#PickerHeaderUseTypeAddress" ("Use this address: ...") suggestion to appear
        await page
          .waitForSelector('#PickerHeaderUseTypeAddress, [id*="PickerHeaderUseTypeAddress"]', {
            timeout: 6000,
          })
          .catch(() => null);
        await new Promise((r) => setTimeout(r, 450));

        // Press Enter to commit the highlighted "#PickerHeaderUseTypeAddress" into a valid recipient pill
        await page.keyboard.press("Enter");
        await new Promise((r) => setTimeout(r, 1500));

        // Check if Outlook displayed inline step-up verification after adding a Forward to rule:
        // "Sign in and verify your account to create a forwarding rule. [Sign in]"
        const inlineForwardSignIn = await page
          .evaluate(() => {
            const dialog = document.querySelector('[role="dialog"]');
            if (!dialog) return false;
            const txt = (dialog as HTMLElement).innerText || "";
            return (
              txt.includes("Sign in and verify your account") ||
              txt.includes("create a forwarding rule")
            );
          })
          .catch(() => false);

        if (inlineForwardSignIn) {
          addLog(
            "INFO",
            "Outlook yêu cầu xác nhận phiên đăng nhập khi tạo Rule Forward ('Sign in and verify your account')..."
          );
          const verified = await handleSettingsTabSignIn(
            browser,
            page,
            mailbox.password,
            addLog
          );
          if (verified === "VERIFY_REQUIRED") {
            const msg = `Bị Microsoft yêu cầu nhập mã từ mail khôi phục khi tạo Rule Forward ('Sign in and verify your account') -> Tự động bỏ qua.`;
            addLog("WARN", msg);
            return {
              success: false,
              ruleStatus: "VERIFY_REQUIRED",
              message: msg,
              ruleForwardAddress,
              forwardingAddress: baseForwardEmail,
              resolvedProxy: effectiveProxyRaw,
              logs,
            };
          }
          await new Promise((r) => setTimeout(r, 3500));
          // Check if the Rule form is still open with Save enabled, or if it reset after Sign in
          const saveUnlockedNow = await page
            .evaluate(() => {
              const dialog = document.querySelector('[role="dialog"]');
              if (!dialog) return false;
              const buttons = Array.from(dialog.querySelectorAll("button")) as HTMLButtonElement[];
              const saveBtn = buttons.find((b) => {
                const t = (b.innerText || "").trim().toLowerCase();
                return t === "save" || t === "lưu";
              });
              return Boolean(
                saveBtn && !saveBtn.disabled && saveBtn.getAttribute("aria-disabled") !== "true"
              );
            })
            .catch(() => false);

          if (!saveUnlockedNow && ruleFormAttempt === 0) {
            addLog("INFO", "Đã xác nhận phiên đăng nhập -> Điền lại form Rule KDP...");
            continue;
          }
        }

        // Verify if Save button is now enabled; if not, click #PickerHeaderUseTypeAddress or press Enter again
        for (let attempt = 0; attempt < 5; attempt++) {
          const isSaveReady = await page
            .evaluate(() => {
              const dialog = document.querySelector('[role="dialog"]');
              if (!dialog) return false;
              const buttons = Array.from(dialog.querySelectorAll("button")) as HTMLButtonElement[];
              const saveBtn = buttons.find((b) => {
                const t = (b.innerText || "").trim().toLowerCase();
                return t === "save" || t === "lưu";
              });
              return Boolean(
                saveBtn && !saveBtn.disabled && saveBtn.getAttribute("aria-disabled") !== "true"
              );
            })
            .catch(() => false);

          if (isSaveReady) break;

          const useAddrHandle = await page.$(
            '#PickerHeaderUseTypeAddress, [id*="PickerHeaderUseTypeAddress"]'
          );
          if (useAddrHandle) {
            await useAddrHandle.click().catch(() => {});
          } else {
            await page.keyboard.press("Enter").catch(() => {});
          }
          await new Promise((r) => setTimeout(r, 900));
        }

        // 3.6: Ensure "Stop processing more rules" checkbox is checked
        await page
          .evaluate(() => {
            const dialog = document.querySelector('[role="dialog"]');
            if (!dialog) return;
            const checkboxes = Array.from(
              dialog.querySelectorAll('input[type="checkbox"]')
            ) as HTMLInputElement[];
            for (const cb of checkboxes) {
              const rect = cb.getBoundingClientRect();
              if ((rect.width > 0 || cb.offsetParent !== null) && !cb.checked) {
                cb.click();
              }
            }
          })
          .catch(() => {});

        // 3.7: Click Save button in Rules tab
        const savedRule = await clickDialogSaveButton(page);
        if (!savedRule) {
          throw new Error("Không bấm được nút Save để lưu Rule KDP.");
        }
        ruleSavedDone = true;
        await new Promise((r) => setTimeout(r, 3000));
        addLog("OK", `Đã lưu Rule "${ruleName}" -> Forward to: ${ruleForwardAddress}!`);
      }
    }

    // =========================================================================
    // STEP 4: SETUP TAB `Forwarding and IMAP`
    // (`Enable forwarding` -> ip.kdpmanagement@gmail.com, POP = ON, IMAP = ON)
    // =========================================================================
    if (setupForwardingTab) {
      addLog(
        "INFO",
        `Bước 4: Mở tab Forwarding and IMAP -> Bật Forwarding (${baseForwardEmail}) + POP + IMAP...`
      );
      await page.click('button[role="tab"][value="forwarding"]');
      await new Promise((r) => setTimeout(r, 3500));

      // Check if Forwarding and IMAP asks for "Sign in and verify your account"
      const fwdNeedsSignIn = await page.evaluate(() => {
        const dialog = document.querySelector('[role="dialog"]');
        if (!dialog) return false;
        const txt = (dialog as HTMLElement).innerText || "";
        return txt.includes("Sign in and verify your account");
      });

      if (fwdNeedsSignIn) {
        addLog("INFO", "Mục Forwarding and IMAP yêu cầu xác nhận lại phiên đăng nhập ('Sign in')...");
        const verified = await handleSettingsTabSignIn(browser, page, mailbox.password, addLog);
        if (verified === "VERIFY_REQUIRED") {
          const msg = setupRulesTab
            ? `Đã lưu Rule "${ruleName}" (${ruleForwardAddress}), nhưng mục Forwarding & IMAP bắt nhập mã từ mail khôi phục.`
            : `Mục Forwarding & IMAP bắt nhập mã xác minh từ mail khôi phục -> Tự động bỏ qua.`;
          addLog("WARN", msg);
          return {
            success: false,
            ruleStatus: "VERIFY_REQUIRED",
            message: msg,
            ruleForwardAddress,
            forwardingAddress: baseForwardEmail,
            resolvedProxy: effectiveProxyRaw,
            logs,
          };
        }
        // Re-ensure Settings -> Forwarding and IMAP tab is open after sign-in
        const dialogStillOpen = await page.$('button[role="tab"][value="forwarding"]');
        if (!dialogStillOpen) {
          const sBtn = await page
            .waitForSelector('button[aria-label="Settings"]', { visible: true, timeout: 15000 })
            .catch(() => null);
          if (sBtn) await sBtn.click();
          await page
            .waitForSelector('button[role="tab"][value="forwarding"]', {
              visible: true,
              timeout: 15000,
            })
            .catch(() => {});
        }
        await page.click('button[role="tab"][value="forwarding"]').catch(() => {});
        await new Promise((r) => setTimeout(r, 4000));
      }

      // Check if Forwarding controls are now unlocked
      const fwdUnlocked = await page.evaluate(() => {
        const dialog = document.querySelector('[role="dialog"]');
        if (!dialog) return false;
        const txt = (dialog as HTMLElement).innerText || "";
        return (
          txt.includes("Enable forwarding") ||
          txt.includes("Let devices and apps use IMAP") ||
          txt.includes("Forward my email to")
        );
      });

      if (!fwdUnlocked) {
        const msg = setupRulesTab
          ? `Đã lưu Rule "${ruleName}" (${ruleForwardAddress}), nhưng mục Forwarding & IMAP vẫn bị khóa xác minh mail khôi phục.`
          : `Mục Forwarding & IMAP bị khóa xác minh mail khôi phục -> Tự động bỏ qua.`;
        addLog("WARN", msg);
        return {
          success: false,
          ruleStatus: "VERIFY_REQUIRED",
          message: msg,
          ruleForwardAddress,
          forwardingAddress: baseForwardEmail,
          resolvedProxy: effectiveProxyRaw,
          logs,
        };
      }

      // Configure Forwarding + POP + IMAP switches
      await page.evaluate(
        ({ targetEmail, keepCopy }) => {
          const dialog = document.querySelector('[role="dialog"]');
          if (!dialog) return;

          // 1. Toggle ON: "Enable forwarding", "Let devices and apps use POP", "Let devices and apps use IMAP"
          const switches = Array.from(
            dialog.querySelectorAll('input[role="switch"], button[role="switch"], input[type="checkbox"]')
          ) as HTMLElement[];

          for (const sw of switches) {
            if (sw.offsetParent === null) continue;
            const labelText = (
              sw.getAttribute("aria-label") ||
              sw.closest("label")?.innerText ||
              sw.parentElement?.innerText ||
              ""
            )
              .trim()
              .toLowerCase();

            const isChecked =
              (sw as HTMLInputElement).checked === true ||
              sw.getAttribute("aria-checked") === "true";

            if (
              labelText.includes("enable forwarding") ||
              labelText.includes("use pop") ||
              labelText.includes("use imap")
            ) {
              if (!isChecked) sw.click();
            } else if (labelText.includes("keep a copy")) {
              if (isChecked !== keepCopy) sw.click();
            }
          }

          // 2. Select radio: "Don't allow devices and apps to delete messages from Outlook"
          const radios = Array.from(
            dialog.querySelectorAll('input[type="radio"], [role="radio"]')
          ) as HTMLElement[];
          for (const r of radios) {
            const rLabel = (
              r.getAttribute("aria-label") ||
              r.closest("label")?.innerText ||
              r.parentElement?.innerText ||
              ""
            )
              .trim()
              .toLowerCase();
            if (rLabel.includes("don't allow") || rLabel.includes("do not allow")) {
              const isSelected =
                (r as HTMLInputElement).checked === true ||
                r.getAttribute("aria-checked") === "true";
              if (!isSelected) r.click();
            }
          }
        },
        { targetEmail: baseForwardEmail, keepCopy: keepForwardedCopy }
      );

      await new Promise((r) => setTimeout(r, 1000));

      // Fill "Forward my email to:" input
      const fwdEmailInput = await page.$(
        '[role="dialog"] input[type="text"]:not([placeholder="Search settings"]), [role="dialog"] input[type="email"]'
      );
      if (fwdEmailInput) {
        await fwdEmailInput.click({ count: 3 }).catch(() => {});
        await page.keyboard.press("Backspace").catch(() => {});
        await fwdEmailInput.type(baseForwardEmail, { delay: 30 });
        await new Promise((r) => setTimeout(r, 500));
      }

      // Ensure "Keep a copy of forwarded messages" matches `keepForwardedCopy`
      await page
        .evaluate((keepCopy) => {
          const dialog = document.querySelector('[role="dialog"]');
          if (!dialog) return;
          const labels = Array.from(dialog.querySelectorAll("label, span")) as HTMLElement[];
          for (const lbl of labels) {
            const t = (lbl.innerText || "").trim().toLowerCase();
            if (t.includes("keep a copy of forwarded messages")) {
              const cb =
                (lbl.querySelector('input[type="checkbox"]') as HTMLInputElement | null) ||
                (lbl.parentElement?.querySelector('input[type="checkbox"]') as HTMLInputElement | null);
              if (cb && cb.checked !== keepCopy) {
                cb.click();
              }
              break;
            }
          }
        }, keepForwardedCopy)
        .catch(() => {});

      const savedFwd = await clickDialogSaveButton(page);
      if (savedFwd) {
        await new Promise((r) => setTimeout(r, 3000));
        addLog("OK", `Đã lưu cài đặt Forwarding (${baseForwardEmail}) + POP + IMAP!`);
      }
    }

    const summaryMsg = `Hoàn tất cài Rule "${ruleName}" (${ruleForwardAddress}) & Forwarding (${baseForwardEmail})!`;
    addLog("OK", summaryMsg);

    return {
      success: true,
      ruleStatus: "SUCCESS",
      message: summaryMsg,
      ruleForwardAddress,
      forwardingAddress: baseForwardEmail,
      resolvedProxy: effectiveProxyRaw,
      logs,
    };
  } catch (err: any) {
    const errMsg = err?.message || "Lỗi không xác định khi cài đặt Rule & Forwarding.";
    addLog("ERROR", errMsg);
    return {
      success: false,
      ruleStatus: "ERROR",
      message: errMsg,
      ruleForwardAddress,
      forwardingAddress: baseForwardEmail,
      logs,
    };
  } finally {
    if (browser) {
      await browser.close().catch(() => {});
    }
  }
}

/**
 * Scroll through an open Fluent UI listbox/menu and click the matching option text.
 */
async function selectDropdownOption(page: Page, keywords: string[]): Promise<boolean> {
  for (let scrollAttempt = 0; scrollAttempt < 8; scrollAttempt++) {
    const clicked = await page
      .evaluate((kws) => {
        const opts = Array.from(
          document.querySelectorAll('[role="option"], [role="menuitem"], [role="menuitemcheckbox"]')
        ) as HTMLElement[];
        for (const opt of opts) {
          const rect = opt.getBoundingClientRect();
          if (rect.width === 0 || rect.height === 0) continue;
          const txt = (opt.innerText || opt.getAttribute("title") || "").trim().toLowerCase();
          if (kws.some((k) => txt === k || txt.includes(k))) {
            opt.scrollIntoView({ block: "center" });
            opt.click();
            return true;
          }
        }
        // Scroll listbox container down to reveal virtualized items if needed
        const listbox = document.querySelector('[role="listbox"], .ms-Callout-main') as HTMLElement | null;
        if (listbox) {
          listbox.scrollTop += 220;
        }
        return false;
      }, keywords)
      .catch(() => false);

    if (clicked) return true;
    await new Promise((r) => setTimeout(r, 400));
  }
  return false;
}

/**
 * Click the primary "Save" button inside the Outlook Settings dialog.
 * Polls up to 8 times (4s) for the Save button to become enabled.
 */
async function clickDialogSaveButton(page: Page): Promise<boolean> {
  for (let attempt = 0; attempt < 8; attempt++) {
    const handles = await page.$$('[role="dialog"] button');
    for (const h of handles) {
      const txt = await h
        .evaluate((el) => (el.innerText || "").trim().toLowerCase())
        .catch(() => "");
      if (txt === "save" || txt === "lưu") {
        const disabled = await h
          .evaluate(
            (el) =>
              (el as HTMLButtonElement).disabled || el.getAttribute("aria-disabled") === "true"
          )
          .catch(() => false);
        if (!disabled) {
          await h.click();
          return true;
        }
      }
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  return false;
}

/**
 * Handle the "Sign in" step-up verification inside Settings -> Forwarding and IMAP / Rules.
 */
async function handleSettingsTabSignIn(
  browser: any,
  mainPage: Page,
  password: string,
  addLog: (level: "INFO" | "OK" | "WARN" | "ERROR", msg: string) => void
): Promise<"AUTHENTICATED" | "VERIFY_REQUIRED"> {
  const pagesBefore = await browser.pages();

  // Click the "Sign in" button/link inside [role="dialog"]
  const handles = await mainPage.$$('[role="dialog"] button, [role="dialog"] a, [role="dialog"] [role="button"]');
  for (const h of handles) {
    const txt = await h
      .evaluate((el) => (el.innerText || "").trim().toLowerCase())
      .catch(() => "");
    if (txt === "sign in" || txt === "đăng nhập" || txt.includes("sign in")) {
      await h.click();
      break;
    }
  }

  await new Promise((r) => setTimeout(r, 2500));
  const pagesAfter = await browser.pages();
  const popupPage = pagesAfter.find((p: Page) => !pagesBefore.includes(p)) || null;
  const targetPage: Page = popupPage || mainPage;

  const status = await checkAndHandleAuthScreen(targetPage, password, addLog);
  if (popupPage && !popupPage.isClosed() && status === "VERIFY_REQUIRED") {
    await popupPage.close().catch(() => {});
  }
  return status === "VERIFY_REQUIRED" ? "VERIFY_REQUIRED" : "AUTHENTICATED";
}
