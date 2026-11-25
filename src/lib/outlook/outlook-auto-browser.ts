/**
 * AUTOMATED OUTLOOK OAUTH2 TOKEN FETCHER VIA PUPPETEER
 * Uses a clean, deterministic step-by-step login sequence identical to human interaction:
 * 1) Open OAuth2 URL & wait for Email input
 * 2) Type Email + press Enter -> wait 3.5s for transition
 * 3) If "Verify your email" appears, click "Use your password" -> wait 2.2s for transition
 * 4) Type Password + verify DOM value + press Enter -> wait for 302 redirect to localhost/?code=...
 * 5) Auto-click Consent / KMSI ("Yes" / "Accept" / "Skip for now") if shown
 */

import fs from "fs";
import puppeteer from "puppeteer-core";
import {
  buildOutlookAuthUrl,
  DEFAULT_OUTLOOK_CLIENT_ID,
  exchangeCodeForTokens,
  OutlookTokenResponse,
} from "./outlook-oauth";
import { parseOutlookProxy, verifyOutlookProxy } from "./outlook-proxy";
import { BitBrowserClient } from "@/lib/bitbrowser/bitbrowser-client";

export interface BrowserTokenCaptureResult extends OutlookTokenResponse {
  logs: string[];
  resolvedProxy?: string;
}

function nowTimeStr(): string {
  const d = new Date();
  return d.toLocaleTimeString("vi-VN", { hour12: false });
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

  throw new Error(
    "Không tìm thấy trình duyệt Chrome hoặc Edge trên máy để chạy tự động. Vui lòng dùng cách Mở Link & Dán URL."
  );
}

function isOAuthCallbackUrl(url?: string | null): boolean {
  if (!url) return false;
  const trimmed = url.trim();
  return (
    (trimmed.startsWith("https://localhost") || trimmed.startsWith("http://localhost")) &&
    (trimmed.includes("code=") || trimmed.includes("error="))
  );
}

export async function captureOutlookTokenViaBrowser(
  email: string,
  password?: string,
  clientId: string = DEFAULT_OUTLOOK_CLIENT_ID,
  proxy?: string,
  accountCode?: string,
  timeoutMs: number = 60000
): Promise<BrowserTokenCaptureResult> {
  const logs: string[] = [];
  const logTag = accountCode ? `${accountCode} | ${email}` : email;
  const addLog = (level: "INFO" | "OK" | "WARN" | "ERROR", msg: string) => {
    const line = `[${nowTimeStr()}] [${level}] [${logTag}] ${msg}`;
    logs.push(line);
    if (level === "ERROR") {
      console.error(line);
    } else {
      console.log(line);
    }
  };

  let browser;
  let openedBbProfileId: string | null = null;
  let pageToClose: any = null;
  const bbClient = new BitBrowserClient();

  try {
    if (!password || !password.trim()) {
      throw new Error(`Tài khoản ${email} chưa có mật khẩu. Vui lòng nhập mật khẩu hoặc dùng Lấy Token Thủ Công.`);
    }

    let effectiveProxyRaw = (proxy || "").trim();
    const authUrl = buildOutlookAuthUrl(email, clientId);
    let page: any = null;

    // 1. Try opening the account's native BitBrowser profile first (preserves trusted Microsoft cookies & fingerprint)
    if (accountCode) {
      try {
        const bbProfile = await bbClient.findProfileByCode(accountCode);
        if (bbProfile?.id) {
          if (bbProfile.proxy?.raw) {
            effectiveProxyRaw = bbProfile.proxy.raw.trim();
          }
          addLog(
            "INFO",
            `Đang mở hồ sơ BitBrowser gốc của ${accountCode} (tận dụng phiên đăng nhập & vân tay tin cậy)...`
          );
          const openRes = await bbClient.openProfile(bbProfile.id);
          if (openRes.ws) {
            openedBbProfileId = bbProfile.id;
            browser = await puppeteer.connect({
              browserWSEndpoint: openRes.ws,
              defaultViewport: null,
            });
            page = await browser.newPage();
            pageToClose = page;
            addLog("OK", `Đã kết nối hồ sơ BitBrowser (${accountCode}).`);
          }
        }
      } catch (bbErr: any) {
        addLog("WARN", `Không mở được BitBrowser (${bbErr?.message || "offline"}), chuyển sang Chrome cục bộ...`);
      }
    }

    // 2. Fallback to local Chrome/Edge if BitBrowser profile was not opened
    if (!page) {
      const executablePath = findLocalBrowserExecutable();
      const parsedProxy = parseOutlookProxy(effectiveProxyRaw);
      let activeProxy = parsedProxy;

      const launchArgs = [
        "--window-size=540,720",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
      ];
      if (parsedProxy) {
        addLog("INFO", `Đang kiểm tra kết nối Proxy riêng (${parsedProxy.displayLabel})...`);
        const check = await verifyOutlookProxy(effectiveProxyRaw, 5500);
        if (check.alive) {
          launchArgs.push(`--proxy-server=${parsedProxy.serverArg}`);
          addLog("OK", `Proxy hoạt động tốt (IP: ${check.ip}) -> Khởi chạy trình duyệt qua Proxy...`);
        } else {
          activeProxy = null;
          addLog(
            "WARN",
            `Proxy lưu trong DB (${parsedProxy.displayLabel}) không phản hồi (${check.error}).`
          );
        }
      }

      if (!activeProxy) {
        addLog("WARN", `Không có Proxy hoạt động -> Chuyển sang mở trình duyệt trực tiếp.`);
      }

      browser = await puppeteer.launch({
        executablePath,
        headless: false,
        defaultViewport: null,
        args: launchArgs,
      });

      const pages = await browser.pages();
      page = pages[0] || (await browser.newPage());

      if (activeProxy?.username && activeProxy?.password) {
        await page.authenticate({
          username: activeProxy.username,
          password: activeProxy.password,
        });
      }
    }

    let capturedUrl: string | null = null;
    const tryCapture = (candidate?: string | null, source: string = "url") => {
      if (!capturedUrl && isOAuthCallbackUrl(candidate)) {
        capturedUrl = candidate!.trim();
        addLog("OK", `Bắt được URL OAuth2 callback (${source}): ${capturedUrl.slice(0, 65)}...`);
      }
    };

    // Passive network listeners (No CDP Network.enable to avoid Microsoft anti-bot triggers)
    page.on("response", (res: any) => {
      if (res.status() >= 300 && res.status() < 400) {
        const headers = res.headers();
        tryCapture(headers["location"] || headers["Location"], `HTTP.${res.status()}`);
      }
    });

    page.on("requestfailed", (req: any) => {
      tryCapture(req.url(), "requestfailed");
    });

    page.on("framenavigated", (frame: any) => {
      if (frame === page.mainFrame()) {
        tryCapture(frame.url(), "framenavigated");
      }
    });

    // STEP 1: Open OAuth2 Login Page
    addLog("INFO", "Bước 1: Mở trang đăng nhập Microsoft OAuth2...");
    await page.goto(authUrl, { waitUntil: "networkidle2", timeout: 30000 }).catch(() => {});

    if (!capturedUrl) {
      await page
        .waitForSelector('input[type="email"], input[name="loginfmt"]', { timeout: 6000 })
        .catch(() => {});
    }

    // STEP 2: Type Email & Press Enter (if not already auto-signed-in via BitBrowser session)
    const emailInput = !capturedUrl
      ? await page.$('input[type="email"], input[name="loginfmt"]').catch(() => null)
      : null;
    if (emailInput && !capturedUrl) {
      try {
        addLog("INFO", `Bước 2: Điền địa chỉ Email (${email}) và bấm Enter...`);
        await emailInput.click();
        await emailInput.type(email.trim(), { delay: 30 });
        await new Promise((r) => setTimeout(r, 500));
        await page.keyboard.press("Enter");
      } catch {
        // Ignore if page navigated away while typing
      }
    }

    // STEP 3: Wait up to 18s (especially over residential proxy) for either:
    //   - "Use your password" / "Other ways to sign in" button on the "Verify your email" screen, OR
    //   - Visible password input (`input[type="password"]`)
    const step3Start = Date.now();
    let clickedUsePassword = false;
    while (!capturedUrl && emailInput && Date.now() - step3Start < 18000) {
      await new Promise((r) => setTimeout(r, 800));
      if (capturedUrl) break;

      const screenState = await page
        .evaluate(() => {
          const passEl = document.querySelector('input[type="password"]') as HTMLElement | null;
          if (passEl && passEl.offsetParent !== null) {
            return { hasPasswordInput: true, clickedUsePass: false };
          }
          const nodes = Array.from(
            document.querySelectorAll('span[role="button"], a, button')
          ) as HTMLElement[];
          for (const el of nodes) {
            if (el.offsetParent === null) continue;
            const txt = (el.innerText || "").trim().toLowerCase();
            if (
              txt.includes("use your password") ||
              txt.includes("dùng mật khẩu") ||
              txt.includes("other ways to sign in")
            ) {
              el.click();
              return { hasPasswordInput: false, clickedUsePass: true };
            }
          }
          return { hasPasswordInput: false, clickedUsePass: false };
        })
        .catch(() => ({ hasPasswordInput: false, clickedUsePass: false }));

      if (screenState.clickedUsePass) {
        clickedUsePassword = true;
        addLog(
          "INFO",
          `Bước 3: Chuyển từ xác minh mail phụ sang dùng mật khẩu ("Use your password")...`
        );
        // Wait 2.4s so Fluent UI slide-in transition finishes completely before typing password
        await new Promise((r) => setTimeout(r, 2400));
        break;
      }

      if (screenState.hasPasswordInput) {
        // Wait 1.2s for input animation to settle
        await new Promise((r) => setTimeout(r, 1200));
        break;
      }
    }

    // STEP 4: Type Password & Press Enter (if not already auto-signed-in via BitBrowser session)
    if (!capturedUrl) {
      await page
        .waitForSelector('input[type="password"]', { visible: true, timeout: 6000 })
        .catch(() => {});
      if (clickedUsePassword) {
        await new Promise((r) => setTimeout(r, 500));
      }
      const passInput = !capturedUrl
        ? await page.$('input[type="password"]').catch(() => null)
        : null;
      if (passInput && !capturedUrl) {
        try {
          addLog("INFO", "Bước 4: Điền mật khẩu và bấm Enter...");
          await passInput.click();
          await passInput.type(password.trim(), { delay: 35 });

          const actualVal = await passInput
            .evaluate((el: any) => (el as HTMLInputElement).value)
            .catch(() => password.trim());
          if (actualVal !== password.trim()) {
            addLog("WARN", "Phát hiện lệch ký tự khi gõ nhanh -> Đang điền lại mật khẩu chính xác...");
            await passInput.click({ count: 3 }).catch(() => {});
            await page.keyboard.press("Backspace").catch(() => {});
            await passInput.type(password.trim(), { delay: 45 });
          }

          await new Promise((r) => setTimeout(r, 500));
          await page.keyboard.press("Enter");
          await new Promise((r) => setTimeout(r, 4000));
        } catch {
          // Ignore if page navigated away while typing password
        }
      }
    }

    // STEP 5: Poll for redirect URL, post-login Consent/KMSI screens, or Microsoft errors
    const startWait = Date.now();
    let lastScreenSummary = "";

    while (!capturedUrl && Date.now() - startWait < timeoutMs) {
      if (page.isClosed()) {
        throw new Error("Cửa sổ trình duyệt đã bị đóng trước khi hoàn tất lấy Token.");
      }

      tryCapture(page.url(), "page.url");
      if (capturedUrl) break;

      // Check chrome-error://chromewebdata/ failedUrl
      const failedRedirectUrl = await page
        .evaluate(() => {
          const href = window.location.href;
          if (href.startsWith("https://localhost") || href.startsWith("http://localhost")) {
            return href;
          }
          const loadData = (window as any).loadTimeData?.data_;
          return loadData?.summary?.failedUrl || loadData?.failedUrl || null;
        })
        .catch(() => null);

      tryCapture(failedRedirectUrl, "chrome-error");
      if (capturedUrl) break;

      const bodyText = await page
        .evaluate(() => (document.body ? document.body.innerText : ""))
        .catch(() => "");
      const lowerText = bodyText.toLowerCase();

      const summary = bodyText
        .split(/\r?\n/)
        .map((s: string) => s.trim())
        .filter(Boolean)
        .slice(0, 4)
        .join(" | ");

      if (summary && summary !== lastScreenSummary) {
        lastScreenSummary = summary;
        addLog("INFO", `Màn hình Microsoft: "${summary}"`);
      }

      // Check explicit Microsoft errors / security checkpoints
      if (
        lowerText.includes("help us protect your account") ||
        lowerText.includes("detected something unusual about this sign-in") ||
        (lowerText.includes("almost there") && lowerText.includes("send a code"))
      ) {
        throw new Error(
          `Tài khoản ${email} bị Microsoft yêu cầu xác minh bảo mật ("Help us protect your account / Send a code"). Cần nhận mã từ mail khôi phục.`
        );
      }

      if (
        lowerText.includes("that password is incorrect") ||
        lowerText.includes("your account or password is incorrect") ||
        lowerText.includes("mật khẩu không đúng")
      ) {
        throw new Error(`Sai mật khẩu Microsoft cho tài khoản ${email} ("That password is incorrect").`);
      }

      if (
        lowerText.includes("tried to sign in too many times") ||
        lowerText.includes("too many times with an incorrect account or password")
      ) {
        throw new Error(
          `Tài khoản ${email} bị Microsoft báo nhập sai quá nhiều lần ("Tried to sign in too many times").`
        );
      }

      if (lowerText.includes("your account has been locked")) {
        throw new Error(`Tài khoản ${email} đã bị khóa ("Your account has been locked").`);
      }

      // Auto-click "Skip for now", "Stay signed in (Yes)", or OAuth "Accept"
      const clickedBtn = await page
        .evaluate(() => {
          const skipLink = document.querySelector("#iShowSkip") as HTMLElement | null;
          if (skipLink && skipLink.offsetParent !== null) {
            skipLink.click();
            return "Skip for now (#iShowSkip)";
          }

          const selectors = [
            "#acceptButton",
            "#idBtn_Accept",
            '[data-testid="appConsentPrimaryButton"]',
            'input[name="ucaccept"]',
          ];
          for (const sel of selectors) {
            const btn = document.querySelector(sel) as HTMLElement | null;
            if (btn && btn.offsetParent !== null) {
              btn.click();
              return sel;
            }
          }

          const buttons = Array.from(
            document.querySelectorAll(
              'button, input[type="submit"], input[type="button"], [data-testid="primaryButton"]'
            )
          ) as HTMLElement[];

          for (const b of buttons) {
            if (b.offsetParent === null) continue;
            const txt = (b.innerText || (b as HTMLInputElement).value || "").trim().toLowerCase();
            if (
              txt === "accept" ||
              txt === "yes" ||
              txt === "đồng ý" ||
              txt === "có" ||
              txt === "chấp nhận"
            ) {
              b.click();
              return `Button("${txt}")`;
            }
          }
          return null;
        })
        .catch(() => null);

      if (clickedBtn) {
        addLog("INFO", `Tự động bấm xác nhận: ${clickedBtn}`);
        await new Promise((r) => setTimeout(r, 1500));
      } else {
        await new Promise((r) => setTimeout(r, 700));
      }
    }

    if (!capturedUrl) {
      throw new Error(
        `Hết thời gian chờ mà chưa nhận được URL OAuth2. Màn hình cuối: "${lastScreenSummary}"`
      );
    }

    addLog("INFO", "Bước 5: Đang đổi mã Authorization Code sang Refresh Token...");
    const tokenRes = await exchangeCodeForTokens(capturedUrl, clientId, effectiveProxyRaw);
    addLog("OK", `Hoàn tất! Đã lưu Refresh Token (${tokenRes.refreshToken.length} ký tự).`);

    return {
      ...tokenRes,
      logs,
      resolvedProxy: effectiveProxyRaw,
    };
  } catch (err: any) {
    addLog("ERROR", err?.message || "Lỗi không xác định khi lấy Token.");
    const enrichedError: any = new Error(err?.message || "Lỗi lấy Token OAuth2");
    enrichedError.logs = logs;
    throw enrichedError;
  } finally {
    if (pageToClose) {
      await pageToClose.close().catch(() => {});
    }
    if (openedBbProfileId) {
      await bbClient.closeProfile(openedBbProfileId).catch(() => {});
    } else if (browser) {
      await browser.close().catch(() => {});
    }
  }
}
