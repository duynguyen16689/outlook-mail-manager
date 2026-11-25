/**
 * OUTLOOK OAUTH2 SERVICE (Public Client ID + Refresh Token Flow)
 * Uses Mozilla Thunderbird's official verified Public Client ID by default.
 * Supports both manual redirect URL code extraction and automated browser token capture.
 */

import { fetchViaOutlookProxy } from "./outlook-proxy";

export const DEFAULT_OUTLOOK_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753";
export const OUTLOOK_REDIRECT_URI = "https://localhost";
export const OUTLOOK_GRAPH_SCOPES = [
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send",
  "offline_access",
].join(" ");
export const OUTLOOK_IMAP_SCOPES = [
  "https://outlook.office.com/IMAP.AccessAsUser.All",
  "https://outlook.office.com/POP.AccessAsUser.All",
  "https://outlook.office.com/SMTP.Send",
  "offline_access",
].join(" ");

export interface OutlookTokenResponse {
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
  scope: string;
  tokenType: string;
}

/**
 * Build the Microsoft OAuth2 Authorization URL for a given email.
 * Requests Microsoft Graph Mail.ReadWrite + Mail.Send + offline_access so that
 * reading INBOX + Junk and Native Forwarding work even when IMAP is disabled on the account.
 */
export function buildOutlookAuthUrl(
  _email?: string,
  clientId: string = DEFAULT_OUTLOOK_CLIENT_ID
): string {
  const params = new URLSearchParams({
    client_id: clientId || DEFAULT_OUTLOOK_CLIENT_ID,
    response_type: "code",
    redirect_uri: OUTLOOK_REDIRECT_URI,
    response_mode: "query",
    scope: OUTLOOK_GRAPH_SCOPES,
  });

  // Note: Do NOT pass `login_hint` in query params!
  // Passing `login_hint` skips Microsoft's initial email input screen (`GetCredentialType`),
  // which causes Microsoft's anti-bot protection on accounts with recovery emails
  // to reject valid passwords with "You've tried to sign in too many times".

  return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
}

/**
 * Extract the `code` query parameter from either a full redirect URL
 * (e.g. `https://localhost/?code=M.R3_BAY...`) or a raw code string.
 */
export function extractAuthCode(urlOrCode: string): string {
  const raw = (urlOrCode || "").trim();
  if (!raw) {
    throw new Error("Vui lòng nhập URL kết quả hoặc mã code OAuth2.");
  }

  if (raw.includes("error=")) {
    try {
      const parsed = new URL(raw.startsWith("http") ? raw : `https://localhost/${raw}`);
      const err = parsed.searchParams.get("error");
      const desc = parsed.searchParams.get("error_description");
      throw new Error(`Microsoft OAuth Error (${err}): ${desc || "Không rõ nguyên nhân"}`);
    } catch (e: any) {
      if (e.message?.startsWith("Microsoft OAuth Error")) throw e;
    }
  }

  if (raw.includes("code=")) {
    try {
      const parsed = new URL(raw.startsWith("http") ? raw : `https://localhost/?${raw.split("?").pop()}`);
      const code = parsed.searchParams.get("code");
      if (code) return code.trim();
    } catch {
      const match = raw.match(/[?&]code=([^&#\s]+)/i);
      if (match?.[1]) return decodeURIComponent(match[1].trim());
    }
  }

  return raw;
}

/**
 * Exchange an OAuth2 authorization code (or full localhost redirect URL) for access_token & refresh_token.
 */
export async function exchangeCodeForTokens(
  codeOrUrl: string,
  clientId: string = DEFAULT_OUTLOOK_CLIENT_ID,
  proxy?: string
): Promise<OutlookTokenResponse> {
  const code = extractAuthCode(codeOrUrl);
  const effectiveClientId = (clientId || DEFAULT_OUTLOOK_CLIENT_ID).trim();

  const body = new URLSearchParams({
    client_id: effectiveClientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: OUTLOOK_REDIRECT_URI,
  });

  const res = await fetchViaOutlookProxy(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    },
    proxy
  );

  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) {
    const errMsg =
      data.error_description ||
      data.error ||
      `Đổi mã code thất bại (HTTP ${res.status})`;
    throw new Error(errMsg);
  }

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || "",
    expiresIn: Number(data.expires_in) || 3600,
    scope: data.scope || "",
    tokenType: data.token_type || "Bearer",
  };
}

/**
 * Exchange a saved refresh_token for a fresh access_token.
 * Requests Microsoft Graph API scope (`https://graph.microsoft.com/.default offline_access`)
 * first so we can read mail via Graph REST API even when IMAP is disabled on the mailbox,
 * and falls back to default scope if needed.
 */
export async function refreshOutlookAccessToken(
  refreshToken: string,
  clientId: string = DEFAULT_OUTLOOK_CLIENT_ID,
  proxy?: string
): Promise<OutlookTokenResponse> {
  const cleanRefresh = (refreshToken || "").trim();
  if (!cleanRefresh) {
    throw new Error("Tài khoản chưa có Refresh Token.");
  }

  const effectiveClientId = (clientId || DEFAULT_OUTLOOK_CLIENT_ID).trim();

  // 1. Try requesting Microsoft Graph scope first (works even if IMAP is disabled on the account!)
  const graphBody = new URLSearchParams({
    client_id: effectiveClientId,
    grant_type: "refresh_token",
    refresh_token: cleanRefresh,
    scope: "https://graph.microsoft.com/.default offline_access",
  });

  const graphRes = await fetchViaOutlookProxy(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: graphBody.toString(),
    },
    proxy
  );

  const graphData = await graphRes.json().catch(() => ({}));
  const graphScope = String(graphData.scope || "");
  if (graphRes.ok && graphData.access_token && /\bMail\.Read(Write)?\b/i.test(graphScope)) {
    return {
      accessToken: graphData.access_token,
      refreshToken: graphData.refresh_token || cleanRefresh,
      expiresIn: Number(graphData.expires_in) || 3600,
      scope: graphScope,
      tokenType: graphData.token_type || "Bearer",
    };
  }

  // 2. Request Outlook Office365 IMAP/SMTP scope (audience: https://outlook.office.com)
  const imapBody = new URLSearchParams({
    client_id: effectiveClientId,
    grant_type: "refresh_token",
    refresh_token: cleanRefresh,
    scope: OUTLOOK_IMAP_SCOPES,
  });

  const imapRes = await fetchViaOutlookProxy(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: imapBody.toString(),
    },
    proxy
  );

  const imapData = await imapRes.json().catch(() => ({}));
  if (imapRes.ok && imapData.access_token) {
    return {
      accessToken: imapData.access_token,
      refreshToken: imapData.refresh_token || cleanRefresh,
      expiresIn: Number(imapData.expires_in) || 3600,
      scope: imapData.scope || OUTLOOK_IMAP_SCOPES,
      tokenType: imapData.token_type || "Bearer",
    };
  }

  // 3. Fallback: request default (original) scope
  const body = new URLSearchParams({
    client_id: effectiveClientId,
    grant_type: "refresh_token",
    refresh_token: cleanRefresh,
  });

  const res = await fetchViaOutlookProxy(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    },
    proxy
  );

  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) {
    const errMsg =
      data.error_description ||
      data.error ||
      `Làm mới Access Token thất bại (HTTP ${res.status})`;
    throw new Error(errMsg);
  }

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || cleanRefresh,
    expiresIn: Number(data.expires_in) || 3600,
    scope: data.scope || "",
    tokenType: data.token_type || "Bearer",
  };
}
