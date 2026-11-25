/**
 * OUTLOOK PROXY PARSER & HTTPS TUNNEL HELPER
 * Supports `host:port:user:pass`, `http://user:pass@host:port`, and `host:port`.
 * Provides:
 * - Chrome/Puppeteer `--proxy-server` + `page.authenticate` config
 * - ImapFlow `proxy` URL (`http://user:pass@host:port`)
 * - Zero-dependency HTTPS over HTTP CONNECT proxy (`fetchViaProxy`) for Microsoft OAuth2 & Graph API calls
 */

import * as http from "http";
import * as https from "https";
import * as tls from "tls";

export interface ParsedProxy {
  raw: string;
  host: string;
  port: number;
  username?: string;
  password?: string;
  proxyUrl: string;
  serverArg: string;
  displayLabel: string;
}

/**
 * Parse a proxy string such as `23.26.93.235:6717:wslpweyo:mnl6j6636mfl`
 * or `http://user:pass@23.26.93.235:6717`
 */
export function parseOutlookProxy(proxyStr?: string | null): ParsedProxy | null {
  const cleaned = (proxyStr || "").trim();
  if (!cleaned) return null;

  try {
    if (cleaned.startsWith("http://") || cleaned.startsWith("https://") || cleaned.startsWith("socks")) {
      const u = new URL(cleaned);
      const host = u.hostname;
      const port = Number(u.port) || 80;
      const username = u.username ? decodeURIComponent(u.username) : undefined;
      const password = u.password ? decodeURIComponent(u.password) : undefined;
      if (!host || !port) return null;
      const authPart =
        username && password
          ? `${encodeURIComponent(username)}:${encodeURIComponent(password)}@`
          : "";
      return {
        raw: cleaned,
        host,
        port,
        username,
        password,
        proxyUrl: `http://${authPart}${host}:${port}`,
        serverArg: `http://${host}:${port}`,
        displayLabel: `${host}:${port}`,
      };
    }

    // Format: host:port:user:pass or host:port
    const parts = cleaned.split(":");
    if (parts.length >= 2) {
      const host = parts[0].trim();
      const port = parseInt(parts[1].trim(), 10);
      if (!host || isNaN(port)) return null;
      const username = parts[2]?.trim() || undefined;
      const password = parts.slice(3).join(":").trim() || undefined;
      const authPart =
        username && password
          ? `${encodeURIComponent(username)}:${encodeURIComponent(password)}@`
          : "";
      return {
        raw: cleaned,
        host,
        port,
        username,
        password,
        proxyUrl: `http://${authPart}${host}:${port}`,
        serverArg: `http://${host}:${port}`,
        displayLabel: `${host}:${port}`,
      };
    }
  } catch {
    return null;
  }

  return null;
}

export interface SimpleFetchResponse {
  ok: boolean;
  status: number;
  text: () => Promise<string>;
  json: () => Promise<any>;
}

/**
 * Test if a proxy is alive and return its outbound IP address.
 */
export async function verifyOutlookProxy(
  proxyStr?: string | null,
  timeoutMs: number = 6000
): Promise<{ alive: boolean; ip?: string; error?: string; displayLabel?: string }> {
  const parsed = parseOutlookProxy(proxyStr);
  if (!parsed) {
    return { alive: false, error: "Không có cấu hình Proxy" };
  }
  try {
    const res = await fetchViaOutlookProxyStrict(
      "https://api.ipify.org?format=json",
      { method: "GET", timeoutMs },
      proxyStr
    );
    const data = await res.json();
    return {
      alive: true,
      ip: data?.ip || parsed.host,
      displayLabel: parsed.displayLabel,
    };
  } catch (err: any) {
    return {
      alive: false,
      error: err?.message || "Lỗi kết nối Proxy",
      displayLabel: parsed.displayLabel,
    };
  }
}

/**
 * Perform an HTTPS request through an HTTP CONNECT proxy (if `proxyStr` is provided),
 * automatically falling back to direct `fetch` if the proxy is expired/unreachable.
 */
export async function fetchViaOutlookProxy(
  urlStr: string,
  init: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
    timeoutMs?: number;
  } = {},
  proxyStr?: string | null
): Promise<SimpleFetchResponse> {
  const parsedProxy = parseOutlookProxy(proxyStr);
  if (!parsedProxy) {
    return fetchDirect(urlStr, init);
  }

  try {
    return await fetchViaOutlookProxyStrict(urlStr, init, proxyStr);
  } catch {
    // Fallback to direct fetch if the proxy is expired (e.g. HTTP 407) or unreachable
    return fetchDirect(urlStr, init);
  }
}

async function fetchDirect(
  urlStr: string,
  init: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
    timeoutMs?: number;
  } = {}
): Promise<SimpleFetchResponse> {
  const res = await fetch(urlStr, {
    method: init.method || "GET",
    headers: init.headers,
    body: init.body,
  });
  const bodyText = await res.text();
  return {
    ok: res.ok,
    status: res.status,
    text: async () => bodyText,
    json: async () => JSON.parse(bodyText),
  };
}

export async function fetchViaOutlookProxyStrict(
  urlStr: string,
  init: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
    timeoutMs?: number;
  } = {},
  proxyStr?: string | null
): Promise<SimpleFetchResponse> {
  const parsedProxy = parseOutlookProxy(proxyStr);
  if (!parsedProxy) {
    return fetchDirect(urlStr, init);
  }

  const targetUrl = new URL(urlStr);
  const targetPort = Number(targetUrl.port) || (targetUrl.protocol === "https:" ? 443 : 80);
  const timeoutMs = init.timeoutMs || 25000;

  return new Promise<SimpleFetchResponse>((resolve, reject) => {
    const connectHeaders: Record<string, string> = {
      Host: `${targetUrl.hostname}:${targetPort}`,
    };
    if (parsedProxy.username && parsedProxy.password) {
      const basicAuth = Buffer.from(
        `${parsedProxy.username}:${parsedProxy.password}`,
        "utf8"
      ).toString("base64");
      connectHeaders["Proxy-Authorization"] = `Basic ${basicAuth}`;
    }

    const connectReq = http.request({
      host: parsedProxy.host,
      port: parsedProxy.port,
      method: "CONNECT",
      path: `${targetUrl.hostname}:${targetPort}`,
      headers: connectHeaders,
      timeout: timeoutMs,
    });

    connectReq.on("timeout", () => {
      connectReq.destroy();
      reject(new Error(`Proxy CONNECT timeout (${parsedProxy.displayLabel})`));
    });

    connectReq.on("error", (err) => {
      reject(new Error(`Lỗi kết nối Proxy (${parsedProxy.displayLabel}): ${err.message}`));
    });

    connectReq.on("connect", (res, socket) => {
      if (res.statusCode !== 200) {
        socket.destroy();
        return reject(
          new Error(
            `Proxy (${parsedProxy.displayLabel}) từ chối kết nối: HTTP ${res.statusCode}`
          )
        );
      }

      const tlsSocket = tls.connect({
        socket,
        servername: targetUrl.hostname,
      });

      const reqHeaders: Record<string, string | number> = {
        Host: targetUrl.hostname,
        ...(init.headers || {}),
      };
      if (init.body) {
        reqHeaders["Content-Length"] = Buffer.byteLength(init.body, "utf8");
      }

      const httpsReq = https.request(
        {
          host: targetUrl.hostname,
          port: targetPort,
          path: targetUrl.pathname + targetUrl.search,
          method: init.method || "GET",
          headers: reqHeaders,
          createConnection: () => tlsSocket,
          timeout: timeoutMs,
        },
        (httpsRes) => {
          const chunks: Buffer[] = [];
          httpsRes.on("data", (chunk) =>
            chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
          );
          httpsRes.on("end", () => {
            const bodyText = Buffer.concat(chunks).toString("utf8");
            const status = httpsRes.statusCode || 500;
            resolve({
              ok: status >= 200 && status < 300,
              status,
              text: async () => bodyText,
              json: async () => JSON.parse(bodyText),
            });
          });
        }
      );

      httpsReq.on("timeout", () => {
        httpsReq.destroy();
        reject(new Error(`Request timeout qua Proxy (${parsedProxy.displayLabel})`));
      });

      httpsReq.on("error", (err) => {
        reject(new Error(`Lỗi HTTPS qua Proxy (${parsedProxy.displayLabel}): ${err.message}`));
      });

      if (init.body) {
        httpsReq.write(init.body);
      }
      httpsReq.end();
    });

    connectReq.end();
  });
}
