/**
 * Standalone BitBrowser Local API Client (default: http://127.0.0.1:54345)
 */

export interface BitBrowserProfileItem {
  id: string;
  name: string;
  accountCode?: string;
  proxy?: {
    raw: string;
  };
}

export class BitBrowserClient {
  private baseUrl: string;
  private apiToken: string;

  constructor(
    baseUrl = process.env.BITBROWSER_BASE_URL || "http://127.0.0.1:54345",
    apiToken = process.env.BITBROWSER_API_TOKEN || ""
  ) {
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.apiToken = apiToken;
  }

  private async request(path: string, body: Record<string, any> = {}): Promise<any> {
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };
    if (this.apiToken) {
      headers["x-api-key"] = this.apiToken;
    }

    const res = await fetch(`${this.baseUrl}${path}`, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      throw new Error(`BitBrowser HTTP ${res.status}`);
    }
    return res.json();
  }

  async checkHealth(): Promise<{ online: boolean }> {
    try {
      await this.request("/browser/list", { page: 0, pageSize: 1 });
      return { online: true };
    } catch {
      return { online: false };
    }
  }

  async listAllProfiles(): Promise<BitBrowserProfileItem[]> {
    const results: BitBrowserProfileItem[] = [];
    for (let page = 0; page < 20; page++) {
      const data = await this.request("/browser/list", { page, pageSize: 100 });
      const list: any[] = data?.data?.list || [];
      if (!list.length) break;

      for (const item of list) {
        const name = String(item.name || "");
        const match = name.match(/\b(KDP[-_\s]*\d+)\b/i);
        const accountCode = match ? match[1].replace(/[_\s]+/g, "-").toUpperCase() : "";
        const host = item.host || "";
        const port = item.port || "";
        const user = item.proxyUserName || "";
        const pass = item.proxyPassword || "";
        const rawProxy =
          host && port
            ? user && pass
              ? `${host}:${port}:${user}:${pass}`
              : `${host}:${port}`
            : "";

        results.push({
          id: item.id,
          name,
          accountCode,
          proxy: rawProxy ? { raw: rawProxy } : undefined,
        });
      }

      if (list.length < 100) break;
    }
    return results;
  }

  async findProfileByCode(accountCode: string): Promise<BitBrowserProfileItem | null> {
    const clean = (accountCode || "").trim().toUpperCase();
    if (!clean) return null;
    const all = await this.listAllProfiles();
    return (
      all.find(
        (p) =>
          p.accountCode?.toUpperCase() === clean ||
          p.name.toUpperCase().includes(clean)
      ) || null
    );
  }

  async openProfile(id: string): Promise<{ ws?: string; http?: string }> {
    const res = await this.request("/browser/open", { id, args: [] });
    return {
      ws: res?.data?.ws,
      http: res?.data?.http,
    };
  }

  async closeProfile(id: string): Promise<void> {
    await this.request("/browser/close", { id });
  }
}
