import type { AssistantSettings, HistoryStore, TaskRecord } from "@loai/core";

const TOKEN_KEY = "loai.companionToken";

function storage(): Storage | null {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
}

export function getToken(): string | null {
  return storage()?.getItem(TOKEN_KEY) ?? null;
}

export function setToken(t: string | null) {
  const s = storage();
  if (!s) return;
  if (t) s.setItem(TOKEN_KEY, t);
  else s.removeItem(TOKEN_KEY);
}

export class CompanionError extends Error {
  constructor(
    message: string,
    public status: number,
  ) {
    super(message);
  }
}

/** Client for the local companion's authenticated API. */
export class CompanionClient {
  constructor(public baseUrl: string) {
    this.baseUrl = baseUrl.replace(/\/+$/, "");
  }

  get paired(): boolean {
    return !!getToken();
  }

  authHeaders(): Record<string, string> {
    const t = getToken();
    return t ? { authorization: `Bearer ${t}` } : {};
  }

  private async call<T>(path: string, init: RequestInit = {}): Promise<T> {
    let res: Response;
    try {
      res = await fetch(this.baseUrl + path, {
        ...init,
        headers: { ...(init.body ? { "content-type": "application/json" } : {}), ...this.authHeaders(), ...(init.headers as any) },
        signal: init.signal ?? AbortSignal.timeout(15_000),
      });
    } catch (e) {
      throw new CompanionError(`Cannot reach the companion at ${this.baseUrl}. Is it running?`, 0);
    }
    const body = await res.json().catch(() => ({}));
    if (res.status === 401 && path !== "/v1/pair") setToken(null);
    if (!res.ok) throw new CompanionError((body as any).error ?? `HTTP ${res.status}`, res.status);
    return body as T;
  }

  health() {
    return this.call<{ ok: boolean; version: string }>("/v1/health", { signal: AbortSignal.timeout(3000) });
  }

  async pair(code: string): Promise<void> {
    const r = await this.call<{ token: string }>("/v1/pair", { method: "POST", body: JSON.stringify({ code, name: `Office add-in (${Office?.context?.host ?? "browser"})` }) });
    setToken(r.token);
  }

  unpair() {
    setToken(null);
  }

  diagnostics() {
    return this.call<Record<string, any>>("/v1/diagnostics");
  }

  getSettings() {
    return this.call<AssistantSettings>("/v1/settings");
  }

  putSettings(s: AssistantSettings) {
    return this.call<AssistantSettings>("/v1/settings", { method: "PUT", body: JSON.stringify(s) });
  }

  /** Explicit, user-initiated model download. Streams progress lines. */
  async pullModel(model: string, onProgress: (status: string, pct?: number) => void, signal?: AbortSignal) {
    const res = await fetch(this.baseUrl + "/v1/models/pull", {
      method: "POST",
      headers: { "content-type": "application/json", ...this.authHeaders() },
      body: JSON.stringify({ model, confirm: true }),
      signal,
    });
    if (!res.ok || !res.body) throw new CompanionError(((await res.json().catch(() => ({}))) as any).error ?? `HTTP ${res.status}`, res.status);
    const reader = res.body.getReader();
    const dec = new TextDecoder();
    let buf = "";
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      let i: number;
      while ((i = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, i).trim();
        buf = buf.slice(i + 1);
        if (!line) continue;
        const m = JSON.parse(line);
        if (m.error) throw new CompanionError(m.error, 500);
        onProgress(m.status ?? "", m.total ? Math.round((100 * (m.completed ?? 0)) / m.total) : undefined);
      }
    }
  }

  history(): HistoryStore {
    const self = this;
    return {
      async add(r: TaskRecord) {
        await self.call("/v1/history", { method: "POST", body: JSON.stringify(r) });
      },
      async list(limit = 50) {
        return (await self.call<{ records: TaskRecord[] }>(`/v1/history?limit=${limit}`)).records;
      },
      async get(id: string) {
        return self.call<TaskRecord>(`/v1/history/${encodeURIComponent(id)}`).catch(() => undefined);
      },
      async clear() {
        await self.call("/v1/history", { method: "DELETE" });
      },
      async prune() {
        return 0; // the companion prunes on write according to retention settings
      },
    };
  }
}
