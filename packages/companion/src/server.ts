import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import type { AddressInfo } from "node:net";
import path from "node:path";
import { Readable } from "node:stream";
import { isLoopbackUrl, minimizeRecord, normalizeSettings, type AcceptanceRecord, type TaskRecord } from "@loai/core";
import { Pairing, authenticate, loadAdminToken, type Principal } from "./auth.js";
import type { CompanionConfig } from "./config.js";
import { CompanionStore } from "./store.js";
import { SourceStore } from "./sources.js";

export const VERSION = "0.3.0";

/** Ollama endpoints the companion forwards. Everything else is refused. */
const OLLAMA_ROUTES: Record<string, "GET" | "POST"> = {
  "/api/version": "GET",
  "/api/tags": "GET",
  "/api/show": "POST",
  "/api/chat": "POST",
};

class HttpError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
  }
}

const MIME: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".xml": "application/xml",
  ".map": "application/json",
  ".woff2": "font/woff2",
};

export interface Companion {
  server: http.Server;
  url: string;
  pairing: Pairing;
  store: CompanionStore;
  sources: SourceStore;
  adminToken: string;
  close(): Promise<void>;
}

export async function startCompanion(cfg: CompanionConfig, opts: { log?: (msg: string) => void; fetchImpl?: typeof fetch; now?: () => number } = {}): Promise<Companion> {
  const log = opts.log ?? ((m: string) => console.log(m));
  const fetchImpl = opts.fetchImpl ?? fetch;
  if (!cfg.allowRemoteOllama && !isLoopbackUrl(cfg.ollamaUrl))
    throw new Error(`Ollama URL ${cfg.ollamaUrl} is not on this machine. Set allowRemoteOllama to true only if you accept that document content leaves this computer.`);

  const store = new CompanionStore(cfg.dbPath);
  const adminToken = loadAdminToken(cfg.home);
  const pairing = new Pairing(store);
  const sources = new SourceStore(opts.now);
  const ollamaBase = cfg.ollamaUrl.replace(/\/+$/, "");
  const tls = cfg.tlsCertPath && cfg.tlsKeyPath && fs.existsSync(cfg.tlsCertPath) && fs.existsSync(cfg.tlsKeyPath);

  const allowedHosts = new Set<string>();
  const allowedOrigins = new Set(cfg.allowedOrigins.map((o) => o.replace(/\/+$/, "").toLowerCase()));

  const handler = async (req: http.IncomingMessage, res: http.ServerResponse) => {
    const origin = typeof req.headers.origin === "string" ? req.headers.origin.toLowerCase() : undefined;
    try {
      // 1. DNS-rebinding protection: the Host header must name loopback.
      const host = (req.headers.host ?? "").toLowerCase();
      if (!allowedHosts.has(host)) throw new HttpError(421, "Misdirected request");

      // 2. Origin policy. Browsers always send Origin on cross-origin and on
      //    non-GET requests. An unknown origin is refused outright — this is a
      //    check in addition to token authentication, not a substitute.
      if (origin && !allowedOrigins.has(origin)) throw new HttpError(403, "Origin not allowed");
      if (origin) {
        res.setHeader("Access-Control-Allow-Origin", origin);
        res.setHeader("Vary", "Origin");
        res.setHeader("Access-Control-Allow-Headers", "authorization, content-type");
        res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
        res.setHeader("Access-Control-Max-Age", "600");
      }
      if (req.method === "OPTIONS") {
        res.writeHead(204).end();
        return;
      }

      const url = new URL(req.url ?? "/", "http://localhost");
      if (url.pathname.startsWith("/v1/")) await api(req, res, url);
      else await serveStatic(req, res, url);
    } catch (e) {
      const status = e instanceof HttpError ? e.status : 500;
      const message = e instanceof HttpError ? e.message : "Internal error";
      if (!(e instanceof HttpError)) log(`error: ${(e as Error).stack ?? e}`);
      if (!res.headersSent) {
        res.writeHead(status, { "content-type": "application/json" });
        res.end(JSON.stringify({ error: message }));
      } else res.destroy();
    }
  };

  async function readJson(req: http.IncomingMessage): Promise<any> {
    const ct = req.headers["content-type"] ?? "";
    // Requiring JSON blocks "simple" cross-site form posts.
    if (!ct.toLowerCase().startsWith("application/json")) throw new HttpError(415, "Content-Type must be application/json");
    const declared = Number(req.headers["content-length"] ?? 0);
    if (declared > cfg.maxBodyBytes) throw new HttpError(413, "Payload too large");
    let size = 0;
    const chunks: Buffer[] = [];
    for await (const c of req) {
      size += (c as Buffer).length;
      if (size > cfg.maxBodyBytes) throw new HttpError(413, "Payload too large");
      chunks.push(c as Buffer);
    }
    if (!size) return {};
    try {
      return JSON.parse(Buffer.concat(chunks).toString("utf8"));
    } catch {
      throw new HttpError(400, "Invalid JSON");
    }
  }

  const json = (res: http.ServerResponse, status: number, body: unknown) => {
    res.writeHead(status, { "content-type": "application/json", "cache-control": "no-store" });
    res.end(JSON.stringify(body));
  };

  async function api(req: http.IncomingMessage, res: http.ServerResponse, url: URL) {
    const route = `${req.method} ${url.pathname}`;

    // Unauthenticated: liveness (no details) and pairing.
    if (route === "GET /v1/health") return json(res, 200, { ok: true, service: "local-office-ai-companion", version: VERSION });
    if (route === "POST /v1/pair") {
      const body = await readJson(req);
      if (typeof body.code !== "string") throw new HttpError(400, "code required");
      const r = pairing.redeem(body.code, typeof body.name === "string" ? body.name : "Office add-in");
      if ("error" in r) return json(res, 401, r);
      return json(res, 200, r);
    }

    const principal = authenticate(req.headers.authorization, adminToken, store);
    if (!principal) throw new HttpError(401, "Not paired. Enter a pairing code in Settings.");

    if (route === "POST /v1/admin/pairing-code") {
      requireAdmin(principal);
      return json(res, 200, pairing.issueCode());
    }
    if (route === "GET /v1/admin/clients") {
      requireAdmin(principal);
      return json(res, 200, { clients: store.listClients() });
    }
    if (route === "DELETE /v1/admin/clients") {
      requireAdmin(principal);
      return json(res, 200, { revoked: store.revokeClients() });
    }

    if (url.pathname.startsWith("/v1/ollama/")) return proxyOllama(req, res, url.pathname.slice("/v1/ollama".length));
    if (route === "POST /v1/models/pull") return pullModel(req, res);

    if (route === "GET /v1/diagnostics") {
      let ollama: unknown;
      try {
        const r = await fetchImpl(ollamaBase + "/api/version", { signal: AbortSignal.timeout(4000) });
        ollama = r.ok ? { reachable: true, ...((await r.json()) as object) } : { reachable: false, status: r.status };
      } catch (e) {
        ollama = { reachable: false, error: (e as Error).message };
      }
      return json(res, 200, {
        companion: VERSION,
        node: process.version,
        platform: `${process.platform}-${process.arch}`,
        tls: !!tls,
        ollamaUrl: ollamaBase,
        ollamaIsLocal: isLoopbackUrl(ollamaBase),
        ollama,
        historyRecords: (await store.list(500)).length,
        pairedClients: store.listClients().length,
      });
    }

    if (route === "GET /v1/settings") return json(res, 200, normalizeSettings(store.getSetting("settings")));
    if (route === "PUT /v1/settings") {
      const s = normalizeSettings(await readJson(req));
      store.setSetting("settings", s);
      return json(res, 200, s);
    }

    if (route === "GET /v1/history") {
      const limit = Number(url.searchParams.get("limit") ?? 50);
      return json(res, 200, { records: await store.list(Number.isFinite(limit) ? limit : 50) });
    }
    if (route === "POST /v1/history") {
      const rec = (await readJson(req)) as TaskRecord;
      if (!rec || typeof rec.id !== "string" || typeof rec.request !== "string" || !Array.isArray(rec.actions))
        throw new HttpError(400, "Invalid task record");
      const settings = normalizeSettings(store.getSetting("settings"));
      if (!settings.historyEnabled || settings.historyRetentionDays === 0) return json(res, 200, { stored: false });
      await store.add(minimizeRecord(rec, settings.historyStoreFullArgs));
      await store.prune(settings.historyRetentionDays);
      return json(res, 201, { stored: true });
    }
    if (route === "DELETE /v1/history") {
      await store.clear();
      return json(res, 200, { cleared: true });
    }
    // Acceptance records (live Office test results) are kept as JSON files for the status matrix.
    if (route === "POST /v1/acceptance") {
      const rec = (await readJson(req)) as AcceptanceRecord;
      const problem = acceptanceProblem(rec);
      if (problem) throw new HttpError(400, problem);
      const dir = path.join(cfg.home, "acceptance");
      fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
      const file = path.join(dir, `${rec.environment.recordedAt.replace(/[^0-9TZ]/g, "")}-${rec.mode}-${rec.scenarioId}.json`);
      fs.writeFileSync(file, JSON.stringify(rec, null, 2), { mode: 0o600 });
      return json(res, 201, { stored: path.basename(file) });
    }
    if (route === "GET /v1/acceptance") return json(res, 200, { records: readAcceptanceRecords(cfg.home) });

    // Cross-app sources: shared on an explicit user action, memory only, expiring.
    if (route === "POST /v1/sources") {
      const r = sources.add(await readJson(req));
      return r.error ? json(res, 400, { error: r.error }) : json(res, 201, r.source);
    }
    if (route === "GET /v1/sources") return json(res, 200, { sources: sources.list() });
    if (route === "DELETE /v1/sources") return json(res, 200, { deleted: sources.clear() });
    const sm = /^(GET|DELETE) \/v1\/sources\/([0-9a-f-]{36})$/.exec(route);
    if (sm) {
      if (sm[1] === "DELETE") return json(res, 200, { deleted: sources.delete(sm[2]!) ? 1 : 0 });
      const s = sources.get(sm[2]!);
      return s ? json(res, 200, s) : json(res, 404, { error: "This shared content has expired or was removed." });
    }

    const hm = /^GET \/v1\/history\/([\w-]{1,80})$/.exec(route);
    if (hm) {
      const r = await store.get(hm[1]!);
      return r ? json(res, 200, r) : json(res, 404, { error: "Not found" });
    }

    throw new HttpError(404, "Unknown endpoint");
  }

  function requireAdmin(p: Principal) {
    if (p.kind !== "admin") throw new HttpError(403, "Admin only");
  }

  async function proxyOllama(req: http.IncomingMessage, res: http.ServerResponse, sub: string) {
    const method = OLLAMA_ROUTES[sub];
    if (!method || method !== req.method) throw new HttpError(404, "Endpoint not available through the companion");
    const body = method === "POST" ? JSON.stringify(await readJson(req)) : undefined;
    const ac = new AbortController();
    res.on("close", () => ac.abort());
    let upstream: Response;
    try {
      upstream = await fetchImpl(ollamaBase + sub, {
        method,
        headers: body ? { "content-type": "application/json" } : undefined,
        body,
        signal: ac.signal,
      });
    } catch (e) {
      throw new HttpError(502, `Cannot reach Ollama at ${ollamaBase}: ${(e as Error).message}`);
    }
    res.writeHead(upstream.status, {
      "content-type": upstream.headers.get("content-type") ?? "application/json",
      "cache-control": "no-store",
    });
    if (!upstream.body) return res.end();
    Readable.fromWeb(upstream.body as any).pipe(res);
  }

  /** Explicit, user-initiated model download (Settings → Install model). */
  async function pullModel(req: http.IncomingMessage, res: http.ServerResponse) {
    const body = await readJson(req);
    if (typeof body.model !== "string" || !/^[\w.\-:/]{1,200}$/.test(body.model)) throw new HttpError(400, "Invalid model name");
    if (body.confirm !== true) throw new HttpError(400, "Model downloads require explicit confirmation");
    const ac = new AbortController();
    res.on("close", () => ac.abort());
    const upstream = await fetchImpl(ollamaBase + "/api/pull", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ model: body.model, stream: true }),
      signal: ac.signal,
    }).catch((e) => {
      throw new HttpError(502, `Cannot reach Ollama: ${(e as Error).message}`);
    });
    res.writeHead(upstream.status, { "content-type": "application/x-ndjson" });
    if (!upstream.body) return res.end();
    Readable.fromWeb(upstream.body as any).pipe(res);
  }

  async function serveStatic(req: http.IncomingMessage, res: http.ServerResponse, url: URL) {
    if (req.method !== "GET" && req.method !== "HEAD") throw new HttpError(405, "Method not allowed");
    if (!cfg.staticDir) throw new HttpError(404, "Not found");
    const root = path.resolve(cfg.staticDir);
    let rel = decodeURIComponent(url.pathname);
    if (rel.endsWith("/")) rel += "index.html";
    const file = path.resolve(root, "." + rel);
    if (!file.startsWith(root + path.sep)) throw new HttpError(403, "Forbidden");
    let stat: fs.Stats;
    try {
      stat = fs.statSync(file);
    } catch {
      throw new HttpError(404, "Not found");
    }
    if (!stat.isFile()) throw new HttpError(404, "Not found");
    res.writeHead(200, {
      "content-type": MIME[path.extname(file)] ?? "application/octet-stream",
      "content-length": stat.size,
      "x-content-type-options": "nosniff",
      "referrer-policy": "no-referrer",
      "cache-control": file.endsWith(".html") ? "no-cache" : "public, max-age=3600",
    });
    if (req.method === "HEAD") return res.end();
    fs.createReadStream(file).pipe(res);
  }

  const server = tls
    ? https.createServer({ cert: fs.readFileSync(cfg.tlsCertPath!), key: fs.readFileSync(cfg.tlsKeyPath!) }, handler)
    : http.createServer(handler);
  server.requestTimeout = 0; // model calls can be long; per-request limits are enforced by the add-in
  server.headersTimeout = 30_000;

  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(cfg.port, cfg.bindHost, () => resolve());
  });
  const port = (server.address() as AddressInfo).port;
  for (const h of ["localhost", "127.0.0.1", "[::1]"]) allowedHosts.add(`${h}:${port}`);
  // When serving its own add-in, the companion's origin is always allowed.
  allowedOrigins.add(`${tls ? "https" : "http"}://localhost:${port}`);

  const url = `${tls ? "https" : "http"}://localhost:${port}`;
  return {
    server,
    url,
    pairing,
    store,
    sources,
    adminToken,
    close: () =>
      new Promise<void>((resolve) => {
        server.closeAllConnections?.();
        server.close(() => {
          store.close();
          resolve();
        });
      }),
  };
}

function acceptanceProblem(r: AcceptanceRecord): string | undefined {
  if (!r || r.schema !== 1) return "Unsupported record schema";
  if (typeof r.scenarioId !== "string" || !/^[\w.-]{1,100}$/.test(r.scenarioId)) return "Invalid scenarioId";
  if (r.mode !== "live" && r.mode !== "mocked") return "Invalid mode";
  if (!["passed", "failed", "blocked", "error"].includes(r.outcome)) return "Invalid outcome";
  if (!r.environment || typeof r.environment.recordedAt !== "string" || typeof r.environment.officeApp !== "string") return "Missing environment";
  if (!Array.isArray(r.checks) || !Array.isArray(r.manual)) return "Missing checks";
  return undefined;
}

export function readAcceptanceRecords(home: string): AcceptanceRecord[] {
  const dir = path.join(home, "acceptance");
  if (!fs.existsSync(dir)) return [];
  return fs
    .readdirSync(dir)
    .filter((f) => f.endsWith(".json"))
    .flatMap((f) => {
      try {
        return [JSON.parse(fs.readFileSync(path.join(dir, f), "utf8")) as AcceptanceRecord];
      } catch {
        return [];
      }
    });
}
