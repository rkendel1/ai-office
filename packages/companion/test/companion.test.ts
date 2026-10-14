import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OllamaProvider } from "@loai/core";
import { loadConfig } from "../src/config.js";
import { startCompanion, type Companion } from "../src/server.js";
import { startFakeOllama } from "../../core/test/fake-ollama.js";

let c: Companion;
let ollama: Awaited<ReturnType<typeof startFakeOllama>>;
let home: string;
const ORIGIN = "https://localhost:3000";

/** Raw request so tests can set Host/Origin headers browsers would send. */
function req(p: { method?: string; path: string; headers?: Record<string, string>; body?: string }): Promise<{ status: number; body: any; text: string }> {
  const u = new URL(c.url);
  return new Promise((resolve, reject) => {
    const r = http.request(
      { host: "127.0.0.1", port: u.port, path: p.path, method: p.method ?? "GET", headers: { host: `localhost:${u.port}`, ...p.headers } },
      (res) => {
        let text = "";
        res.on("data", (d) => (text += d));
        res.on("end", () => {
          let body: any;
          try {
            body = JSON.parse(text);
          } catch {
            body = undefined;
          }
          resolve({ status: res.statusCode!, body, text });
        });
      },
    );
    r.on("error", reject);
    if (p.body) r.write(p.body);
    r.end();
  });
}

async function pairClient(): Promise<string> {
  const { code } = c.pairing.issueCode();
  const r = await req({ method: "POST", path: "/v1/pair", headers: { origin: ORIGIN, "content-type": "application/json" }, body: JSON.stringify({ code, name: "test" }) });
  expect(r.status).toBe(200);
  return r.body.token;
}

beforeEach(async () => {
  home = fs.mkdtempSync(path.join(os.tmpdir(), "loai-"));
  ollama = await startFakeOllama({ replies: [{ content: "Hello from local model" }] });
  fs.mkdirSync(path.join(home, "static"));
  fs.writeFileSync(path.join(home, "static", "index.html"), "<h1>pane</h1>");
  fs.writeFileSync(path.join(home, "secret.txt"), "secret");
  c = await startCompanion(
    loadConfig({ home, port: 0, ollamaUrl: ollama.url, tlsCertPath: undefined, tlsKeyPath: undefined, allowedOrigins: [ORIGIN], staticDir: path.join(home, "static"), maxBodyBytes: 10_000 }),
    { log: () => {} },
  );
});

afterEach(async () => {
  await c.close();
  await ollama.close();
  fs.rmSync(home, { recursive: true, force: true });
});

describe("companion security", () => {
  it("binds to loopback only", () => {
    const addr = c.server.address() as any;
    expect(addr.address).toBe("127.0.0.1");
    expect(() => loadConfig({ home, bindHost: "0.0.0.0" })).toThrow(/Refusing to bind/);
  });

  it("rejects requests without a token", async () => {
    const r = await req({ path: "/v1/history", headers: { origin: ORIGIN } });
    expect(r.status).toBe(401);
  });

  it("rejects unknown origins even with a valid token (malicious website)", async () => {
    const token = await pairClient();
    const r = await req({ path: "/v1/history", headers: { origin: "https://evil.test", authorization: `Bearer ${token}` } });
    expect(r.status).toBe(403);
  });

  it("rejects DNS-rebinding style Host headers", async () => {
    const token = await pairClient();
    const r = await req({ path: "/v1/history", headers: { host: "evil.test", authorization: `Bearer ${token}` } });
    expect(r.status).toBe(421);
  });

  it("requires JSON content type on writes and limits payload size", async () => {
    const token = await pairClient();
    const auth = { origin: ORIGIN, authorization: `Bearer ${token}` };
    expect((await req({ method: "PUT", path: "/v1/settings", headers: { ...auth, "content-type": "text/plain" }, body: "{}" })).status).toBe(415);
    expect((await req({ method: "PUT", path: "/v1/settings", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ x: "y".repeat(20_000) }) })).status).toBe(413);
  });

  it("pairing codes are single-use and lock after wrong guesses", async () => {
    const { code } = c.pairing.issueCode();
    const post = (body: unknown) => req({ method: "POST", path: "/v1/pair", headers: { origin: ORIGIN, "content-type": "application/json" }, body: JSON.stringify(body) });
    for (let i = 0; i < 4; i++) expect((await post({ code: "AAAA-AAAA" })).status).toBe(401);
    const last = await post({ code: "AAAA-AAAA" });
    expect(last.body.error).toMatch(/Too many/);
    expect((await post({ code })).status).toBe(401); // invalidated
    const fresh = c.pairing.issueCode();
    expect((await post({ code: fresh.code.toLowerCase() })).status).toBe(200);
    expect((await post({ code: fresh.code })).status).toBe(401); // single use
  });

  it("admin endpoints need the admin token, not a client token", async () => {
    const token = await pairClient();
    expect((await req({ method: "POST", path: "/v1/admin/pairing-code", headers: { authorization: `Bearer ${token}` } })).status).toBe(403);
    expect((await req({ method: "POST", path: "/v1/admin/pairing-code", headers: { authorization: `Bearer ${c.adminToken}` } })).status).toBe(200);
    const mode = fs.statSync(path.join(home, "admin.token")).mode & 0o777;
    if (process.platform !== "win32") expect(mode).toBe(0o600);
  });

  it("only forwards an allowlist of Ollama endpoints", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    for (const p of ["/v1/ollama/api/pull", "/v1/ollama/api/delete", "/v1/ollama/api/create", "/v1/ollama/api/generate"])
      expect((await req({ method: "POST", path: p, headers: h, body: "{}" })).status).toBe(404);
    expect((await req({ method: "POST", path: "/v1/models/pull", headers: h, body: JSON.stringify({ model: "llama3.2" }) })).status).toBe(400);
  });

  it("does not serve files outside the add-in directory", async () => {
    expect((await req({ path: "/index.html" })).text).toContain("pane");
    expect((await req({ path: "/../secret.txt" })).status).not.toBe(200);
    expect((await req({ path: "/%2e%2e/secret.txt" })).status).not.toBe(200);
    expect((await req({ path: "/..%2fsecret.txt" })).status).toBe(403);
  });

  it("refuses a remote Ollama URL unless explicitly allowed", async () => {
    await expect(startCompanion(loadConfig({ home, port: 0, ollamaUrl: "https://ollama.example.com" }), { log: () => {} })).rejects.toThrow(/not on this machine/);
  });
});

describe("companion API", () => {
  it("proxies Ollama for an authenticated add-in (provider works through the companion)", async () => {
    const token = await pairClient();
    const p = new OllamaProvider({
      baseUrl: c.url + "/v1/ollama",
      headers: { authorization: `Bearer ${token}` },
      fetchImpl: (u, init) => fetch(u, init),
    });
    expect((await p.listModels())[0]!.name).toBe("llama3.1:8b");
    const deltas: string[] = [];
    const r = await p.chat({ model: "llama3.1:8b", messages: [{ role: "user", content: "hi" }], onDelta: (d) => deltas.push(d) });
    expect(r.message.content).toBe("Hello from local model");
    expect(deltas.length).toBeGreaterThan(1);
  });

  it("stores, lists, and clears history; minimises content", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    const rec = {
      id: "task_1", request: "r", model: "m", host: "word", contextRefs: [], outcome: "completed", finalText: "ok", errors: [], startedAt: Date.now(), finishedAt: Date.now(),
      actions: [{ id: "a", capability: "document.replace_text", args: { newText: "CONFIDENTIAL ".repeat(100) }, target: "", risk: "modify", decision: "approve", status: "succeeded", startedAt: 0, finishedAt: 0 }],
    };
    expect((await req({ method: "POST", path: "/v1/history", headers: h, body: JSON.stringify(rec) })).status).toBe(201);
    const list = await req({ path: "/v1/history", headers: h });
    expect(list.body.records).toHaveLength(1);
    expect(JSON.stringify(list.body).length).toBeLessThan(1500);
    expect((await req({ path: "/v1/history/task_1", headers: h })).status).toBe(200);
    expect((await req({ method: "DELETE", path: "/v1/history", headers: h })).status).toBe(200);
    expect((await req({ path: "/v1/history", headers: h })).body.records).toHaveLength(0);
  });

  it("normalises settings and keeps policy floors", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    const r = await req({ method: "PUT", path: "/v1/settings", headers: h, body: JSON.stringify({ model: "qwen", policy: { send: "allow", prohibited: "allow" } }) });
    expect(r.body.model).toBe("qwen");
    expect(r.body.policy.send).toBe("confirm");
    expect(r.body.policy.prohibited).toBe("deny");
    expect((await req({ path: "/v1/settings", headers: h })).body.model).toBe("qwen");
  });

  it("diagnostics report connectivity without document content", async () => {
    const token = await pairClient();
    const r = await req({ path: "/v1/diagnostics", headers: { origin: ORIGIN, authorization: `Bearer ${token}` } });
    expect(r.body.ollama.reachable).toBe(true);
    expect(r.body.ollamaIsLocal).toBe(true);
    expect(r.body.pairedClients).toBe(1);
  });
});

describe("cross-app sources", () => {
  const source = { app: "word", documentTitle: "Strategy.docx", scopeLabel: "Selection (paragraphs 1–2)", content: "Customer problem: reports take 3 days.", truncated: false, retrievedAt: Date.now() };

  it("shares, lists without content, fetches and deletes; requires pairing", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    expect((await req({ method: "POST", path: "/v1/sources", headers: { origin: ORIGIN, "content-type": "application/json" }, body: JSON.stringify(source) })).status).toBe(401);
    const created = await req({ method: "POST", path: "/v1/sources", headers: h, body: JSON.stringify(source) });
    expect(created.status).toBe(201);
    expect(created.body).toMatchObject({ app: "word", chars: source.content.length });
    expect(created.body.content).toBeUndefined();
    const list = await req({ path: "/v1/sources", headers: h });
    expect(list.body.sources).toHaveLength(1);
    expect(list.text).not.toContain("Customer problem");
    const full = await req({ path: `/v1/sources/${created.body.id}`, headers: h });
    expect(full.body.content).toBe(source.content);
    expect((await req({ method: "DELETE", path: `/v1/sources/${created.body.id}`, headers: h })).body.deleted).toBe(1);
    expect((await req({ path: `/v1/sources/${created.body.id}`, headers: h })).status).toBe(404);
  });

  it("rejects invalid or oversized content and foreign origins", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    expect((await req({ method: "POST", path: "/v1/sources", headers: h, body: JSON.stringify({ ...source, app: "notepad" }) })).status).toBe(400);
    expect((await req({ method: "POST", path: "/v1/sources", headers: h, body: JSON.stringify({ ...source, content: "" }) })).status).toBe(400);
    expect((await req({ path: "/v1/sources", headers: { ...h, origin: "https://evil.test" } })).status).toBe(403);
  });

  it("keeps shared content in memory only", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    await req({ method: "POST", path: "/v1/sources", headers: h, body: JSON.stringify({ ...source, content: "UNIQUE-MARKER-7f3a" }) });
    c.store.db.exec("PRAGMA wal_checkpoint(TRUNCATE)");
    for (const f of fs.readdirSync(home)) {
      const p = path.join(home, f);
      if (fs.statSync(p).isFile()) expect(fs.readFileSync(p).includes("UNIQUE-MARKER-7f3a")).toBe(false);
    }
  });
});

describe("source expiry", () => {
  it("expires shared content after the time limit", async () => {
    let t = Date.now();
    const h2 = fs.mkdtempSync(path.join(os.tmpdir(), "loai-"));
    const comp = await startCompanion(loadConfig({ home: h2, port: 0, ollamaUrl: "http://127.0.0.1:9", tlsCertPath: undefined, tlsKeyPath: undefined }), { log: () => {}, now: () => t });
    const s = comp.sources.add({ app: "excel", scopeLabel: "Sales!A1:B3", content: "a\t1", retrievedAt: t });
    expect(comp.sources.list()).toHaveLength(1);
    t += 31 * 60_000;
    expect(comp.sources.list()).toHaveLength(0);
    expect(comp.sources.get(s.source!.id)).toBeUndefined();
    await comp.close();
    fs.rmSync(h2, { recursive: true, force: true });
  });
});

describe("acceptance records", () => {
  it("stores valid live records as files and lists them; rejects malformed ones", async () => {
    const token = await pairClient();
    const h = { origin: ORIGIN, authorization: `Bearer ${token}`, "content-type": "application/json" };
    const rec = {
      schema: 1, scenarioId: "word.editor.scoped-edit", title: "t", host: "word", mode: "live", model: { kind: "scripted" },
      environment: { officeApp: "Word", officeVersion: "16.0.18129", platform: "PC", addinVersion: "0.3.0", recordedAt: "2026-10-13T12:00:00.000Z" },
      outcome: "passed", checks: [{ id: "a", description: "d", passed: true, evidence: "e" }], visual: "not-applicable", manual: [], durationMs: 10,
    };
    const ok = await req({ method: "POST", path: "/v1/acceptance", headers: h, body: JSON.stringify(rec) });
    expect(ok.status).toBe(201);
    expect(fs.existsSync(path.join(home, "acceptance", ok.body.stored))).toBe(true);
    expect((await req({ method: "POST", path: "/v1/acceptance", headers: h, body: JSON.stringify({ ...rec, scenarioId: "../../etc/passwd" }) })).status).toBe(400);
    expect((await req({ method: "POST", path: "/v1/acceptance", headers: h, body: JSON.stringify({ ...rec, mode: "imagined" }) })).status).toBe(400);
    const list = await req({ path: "/v1/acceptance", headers: h });
    expect(list.body.records).toHaveLength(1);
  });
});
