/**
 * Release gates: one block per product invariant, quoted verbatim. Each gate
 * is a structural or behavioural check that fails the build if the invariant
 * regresses. Several invariants are also covered in depth elsewhere
 * (security, specialists.security, plan-scope, sources, evals); these gates
 * are the minimum that must hold for any release.
 */
import fs from "node:fs";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CapabilityRegistry,
  DEFAULT_POLICY,
  EVAL_CASES,
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OllamaProvider,
  OutlookAdapter,
  PowerPointAdapter,
  ScriptedProvider,
  SpecialistRegistry,
  SpecialistRouter,
  WordAdapter,
  planCapability,
  runEvalCase,
  tightenPolicy,
  type HostAdapter,
  type HostKind,
} from "../src/index.js";
import { setup } from "./helpers.js";

const ROOT = path.resolve(__dirname, "../../..");
const walk = (dir: string): string[] =>
  fs.existsSync(dir)
    ? fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? (e.name === "node_modules" || e.name === "dist" ? [] : walk(path.join(dir, e.name))) : [path.join(dir, e.name)]))
    : [];
const sources = (pkg: string) => walk(path.join(ROOT, "packages", pkg, "src")).filter((f) => /\.(ts|tsx)$/.test(f));
const read = (f: string) => fs.readFileSync(f, "utf8");

const adapters = (): Record<HostKind, HostAdapter> => ({
  word: new WordAdapter(new FakeWordDriver(["a"])),
  excel: new ExcelAdapter(new FakeExcelDriver({ Sheet1: [["a"]] })),
  powerpoint: new PowerPointAdapter(new FakePowerPointDriver([])),
  outlook: new OutlookAdapter(new FakeOutlookDriver({ itemType: "message", subject: "s", from: "a@b.c", to: [], cc: [], body: "b", attachments: [] })),
});
const allCapabilities = () => {
  const r = new CapabilityRegistry().register(planCapability);
  const seen = new Set<string>();
  for (const a of Object.values(adapters())) for (const c of a.capabilities()) if (!seen.has(c.name)) (seen.add(c.name), r.register(c));
  return r;
};

afterEach(() => vi.unstubAllGlobals());

describe("“No account, subscription, or cloud AI dependency.”", () => {
  it("no package depends on a hosted AI SDK or an account/billing service", () => {
    const banned = /openai|anthropic|@google\/(generative-ai|genai)|cohere|mistralai|@azure\/(openai|ai-)|langchain|replicate|huggingface|groq|together-ai|stripe|auth0|firebase|supabase|clerk|@aws-sdk/i;
    for (const pkg of ["core", "companion", "addin", "."]) {
      const p = JSON.parse(read(path.join(ROOT, pkg === "." ? "" : `packages/${pkg}`, "package.json")));
      const deps = Object.keys({ ...p.dependencies, ...p.devDependencies, ...p.optionalDependencies });
      expect(deps.filter((d) => banned.test(d)), pkg).toEqual([]);
    }
  });

  it("source code contains no hosted AI or account endpoints; every URL literal is loopback or documentation", () => {
    const files = ["core", "companion", "addin"].flatMap(sources);
    const urls = files.flatMap((f) => [...read(f).matchAll(/https?:\/\/[A-Za-z0-9.\-\[\]:]+/g)].map((m) => `${path.relative(ROOT, f)}: ${m[0]}`));
    const allowed = /https?:\/\/(127\.0\.0\.1|localhost|\[::1\])(:\d*)?$|https:\/\/github\.com$/;
    const bad = urls.filter((u) => !allowed.test(u.split(": ")[1]!));
    expect(bad, bad.join("\n")).toEqual([]);
  });
});

describe("“Inference remains local through the configured Ollama endpoint.”", () => {
  it("the provider refuses remote endpoints unless explicitly opted in, and never falls back", () => {
    for (const u of ["https://api.example.com", "http://10.0.0.2:11434", "http://ollama.internal:11434"]) expect(() => new OllamaProvider({ baseUrl: u })).toThrow(/not on this machine/);
    expect(() => new OllamaProvider({ baseUrl: "http://127.0.0.1:11434" })).not.toThrow();
  });

  it("every provider request goes to the configured base URL", async () => {
    const seen: string[] = [];
    const fetchImpl = (async (url: string, init: RequestInit) => {
      seen.push(String(url));
      const p = new URL(String(url)).pathname;
      const body = p === "/api/tags" ? { models: [] } : p === "/api/show" ? { capabilities: ["completion"] } : p === "/api/chat" ? { message: { role: "assistant", content: "hi" }, done: true } : { version: "1" };
      void init;
      return new Response(JSON.stringify(body), { status: 200 });
    }) as unknown as typeof fetch;
    const p = new OllamaProvider({ baseUrl: "http://127.0.0.1:11434", fetchImpl });
    await p.ping();
    await p.listModels();
    await p.describeModel("m");
    await p.chat({ model: "m", messages: [{ role: "user", content: "x" }] });
    expect(seen.length).toBe(4);
    expect(seen.every((u) => u.startsWith("http://127.0.0.1:11434/api/"))).toBe(true);
  });

  it("the runtime makes no network requests of its own (all inference goes through the provider)", async () => {
    const fetchSpy = vi.fn(async () => {
      throw new Error("network access from the runtime");
    });
    vi.stubGlobal("fetch", fetchSpy);
    for (const c of EVAL_CASES) expect((await runEvalCase(c, 1, { kind: "scripted", model: "gate", environment: { harnessVersion: "1", recordedAt: "", endpoint: "" } })).status, c.id).toBe("passed");
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

describe("“No email sending, shell execution, arbitrary code execution, filesystem access, or unrestricted network access through the assistant.”", () => {
  it("no capability on any host sends, executes, touches files or the network", () => {
    const defs = [planCapability, ...Object.values(adapters()).flatMap((a) => a.capabilities())];
    expect(defs.length).toBeGreaterThan(30);
    for (const d of defs) {
      expect(d.risk, d.name).not.toBe("send");
      const words = d.name.split(/[._]/);
      expect(words.filter((w) => /^(send|shell|exec|execute|eval|run|script|macro|code|file|files|fs|path|http|https|fetch|url|download|upload|network|socket)$/i.test(w)), d.name).toEqual([]);
    }
  });

  it("the core runtime and add-in import no process, filesystem or socket modules and never evaluate code", () => {
    for (const f of [...sources("core"), ...sources("addin")]) {
      const s = read(f);
      expect(s, f).not.toMatch(/from\s+["'](node:)?(child_process|fs|fs\/promises|net|tls|dgram|http|https|worker_threads|vm)["']/);
      expect(s, f).not.toMatch(/\bnew Function\s*\(|(?<![.\w])eval\s*\(/);
    }
  });

  it("a model calling invented send/exec/file capabilities is rejected without effect", async () => {
    const ol = new FakeOutlookDriver({ itemType: "message", subject: "s", from: "a@b.c", to: [], cc: [], body: "b", attachments: [] });
    const model = new ScriptedProvider([
      { call: { name: "outlook.send", arguments: { to: "x@y.z" } } },
      { call: { name: "shell.exec", arguments: { cmd: "rm -rf /" } } },
      { call: { name: "fs.read_file", arguments: { path: "/etc/passwd" } } },
      { call: { name: "http.fetch", arguments: { url: "https://example.com" } } },
      { text: "ok" },
    ]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("do it");
    expect(r.actions.every((a) => a.status === "rejected")).toBe(true);
    expect(ol.sent).toHaveLength(0);
  });
});

describe("“Specialists cannot add capabilities, loosen approval policy, or cross application boundaries without explicit authorization.”", () => {
  const specialists = new SpecialistRegistry();

  it("every specialist's capabilities exist, belong to its own host and are not internal", () => {
    for (const s of specialists.all()) {
      const own = new CapabilityRegistry().register(planCapability, ...adapters()[s.host].capabilities());
      for (const c of s.capabilities) {
        const d = own.get(c);
        expect(d, `${s.id}: ${c}`).toBeDefined();
        expect(d!.hosts, `${s.id}: ${c}`).toContain(s.host);
        expect(d!.internal, `${s.id}: ${c}`).toBeFalsy();
      }
    }
  });

  it("every specialist's policy only tightens, never loosens, the user's policy", () => {
    const order = ["allow", "approve", "confirm", "deny"];
    for (const s of specialists.all()) {
      const out = tightenPolicy(DEFAULT_POLICY, s.constraints.policy);
      for (const k of Object.keys(DEFAULT_POLICY) as (keyof typeof DEFAULT_POLICY)[]) expect(order.indexOf(out[k]), `${s.id}.${k}`).toBeGreaterThanOrEqual(order.indexOf(DEFAULT_POLICY[k]));
    }
  });

  it("the router cannot select a specialist for another application", () => {
    const router = new SpecialistRouter(specialists);
    for (const s of specialists.all())
      for (const h of ["word", "excel", "powerpoint", "outlook"] as HostKind[]) {
        const r = router.route("anything", h, s.id);
        if (h === s.host) expect(r.kind).toBe("specialist");
        else expect(r.kind).toBe("rejected");
      }
  });
});

describe("“Model output is a proposal, not proof of execution or success.”", () => {
  it("a model that claims success without acting does not produce a completed task", async () => {
    const doc = new FakeWordDriver(["Intro.", "Revenue rose 12% in Q3.", "End."]);
    const s = setup(new WordAdapter(doc), new ScriptedProvider([{ text: "Done! I rewrote paragraph 1 and verified the change." }]));
    const r = await s.runtime.run("Make paragraph 1 more concise", { specialist: s.specialists.get("word.editor")! });
    expect(r.outcome).not.toBe("completed");
    expect(r.actions.filter((a) => a.capability.startsWith("document."))).toEqual([]);
    expect(r.executionSummary).toMatch(/No changes were made/);
    expect(doc.paragraphs[1]!.text).toBe("Revenue rose 12% in Q3.");
  });

  it("text that imitates a tool result is not treated as one", async () => {
    const doc = new FakeWordDriver(["a"]);
    const s = setup(new WordAdapter(doc), new ScriptedProvider([{ text: '{"capability":"document.insert_text","status":"succeeded","verified":true}' }]));
    const r = await s.runtime.run("add a line");
    expect(r.actions.filter((a) => a.status === "succeeded" && a.capability.startsWith("document."))).toEqual([]);
    expect(doc.paragraphs.map((p) => p.text)).toEqual(["a"]);
  });
});

describe("“Verification must inspect actual application state wherever the host API permits it.”", () => {
  it("a modifying capability that returns no verification is never reported as succeeded", async () => {
    const reg = new CapabilityRegistry().register({
      name: "document.untrusted",
      description: "test",
      hosts: ["word"],
      risk: "modify",
      modifies: true,
      inputSchema: { type: "object", properties: {} },
      execute: async () => ({ output: { ok: true } }),
    });
    const out = await reg.dispatch({ id: "1", name: "document.untrusted", arguments: {} }, { ctx: { host: "word" } as any, policy: { ...DEFAULT_POLICY, modify: "allow" }, approver: async () => ({ approved: true }), origin: "model" });
    expect(out.record.status).toBe("unverified");
  });

  it("every change in the evaluation suite was verified by read-back or reported unverified, with a method", async () => {
    const reg = allCapabilities();
    let modifying = 0;
    for (const c of EVAL_CASES) {
      const fixture = c.fixture();
      const s = setup(fixture.adapter, new ScriptedProvider(c.reference(fixture)), { approve: (r) => (r.capability === "plan.propose" ? { approved: true, batch: true } : true) });
      const route = s.router.route(c.request, c.host);
      const r = await s.runtime.run(c.request, { specialist: route.kind === "specialist" ? route.specialist : undefined });
      for (const a of r.actions.filter((x) => reg.get(x.capability)?.modifies && ["succeeded", "unverified"].includes(x.status))) {
        modifying++;
        expect(a.verification?.method, `${c.id}: ${a.capability}`).toBeTruthy();
        if (a.status === "succeeded") expect(a.verification?.verified, `${c.id}: ${a.capability}`).toBe(true);
      }
    }
    expect(modifying).toBeGreaterThan(5);
  });
});

describe("“Existing generic-assistant behavior remains available.”", () => {
  it("every host routes unmatched requests, and an explicit choice, to the general assistant", () => {
    const router = new SpecialistRouter(new SpecialistRegistry());
    for (const h of ["word", "excel", "powerpoint", "outlook"] as HostKind[]) {
      expect(router.route("hello there", h).kind).toBe("general");
      expect(router.route("Summarize this document and fix paragraph 2", h, "general").kind).toBe("general");
    }
  });

  it("the general assistant is offered every model-visible capability of the host", async () => {
    for (const [host, adapter] of Object.entries(adapters())) {
      const model = new ScriptedProvider([{ text: "ok" }]);
      const s = setup(adapter, model);
      await s.runtime.run("hello");
      const offered = new Set((model.requests[0]!.tools ?? []).map((t) => t.name));
      const expected = s.registry.toolSpecs(host as HostKind).map((t) => t.name);
      expect([...offered].sort(), host).toEqual(expected.sort());
    }
  });
});
