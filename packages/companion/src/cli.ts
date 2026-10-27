#!/usr/bin/env node
import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { loadAdminToken } from "./auth.js";
import { loadConfig } from "./config.js";
import { CompanionStore } from "./store.js";
import { VERSION, readAcceptanceRecords, startCompanion, type Companion } from "./server.js";
import { ACCEPTANCE_SCENARIOS, acceptanceMatrix } from "@loai/core";
import { evalMatrix, readEvalRuns, runModelEvaluation } from "./evals.js";

const HELP = `Local Office AI companion ${VERSION}

Usage: loai-companion <command> [options]

Commands:
  start [--dev]        Start the companion (serves the add-in and the local API)
  pair                 Print a one-time pairing code for the Office add-in
  clients              List paired add-in clients
  revoke               Revoke all paired clients
  clear-history        Delete all local task history
  acceptance [--out f] Print the live/mocked acceptance status matrix from saved records
  eval --model <name>  Evaluate local Ollama model(s) on the evaluation suite (local endpoint only)
       [--cases ids|categories] [--repeat n] [--timeout seconds] [--out dir] [--gpu "description"]
       [--self-test]        Run the harness against its reference scripts (no model needed)
  eval-matrix [--dir d] [--out f]  Print the model compatibility matrix from saved run files
  config               Show the effective configuration and file locations

Options:
  --port <n>           Port (default 47615)
  --llm <url>          Local model base URL (default http://127.0.0.1:1337/v1)
  --provider <kind>    openai (Atomic Chat) or ollama (default openai)
  --ollama <url>       Shorthand for --llm <url> --provider ollama (default Ollama URL)
  --static <dir>       Directory of the built add-in to serve
`;

function arg(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

/** The dev-certs tool writes its CA next to the leaf certificate; trust it for CLI calls. */
function caFor(certPath: string): Buffer {
  const ca = path.join(path.dirname(certPath), "ca.crt");
  return fs.readFileSync(fs.existsSync(ca) ? ca : certPath);
}

function admin(pathname: string, method = "GET"): Promise<any> {
  const cfg = loadConfig(arg("--port") ? { port: Number(arg("--port")) } : {});
  const token = loadAdminToken(cfg.home);
  const tls = !!(cfg.tlsCertPath && cfg.tlsKeyPath);
  const mod = tls ? https : http;
  return new Promise((resolve) => {
    const req = mod.request(
      {
        host: "127.0.0.1",
        port: cfg.port,
        path: pathname,
        method,
        headers: { authorization: `Bearer ${token}` },
        ...(tls ? { ca: [caFor(cfg.tlsCertPath!), fs.readFileSync(cfg.tlsCertPath!)], servername: "localhost" } : {}),
      },
      (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => {
          const body = data ? JSON.parse(data) : {};
          if ((res.statusCode ?? 500) >= 400) {
            console.error(body.error ?? res.statusCode);
            process.exit(1);
          }
          resolve(body);
        });
      },
    );
    req.on("error", (e) => {
      console.error(`Cannot reach the companion on port ${cfg.port}. Is it running? (${e.message})`);
      process.exit(1);
    });
    req.end();
  });
}

/**
 * Identify a companion already listening on this port. Its health endpoint is
 * unauthenticated, which makes it safe to probe before deciding how to report
 * an EADDRINUSE error.
 */
function probeRunningCompanion(
  port: number,
  cfg: { tlsCertPath?: string; tlsKeyPath?: string },
): Promise<{ url: string; version?: string } | undefined> {
  const tls = !!(cfg.tlsCertPath && cfg.tlsKeyPath && fs.existsSync(cfg.tlsCertPath) && fs.existsSync(cfg.tlsKeyPath));
  const mod = tls ? https : http;
  const scheme = tls ? "https" : "http";
  return new Promise((resolve) => {
    const req = mod.get(
      {
        host: "127.0.0.1",
        port,
        path: "/v1/health",
        timeout: 1500,
        ...(tls
          ? { ca: [caFor(cfg.tlsCertPath!), fs.readFileSync(cfg.tlsCertPath!)], servername: "localhost" }
          : {}),
      },
      (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => {
          try {
            const body = JSON.parse(data);
            if (body?.ok === true && body?.service === "local-office-ai-companion")
              resolve({ url: `${scheme}://localhost:${port}`, version: typeof body.version === "string" ? body.version : undefined });
            else resolve(undefined);
          } catch {
            resolve(undefined);
          }
        });
      },
    );
    req.on("timeout", () => req.destroy());
    req.on("error", () => resolve(undefined));
  });
}

async function main() {
  const cmd = process.argv[2] ?? "help";
  const here = path.dirname(fileURLToPath(import.meta.url));
  switch (cmd) {
    case "start": {
      const defaultStatic = path.resolve(here, "../../addin/dist");
      const cfg = loadConfig({
        ...(arg("--port") ? { port: Number(arg("--port")) } : {}),
        ...(arg("--llm") ? { llmUrl: arg("--llm") } : {}),
        ...(arg("--provider") ? { provider: arg("--provider") as "openai" | "ollama" } : {}),
        ...(arg("--ollama") ? { llmUrl: arg("--ollama"), provider: "ollama" as const } : {}),
        staticDir: arg("--static") ?? (fs.existsSync(defaultStatic) ? defaultStatic : undefined),
      });
      let c: Companion;
      try {
        c = await startCompanion(cfg);
      } catch (e) {
        if ((e as NodeJS.ErrnoException).code !== "EADDRINUSE") throw e;
        const altPort = cfg.port === 65535 ? 47616 : cfg.port + 1;
        const running = await probeRunningCompanion(cfg.port, cfg);
        if (running) {
          console.log(`A Local Office AI companion${running.version ? ` ${running.version}` : ""} is already running at ${running.url}.`);
          console.log("The installed app starts its companion at login, so this development build was not started.");
          if (process.platform === "darwin") {
            console.log("Stop the installed companion for this login session, then run this command again:");
            console.log("  launchctl bootout gui/$(id -u)/ai.localoffice.companion");
          } else {
            console.log("Stop the installed companion, then run this command again.");
          }
          console.log(`Or start this build on another port: npm start -- --port ${altPort}`);
          break;
        }
        console.error(`Port ${cfg.port} is already in use by another process.`);
        console.error(`Stop that process, or start this build on another port: npm start -- --port ${altPort}`);
        process.exitCode = 1;
        break;
      }
      console.log(`Local Office AI companion ${VERSION}`);
      console.log(`  Listening:  ${c.url} (loopback only)`);
      console.log(`  Model app:  ${cfg.provider === "ollama" ? "Ollama" : "Atomic Chat (OpenAI-compatible)"} at ${cfg.llmUrl}`);
      console.log(`  Add-in:     ${cfg.staticDir ? cfg.staticDir : "(not served; use the dev server)"}`);
      console.log(`  Data:       ${cfg.home}`);
      if (!c.url.startsWith("https")) console.log("  WARNING: no TLS certificate found. Office on macOS requires HTTPS. Run: npx office-addin-dev-certs install");
      if (c.store.listClients().length === 0) {
        const p = c.pairing.issueCode();
        console.log(`\n  Pairing code: ${p.code}   (enter it in the add-in's Settings; valid 10 minutes)\n`);
      }
      const stop = async () => {
        await c.close();
        process.exit(0);
      };
      process.on("SIGINT", stop);
      process.on("SIGTERM", stop);
      break;
    }
    case "pair": {
      const r = await admin("/v1/admin/pairing-code", "POST");
      console.log(`Pairing code: ${r.code} (valid until ${new Date(r.expiresAt).toLocaleTimeString()})`);
      break;
    }
    case "clients": {
      const r = await admin("/v1/admin/clients");
      for (const c of r.clients) console.log(`${c.id}  ${c.name}  paired ${new Date(c.created_at).toLocaleString()}  last seen ${c.last_seen ? new Date(c.last_seen).toLocaleString() : "never"}`);
      if (!r.clients.length) console.log("No paired clients.");
      break;
    }
    case "revoke": {
      const r = await admin("/v1/admin/clients", "DELETE");
      console.log(`Revoked ${r.revoked} client(s).`);
      break;
    }
    case "clear-history": {
      const cfg = loadConfig();
      const s = new CompanionStore(cfg.dbPath);
      await s.clear();
      s.close();
      console.log("Local history cleared.");
      break;
    }
    case "acceptance": {
      const cfg = loadConfig();
      const md = acceptanceMatrix(readAcceptanceRecords(cfg.home), ACCEPTANCE_SCENARIOS);
      const out = arg("--out");
      if (out) {
        fs.writeFileSync(out, md + "\n");
        console.log(`Wrote ${out}`);
      } else console.log(md);
      break;
    }
    case "eval": {
      const cfg = loadConfig();
      const endpoint = arg("--ollama") ?? (cfg.provider === "ollama" ? cfg.llmUrl : "http://127.0.0.1:11434");
      const list = (v?: string) => (v ? v.split(",").map((x) => x.trim()).filter(Boolean) : []);
      const { runs, files } = await runModelEvaluation({
        endpoint,
        models: list(arg("--model")),
        select: list(arg("--cases")),
        repetitions: arg("--repeat") ? Number(arg("--repeat")) : 1,
        timeoutMs: arg("--timeout") ? Number(arg("--timeout")) * 1000 : undefined,
        outDir: arg("--out") ?? path.join(cfg.home, "evals"),
        gpu: arg("--gpu"),
        selfTest: process.argv.includes("--self-test"),
        log: (l) => console.log(l),
      });
      for (const f of files) console.log(`Wrote ${f}`);
      if (runs.some((r) => r.summary.errors === r.summary.runs)) process.exitCode = 2;
      break;
    }
    case "eval-matrix": {
      const dir = arg("--dir") ?? "evals/results";
      const { invalid } = readEvalRuns(dir);
      for (const i of invalid) console.error(`Skipped invalid run file ${i}`);
      const md = evalMatrix(dir);
      const out = arg("--out");
      if (out) {
        fs.writeFileSync(out, md + "\n");
        console.log(`Wrote ${out}`);
      } else console.log(md);
      break;
    }
    case "config": {
      console.log(JSON.stringify(loadConfig(), null, 2));
      break;
    }
    default:
      console.log(HELP);
  }
}

main().catch((e) => {
  console.error(e.message ?? e);
  process.exit(1);
});
