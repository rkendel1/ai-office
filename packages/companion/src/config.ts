import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { LocalProviderKind } from "@loai/core";

export interface CompanionConfig {
  /** Directory for config, token, database. */
  home: string;
  /** Interface to bind. Loopback only unless explicitly overridden. */
  bindHost: string;
  port: number;
  /** Local model app base URL (Atomic Chat by default, or Ollama). */
  llmUrl: string;
  /** API language of the local model app. */
  provider: LocalProviderKind;
  /** Allow a non-loopback model URL. Off by default. */
  allowRemoteLlm: boolean;
  /** @deprecated Use llmUrl. Kept so older config files and callers still work. */
  ollamaUrl?: string;
  /** @deprecated Use allowRemoteLlm. */
  allowRemoteOllama?: boolean;
  /** Browser origins allowed to call the API (the add-in's origin). */
  allowedOrigins: string[];
  tlsCertPath?: string;
  tlsKeyPath?: string;
  /** Built add-in to serve at "/". */
  staticDir?: string;
  maxBodyBytes: number;
  dbPath: string;
}

export const DEFAULT_PORT = 47615;

export function defaultHome(): string {
  if (process.env.LOAI_HOME) return process.env.LOAI_HOME;
  if (process.platform === "win32") return path.join(process.env.APPDATA ?? path.join(os.homedir(), "AppData", "Roaming"), "LocalOfficeAI");
  if (process.platform === "darwin") return path.join(os.homedir(), "Library", "Application Support", "LocalOfficeAI");
  return path.join(process.env.XDG_CONFIG_HOME ?? path.join(os.homedir(), ".config"), "local-office-ai");
}

const LOOPBACK_BIND = new Set(["127.0.0.1", "::1", "localhost"]);

export function loadConfig(overrides: Partial<CompanionConfig> = {}): CompanionConfig {
  const home = overrides.home ?? defaultHome();
  fs.mkdirSync(home, { recursive: true, mode: 0o700 });
  const file = path.join(home, "config.json");
  let fromFile: Partial<CompanionConfig> = {};
  if (fs.existsSync(file)) {
    try {
      fromFile = JSON.parse(fs.readFileSync(file, "utf8"));
    } catch (e) {
      throw new Error(`Invalid configuration file ${file}: ${(e as Error).message}`);
    }
  }
  const devCerts = path.join(os.homedir(), ".office-addin-dev-certs");
  const port = Number(overrides.port ?? fromFile.port ?? process.env.LOAI_PORT ?? DEFAULT_PORT);

  // `ollamaUrl` is the pre-Atomic-Chat name for the same setting. An explicit
  // legacy value keeps the Ollama protocol; new installs default to Atomic Chat.
  const legacyUrl = overrides.ollamaUrl ?? fromFile.ollamaUrl;
  const envOllamaUrl = process.env.OLLAMA_HOST_URL;
  const llmUrl =
    overrides.llmUrl ??
    legacyUrl ??
    fromFile.llmUrl ??
    process.env.LOAI_LLM_URL ??
    process.env.ATOMIC_CHAT_URL ??
    envOllamaUrl ??
    "http://127.0.0.1:1337/v1";
  const rawProvider = overrides.provider ?? fromFile.provider ?? process.env.LOAI_LLM_PROVIDER;
  const provider: LocalProviderKind =
    rawProvider === "ollama" || rawProvider === "openai"
      ? rawProvider
      : legacyUrl || (!fromFile.llmUrl && envOllamaUrl)
        ? "ollama"
        : "openai";

  const cfg: CompanionConfig = {
    home,
    bindHost: overrides.bindHost ?? fromFile.bindHost ?? "127.0.0.1",
    port,
    llmUrl,
    provider,
    allowRemoteLlm: overrides.allowRemoteLlm ?? overrides.allowRemoteOllama ?? fromFile.allowRemoteLlm ?? fromFile.allowRemoteOllama ?? false,
    allowedOrigins: overrides.allowedOrigins ??
      fromFile.allowedOrigins ?? [`https://localhost:${port}`, "https://localhost:3000"],
    tlsCertPath: "tlsCertPath" in overrides ? overrides.tlsCertPath : (fromFile.tlsCertPath ?? existing(path.join(devCerts, "localhost.crt"))),
    tlsKeyPath: "tlsKeyPath" in overrides ? overrides.tlsKeyPath : (fromFile.tlsKeyPath ?? existing(path.join(devCerts, "localhost.key"))),
    staticDir: overrides.staticDir ?? fromFile.staticDir,
    maxBodyBytes: overrides.maxBodyBytes ?? fromFile.maxBodyBytes ?? 4 * 1024 * 1024,
    dbPath: overrides.dbPath ?? fromFile.dbPath ?? path.join(home, "companion.db"),
  };
  if (!LOOPBACK_BIND.has(cfg.bindHost) && process.env.LOAI_ALLOW_NONLOOPBACK_BIND !== "1") {
    throw new Error(
      `Refusing to bind to ${cfg.bindHost}. The companion only listens on loopback. Set LOAI_ALLOW_NONLOOPBACK_BIND=1 only for a documented, firewalled use case.`,
    );
  }
  if (!Number.isInteger(cfg.port) || cfg.port < 0 || cfg.port > 65535) throw new Error(`Invalid port ${cfg.port}`);
  return cfg;
}

function existing(p: string): string | undefined {
  return fs.existsSync(p) ? p : undefined;
}
