import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { AssistantSettings, TaskState } from "@loai/core";
import { CompanionClient } from "../services/companion";
import { createProvider, type HostInfo } from "../services/host";
import { loadSettings, saveSettings } from "../services/settings";
import { ChatView } from "./ChatView";
import { ActionsView, ContextView, HistoryView } from "./Panels";
import { SettingsView } from "./SettingsView";
import { useAssistant } from "./useAssistant";

type Tab = "chat" | "context" | "actions" | "history" | "settings";
const TABS: { id: Tab; label: string }[] = [
  { id: "chat", label: "Chat" },
  { id: "context", label: "Context" },
  { id: "actions", label: "Actions" },
  { id: "history", label: "History" },
  { id: "settings", label: "Settings" },
];

const IDLE: TaskState[] = ["idle", "completed", "partial", "failed", "cancelled"];
export type Connection = { state: "checking" | "ok" | "error" | "unpaired"; detail?: string };

export function App({ host }: { host: HostInfo }) {
  const [settings, setSettingsState] = useState<AssistantSettings>(loadSettings);
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const [tab, setTab] = useState<Tab>("chat");
  const [conn, setConn] = useState<Connection>({ state: "checking" });
  const [pairVersion, setPairVersion] = useState(0);

  const companion = useMemo(() => new CompanionClient(settings.companionUrl), [settings.companionUrl]);
  const companionRef = useRef(companion);
  companionRef.current = companion;

  const history = useMemo(() => (settings.transport === "companion" && companion.paired ? companion.history() : null), [companion, settings.transport, pairVersion]);

  const getProvider = useCallback(() => createProvider(settingsRef.current, companionRef.current), []);
  const getSettings = useCallback(() => settingsRef.current, []);
  const a = useAssistant(host, getProvider, getSettings, history);
  const busy = !IDLE.includes(a.state);

  const setSettings = useCallback((s: AssistantSettings) => setSettingsState(saveSettings(s)), []);

  // Connection status is shown without waiting for a generation request.
  const checkConnection = useCallback(async () => {
    const s = settingsRef.current;
    setConn({ state: "checking" });
    if (s.transport === "companion" && !companionRef.current.paired) {
      try {
        await companionRef.current.health();
        setConn({ state: "unpaired", detail: "Companion running — pair it in Settings" });
      } catch (e) {
        setConn({ state: "error", detail: (e as Error).message });
      }
      return;
    }
    try {
      const r = await getProvider().ping();
      setConn(r.ok ? { state: "ok", detail: `Ollama ${r.version ?? ""}`.trim() } : { state: "error", detail: r.error });
    } catch (e) {
      setConn({ state: "error", detail: (e as Error).message });
    }
  }, [getProvider]);

  useEffect(() => {
    checkConnection();
    const t = setInterval(checkConnection, 30_000);
    return () => clearInterval(t);
  }, [checkConnection, settings.transport, settings.ollamaEndpoint, settings.companionUrl, pairVersion]);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="logo" aria-hidden>
            ◆
          </span>
          Local Office AI
          {host.demo && <span className="pill">{host.label}</span>}
        </div>
        <button className="status" onClick={() => setTab("settings")} title={conn.detail}>
          <span className={`dot dot-${conn.state}`} />
          <span className="status-text">{settings.model || "No model"}</span>
        </button>
      </header>
      <nav className="tabs" role="tablist">
        {TABS.map((t) => (
          <button key={t.id} role="tab" aria-selected={tab === t.id} className={tab === t.id ? "tab active" : "tab"} onClick={() => setTab(t.id)}>
            {t.label}
            {t.id === "actions" && a.pending && <span className="tab-dot" />}
          </button>
        ))}
      </nav>
      <main className="content">
        {tab === "chat" && <ChatView a={a} busy={busy} settings={settings} conn={conn} onOpenSettings={() => setTab("settings")} />}
        {tab === "context" && <ContextView a={a} host={host} />}
        {tab === "actions" && <ActionsView a={a} />}
        {tab === "history" && <HistoryView history={history} settings={settings} />}
        {tab === "settings" && (
          <SettingsView
            settings={settings}
            setSettings={setSettings}
            companion={companion}
            conn={conn}
            onPairChanged={() => setPairVersion((v) => v + 1)}
            recheck={checkConnection}
            getProvider={getProvider}
            host={host}
          />
        )}
      </main>
    </div>
  );
}
