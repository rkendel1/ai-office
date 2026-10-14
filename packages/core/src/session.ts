import { ContextManager } from "./context.js";
import type { HistoryStore } from "./history.js";
import type { HostAdapter } from "./hosts/adapter.js";
import type { Approver } from "./permissions.js";
import { CapabilityRegistry } from "./registry.js";
import { AgentRuntime } from "./runtime.js";
import type { AssistantSettings } from "./settings.js";
import type { InferenceProvider } from "./types.js";

export interface Session {
  adapter: HostAdapter;
  registry: CapabilityRegistry;
  context: ContextManager;
  runtime: AgentRuntime;
}

/** Wire an adapter, provider and approver into a ready-to-run assistant session. */
export function createSession(p: {
  adapter: HostAdapter;
  provider: InferenceProvider;
  approver: Approver;
  settings: () => AssistantSettings;
  history?: HistoryStore;
}): Session {
  const registry = new CapabilityRegistry().register(...p.adapter.capabilities());
  const context = new ContextManager();
  const runtime = new AgentRuntime({
    provider: p.provider,
    adapter: p.adapter,
    registry,
    context,
    approver: p.approver,
    settings: p.settings,
    history: p.history,
  });
  return { adapter: p.adapter, registry, context, runtime };
}
