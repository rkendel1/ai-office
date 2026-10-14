import { ContextManager } from "./context.js";
import type { HistoryStore } from "./history.js";
import type { HostAdapter } from "./hosts/adapter.js";
import type { Approver } from "./permissions.js";
import { CapabilityRegistry } from "./registry.js";
import { AgentRuntime } from "./runtime.js";
import type { AssistantSettings } from "./settings.js";
import type { InferenceProvider } from "./types.js";
import { SpecialistRegistry } from "./specialists/registry.js";
import { SpecialistRouter } from "./specialists/router.js";
import { planCapability } from "./specialists/shared/planning.js";
import type { Specialist } from "./specialists/types.js";

export interface Session {
  adapter: HostAdapter;
  registry: CapabilityRegistry;
  context: ContextManager;
  runtime: AgentRuntime;
  specialists: SpecialistRegistry;
  router: SpecialistRouter;
}

/** Wire an adapter, provider and approver into a ready-to-run assistant session. */
export function createSession(p: {
  adapter: HostAdapter;
  provider: InferenceProvider;
  approver: Approver;
  settings: () => AssistantSettings;
  history?: HistoryStore;
  specialists?: readonly Specialist[];
}): Session {
  const registry = new CapabilityRegistry().register(...p.adapter.capabilities(), planCapability);
  const specialists = new SpecialistRegistry(p.specialists);
  for (const s of specialists.forHost(p.adapter.host)) {
    const problems = specialists.validate(s, registry);
    if (problems.length) throw new Error(`Invalid specialist configuration: ${problems.join("; ")}`);
  }
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
  return { adapter: p.adapter, registry, context, runtime, specialists, router: new SpecialistRouter(specialists) };
}
