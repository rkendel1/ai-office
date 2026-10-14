import type { CapabilityRegistry } from "../registry.js";
import type { HostKind } from "../types.js";
import { EXCEL_SPECIALISTS } from "./excel/index.js";
import { OUTLOOK_SPECIALISTS } from "./outlook/index.js";
import { POWERPOINT_SPECIALISTS } from "./powerpoint/index.js";
import type { Specialist } from "./types.js";
import { WORD_SPECIALISTS } from "./word/index.js";

export const ALL_SPECIALISTS: readonly Specialist[] = Object.freeze([...WORD_SPECIALISTS, ...EXCEL_SPECIALISTS, ...POWERPOINT_SPECIALISTS, ...OUTLOOK_SPECIALISTS]);

export interface EffectiveCapabilities {
  /** Allowlisted and registered for the host: what the specialist can actually use. */
  available: string[];
  /** Allowlisted but not available in this Office version. */
  unavailable: string[];
}

export class SpecialistRegistry {
  private readonly byId = new Map<string, Specialist>();

  constructor(list: readonly Specialist[] = ALL_SPECIALISTS) {
    for (const s of list) {
      if (this.byId.has(s.id)) throw new Error(`Duplicate specialist id ${s.id}`);
      if (!s.id.startsWith(`${s.host}.`)) throw new Error(`Specialist id ${s.id} must start with its host "${s.host}."`);
      this.byId.set(s.id, s);
    }
  }

  get(id: string): Specialist | undefined {
    return this.byId.get(id);
  }

  all(): Specialist[] {
    return [...this.byId.values()];
  }

  forHost(host: HostKind): Specialist[] {
    return this.all().filter((s) => s.host === host);
  }

  /**
   * The specialist's allowlist intersected with what the host registered.
   * Internal capabilities are never included, whatever the allowlist says.
   */
  effectiveCapabilities(s: Specialist, caps: CapabilityRegistry): EffectiveCapabilities {
    const registered = new Set(caps.list(s.host).map((d) => d.name));
    return { available: s.capabilities.filter((c) => registered.has(c)), unavailable: s.capabilities.filter((c) => !registered.has(c)) };
  }

  /** Configuration errors (a read-only specialist that could modify content, etc.). */
  validate(s: Specialist, caps: CapabilityRegistry): string[] {
    const problems: string[] = [];
    const { available } = this.effectiveCapabilities(s, caps);
    if (s.constraints.readOnly) {
      const mods = available.filter((c) => caps.get(c)?.modifies);
      if (mods.length) problems.push(`${s.id} is read-only but allows modifying capabilities: ${mods.join(", ")}`);
    }
    if (s.constraints.requirePlan && !s.capabilities.includes("plan.propose")) problems.push(`${s.id} requires a plan but cannot propose one`);
    for (const c of s.capabilities) if (caps.get(c)?.internal) problems.push(`${s.id} allowlists internal capability ${c}`);
    return problems;
  }
}
