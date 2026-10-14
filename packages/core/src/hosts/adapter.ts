import type { Freshness } from "../context.js";
import type { CapabilityDefinition } from "../registry.js";
import type { ContextItem, HostKind, ItemIdentity } from "../types.js";

export interface ApiSupportEntry {
  /** e.g. "WordApi 1.3" */
  requirement: string;
  supported: boolean;
  /** Capabilities that need it. */
  capabilities: string[];
}

/**
 * One adapter per Office host. Adapters expose only typed, supported
 * operations via capabilities; they never hand raw Office.js objects to the model.
 */
export interface HostAdapter {
  readonly host: HostKind;
  identity(): Promise<ItemIdentity>;
  /** Capabilities available given the host's supported API sets. */
  capabilities(): CapabilityDefinition<any>[];
  /**
   * Minimal context gathered at the start of a task: the selection (if any)
   * plus lightweight metadata. Never the whole document by default.
   */
  initialContext(): Promise<ContextItem[]>;
  /** Re-check whether a previously inspected item still matches Office state. */
  checkFreshness(item: ContextItem): Promise<Freshness>;
  apiSupport(): ApiSupportEntry[];
}

export interface RequirementChecker {
  isSetSupported(name: string, version: string): boolean;
}

/** Default checker used outside Office (tests): everything supported. */
export const allSupported: RequirementChecker = { isSetSupported: () => true };
