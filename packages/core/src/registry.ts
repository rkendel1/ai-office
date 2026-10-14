import { decide, type ApprovalPolicy, type Approver } from "./permissions.js";
import { validate } from "./schema.js";
import type {
  ActionPreview,
  ContextItem,
  ExecutionRecord,
  HandlerResult,
  HostKind,
  ItemIdentity,
  Json,
  JsonSchema,
  RiskClass,
  ToolCall,
  ToolSpec,
} from "./types.js";
import { errorMessage, newId } from "./util.js";

export interface ExecutionContext {
  host: HostKind;
  identity: ItemIdentity;
  signal?: AbortSignal;
}

/** Thrown by `resolveTarget` when arguments point outside the permitted scope. */
export class ScopeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ScopeError";
  }
}

/** Thrown by handlers when the host does not support an operation. */
export class CapabilityUnsupportedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CapabilityUnsupportedError";
  }
}

export interface CapabilityDefinition<A extends Record<string, Json> = Record<string, Json>> {
  /** Internal identifier, e.g. "document.replace_text". */
  name: string;
  /** Description given to the model. */
  description: string;
  hosts: HostKind[];
  /** Base risk; may be escalated by `assessRisk`. */
  risk: RiskClass;
  /** Whether the capability changes Office state. Modifying actions must verify. */
  modifies: boolean;
  inputSchema: JsonSchema;
  /** Hidden from the model (e.g. compensation-only operations). */
  internal?: boolean;
  /**
   * Always ask the user, even if policy would allow (e.g. plan approval).
   * Can only make a decision stricter.
   */
  alwaysAsk?: boolean;
  /** Validates scope and returns a user-facing target description. Throw ScopeError to reject. */
  resolveTarget?(args: A, ctx: ExecutionContext): Promise<string> | string;
  /** May escalate risk based on actual Office state (e.g. overwriting substantial content). */
  assessRisk?(args: A, ctx: ExecutionContext): Promise<RiskClass> | RiskClass;
  /** Describe the change before approval. Must not modify Office state. */
  preview?(args: A, ctx: ExecutionContext): Promise<ActionPreview>;
  execute(args: A, ctx: ExecutionContext): Promise<HandlerResult>;
}

export interface DispatchResult {
  record: ExecutionRecord;
  /** Context items the capability inspected (registered by the runtime). */
  context: ContextItem[];
}

export interface DispatchOptions {
  ctx: ExecutionContext;
  policy: ApprovalPolicy;
  approver: Approver;
  /** "model" calls may not reach internal capabilities. */
  origin: "model" | "user";
  /**
   * Optional narrowing (e.g. the active specialist's allowlist). A capability
   * outside it is rejected. It can never add capabilities.
   */
  allowlist?: ReadonlySet<string>;
  /** Label for the narrowing source, used in rejection messages. */
  allowlistLabel?: string;
}

const RISK_ORDER: RiskClass[] = [
  "read_selection",
  "read_active",
  "read_external",
  "modify",
  "destructive",
  "send",
  "prohibited",
];

function maxRisk(a: RiskClass, b: RiskClass): RiskClass {
  return RISK_ORDER.indexOf(a) >= RISK_ORDER.indexOf(b) ? a : b;
}

/**
 * Host-controlled registry. The only path from a model proposal to Office
 * state: unknown / invalid / out-of-scope / unauthorized calls are rejected.
 */
export class CapabilityRegistry {
  private readonly defs = new Map<string, CapabilityDefinition>();

  register(...defs: CapabilityDefinition<any>[]): this {
    for (const d of defs) {
      if (this.defs.has(d.name)) throw new Error(`Capability already registered: ${d.name}`);
      if (d.inputSchema.type !== "object") throw new Error(`Capability ${d.name} must take an object`);
      this.defs.set(d.name, d as CapabilityDefinition);
    }
    return this;
  }

  get(name: string): CapabilityDefinition | undefined {
    return this.defs.get(name);
  }

  list(host?: HostKind, includeInternal = false): CapabilityDefinition[] {
    return [...this.defs.values()].filter(
      (d) => (!host || d.hosts.includes(host)) && (includeInternal || !d.internal),
    );
  }

  toolSpecs(host: HostKind, allowlist?: ReadonlySet<string>): ToolSpec[] {
    return this.list(host)
      .filter((d) => !allowlist || allowlist.has(d.name))
      .map((d) => ({ name: d.name, description: d.description, parameters: d.inputSchema }));
  }

  async dispatch(call: ToolCall, opts: DispatchOptions): Promise<DispatchResult> {
    const out = await this.dispatchInner(call, opts);
    return out;
  }

  private async dispatchInner(call: ToolCall, opts: DispatchOptions): Promise<DispatchResult> {
    const startedAt = Date.now();
    const base = {
      id: newId("act"),
      capability: call.name,
      args: call.arguments ?? {},
      target: "",
      risk: "prohibited" as RiskClass,
      decision: "deny" as const,
      startedAt,
    };
    const reject = (error: string, extra: Partial<ExecutionRecord> = {}): DispatchResult => ({
      record: { ...base, ...extra, status: "rejected", error, finishedAt: Date.now() },
      context: [],
    });
    const done = (record: ExecutionRecord, context: ContextItem[] = []): DispatchResult => ({ record, context });

    const def = this.defs.get(call.name);
    if (!def) return reject(`Unknown capability "${call.name}"`);
    if (def.internal && opts.origin === "model") return reject(`Capability "${call.name}" is not available to the assistant`);
    if (!def.hosts.includes(opts.ctx.host)) return reject(`Capability "${call.name}" is not supported in ${opts.ctx.host}`);
    if (opts.allowlist && !opts.allowlist.has(def.name))
      return reject(`Capability "${call.name}" is not available to ${opts.allowlistLabel ?? "the active specialist"}`);
    if (typeof call.arguments !== "object" || call.arguments === null || Array.isArray(call.arguments))
      return reject("Arguments must be an object");

    const issues = validate(def.inputSchema, call.arguments);
    if (issues.length) return reject("Invalid arguments: " + issues.map((i) => `${i.path} ${i.message}`).join("; "));

    let target = def.name;
    try {
      if (def.resolveTarget) target = await def.resolveTarget(call.arguments, opts.ctx);
    } catch (e) {
      return reject(e instanceof ScopeError ? `Out of scope: ${e.message}` : `Could not resolve target: ${errorMessage(e)}`);
    }

    let risk = def.risk;
    try {
      if (def.assessRisk) risk = maxRisk(risk, await def.assessRisk(call.arguments, opts.ctx));
    } catch {
      // If risk can't be assessed, assume the worst non-prohibited class for modifying actions.
      if (def.modifies) risk = maxRisk(risk, "destructive");
    }
    let decision = decide(opts.policy, risk);
    if (def.alwaysAsk && decision === "allow") decision = "approve";
    const withMeta = { target, risk, decision };

    if (decision === "deny") return reject(`Not permitted by policy (${risk})`, withMeta);

    let approval: ExecutionRecord["approval"] = "not_required";
    if (decision === "approve" || decision === "confirm") {
      let preview: ActionPreview;
      try {
        preview = def.preview ? await def.preview(call.arguments, opts.ctx) : { summary: def.name, target };
      } catch (e) {
        return reject(`Could not prepare preview: ${errorMessage(e)}`, withMeta);
      }
      if (opts.ctx.signal?.aborted) return done({ ...base, ...withMeta, status: "cancelled", finishedAt: Date.now() });
      const resp = await opts.approver({ capability: def.name, risk, decision, preview, args: call.arguments }, opts.ctx.signal);
      if (!resp.approved) {
        return reject(resp.reason ? `Rejected by user: ${resp.reason}` : "Rejected by user", {
          ...withMeta,
          approval: "rejected",
        });
      }
      approval = resp.via === "plan" ? "plan" : "approved";
    }

    if (opts.ctx.signal?.aborted)
      return done({ ...base, ...withMeta, approval, status: "cancelled", finishedAt: Date.now() });

    let result: HandlerResult;
    try {
      result = await def.execute(call.arguments, opts.ctx);
    } catch (e) {
      return done({
        ...base,
        ...withMeta,
        approval,
        status: "failed",
        error: errorMessage(e),
        finishedAt: Date.now(),
      });
    }

    let status: ExecutionRecord["status"];
    if (result.partial) status = "partial";
    else if (!def.modifies) status = "succeeded";
    else if (result.verification?.verified === true) status = "succeeded";
    else if (result.verification?.verified === false) status = "failed";
    else status = "unverified";

    return done(
      {
      ...base,
      ...withMeta,
      approval,
      status,
      output: result.output,
      verification: result.verification,
      compensation: result.compensation,
      error:
        status === "failed"
          ? result.verification?.mismatch ?? result.message ?? "Verification failed"
          : undefined,
      finishedAt: Date.now(),
      },
      result.context ?? [],
    );
  }
}
