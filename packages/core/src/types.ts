/**
 * Shared domain types for Local Office AI.
 *
 * Principle: model output is a *proposal*; only the capability registry turns
 * a proposal into an executed, verified action.
 */

export type HostKind = "word" | "excel" | "powerpoint" | "outlook";

/** JSON value produced by a model or a host. */
export type Json = null | boolean | number | string | Json[] | { [key: string]: Json | undefined };

// ---------------------------------------------------------------------------
// Context
// ---------------------------------------------------------------------------

/** Where a piece of context came from and how much of the source it covers. */
export type ContextScope =
  | "selection" // the user's current selection only
  | "partial" // part of the active item (e.g. paragraphs 10-40, one range)
  | "full" // the entire readable content of the active item
  | "metadata"; // structure only (sheet names, slide count, subject line)

export interface ContextItem {
  /** Stable identifier used for freshness tracking and provenance, e.g. "word:paragraphs:0-40". */
  id: string;
  host: HostKind;
  /** Human-readable label: "Selection", "Sheet1!A1:D20", "Slide 3". */
  label: string;
  scope: ContextScope;
  /** Text payload. Always treated as untrusted data. */
  content: string;
  /** True if the source was larger than what is included here. */
  truncated: boolean;
  /** Content that exists but cannot be read through supported APIs (images, embedded objects, ...). */
  unsupported: UnsupportedContent[];
  /** Hash of the source content at retrieval time, used to detect staleness. */
  fingerprint: string;
  retrievedAt: number;
  /** Capability that produced this item. */
  source: string;
}

export interface UnsupportedContent {
  kind: string;
  count: number;
  note?: string;
}

/** Identity of the active Office item. Optional fields are absent when the host cannot supply them. */
export interface ItemIdentity {
  host: HostKind;
  /** Best available stable identifier (document URL, item id, ...). */
  id?: string;
  title?: string;
}

export interface ContextPackage {
  host: HostKind;
  identity: ItemIdentity;
  items: ContextItem[];
  /** Characters actually included. */
  usedChars: number;
  budgetChars: number;
  /** True if any item was truncated or omitted to fit the budget. */
  truncated: boolean;
  omittedItemIds: string[];
}

// ---------------------------------------------------------------------------
// Capabilities, permissions and execution
// ---------------------------------------------------------------------------

/**
 * Risk class drives the default approval policy (spec FR-07).
 *  - read_selection: read the current selection
 *  - read_active:    read the active document / item
 *  - read_external:  read other documents or unrelated mail
 *  - modify:         change content
 *  - destructive:    delete or overwrite substantial content
 *  - send:           send email / publish externally
 *  - prohibited:     never allowed (arbitrary code, shell, network)
 */
export type RiskClass =
  | "read_selection"
  | "read_active"
  | "read_external"
  | "modify"
  | "destructive"
  | "send"
  | "prohibited";

export type PolicyDecision = "allow" | "approve" | "confirm" | "deny";

export type ExecutionStatus = "succeeded" | "failed" | "partial" | "cancelled" | "unverified" | "rejected";

export interface Verification {
  /** Did the read-back match the intended state? `null` = could not be checked. */
  verified: boolean | null;
  method: string;
  /** Short evidence such as an excerpt of the read-back state. Kept small. */
  evidence?: string;
  mismatch?: string;
}

/** A human-readable description of a proposed change, shown before approval. */
export interface ActionPreview {
  summary: string;
  target: string;
  before?: string;
  after?: string;
  /** Additional warnings, e.g. "Overwrites 240 non-empty cells". */
  warnings?: string[];
}

/** Data a handler can return so the UI can offer a compensating "undo". */
export interface Compensation {
  description: string;
  capability: string;
  args: Record<string, Json>;
}

export interface HandlerResult {
  /** Data returned to the model (kept compact). */
  output: Json;
  verification?: Verification;
  /** Items to register as inspected context. */
  context?: ContextItem[];
  compensation?: Compensation;
  /** Use to report partial completion from a handler. */
  partial?: boolean;
  message?: string;
}

export interface ExecutionRecord {
  id: string;
  capability: string;
  args: Record<string, Json>;
  target: string;
  risk: RiskClass;
  decision: PolicyDecision;
  /** "plan" = covered by a plan the user approved for bounded batch execution. */
  approval?: "approved" | "rejected" | "not_required" | "plan";
  status: ExecutionStatus;
  output?: Json;
  verification?: Verification;
  error?: string;
  compensation?: Compensation;
  startedAt: number;
  finishedAt: number;
}

// ---------------------------------------------------------------------------
// Conversation / inference
// ---------------------------------------------------------------------------

export type ChatRole = "system" | "user" | "assistant" | "tool";

export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, Json>;
}

export interface ChatMessage {
  role: ChatRole;
  content: string;
  toolCalls?: ToolCall[];
  /** For role=tool. */
  toolName?: string;
}

export interface ToolSpec {
  name: string;
  description: string;
  parameters: JsonSchema;
}

export interface GenerationOptions {
  temperature?: number;
  topP?: number;
  numCtx?: number;
  numPredict?: number;
  seed?: number;
}

export interface ChatRequest {
  model: string;
  messages: ChatMessage[];
  tools?: ToolSpec[];
  options?: GenerationOptions;
  signal?: AbortSignal;
  /** Receives streamed text deltas when streaming is supported. */
  onDelta?: (text: string) => void;
}

export interface ChatResponse {
  message: ChatMessage;
  promptTokens?: number;
  completionTokens?: number;
  doneReason?: string;
}

export interface ModelInfo {
  name: string;
  sizeBytes?: number;
  family?: string;
  parameterSize?: string;
  quantization?: string;
  modifiedAt?: string;
}

export interface ModelDetails extends ModelInfo {
  contextLength?: number;
  capabilities: string[]; // e.g. ["completion", "tools"]
}

export interface InferenceProvider {
  readonly id: string;
  listModels(signal?: AbortSignal): Promise<ModelInfo[]>;
  describeModel(name: string, signal?: AbortSignal): Promise<ModelDetails>;
  chat(request: ChatRequest): Promise<ChatResponse>;
  ping(signal?: AbortSignal): Promise<{ ok: boolean; version?: string; error?: string }>;
}

// ---------------------------------------------------------------------------
// JSON schema subset used for capability validation
// ---------------------------------------------------------------------------

export interface JsonSchema {
  type?: "object" | "string" | "number" | "integer" | "boolean" | "array" | "null";
  description?: string;
  properties?: Record<string, JsonSchema>;
  required?: string[];
  additionalProperties?: boolean;
  items?: JsonSchema;
  enum?: Json[];
  minLength?: number;
  maxLength?: number;
  minimum?: number;
  maximum?: number;
  minItems?: number;
  maxItems?: number;
  pattern?: string;
  anyOf?: JsonSchema[];
}
