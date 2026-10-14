import type { Json, JsonSchema } from "./types.js";

/**
 * Minimal, strict JSON-schema validator for capability arguments.
 * Supports the subset declared in `JsonSchema`. Objects default to
 * `additionalProperties: false` — unexpected arguments are rejected.
 */
export interface ValidationIssue {
  path: string;
  message: string;
}

export function validate(schema: JsonSchema, value: unknown, path = "$"): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  check(schema, value, path, issues);
  return issues;
}

function typeOf(v: unknown): string {
  if (v === null) return "null";
  if (Array.isArray(v)) return "array";
  if (typeof v === "number") return Number.isInteger(v) ? "integer" : "number";
  return typeof v;
}

function check(schema: JsonSchema, value: unknown, path: string, issues: ValidationIssue[]): void {
  if (schema.anyOf) {
    const ok = schema.anyOf.some((s) => validate(s, value, path).length === 0);
    if (!ok) issues.push({ path, message: "does not match any allowed shape" });
    return;
  }

  if (schema.type) {
    const actual = typeOf(value);
    const matches =
      actual === schema.type || (schema.type === "number" && actual === "integer");
    if (!matches) {
      issues.push({ path, message: `expected ${schema.type}, got ${actual}` });
      return;
    }
  }

  if (schema.enum && !schema.enum.some((e) => e === value)) {
    issues.push({ path, message: `must be one of ${schema.enum.map((e) => JSON.stringify(e)).join(", ")}` });
  }

  if (typeof value === "string") {
    if (schema.minLength !== undefined && value.length < schema.minLength)
      issues.push({ path, message: `must have at least ${schema.minLength} characters` });
    if (schema.maxLength !== undefined && value.length > schema.maxLength)
      issues.push({ path, message: `must have at most ${schema.maxLength} characters` });
    if (schema.pattern && !new RegExp(schema.pattern).test(value))
      issues.push({ path, message: `does not match pattern ${schema.pattern}` });
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value)) issues.push({ path, message: "must be finite" });
    if (schema.minimum !== undefined && value < schema.minimum)
      issues.push({ path, message: `must be >= ${schema.minimum}` });
    if (schema.maximum !== undefined && value > schema.maximum)
      issues.push({ path, message: `must be <= ${schema.maximum}` });
  }

  if (Array.isArray(value)) {
    if (schema.minItems !== undefined && value.length < schema.minItems)
      issues.push({ path, message: `must have at least ${schema.minItems} items` });
    if (schema.maxItems !== undefined && value.length > schema.maxItems)
      issues.push({ path, message: `must have at most ${schema.maxItems} items` });
    if (schema.items) value.forEach((v, i) => check(schema.items!, v, `${path}[${i}]`, issues));
  }

  if (schema.type === "object" && value && typeof value === "object" && !Array.isArray(value)) {
    const obj = value as Record<string, unknown>;
    const props = schema.properties ?? {};
    for (const req of schema.required ?? []) {
      if (!(req in obj) || obj[req] === undefined) issues.push({ path: `${path}.${req}`, message: "is required" });
    }
    for (const [k, v] of Object.entries(obj)) {
      const sub = props[k];
      if (!sub) {
        if (schema.additionalProperties !== true) issues.push({ path: `${path}.${k}`, message: "is not an allowed argument" });
        continue;
      }
      check(sub, v, `${path}.${k}`, issues);
    }
  }
}

/** Throws-free JSON parse returning undefined on failure. */
export function tryParseJson(text: string): Json | undefined {
  try {
    return JSON.parse(text) as Json;
  } catch {
    return undefined;
  }
}
