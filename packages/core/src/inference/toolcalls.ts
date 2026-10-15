import type { ChatMessage, Json, ToolCall, ToolSpec } from "../types.js";
import { newId } from "../util.js";

/**
 * Text-based tool-call protocol for local models without native tool calling.
 * The model replies with a fenced block:
 *
 * ```tool
 * {"name": "document.replace_text", "arguments": {...}}
 * ```
 *
 * Only the model's own reply is parsed — never document content — and the
 * registry still validates and authorizes everything parsed here.
 */
export const TOOL_FENCE = "tool";

export function describeToolsAsText(tools: ToolSpec[]): string {
  const lines = [
    "You can use capabilities. To call one, reply with ONLY a fenced block and nothing else:",
    "```tool",
    '{"name": "<capability name>", "arguments": { ... }}',
    "```",
    "Call one capability per reply. You will receive the result, then decide the next step.",
    "Available capabilities:",
  ];
  for (const t of tools) lines.push(`- ${t.name}: ${t.description}\n  arguments schema: ${JSON.stringify(t.parameters)}`);
  return lines.join("\n");
}

function asCall(obj: unknown): ToolCall | undefined {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) return undefined;
  const o = obj as Record<string, unknown>;
  const name = typeof o.name === "string" ? o.name : typeof o.tool === "string" ? o.tool : undefined;
  if (!name) return undefined;
  const args = o.arguments ?? o.args ?? o.parameters ?? {};
  if (!args || typeof args !== "object" || Array.isArray(args)) return undefined;
  return { id: newId("call"), name, arguments: args as Record<string, Json> };
}

export interface ParsedReply {
  toolCalls: ToolCall[];
  /** Text with tool blocks removed. */
  text: string;
  /** Text that looked like a tool call but could not be parsed. */
  malformed: string[];
}

export function parseTextToolCalls(content: string): ParsedReply {
  const toolCalls: ToolCall[] = [];
  const malformed: string[] = [];
  const fence = /```(?:tool|tool_call|json)?\s*\n([\s\S]*?)```/g;
  let text = content;
  let m: RegExpExecArray | null;
  const blocks: string[] = [];
  while ((m = fence.exec(content))) {
    const raw = m[1]!.trim();
    const isToolFence = /^```(tool|tool_call)/.test(m[0]);
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      if (isToolFence) malformed.push(raw);
      continue;
    }
    const call = asCall(parsed);
    if (call) {
      toolCalls.push(call);
      blocks.push(m[0]);
    } else if (isToolFence) malformed.push(raw);
  }
  // Whole reply is a bare JSON call.
  if (!toolCalls.length) {
    const trimmed = content.trim();
    if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
      try {
        const call = asCall(JSON.parse(trimmed));
        if (call) {
          toolCalls.push(call);
          blocks.push(content);
        }
      } catch {
        /* not JSON */
      }
    }
  }
  // Qwen-style XML calls: <tool_call><function=name><parameter=key>value</parameter></function></tool_call>
  if (!toolCalls.length) {
    const xml = /<tool_call>([\s\S]*?)(?:<\/tool_call>|$)/g;
    while ((m = xml.exec(content))) {
      const fn = /<function=([\w.\-]+)>/.exec(m[1]!);
      if (!fn) {
        malformed.push(m[1]!.trim());
        continue;
      }
      const args: Record<string, Json> = {};
      const param = /<parameter(?:=([\w.\-]+)>|>\s*([\w.\-]+)"?>)([\s\S]*?)<\/parameter>/g;
      let pm: RegExpExecArray | null;
      while ((pm = param.exec(m[1]!))) {
        const raw = pm[3]!.trim();
        let v: Json = raw;
        try {
          v = JSON.parse(raw) as Json;
        } catch {
          /* plain string */
        }
        args[(pm[1] ?? pm[2])!] = v;
      }
      toolCalls.push({ id: newId("call"), name: fn[1]!, arguments: args });
      blocks.push(m[0]);
    }
  }
  for (const b of blocks) text = text.replace(b, "");
  // Drop stray closing tags left by malformed XML replies.
  if (toolCalls.length) text = text.replace(/<\/?(?:tool_call|function|parameter)[^>]*>/g, "");
  return { toolCalls, text: text.trim(), malformed };
}

/** Normalise a model reply into tool calls regardless of protocol. */
export function extractToolCalls(message: ChatMessage, nativeTools: boolean): ParsedReply {
  if (message.toolCalls?.length) return { toolCalls: message.toolCalls, text: message.content.trim(), malformed: [] };
  // Some models emit text-protocol calls even when native tools are offered.
  const parsed = parseTextToolCalls(message.content);
  if (!nativeTools || parsed.toolCalls.length) return parsed;
  return { toolCalls: [], text: message.content.trim(), malformed: [] };
}
