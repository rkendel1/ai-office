import { describe, expect, it } from "vitest";
import { parseTextToolCalls } from "../src/inference/toolcalls";

describe("XML-style tool calls (Qwen family)", () => {
  it("parses function and parameters, coercing numbers", () => {
    const r = parseTextToolCalls(
      "<tool_call>\n<function=document.replace_text>\n<parameter=target>\nparagraphs\n</parameter>\n<parameter=start>\n5\n</parameter>\n<parameter=newText>\nHello \"there\"\n</parameter>\n</function>\n</tool_call>",
    );
    expect(r.toolCalls).toHaveLength(1);
    expect(r.toolCalls[0]!.name).toBe("document.replace_text");
    expect(r.toolCalls[0]!.arguments).toEqual({ target: "paragraphs", start: 5, newText: 'Hello "there"' });
    expect(r.text).toBe("");
  });
  it("tolerates an unclosed call and a mangled parameter tag", () => {
    const r = parseTextToolCalls('Applying.\n<tool_call>\n<function=document.replace_text>\n<parameter>expectedText">old text\n</parameter>\n</function>');
    expect(r.toolCalls[0]!.arguments).toEqual({ expectedText: "old text" });
    expect(r.text).toBe("Applying.");
  });
});
