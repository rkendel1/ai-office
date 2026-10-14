// @vitest-environment jsdom
/**
 * Task-pane smoke test: renders the real React UI in demo mode (in-memory Word
 * document) against an HTTP Ollama stand-in, and drives the full
 * ask → propose → approve → apply → verify flow through the interface.
 */
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { App } from "../src/ui/App";
import { createHost } from "../src/services/host";
import { startFakeOllama } from "../../core/test/fake-ollama";
import type { FakeWordDriver } from "@loai/core";

let ollama: Awaited<ReturnType<typeof startFakeOllama>>;
const nodeFetch = globalThis.fetch;

beforeAll(async () => {
  ollama = await startFakeOllama({
    replies: [
      { content: "", tool_calls: [{ function: { name: "context.read", arguments: {} } }] },
      {
        content: "I'll tighten paragraph 3.",
        tool_calls: [
          {
            function: {
              name: "document.replace_text",
              arguments: {
                target: "paragraphs",
                start: 3,
                expectedText: "Risks: the data export takes too long and the onboarding guide is out of date.",
                newText: "Risks: slow data exports and an outdated onboarding guide.",
              },
            },
          },
        ],
      },
      { content: "Paragraph 3 now reads more concisely." },
    ],
  });
  // jsdom's AbortSignal is not accepted by Node's fetch; strip it for the test transport.
  globalThis.fetch = ((u: any, init: any = {}) => nodeFetch(u, { ...init, signal: undefined })) as typeof fetch;
  Element.prototype.scrollIntoView = () => {};
  localStorage.setItem(
    "loai.settings",
    JSON.stringify({ transport: "direct", ollamaEndpoint: ollama.url, model: "llama3.1:8b", companionUrl: "https://localhost:1" }),
  );
});

afterAll(async () => {
  cleanup();
  globalThis.fetch = nodeFetch;
  await ollama.close();
});

describe("task pane (demo Word document)", () => {
  it("connects, proposes a change, applies it after approval, and shows verified status", async () => {
    const host = createHost(null, "word");
    const doc = (host.adapter as any).d as FakeWordDriver;
    render(<App host={host} />);

    await waitFor(() => expect(document.querySelector(".dot-ok")).not.toBeNull());
    expect(screen.getByText("llama3.1:8b")).toBeTruthy();

    fireEvent.change(screen.getByLabelText("Message"), { target: { value: "Tighten the risks paragraph" } });
    fireEvent.click(screen.getByText("Send"));

    const apply = await screen.findByText("Apply", {}, { timeout: 5000 });
    expect(screen.getByText("Review this change")).toBeTruthy();
    expect(document.querySelector(".diff ins")?.textContent).toContain("slow");
    // Document unchanged until approved
    expect(doc.paragraphs[3]!.text).toMatch(/takes too long/);

    await act(async () => fireEvent.click(apply));
    await screen.findByText("Done", {}, { timeout: 5000 });
    expect(doc.paragraphs[3]!.text).toBe("Risks: slow data exports and an outdated onboarding guide.");
    expect(screen.getAllByText("Verified").length).toBeGreaterThan(0);
    expect(screen.getByText(/Context used: Document \(all 5 paragraphs\)/)).toBeTruthy();

    fireEvent.click(screen.getByRole("tab", { name: "Context" }));
    expect(screen.getByText("Document (all 5 paragraphs)")).toBeTruthy();
    expect(screen.getByText("may have changed")).toBeTruthy();

    fireEvent.click(screen.getByRole("tab", { name: "Actions" }));
    expect(screen.getByText("Undo")).toBeTruthy();

    fireEvent.click(screen.getByRole("tab", { name: "Settings" }));
    expect(screen.getByText("Approvals")).toBeTruthy();
    // Floors are enforced in the UI: "send" offers no "Allow".
    const send = screen.getByLabelText("Send or publish") as HTMLSelectElement;
    expect([...send.options].map((o) => o.value)).toEqual(["confirm", "deny"]);
  });
});
