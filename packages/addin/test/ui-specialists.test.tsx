// @vitest-environment jsdom
/**
 * Task-pane specialist UI: selector, routing preview, editable examples,
 * clarification, and the completion checklist in the result.
 */
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { App } from "../src/ui/App";
import { createHost } from "../src/services/host";
import { startFakeOllama } from "../../core/test/fake-ollama";

let ollama: Awaited<ReturnType<typeof startFakeOllama>>;
const nodeFetch = globalThis.fetch;

beforeAll(async () => {
  ollama = await startFakeOllama({
    replies: [{ content: "1. [Observation] ¶3 says exports are slow but gives no fix date.\n   [Recommendation] Add a date for the export fix." }],
  });
  globalThis.fetch = ((u: any, init: any = {}) => nodeFetch(u, { ...init, signal: undefined })) as typeof fetch;
  Element.prototype.scrollIntoView = () => {};
  localStorage.setItem("loai.settings", JSON.stringify({ transport: "direct", ollamaEndpoint: ollama.url, model: "llama3.1:8b", companionUrl: "https://localhost:1" }));
});

afterAll(async () => {
  cleanup();
  globalThis.fetch = nodeFetch;
  await ollama.close();
});

describe("specialists in the task pane", () => {
  it("previews routing, lets the user pick a specialist, edit an example, and resolve ambiguity", async () => {
    const host = createHost(null, "word");
    const doc = (host.adapter as any).d;
    render(<App host={host} />);
    await waitFor(() => expect(document.querySelector(".dot-ok")).not.toBeNull());

    // Auto mode: every Word specialist offered as a starting point.
    for (const name of ["Document Writer", "Document Editor", "Document Reviewer", "Document Summarizer"]) expect(screen.getAllByText(name).length).toBeGreaterThan(0);

    const box = screen.getByLabelText("Message") as HTMLTextAreaElement;
    fireEvent.change(box, { target: { value: "Proofread the selected paragraph" } });
    expect(screen.getByText(/Will use/).textContent).toMatch(/Document Editor/);

    // Choosing a specialist shows its purpose and examples; examples fill the box but do not send.
    fireEvent.change(screen.getByLabelText("Specialist"), { target: { value: "word.summarizer" } });
    expect(screen.getAllByText(/audience-specific summaries/).length).toBeGreaterThan(0);
    fireEvent.click(screen.getAllByText("Give me a five-bullet brief of the key findings")[0]!);
    expect(box.value).toBe("Give me a five-bullet brief of the key findings");
    expect(doc.paragraphs.length).toBe(5);
    expect(screen.queryByText(/Reading context/)).toBeNull();

    // Back to auto; an ambiguous request asks which specialist to use before running.
    fireEvent.change(screen.getByLabelText("Specialist"), { target: { value: "auto" } });
    fireEvent.change(box, { target: { value: "Review and rewrite this document" } });
    fireEvent.click(screen.getByText("Send"));
    const group = await screen.findByRole("group", { name: "Choose a specialist" });
    expect(group.textContent).toMatch(/Document Reviewer/);
    expect(group.textContent).toMatch(/Document Editor/);
    fireEvent.click(screen.getByRole("button", { name: "Document Reviewer" }));

    await screen.findByText("Done", { exact: false }, { timeout: 5000 });
    await waitFor(() => expect(screen.getByLabelText("Completion checks")).toBeTruthy());
    const checks = screen.getByLabelText("Completion checks").textContent!;
    expect(checks).toMatch(/Findings separate observations from recommendations/);
    expect(checks).toMatch(/The source was not modified/);
    expect(document.querySelectorAll(".checks .bad")).toHaveLength(0);
    expect(screen.getAllByText(/Document Reviewer/).length).toBeGreaterThan(0);
    // Implementation details are not shown in the normal UI.
    expect(document.body.textContent).not.toMatch(/inputSchema|Active specialist:|plan\.propose/);
  });
});
