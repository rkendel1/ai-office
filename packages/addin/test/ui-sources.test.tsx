// @vitest-environment jsdom
/**
 * Cross-app handoff through the real UI and a real companion: share a Word
 * selection, attach it in PowerPoint, and check exactly that content (with
 * provenance) reaches the local model.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { App } from "../src/ui/App";
import { createHost } from "../src/services/host";
import { startFakeOllama } from "../../core/test/fake-ollama";
import { loadConfig } from "../../companion/src/config";
import { startCompanion, type Companion } from "../../companion/src/server";

let ollama: Awaited<ReturnType<typeof startFakeOllama>>;
let companion: Companion;
let home: string;
const nodeFetch = globalThis.fetch;

beforeAll(async () => {
  ollama = await startFakeOllama({ replies: [{ content: "Who is the audience for this deck?" }] });
  home = fs.mkdtempSync(path.join(os.tmpdir(), "loai-ui-"));
  // jsdom's AbortSignal is not an instance of Node's, which undici's fetch requires.
  const stripSignal = ((u: any, init: any = {}) => nodeFetch(u, { ...init, signal: undefined })) as typeof fetch;
  companion = await startCompanion(loadConfig({ home, port: 0, ollamaUrl: ollama.url, tlsCertPath: undefined, tlsKeyPath: undefined }), { log: () => {}, fetchImpl: stripSignal });
  const { code } = companion.pairing.issueCode();
  const token = (companion.pairing.redeem(code, "ui test") as { token: string }).token;
  globalThis.fetch = ((u: any, init: any = {}) => nodeFetch(u, { ...init, signal: undefined })) as typeof fetch;
  Element.prototype.scrollIntoView = () => {};
  localStorage.setItem("loai.companionToken", token);
  localStorage.setItem("loai.settings", JSON.stringify({ transport: "companion", companionUrl: companion.url, model: "llama3.1:8b" }));
});

afterAll(async () => {
  cleanup();
  globalThis.fetch = nodeFetch;
  await companion.close();
  await ollama.close();
  fs.rmSync(home, { recursive: true, force: true });
});

describe("sharing between Office apps", () => {
  it("shares a Word selection and uses only that content in PowerPoint", async () => {
    // --- Word pane: share the selection
    const word = createHost(null, "word");
    (word.adapter as any).d.selectionRange = { start: 1, end: 2 };
    render(<App host={word} />);
    await waitFor(() => expect(document.querySelector(".dot-ok")).not.toBeNull());
    fireEvent.click(screen.getByRole("tab", { name: "Context" }));
    fireEvent.click(await screen.findByText("Share my selection…"));
    const dialog = await screen.findByRole("dialog", { name: "Confirm sharing" });
    expect(dialog.textContent).toContain("The pilot launched in March with 40 users");
    expect(dialog.textContent).not.toContain("Next steps"); // outside the selection
    expect(companion.sources.list()).toHaveLength(0); // nothing shared before confirming
    fireEvent.click(screen.getByRole("button", { name: "Share" }));
    await waitFor(() => expect(companion.sources.list()).toHaveLength(1));
    expect(companion.sources.list()[0]).toMatchObject({ app: "word", documentTitle: "Demo document", scopeLabel: "Selection (paragraphs 1–2)" });
    cleanup();

    // --- PowerPoint pane: attach it and ask
    const ppt = createHost(null, "powerpoint");
    render(<App host={ppt} />);
    await waitFor(() => expect(document.querySelector(".dot-ok")).not.toBeNull());
    fireEvent.click(screen.getByRole("button", { name: /^Sources/ }));
    const box = await screen.findByRole("checkbox");
    fireEvent.click(box);
    expect(screen.getByText(/From Word “Demo document” — Selection \(paragraphs 1–2\)/)).toBeTruthy();
    fireEvent.change(screen.getByLabelText("Message"), { target: { value: "Create a 3-slide deck from the shared pilot notes" } });
    fireEvent.click(screen.getByText("Send"));
    await screen.findByText(/Needs your answer/, {}, { timeout: 5000 });

    const chat = ollama.requests.filter((r) => r.path === "/api/chat").at(-1)!.body;
    const prompt = chat.messages.map((m: any) => m.content).join("\n");
    expect(prompt).toContain("The pilot launched in March with 40 users across two teams.");
    expect(prompt).toContain('from_app="word"');
    expect(prompt).not.toContain("Next steps: fix exports"); // not shared
    expect(chat.tools.map((t: any) => t.function.name).some((n: string) => /source|share/.test(n))).toBe(false);
    expect(screen.getByText(/Context used:.*From Word/)).toBeTruthy();
  });
});
