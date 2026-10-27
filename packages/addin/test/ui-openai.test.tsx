// @vitest-environment jsdom
/**
 * Atomic Chat is the default local model app. With no model selected the pane
 * discovers the running app's models and picks the first one automatically.
 */
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { App } from "../src/ui/App";
import { createHost } from "../src/services/host";
import { startFakeOpenAI } from "../../core/test/fake-openai";

let ai: Awaited<ReturnType<typeof startFakeOpenAI>>;
const nodeFetch = globalThis.fetch;

beforeAll(async () => {
  ai = await startFakeOpenAI({ models: ["AtomicChat/Qwen3-8B", "AtomicChat/Llama-3.2-3B"] });
  globalThis.fetch = ((u: any, init: any = {}) => nodeFetch(u, { ...init, signal: undefined })) as typeof fetch;
  Element.prototype.scrollIntoView = () => {};
  localStorage.setItem(
    "loai.settings",
    JSON.stringify({ transport: "direct", provider: "openai", endpoint: ai.url, model: "", companionUrl: "https://localhost:1" }),
  );
});

afterAll(async () => {
  cleanup();
  globalThis.fetch = nodeFetch;
  await ai.close();
});

describe("Atomic Chat model discovery", () => {
  it("finds the running app's models and selects the first one", async () => {
    render(<App host={createHost(null, "word")} />);
    await waitFor(() => expect(document.querySelector(".dot-ok")).not.toBeNull());
    await waitFor(() => expect(screen.getAllByText("AtomicChat/Qwen3-8B").length).toBeGreaterThan(0));

    fireEvent.click(screen.getByRole("tab", { name: "Settings" }));
    const select = (await screen.findByLabelText("Model")) as HTMLSelectElement;
    await waitFor(() => expect(select.value).toBe("AtomicChat/Qwen3-8B"));
    expect([...select.options].map((o) => o.value)).toContain("AtomicChat/Llama-3.2-3B");
  });
});
