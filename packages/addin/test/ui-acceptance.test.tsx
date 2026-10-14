// @vitest-environment jsdom
/** The acceptance runner panel: safety gate, recording mode, and results. */
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { App } from "../src/ui/App";
import { createHost } from "../src/services/host";

beforeAll(() => {
  Element.prototype.scrollIntoView = () => {};
  localStorage.setItem("loai.settings", JSON.stringify({ transport: "direct", ollamaEndpoint: "http://127.0.0.1:9", model: "" }));
});
afterAll(() => cleanup());

const openSettings = (host: string) => {
  render(<App host={createHost(null, host)} />);
  fireEvent.click(screen.getByRole("tab", { name: "Settings" }));
};

describe("acceptance panel", () => {
  it("requires confirming a blank test document, and records demo runs as mocked", async () => {
    openSettings("excel");
    const run = screen.getByText("Excel: dashboard with Excel-calculated, traceable metrics and a chart").closest("details")!;
    fireEvent.click(run.querySelector("summary")!);
    const button = screen.getAllByText("Run (scripted)")[0]! as HTMLButtonElement;
    expect(button.disabled).toBe(true);
    fireEvent.click(screen.getByLabelText(/This is a blank test document/));
    expect(button.disabled).toBe(false);
    fireEvent.click(button);
    const result = await screen.findByLabelText("Acceptance result", {}, { timeout: 5000 });
    expect(result.textContent).toMatch(/Passed \(automated checks\) · mocked · scripted responses/);
    expect(result.textContent).toMatch(/Values Excel calculated match independently computed totals/);
    expect(screen.getByText("Download JSON")).toBeTruthy();
    cleanup();
  });

  it("blocks a scenario when the document is not blank", async () => {
    openSettings("word");
    fireEvent.click(screen.getByLabelText(/This is a blank test document/));
    fireEvent.click(screen.getAllByText("Run (scripted)")[0]!);
    const result = await screen.findByLabelText("Acceptance result");
    await waitFor(() => expect(result.textContent).toMatch(/Blocked/));
    expect(result.textContent).toMatch(/not blank/);
    cleanup();
  });
});
