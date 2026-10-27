import { describe, expect, it } from "vitest";
import { companionOriginForPage } from "../src/services/settings";

const page = (overrides: Partial<Parameters<typeof companionOriginForPage>[0]> = {}) => ({
  protocol: "https:",
  hostname: "localhost",
  port: "47615",
  origin: "https://localhost:47615",
  ...overrides,
});

describe("companion URL defaults", () => {
  it("uses the origin of a pane served by the companion", () => {
    expect(companionOriginForPage(page())).toBe("https://localhost:47615");
    expect(companionOriginForPage(page({ port: "47616", origin: "https://localhost:47616" }))).toBe("https://localhost:47616");
    expect(companionOriginForPage(page({ hostname: "127.0.0.1", origin: "https://127.0.0.1:47615" }))).toBe("https://127.0.0.1:47615");
  });

  it("keeps the default companion URL for the Vite dev server and non-local pages", () => {
    expect(companionOriginForPage(page({ port: "3000", origin: "https://localhost:3000" }))).toBeUndefined();
    expect(companionOriginForPage(page({ protocol: "http:" }))).toBeUndefined();
    expect(companionOriginForPage(page({ hostname: "office.example.com", origin: "https://office.example.com" }))).toBeUndefined();
    expect(companionOriginForPage(page({ port: "", origin: "https://localhost" }))).toBeUndefined();
  });
});
