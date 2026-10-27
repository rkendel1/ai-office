import { DEFAULT_SETTINGS, normalizeSettings, type AssistantSettings } from "@loai/core";

const KEY = "loai.settings";

/**
 * When the pane is served by a companion (including one started on an
 * alternate port to avoid the installed app), point at that same origin.
 * The Vite dev server on port 3000 still uses the default companion port.
 */
export function companionOriginForPage(loc: { protocol: string; hostname: string; port: string; origin: string }): string | undefined {
  if (loc.protocol !== "https:") return undefined;
  if (loc.hostname !== "localhost" && loc.hostname !== "127.0.0.1") return undefined;
  if (!loc.port || loc.port === "3000") return undefined;
  return loc.origin;
}

/** Settings persist per add-in origin in local storage; nothing leaves the machine. */
export function loadSettings(): AssistantSettings {
  try {
    const raw = window.localStorage.getItem(KEY);
    // normalizeSettings migrates settings written before Atomic Chat support.
    const s = normalizeSettings(raw ? JSON.parse(raw) : DEFAULT_SETTINGS);
    // When the add-in is served by the companion, default to that origin.
    const companionOrigin = companionOriginForPage(location);
    if (!raw && companionOrigin) s.companionUrl = companionOrigin;
    return s;
  } catch {
    return normalizeSettings(DEFAULT_SETTINGS);
  }
}

export function saveSettings(s: AssistantSettings): AssistantSettings {
  const n = normalizeSettings(s);
  try {
    window.localStorage.setItem(KEY, JSON.stringify(n));
  } catch {
    /* storage unavailable: settings last for this session */
  }
  return n;
}
