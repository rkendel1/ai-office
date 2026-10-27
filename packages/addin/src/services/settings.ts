import { DEFAULT_SETTINGS, normalizeSettings, type AssistantSettings } from "@loai/core";

const KEY = "loai.settings";

/** Settings persist per add-in origin in local storage; nothing leaves the machine. */
export function loadSettings(): AssistantSettings {
  try {
    const raw = window.localStorage.getItem(KEY);
    // normalizeSettings migrates settings written before Atomic Chat support.
    const s = normalizeSettings(raw ? JSON.parse(raw) : DEFAULT_SETTINGS);
    // When the add-in is served by the companion, default to that origin.
    if (!raw && location.port === "47615") s.companionUrl = location.origin;
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
