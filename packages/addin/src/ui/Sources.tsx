import { useEffect, useState } from "react";
import { APP_NAMES, sourceLabel, type SourceCapture, type SourceSummary } from "@loai/core";
import type { CompanionClient } from "../services/companion";
import type { useAssistant } from "./useAssistant";

const age = (t: number) => {
  const m = Math.round((Date.now() - t) / 60_000);
  return m < 1 ? "just now" : `${m} min ago`;
};

/**
 * Source app: share the current selection with other Office apps. Reads the
 * selection only when the user clicks, shows exactly what will be shared, and
 * sends nothing until the user confirms.
 */
export function SharePanel({ a, companion, available }: { a: ReturnType<typeof useAssistant>; companion: CompanionClient; available: boolean }) {
  const [preview, setPreview] = useState<SourceCapture | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [mine, setMine] = useState<SourceSummary[]>([]);
  const app = a.session.adapter.host;

  const refresh = async () => {
    if (!available) return;
    try {
      setMine((await companion.listSources()).filter((s) => s.app === app));
    } catch {
      /* shown on next action */
    }
  };
  useEffect(() => {
    refresh();
  }, [available]);

  if (!available)
    return <p className="muted small">Sharing content between Office apps needs the paired companion (Settings ▸ Connection). You can always paste content into a request instead.</p>;

  return (
    <div className="share">
      <div className="row between">
        <b className="small">Share with other Office apps</b>
        <button
          onClick={async () => {
            setMsg(null);
            const r = await a.captureForSharing();
            if (r.capture) setPreview(r.capture);
            else setMsg(r.reason ?? "Nothing to share.");
          }}
        >
          Share my selection…
        </button>
      </div>
      {msg && <p className="small warn">{msg}</p>}
      {preview && (
        <div className="share-preview" role="dialog" aria-label="Confirm sharing">
          <div className="small">
            <b>{sourceLabel(preview)}</b> · {preview.content.length.toLocaleString()} characters{preview.truncated ? " (truncated)" : ""}
          </div>
          <pre className="ctx-body">{preview.content.slice(0, 1500)}{preview.content.length > 1500 ? "\n…" : ""}</pre>
          <p className="small muted">Only this text is shared, kept in the companion's memory on this computer for 30 minutes. Other apps can use it only when you attach it to a request there.</p>
          <div className="row gap">
            <button
              className="primary"
              onClick={async () => {
                try {
                  await companion.shareSource(preview);
                  setPreview(null);
                  setMsg("Shared. Attach it from the Sources button in the other app.");
                  refresh();
                } catch (e) {
                  setMsg((e as Error).message);
                }
              }}
            >
              Share
            </button>
            <button onClick={() => setPreview(null)}>Cancel</button>
          </div>
        </div>
      )}
      {mine.length > 0 && (
        <ul className="small share-list">
          {mine.map((s) => (
            <li key={s.id} className="row between">
              <span>
                {s.scopeLabel} <span className="muted">· {age(s.sharedAt)}</span>
              </span>
              <button className="link small" onClick={async () => (await companion.deleteSource(s.id), refresh())}>
                Stop sharing
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/** Destination app: pick shared sources to attach to requests. Nothing is attached by default. */
export function SourcesBar({
  companion,
  available,
  attached,
  setAttached,
  host,
}: {
  companion: CompanionClient;
  available: boolean;
  attached: SourceSummary[];
  setAttached: (s: SourceSummary[]) => void;
  host: string;
}) {
  const [open, setOpen] = useState(false);
  const [list, setList] = useState<SourceSummary[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  useEffect(() => {
    if (!open || !available) return;
    const load = () =>
      companion.listSources().then(
        (l) => (setList(l.filter((s) => s.app !== host)), setErr(null)),
        (e) => setErr((e as Error).message),
      );
    load();
    const t = setInterval(load, 15_000);
    return () => clearInterval(t);
  }, [open, available]);

  const toggle = (s: SourceSummary) => setAttached(attached.some((x) => x.id === s.id) ? attached.filter((x) => x.id !== s.id) : [...attached, s]);

  return (
    <div className="sources">
      <div className="row gap">
        <button className="small-btn" onClick={() => setOpen((v) => !v)} aria-expanded={open}>
          Sources{attached.length ? ` (${attached.length})` : ""}
        </button>
        {attached.map((s) => (
          <span key={s.id} className="chip-attached small" title={`Shared ${age(s.sharedAt)}`}>
            {sourceLabel(s)}
            <button className="link" aria-label={`Detach ${sourceLabel(s)}`} onClick={() => toggle(s)}>
              ×
            </button>
          </span>
        ))}
      </div>
      {open && (
        <div className="sources-panel small" role="dialog" aria-label="Shared sources">
          {!available ? (
            <p className="muted">Sharing between Office apps needs the paired companion. Paste content into your message instead.</p>
          ) : err ? (
            <p className="err">{err}</p>
          ) : !list ? (
            <p className="muted">Loading…</p>
          ) : !list.length ? (
            <p className="muted">Nothing shared yet. In Word, Excel, PowerPoint or Outlook, open this assistant, go to Context, and use “Share my selection”.</p>
          ) : (
            <ul>
              {list.map((s) => (
                <li key={s.id}>
                  <label className="check">
                    <input type="checkbox" checked={attached.some((x) => x.id === s.id)} onChange={() => toggle(s)} />
                    <span>
                      <b>{APP_NAMES[s.app]}</b> {s.documentTitle ? `“${s.documentTitle}” ` : ""}— {s.scopeLabel}
                      <span className="muted"> · {s.chars.toLocaleString()} chars · {age(s.sharedAt)}</span>
                    </span>
                  </label>
                </li>
              ))}
            </ul>
          )}
          <p className="muted">Attached content is sent to your local model with your request. The assistant cannot see anything you have not attached.</p>
        </div>
      )}
    </div>
  );
}
