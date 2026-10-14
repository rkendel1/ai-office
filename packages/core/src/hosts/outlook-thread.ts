/**
 * Split an email body (with its quoted reply chain) into attributed messages.
 * Deterministic and conservative: if a header cannot be recognised, the text
 * stays with the previous message rather than being attributed to anyone.
 */
import type { MailItem } from "./outlook.js";

export interface ThreadMessage {
  /** M1 = the open message (newest), M2 = the message it quotes, ... */
  id: string;
  author: string;
  sent?: string;
  text: string;
}

const FROM = /^\s*\*?From:\*?\s*(.+?)\s*$/i;
const SENT = /^\s*\*?(Sent|Date):\*?\s*(.+?)\s*$/i;
const HEADER_FIELD = /^\s*\*?(To|Cc|Bcc|Subject):\*?/i;
const ON_WROTE_TIME = /^\s*On\s+(.+\d{1,2}:\d{2}(?:\s*[AP]M)?)\s*,?\s+(.+?)\s+wrote:\s*$/i;
const ON_WROTE_DATE = /^\s*On\s+(.+?\d{4})\s*,?\s+(.+?)\s+wrote:\s*$/i;
const ORIGINAL = /^\s*-{2,}\s*Original Message\s*-{2,}\s*$/i;

export function displayName(addr: string): string {
  const m = /^\s*"?([^"<]+?)"?\s*<[^>]+>\s*$/.exec(addr);
  if (m) return m[1]!.trim();
  return addr.replace(/[<>]/g, "").trim();
}

export function splitThread(item: Pick<MailItem, "from" | "body" | "sent">, selfLabel = "You (draft)"): ThreadMessage[] {
  const lines = item.body.replace(/\r\n?/g, "\n").split("\n");
  const msgs: { author: string; sent?: string; lines: string[] }[] = [{ author: item.from ? displayName(item.from) : selfLabel, sent: item.sent, lines: [] }];
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i]!.replace(/^(\s*>)+\s?/, "");
    if (ORIGINAL.test(raw)) continue;
    const from = FROM.exec(raw);
    if (from) {
      // Require a Sent/Date line within the next few lines to treat this as a header.
      let sent: string | undefined;
      let j = i + 1;
      for (; j < Math.min(lines.length, i + 6); j++) {
        const l = lines[j]!.replace(/^(\s*>)+\s?/, "");
        const s = SENT.exec(l);
        if (s) sent = s[2];
        else if (!HEADER_FIELD.test(l)) break;
      }
      if (sent) {
        msgs.push({ author: displayName(from[1]!), sent, lines: [] });
        i = j - 1;
        continue;
      }
    }
    const on = ON_WROTE_TIME.exec(raw) ?? ON_WROTE_DATE.exec(raw);
    if (on) {
      msgs.push({ author: displayName(on[2]!), sent: on[1], lines: [] });
      continue;
    }
    msgs[msgs.length - 1]!.lines.push(raw);
  }
  return msgs
    .map((m) => ({ ...m, text: m.lines.join("\n").replace(/\n{3,}/g, "\n\n").trim() }))
    .filter((m, i) => i === 0 || m.text)
    .map((m, i) => ({ id: `M${i + 1}`, author: m.author, sent: m.sent, text: m.text }));
}

export function renderThread(msgs: ThreadMessage[]): string {
  return msgs.map((m) => `[${m.id}] From: ${m.author}${m.sent ? ` — Sent: ${m.sent}` : ""}\n${m.text || "(no text)"}`).join("\n\n");
}
