/**
 * Deterministic specialist routing. Uses only the user's request text, the
 * active host and an explicit user selection — never model output — so a
 * model or a document cannot select a specialist, let alone one for another host.
 */
import type { HostKind } from "../types.js";
import type { SpecialistRegistry } from "./registry.js";
import type { Specialist } from "./types.js";

export const MIN_SCORE = 2;
/** Scores closer than this between unrelated specialists are treated as ambiguous. */
export const AMBIGUITY_MARGIN = 1;

export type RouteResult =
  | { kind: "specialist"; specialist: Specialist; how: "selected" | "matched"; score: number; reason: string }
  | { kind: "clarify"; question: string; options: Specialist[] }
  | { kind: "general"; reason: string }
  | { kind: "rejected"; reason: string };

export interface Scored {
  specialist: Specialist;
  score: number;
  matched: string[];
}

export function scoreRequest(request: string, candidates: Specialist[]): Scored[] {
  return candidates
    .map((s) => {
      const matched = s.routing.patterns.filter((p) => p.re.test(request));
      return { specialist: s, score: matched.reduce((n, p) => n + p.weight, 0), matched: matched.map((p) => (request.match(p.re)?.[0] ?? "").trim()) };
    })
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score || a.specialist.id.localeCompare(b.specialist.id));
}

export class SpecialistRouter {
  constructor(private readonly specialists: SpecialistRegistry) {}

  /**
   * @param selectedId explicit user choice from the task pane ("auto" / undefined = route by intent;
   *                   "general" = the general-purpose assistant).
   */
  route(request: string, host: HostKind, selectedId?: string): RouteResult {
    if (selectedId === "general") return { kind: "general", reason: "You chose the general assistant." };
    if (selectedId && selectedId !== "auto") {
      const s = this.specialists.get(selectedId);
      if (!s) return { kind: "rejected", reason: `Unknown specialist "${selectedId}".` };
      if (s.host !== host) return { kind: "rejected", reason: `${s.name} works in ${s.host}, not in ${host}.` };
      return { kind: "specialist", specialist: s, how: "selected", score: Number.POSITIVE_INFINITY, reason: "You selected this specialist." };
    }

    const scored = scoreRequest(request, this.specialists.forHost(host)).filter((x) => x.score >= MIN_SCORE);
    if (!scored.length) return { kind: "general", reason: "No specialist clearly matches this request." };

    let top = scored[0]!;
    // A specialist whose workflow includes the top match's work (e.g. a reply writer analyses the
    // thread first) wins when it also matched: it is the only one that can produce everything asked for.
    const covering = scored.find((x) => x !== top && x.specialist.routing.subsumes?.includes(top.specialist.id));
    if (covering) top = covering;

    const rivals = scored.filter(
      (x) =>
        x !== top &&
        !top.specialist.routing.subsumes?.includes(x.specialist.id) &&
        !x.specialist.routing.subsumes?.includes(top.specialist.id) &&
        top.score - x.score < AMBIGUITY_MARGIN,
    );
    if (rivals.length) {
      const options = [top, ...rivals].map((x) => x.specialist);
      return {
        kind: "clarify",
        question: `This could be handled in different ways: ${options.map((o) => `${o.name} (${o.description.split(/[.—]/)[0]!.trim().toLowerCase()})`).join(" or ")}. Which do you want?`,
        options,
      };
    }
    return {
      kind: "specialist",
      specialist: top.specialist,
      how: "matched",
      score: top.score,
      reason: covering ? `Your request includes work the ${top.specialist.name} covers end to end.` : `Matched: ${top.matched.filter(Boolean).slice(0, 3).map((m) => `“${m}”`).join(", ")}.`,
    };
  }
}
