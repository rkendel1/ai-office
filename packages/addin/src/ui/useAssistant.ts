import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  createSession,
  MemoryHistoryStore,
  type ApprovalRequest,
  type ApprovalResponse,
  type AssistantSettings,
  type ExecutionRecord,
  type HistoryStore,
  type InferenceProvider,
  type Plan,
  type RuntimeEvent,
  type SharedSource,
  type Specialist,
  type TaskReport,
  type TaskState,
  type ToolCall,
} from "@loai/core";
import type { HostInfo } from "../services/host";

export type ChatItem =
  | { id: string; kind: "user"; text: string }
  | { id: string; kind: "assistant"; text: string; streaming: boolean }
  | { id: string; kind: "action"; record: ExecutionRecord }
  | { id: string; kind: "report"; report: TaskReport; request: string; specialist?: Specialist }
  | { id: string; kind: "notice"; level: "info" | "warning" | "error"; text: string }
  | { id: string; kind: "specialist"; specialist: Specialist; reason?: string }
  | { id: string; kind: "plan"; plan: Plan; batch: boolean };

export interface PendingApproval {
  request: ApprovalRequest;
  resolve: (r: ApprovalResponse) => void;
}

let seq = 0;
const nid = () => `i${++seq}`;

/** Wires the core runtime to React state: chat stream, actions, approvals, cancellation. */
export function useAssistant(host: HostInfo, getProvider: () => InferenceProvider, getSettings: () => AssistantSettings, history: HistoryStore | null) {
  const [items, setItems] = useState<ChatItem[]>([]);
  const [state, setState] = useState<TaskState>("idle");
  const [stateDetail, setStateDetail] = useState<string | undefined>();
  const [pending, setPending] = useState<PendingApproval | null>(null);
  const [actions, setActions] = useState<ExecutionRecord[]>([]);
  const [contextVersion, setContextVersion] = useState(0);
  const abortRef = useRef<AbortController | null>(null);
  const historyRef = useRef<HistoryStore>(history ?? new MemoryHistoryStore());
  historyRef.current = history ?? historyRef.current;

  const session = useMemo(() => {
    // Delegating provider: settings changes take effect on the next task without losing context.
    const provider: InferenceProvider = {
      id: "delegate",
      listModels: (s) => getProvider().listModels(s),
      describeModel: (n, s) => getProvider().describeModel(n, s),
      chat: (r) => getProvider().chat(r),
      ping: (s) => getProvider().ping(s),
    };
    const delegatingHistory: HistoryStore = {
      add: (r) => historyRef.current.add(r),
      list: (l) => historyRef.current.list(l),
      get: (id) => historyRef.current.get(id),
      clear: () => historyRef.current.clear(),
      prune: (d) => historyRef.current.prune(d),
    };
    return createSession({
      adapter: host.adapter,
      provider,
      settings: getSettings,
      history: delegatingHistory,
      approver: (request, signal) =>
        new Promise<ApprovalResponse>((resolve) => {
          const done = (r: ApprovalResponse) => {
            setPending(null);
            resolve(r);
          };
          signal?.addEventListener("abort", () => done({ approved: false, reason: "Task cancelled" }));
          setPending({ request, resolve: done });
        }),
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [host.adapter]);

  useEffect(() => session.context.subscribe(() => setContextVersion((v) => v + 1)), [session]);

  const run = useCallback(
    async (text: string, specialist?: Specialist, reason?: string, sources?: SharedSource[]) => {
      if (session.runtime.busy) return;
      const ac = new AbortController();
      abortRef.current = ac;
      setItems((xs) => [...xs, { id: nid(), kind: "user", text }, ...(specialist ? [{ id: nid(), kind: "specialist" as const, specialist, reason }] : [])]);
      let open: string | null = null;
      const onEvent = (e: RuntimeEvent) => {
        switch (e.type) {
          case "state":
            setState(e.state);
            setStateDetail(e.detail);
            break;
          case "delta":
            if (!open) {
              const id = nid();
              open = id;
              setItems((xs) => [...xs, { id, kind: "assistant", text: e.text, streaming: true }]);
            } else {
              const id = open;
              setItems((xs) => xs.map((x) => (x.id === id && x.kind === "assistant" ? { ...x, text: x.text + e.text } : x)));
            }
            break;
          case "assistant": {
            const id = open;
            open = null;
            if (id) setItems((xs) => xs.map((x) => (x.id === id && x.kind === "assistant" ? { ...x, text: e.text, streaming: false } : x)));
            else setItems((xs) => [...xs, { id: nid(), kind: "assistant", text: e.text, streaming: false }]);
            break;
          }
          case "action_started": {
            const id = open;
            open = null;
            // Drop a bubble that only streamed the raw tool-call block.
            if (id) setItems((xs) => xs.filter((x) => !(x.id === id && x.kind === "assistant" && (!x.text.trim() || /```tool|^\s*\{/.test(x.text)))));
            break;
          }
          case "action":
            setItems((xs) => [...xs, { id: nid(), kind: "action", record: e.record }]);
            setActions((as) => [e.record, ...as].slice(0, 200));
            break;
          case "notice":
            setItems((xs) => [...xs, { id: nid(), kind: "notice", level: e.level, text: e.text }]);
            break;
          case "plan":
            setItems((xs) => [...xs, { id: nid(), kind: "plan", plan: e.plan, batch: e.batch }]);
            break;
        }
      };
      try {
        const report = await session.runtime.run(text, { signal: ac.signal, onEvent, specialist, sources });
        const id = open;
        if (id) setItems((xs) => xs.map((x) => (x.id === id && x.kind === "assistant" ? { ...x, streaming: false } : x)));
        setItems((xs) => [...xs, { id: nid(), kind: "report", report, request: text, specialist }]);
      } catch (e) {
        setItems((xs) => [...xs, { id: nid(), kind: "notice", level: "error", text: (e as Error).message }]);
        setState("failed");
      } finally {
        abortRef.current = null;
      }
    },
    [session],
  );

  const cancel = useCallback(() => abortRef.current?.abort(), []);

  /** User-initiated re-execution of one capability call (retry or undo). Requires approval like any change. */
  const execute = useCallback(
    async (call: ToolCall) => {
      const rec = await session.runtime.executeDirect(call);
      setActions((as) => [rec, ...as]);
      setItems((xs) => [...xs, { id: nid(), kind: "action", record: rec }]);
      return rec;
    },
    [session],
  );

  const captureForSharing = useCallback(() => session.runtime.captureForSharing(), [session]);

  const newConversation = useCallback(() => {
    session.runtime.resetConversation();
    setItems([]);
  }, [session]);

  return { session, items, state, stateDetail, pending, actions, contextVersion, run, cancel, execute, newConversation, captureForSharing };
}
