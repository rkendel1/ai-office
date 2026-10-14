import {
  DEFAULT_SETTINGS,
  MemoryHistoryStore,
  createSession,
  normalizeSettings,
  type ApprovalRequest,
  type Approver,
  type AssistantSettings,
  type HostAdapter,
  type InferenceProvider,
} from "../src/index.js";

export function settings(over: Partial<AssistantSettings> = {}): AssistantSettings {
  return normalizeSettings({ ...DEFAULT_SETTINGS, model: "test-model", ...over });
}

export class RecordingApprover {
  readonly requests: ApprovalRequest[] = [];
  constructor(public answer: boolean | ((r: ApprovalRequest) => boolean) = true) {}
  readonly fn: Approver = async (req) => {
    this.requests.push(req);
    const ok = typeof this.answer === "function" ? this.answer(req) : this.answer;
    return { approved: ok, reason: ok ? undefined : "no thanks" };
  };
}

export function setup(adapter: HostAdapter, provider: InferenceProvider, opts: { approve?: boolean | ((r: ApprovalRequest) => boolean); settings?: Partial<AssistantSettings> } = {}) {
  const approver = new RecordingApprover(opts.approve ?? true);
  const history = new MemoryHistoryStore();
  const s = settings(opts.settings);
  const session = createSession({ adapter, provider, approver: approver.fn, settings: () => s, history });
  return { ...session, approver, history, settings: s };
}
