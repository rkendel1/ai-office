import type { Freshness } from "../context.js";
import { ScopeError, type CapabilityDefinition } from "../registry.js";
import type { ContextItem, ItemIdentity } from "../types.js";
import { normalizeText } from "../util.js";
import { allSupported, type ApiSupportEntry, type HostAdapter, type RequirementChecker } from "./adapter.js";
import { excerpt, makeItem } from "./common.js";

export interface MailItem {
  itemId?: string;
  conversationId?: string;
  itemType: "message" | "appointment" | "other";
  subject: string;
  from?: string;
  to: string[];
  cc: string[];
  sent?: string;
  /** Plain-text body (Outlook returns the whole conversation quote chain when present). */
  body: string;
  attachments: { name: string; size: number }[];
}

export interface OutlookDriver {
  mode(): "read" | "compose";
  item(): Promise<MailItem>;
  /** Compose mode: read the current body. */
  composeBody(): Promise<string>;
  /** Compose mode: insert text at the cursor (or append). */
  insertComposeText(text: string): Promise<void>;
  /**
   * Read mode: open a reply form with the given body. Outlook shows a draft
   * window to the user; it is NOT sent. Office.js cannot read the new form back.
   */
  displayReplyForm(body: string, replyAll: boolean): Promise<void>;
}

/**
 * Outlook adapter. Deliberately has no "send" capability: the user always
 * sends from Outlook's own UI, so no message can be sent without their action.
 */
export class OutlookAdapter implements HostAdapter {
  readonly host = "outlook" as const;

  constructor(
    private readonly d: OutlookDriver,
    private readonly req: RequirementChecker = allSupported,
  ) {}

  async identity(): Promise<ItemIdentity> {
    const i = await this.d.item();
    return { host: "outlook", id: i.itemId ?? i.conversationId, title: i.subject || "(no subject)" };
  }

  apiSupport(): ApiSupportEntry[] {
    return [
      { requirement: "Mailbox 1.1", supported: this.req.isSetSupported("Mailbox", "1.1"), capabilities: ["outlook.read_message", "outlook.create_draft"] },
      { requirement: "Mailbox 1.2", supported: this.req.isSetSupported("Mailbox", "1.2"), capabilities: ["insert into compose body"] },
      { requirement: "Send mail", supported: false, capabilities: ["Intentionally not offered — you send drafts yourself"] },
    ];
  }

  private async messageItem(): Promise<ContextItem> {
    const m = await this.d.item();
    const header = [
      `Subject: ${m.subject}`,
      m.from ? `From: ${m.from}` : undefined,
      m.to.length ? `To: ${m.to.join(", ")}` : undefined,
      m.cc.length ? `Cc: ${m.cc.join(", ")}` : undefined,
      m.sent ? `Sent: ${m.sent}` : undefined,
      m.attachments.length ? `Attachments (not read): ${m.attachments.map((a) => a.name).join(", ")}` : undefined,
    ].filter(Boolean);
    return makeItem({
      host: "outlook",
      id: "outlook:message",
      label: `Message “${m.subject || "(no subject)"}”`,
      scope: "full",
      content: `${header.join("\n")}\n\n${m.body}`,
      unsupported: m.attachments.length ? [{ kind: "attachments", count: m.attachments.length }] : [],
      source: "outlook.read_message",
    });
  }

  async initialContext(): Promise<ContextItem[]> {
    const m = await this.d.item();
    return [
      makeItem({
        host: "outlook",
        id: "outlook:meta",
        label: "Open item",
        scope: "metadata",
        content: `Mode: ${this.d.mode()}\nType: ${m.itemType}\nSubject: ${m.subject}\nFrom: ${m.from ?? "(n/a)"}\nAttachments: ${m.attachments.length}`,
        source: "initial",
      }),
    ];
  }

  async checkFreshness(item: ContextItem): Promise<Freshness> {
    if (item.id !== "outlook:message") return "unknown";
    return (await this.messageItem()).fingerprint === item.fingerprint ? "fresh" : "stale";
  }

  capabilities(): CapabilityDefinition<any>[] {
    const d = this.d;
    const self = this;
    const read: CapabilityDefinition<any> = {
      name: "outlook.read_message",
      description: "Read the open email: headers and body (including quoted earlier messages in the thread). Attachments are listed but not read.",
      hosts: ["outlook"],
      risk: "read_active",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const it = await self.messageItem();
        return { output: { item: it.id }, context: [it] };
      },
    };

    type DraftArgs = { body: string; replyAll?: boolean };
    const draft: CapabilityDefinition<any> = {
      name: "outlook.create_draft",
      description:
        "Create a reply draft with the given plain-text body. In read mode this opens Outlook's reply window (replyAll optional); in compose mode it inserts the text into the message being written. It never sends.",
      hosts: ["outlook"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["body"],
        properties: { body: { type: "string", minLength: 1, maxLength: 30_000 }, replyAll: { type: "boolean" } },
      },
      async resolveTarget(a: DraftArgs) {
        const m = await d.item();
        if (m.itemType !== "message") throw new ScopeError("Drafts can only be created for email messages.");
        return d.mode() === "compose" ? "the message you are writing" : `${a.replyAll ? "reply-all" : "reply"} to “${m.subject}”`;
      },
      async preview(a: DraftArgs) {
        const m = await d.item();
        const recipients = d.mode() === "compose" ? "(unchanged)" : a.replyAll ? [m.from, ...m.to, ...m.cc].filter(Boolean).join(", ") : m.from ?? "(sender)";
        return {
          summary: d.mode() === "compose" ? "Insert draft text into this message" : `Open a ${a.replyAll ? "reply-all" : "reply"} draft (not sent)`,
          target: `Recipients: ${recipients}`,
          after: excerpt(a.body, 4000),
          warnings: ["Nothing will be sent. You review and send the draft yourself in Outlook."],
        };
      },
      async execute(a: DraftArgs) {
        if (d.mode() === "compose") {
          await d.insertComposeText(a.body);
          const body = await d.composeBody();
          const ok = normalizeText(body).includes(normalizeText(a.body).slice(0, 2000));
          return {
            output: { inserted: a.body.length },
            verification: { verified: ok, method: "re-read compose body", mismatch: ok ? undefined : "The draft body does not contain the inserted text." },
          };
        }
        await d.displayReplyForm(a.body, !!a.replyAll);
        return {
          output: { opened: a.replyAll ? "reply-all form" : "reply form", note: "A draft window was opened for the user. It has not been sent." },
          verification: {
            verified: null,
            method: "none available",
            evidence: "Outlook does not let add-ins read back a newly opened reply window. Check the draft window.",
          },
        };
      },
    };

    return [read, { ...read, name: "context.read" }, draft];
  }
}

