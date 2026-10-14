import type { MailItem, OutlookDriver } from "@loai/core";

function asyncCall<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error?.message ?? "Outlook operation failed")))),
  );
}

function fmt(a?: Office.EmailAddressDetails): string {
  if (!a) return "";
  return a.displayName && a.displayName !== a.emailAddress ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress;
}

function escapeHtml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Outlook driver over Office.js (Mailbox 1.1+). It cannot send mail. */
export class OfficeOutlookDriver implements OutlookDriver {
  private get current(): any {
    return Office.context.mailbox.item;
  }

  mode(): "read" | "compose" {
    // In compose mode, subject is an object with getAsync; in read mode a string.
    return typeof this.current?.subject === "string" ? "read" : "compose";
  }

  async item(): Promise<MailItem> {
    const it = this.current;
    if (!it) throw new Error("No Outlook item is open.");
    const body = await asyncCall<string>((cb) => it.body.getAsync(Office.CoercionType.Text, cb));
    if (this.mode() === "read") {
      return {
        itemId: it.itemId,
        conversationId: it.conversationId,
        itemType: it.itemType === Office.MailboxEnums.ItemType.Message ? "message" : it.itemType === Office.MailboxEnums.ItemType.Appointment ? "appointment" : "other",
        subject: it.subject ?? "",
        from: fmt(it.from),
        to: (it.to ?? []).map(fmt),
        cc: (it.cc ?? []).map(fmt),
        sent: it.dateTimeCreated ? new Date(it.dateTimeCreated).toLocaleString() : undefined,
        body,
        attachments: (it.attachments ?? []).map((a: Office.AttachmentDetails) => ({ name: a.name, size: a.size })),
      };
    }
    const subject = await asyncCall<string>((cb) => it.subject.getAsync(cb));
    const to = await asyncCall<Office.EmailAddressDetails[]>((cb) => it.to.getAsync(cb)).catch(() => []);
    const cc = await asyncCall<Office.EmailAddressDetails[]>((cb) => it.cc.getAsync(cb)).catch(() => []);
    return {
      conversationId: it.conversationId,
      itemType: it.itemType === Office.MailboxEnums.ItemType.Message ? "message" : "other",
      subject,
      to: to.map(fmt),
      cc: cc.map(fmt),
      body,
      attachments: [],
    };
  }

  async composeBody() {
    return asyncCall<string>((cb) => this.current.body.getAsync(Office.CoercionType.Text, cb));
  }

  async insertComposeText(text: string) {
    await asyncCall<void>((cb) => this.current.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb));
  }

  async displayReplyForm(body: string, replyAll: boolean) {
    const html = escapeHtml(body).replace(/\r?\n/g, "<br>");
    if (replyAll) this.current.displayReplyAllForm(html);
    else this.current.displayReplyForm(html);
  }
}
