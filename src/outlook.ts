import type { AppConfig } from "./config.js";
import { graphGet, type GraphPage } from "./graph.js";

export type MailSummary = {
  id: string;
  receivedDateTime?: string;
  from?: string;
  subject?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  bodyPreview?: string;
  webLink?: string;
};

type GraphMail = {
  id: string;
  receivedDateTime?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  subject?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  bodyPreview?: string;
  webLink?: string;
};

export async function listInbox(config: AppConfig, limit: number): Promise<MailSummary[]> {
  const top = Math.min(limit, config.policy.maxMailFetchLimit);
  const select = [
    "id",
    "receivedDateTime",
    "from",
    "subject",
    "isRead",
    "hasAttachments",
    "importance",
    "bodyPreview",
    "webLink"
  ].join(",");
  const page = await graphGet<GraphPage<GraphMail>>(
    config,
    `/me/mailFolders/Inbox/messages?$top=${top}&$orderby=receivedDateTime desc&$select=${select}`
  );

  return (page.value ?? []).map((mail) => ({
    id: mail.id,
    receivedDateTime: mail.receivedDateTime,
    from: mail.from?.emailAddress?.address ?? mail.from?.emailAddress?.name,
    subject: mail.subject,
    isRead: mail.isRead,
    hasAttachments: mail.hasAttachments,
    importance: mail.importance,
    bodyPreview: mail.bodyPreview,
    webLink: mail.webLink
  }));
}
