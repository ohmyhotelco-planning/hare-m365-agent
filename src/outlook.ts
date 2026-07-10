import type { AppConfig } from "./config.js";
import { graphGet, type GraphPage } from "./graph.js";
import { resolveSearchRange, type SearchRange } from "./search-range.js";

export type MailSummary = {
  id: string;
  receivedDateTime?: string;
  sentDateTime?: string;
  from?: string;
  to?: string[];
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
  sentDateTime?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  subject?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  bodyPreview?: string;
  webLink?: string;
};

export type MailFolderScope = "all" | "inbox" | "sent";

export type MailSearchResult = {
  search: {
    query: string;
    folderScope: MailFolderScope;
    range: SearchRange;
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
  };
  messages: MailSummary[];
};

export type MailCountResult = {
  count: {
    subjectContains?: string;
    fromContains?: string;
    folderScope: MailFolderScope;
    range: SearchRange;
    scannedCount: number;
    matchedCount: number;
    earliestReceivedDateTime?: string;
    latestReceivedDateTime?: string;
    complete: true;
  };
  breakdownBySender: Array<{ sender: string; count: number }>;
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
    sentDateTime: mail.sentDateTime,
    from: mail.from?.emailAddress?.address ?? mail.from?.emailAddress?.name,
    to: formatRecipients(mail.toRecipients),
    subject: mail.subject,
    isRead: mail.isRead,
    hasAttachments: mail.hasAttachments,
    importance: mail.importance,
    bodyPreview: mail.bodyPreview,
    webLink: mail.webLink
  }));
}

export async function searchMailbox(
  config: AppConfig,
  query: string,
  since: string | undefined,
  until: string | undefined,
  folderScope: MailFolderScope,
  requestedLimit: number
): Promise<MailSearchResult> {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) {
    throw new Error("query must not be empty.");
  }

  const range = resolveSearchRange(since, until, config.policy.defaultSearchLookbackDays);
  const maxResults = normalizeSearchLimit(requestedLimit, config.policy.maxSearchResults);
  const select = [
    "id",
    "receivedDateTime",
    "sentDateTime",
    "from",
    "toRecipients",
    "subject",
    "isRead",
    "hasAttachments",
    "importance",
    "bodyPreview",
    "webLink"
  ].join(",");
  const folderPath = getFolderPath(folderScope);
  const kql = `${trimmedQuery} AND sent>=${range.since} AND sent<=${range.until}`;
  const params = new URLSearchParams({
    "$top": String(Math.min(100, maxResults)),
    "$search": `"${escapeSearchValue(kql)}"`,
    "$select": select
  });

  let nextUrl: string | undefined = `${folderPath}?${params.toString()}`;
  let hasMore = false;
  const messages: GraphMail[] = [];

  while (nextUrl && messages.length < maxResults) {
    const page: GraphPage<GraphMail> = await graphGet<GraphPage<GraphMail>>(config, nextUrl);
    const remaining = maxResults - messages.length;
    messages.push(...(page.value ?? []).slice(0, remaining));
    nextUrl = page["@odata.nextLink"];
    hasMore = Boolean(nextUrl);
  }

  const summaries = messages.map(toMailSummary);
  return {
    search: {
      query: trimmedQuery,
      folderScope,
      range,
      returnedCount: summaries.length,
      maxResults,
      limitReached: hasMore
    },
    messages: summaries
  };
}

export async function countMailboxMessages(
  config: AppConfig,
  subjectContains: string | undefined,
  fromContains: string | undefined,
  since: string | undefined,
  until: string | undefined,
  folderScope: MailFolderScope
): Promise<MailCountResult> {
  const subjectNeedle = subjectContains?.trim().toLocaleLowerCase();
  const fromNeedle = fromContains?.trim().toLocaleLowerCase();
  const range = resolveSearchRange(since, until, config.policy.defaultSearchLookbackDays);
  const folderPath = getFolderPath(folderScope);
  const filter = [
    `receivedDateTime ge ${range.since}T00:00:00Z`,
    `receivedDateTime lt ${nextDate(range.until)}T00:00:00Z`
  ].join(" and ");
  const params = new URLSearchParams({
    "$top": "500",
    "$filter": filter,
    "$orderby": "receivedDateTime desc",
    "$select": "id,subject,receivedDateTime,from"
  });

  let nextUrl: string | undefined = `${folderPath}?${params.toString()}`;
  let scannedCount = 0;
  let matchedCount = 0;
  let earliestReceivedDateTime: string | undefined;
  let latestReceivedDateTime: string | undefined;
  const senderCounts = new Map<string, number>();

  while (nextUrl) {
    const page: GraphPage<GraphMail> = await graphGet<GraphPage<GraphMail>>(config, nextUrl);
    for (const mail of page.value ?? []) {
      scannedCount += 1;
      const subject = mail.subject?.toLocaleLowerCase() ?? "";
      const senderAddress = mail.from?.emailAddress?.address ?? "";
      const senderName = mail.from?.emailAddress?.name ?? "";
      const senderText = `${senderName} ${senderAddress}`.toLocaleLowerCase();
      if (subjectNeedle && !subject.includes(subjectNeedle)) continue;
      if (fromNeedle && !senderText.includes(fromNeedle)) continue;

      matchedCount += 1;
      const received = mail.receivedDateTime;
      if (received && (!earliestReceivedDateTime || received < earliestReceivedDateTime)) {
        earliestReceivedDateTime = received;
      }
      if (received && (!latestReceivedDateTime || received > latestReceivedDateTime)) {
        latestReceivedDateTime = received;
      }
      const sender = senderAddress || senderName || "(unknown)";
      senderCounts.set(sender, (senderCounts.get(sender) ?? 0) + 1);
    }
    nextUrl = page["@odata.nextLink"];
  }

  return {
    count: {
      subjectContains: subjectContains?.trim() || undefined,
      fromContains: fromContains?.trim() || undefined,
      folderScope,
      range,
      scannedCount,
      matchedCount,
      earliestReceivedDateTime,
      latestReceivedDateTime,
      complete: true
    },
    breakdownBySender: [...senderCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([sender, count]) => ({ sender, count }))
  };
}

function toMailSummary(mail: GraphMail): MailSummary {
  return {
    id: mail.id,
    receivedDateTime: mail.receivedDateTime,
    sentDateTime: mail.sentDateTime,
    from: mail.from?.emailAddress?.address ?? mail.from?.emailAddress?.name,
    to: formatRecipients(mail.toRecipients),
    subject: mail.subject,
    isRead: mail.isRead,
    hasAttachments: mail.hasAttachments,
    importance: mail.importance,
    bodyPreview: mail.bodyPreview,
    webLink: mail.webLink
  };
}

function formatRecipients(recipients: GraphMail["toRecipients"]): string[] | undefined {
  if (!recipients) return undefined;
  return recipients
    .map((recipient) => recipient.emailAddress?.address ?? recipient.emailAddress?.name)
    .filter((value): value is string => Boolean(value));
}

function getFolderPath(folderScope: MailFolderScope): string {
  if (folderScope === "inbox") return "/me/mailFolders/inbox/messages";
  if (folderScope === "sent") return "/me/mailFolders/sentitems/messages";
  return "/me/messages";
}

function escapeSearchValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function normalizeSearchLimit(requested: number, policyMaximum: number): number {
  if (!Number.isFinite(requested) || requested < 1) {
    throw new Error("limit must be a positive number.");
  }
  return Math.min(Math.floor(requested), policyMaximum, 1000);
}

function nextDate(value: string): string {
  const date = new Date(`${value}T00:00:00Z`);
  return new Date(date.getTime() + 86_400_000).toISOString().slice(0, 10);
}
