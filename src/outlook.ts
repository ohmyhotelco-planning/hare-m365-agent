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
  flagStatus?: "notFlagged" | "complete" | "flagged";
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
  flag?: { flagStatus?: "notFlagged" | "complete" | "flagged" };
  bodyPreview?: string;
  webLink?: string;
  parentFolderId?: string;
};

type GraphMailFolder = { id: string };

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

export type MailListResult = {
  list: {
    folderScope: MailFolderScope;
    returnedCount: number;
    excludedDeletedItems: boolean;
  };
  messages: MailSummary[];
};

export type FlaggedMailResult = {
  flagged: {
    folderScope: MailFolderScope;
    range: SearchRange;
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
    excludedDeletedItems: boolean;
  };
  messages: MailSummary[];
};

const mailSelect = [
  "id",
  "receivedDateTime",
  "sentDateTime",
  "from",
  "toRecipients",
  "subject",
  "isRead",
  "hasAttachments",
  "importance",
  "flag",
  "bodyPreview",
  "webLink",
  "parentFolderId"
].join(",");

export type MailCountResult = {
  count: {
    subjectContains?: string;
    fromContains?: string;
    folderScope: MailFolderScope;
    range: SearchRange;
    scannedCount: number;
    matchedCount: number;
    earliestDateTime?: string;
    latestDateTime?: string;
    earliestReceivedDateTime?: string;
    latestReceivedDateTime?: string;
    dateProperty: "receivedDateTime" | "sentDateTime";
    excludedDeletedItems: boolean;
    complete: true;
  };
  breakdownBySender: Array<{ sender: string; count: number }>;
};

export async function listInbox(config: AppConfig, limit: number): Promise<MailSummary[]> {
  const top = normalizeSearchLimit(limit, config.policy.maxMailFetchLimit);
  const page = await graphGet<GraphPage<GraphMail>>(
    config,
    buildRecentMessagesPath("inbox", top)
  );

  return (page.value ?? []).map(toMailSummary);
}

export async function listRecentMailbox(
  config: AppConfig,
  folderScope: MailFolderScope,
  limit: number
): Promise<MailListResult> {
  const maxResults = normalizeSearchLimit(limit, config.policy.maxMailFetchLimit);
  const deletedItemsFolderId = folderScope === "all" ? await getDeletedItemsFolderId(config) : undefined;
  const collected = await collectMessages(
    config,
    buildRecentMessagesPath(folderScope, maxResults),
    maxResults,
    deletedItemsFolderId
  );

  return {
    list: {
      folderScope,
      returnedCount: collected.messages.length,
      excludedDeletedItems: folderScope === "all"
    },
    messages: collected.messages.map(toMailSummary)
  };
}

export async function listFlaggedMessages(
  config: AppConfig,
  since: string | undefined,
  until: string | undefined,
  folderScope: MailFolderScope,
  requestedLimit: number
): Promise<FlaggedMailResult> {
  const range = resolveSearchRange(
    since,
    until,
    config.policy.defaultSearchLookbackDays,
    new Date(),
    config.timeZone
  );
  const maxResults = normalizeSearchLimit(requestedLimit, config.policy.maxSearchResults);
  const deletedItemsFolderId = folderScope === "all" ? await getDeletedItemsFolderId(config) : undefined;
  const collected = await collectMessages(
    config,
    buildFlaggedMessagesPath(folderScope, range, maxResults),
    maxResults,
    deletedItemsFolderId
  );

  return {
    flagged: {
      folderScope,
      range,
      returnedCount: collected.messages.length,
      maxResults,
      limitReached: collected.hasMore,
      excludedDeletedItems: folderScope === "all"
    },
    messages: collected.messages.map(toMailSummary)
  };
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

  const range = resolveSearchRange(
    since,
    until,
    config.policy.defaultSearchLookbackDays,
    new Date(),
    config.timeZone
  );
  const maxResults = normalizeSearchLimit(requestedLimit, config.policy.maxSearchResults);
  const folderPath = getFolderPath(folderScope);
  const dateProperty = getDateProperty(folderScope);
  const kqlDateProperty = dateProperty === "sentDateTime" ? "sent" : "received";
  const kql = `${trimmedQuery} AND ${kqlDateProperty}>=${range.since} AND ${kqlDateProperty}<=${range.until}`;
  const deletedItemsFolderId = folderScope === "all" ? await getDeletedItemsFolderId(config) : undefined;
  const params = new URLSearchParams({
    "$top": String(Math.min(100, maxResults)),
    "$search": `"${escapeSearchValue(kql)}"`,
    "$select": mailSelect
  });

  let nextUrl: string | undefined = `${folderPath}?${params.toString()}`;
  let hasMore = false;
  const messages: GraphMail[] = [];

  while (nextUrl && messages.length < maxResults) {
    const page: GraphPage<GraphMail> = await graphGet<GraphPage<GraphMail>>(config, nextUrl);
    const eligible = (page.value ?? []).filter(
      (mail) => !deletedItemsFolderId || mail.parentFolderId !== deletedItemsFolderId
    );
    const remaining = maxResults - messages.length;
    messages.push(...eligible.slice(0, remaining));
    nextUrl = page["@odata.nextLink"];
    hasMore = Boolean(nextUrl) || eligible.length > remaining;
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
  const range = resolveSearchRange(
    since,
    until,
    config.policy.defaultSearchLookbackDays,
    new Date(),
    config.timeZone
  );
  const folderPath = getFolderPath(folderScope);
  const dateProperty = getDateProperty(folderScope);
  const deletedItemsFolderId = folderScope === "all" ? await getDeletedItemsFolderId(config) : undefined;
  const filter = [
    `${dateProperty} ge ${range.startDateTime}`,
    `${dateProperty} lt ${range.endDateTimeExclusive}`
  ].join(" and ");
  const params = new URLSearchParams({
    "$top": "500",
    "$filter": filter,
    "$orderby": `${dateProperty} desc`,
    "$select": "id,subject,receivedDateTime,sentDateTime,from,parentFolderId"
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
      if (deletedItemsFolderId && mail.parentFolderId === deletedItemsFolderId) continue;
      scannedCount += 1;
      const subject = mail.subject?.toLocaleLowerCase() ?? "";
      const senderAddress = mail.from?.emailAddress?.address ?? "";
      const senderName = mail.from?.emailAddress?.name ?? "";
      const senderText = `${senderName} ${senderAddress}`.toLocaleLowerCase();
      if (subjectNeedle && !subject.includes(subjectNeedle)) continue;
      if (fromNeedle && !senderText.includes(fromNeedle)) continue;

      matchedCount += 1;
      const messageDate = mail[dateProperty];
      if (messageDate && (!earliestReceivedDateTime || messageDate < earliestReceivedDateTime)) {
        earliestReceivedDateTime = messageDate;
      }
      if (messageDate && (!latestReceivedDateTime || messageDate > latestReceivedDateTime)) {
        latestReceivedDateTime = messageDate;
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
      earliestDateTime: earliestReceivedDateTime,
      latestDateTime: latestReceivedDateTime,
      earliestReceivedDateTime,
      latestReceivedDateTime,
      dateProperty,
      excludedDeletedItems: folderScope === "all",
      complete: true
    },
    breakdownBySender: [...senderCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([sender, count]) => ({ sender, count }))
  };
}

export function toMailSummary(mail: GraphMail): MailSummary {
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
    flagStatus: mail.flag?.flagStatus,
    bodyPreview: mail.bodyPreview,
    webLink: mail.webLink
  };
}

export function buildRecentMessagesPath(folderScope: MailFolderScope, limit: number): string {
  const dateProperty = getDateProperty(folderScope);
  const params = new URLSearchParams({
    "$top": String(Math.min(100, limit)),
    "$orderby": `${dateProperty} desc`,
    "$select": mailSelect
  });
  return `${getFolderPath(folderScope)}?${params.toString()}`;
}

export function buildFlaggedMessagesPath(
  folderScope: MailFolderScope,
  range: SearchRange,
  limit: number
): string {
  const dateProperty = getDateProperty(folderScope);
  const params = new URLSearchParams({
    "$top": String(Math.min(100, limit)),
    "$filter": [
      `${dateProperty} ge ${range.startDateTime}`,
      `${dateProperty} lt ${range.endDateTimeExclusive}`,
      "flag/flagStatus eq 'flagged'"
    ].join(" and "),
    "$orderby": `${dateProperty} desc`,
    "$select": mailSelect
  });
  return `${getFolderPath(folderScope)}?${params.toString()}`;
}

async function collectMessages(
  config: AppConfig,
  initialUrl: string,
  maxResults: number,
  deletedItemsFolderId?: string
): Promise<{ messages: GraphMail[]; hasMore: boolean }> {
  let nextUrl: string | undefined = initialUrl;
  let hasMore = false;
  const messages: GraphMail[] = [];

  while (nextUrl && messages.length < maxResults) {
    const page: GraphPage<GraphMail> = await graphGet<GraphPage<GraphMail>>(config, nextUrl);
    const eligible = (page.value ?? []).filter(
      (mail) => !deletedItemsFolderId || mail.parentFolderId !== deletedItemsFolderId
    );
    const remaining = maxResults - messages.length;
    messages.push(...eligible.slice(0, remaining));
    nextUrl = page["@odata.nextLink"];
    hasMore = Boolean(nextUrl) || eligible.length > remaining;
  }

  return { messages, hasMore };
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

function getDateProperty(folderScope: MailFolderScope): "receivedDateTime" | "sentDateTime" {
  return folderScope === "sent" ? "sentDateTime" : "receivedDateTime";
}

async function getDeletedItemsFolderId(config: AppConfig): Promise<string> {
  const folder = await graphGet<GraphMailFolder>(config, "/me/mailFolders/deleteditems?$select=id");
  return folder.id;
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
