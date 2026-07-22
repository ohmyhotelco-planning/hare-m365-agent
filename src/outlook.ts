import type { AppConfig } from "./config.js";
import { decodeGraphContinuation, encodeGraphContinuation } from "./continuation.js";
import { graphGet, type GraphPage, type GraphRequestOptions } from "./graph.js";
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
  body?: string;
  bodyHtml?: string;
  bodyContentType?: string;
  fullBodyAvailable: boolean;
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
  body?: { contentType?: string; content?: string };
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
    continuationAvailable: boolean;
    nextCursor?: string;
    partialResult: boolean;
    partialReason?: "time-budget-exceeded";
    timeBudgetMs: number;
    fullBodyReturnedCount: number;
    fullBodyUnavailableCount: number;
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

const mailSearchSelect = `${mailSelect},body`;

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
    complete: boolean;
    continuationAvailable: boolean;
    nextCursor?: string;
    partialResult: boolean;
    partialReason?: "time-budget-exceeded";
    timeBudgetMs: number;
    continuationMode: "cumulative";
  };
  breakdownBySender: Array<{ sender: string; count: number }>;
};

export type MailSearchOptions = {
  cursor?: string;
  timeBudgetMs?: number;
  now?: () => number;
};

export type MailCountOptions = MailSearchOptions;

type MailSearchCursorState = {
  version: 1;
  nextUrl: string;
  criteriaKey: string;
};

type MailCountCursorState = {
  version: 1;
  nextUrl: string;
  criteriaKey: string;
  scannedCount: number;
  matchedCount: number;
  earliestDateTime?: string;
  latestDateTime?: string;
  senderCounts: Array<[string, number]>;
};

export type OutlookGraphClient = {
  get<T>(pathOrUrl: string, options?: GraphRequestOptions): Promise<T>;
};

const defaultSearchTimeBudgetMs = 35_000;

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
  requestedLimit: number,
  options: MailSearchOptions = {},
  client: OutlookGraphClient = defaultOutlookClient(config)
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
  const maxResults = normalizePageLimit(requestedLimit, config.policy.maxSearchResults);
  const folderPath = getFolderPath(folderScope);
  const dateProperty = getDateProperty(folderScope);
  const kqlDateProperty = dateProperty === "sentDateTime" ? "sent" : "received";
  const kql = `${trimmedQuery} AND ${kqlDateProperty}>=${range.since} AND ${kqlDateProperty}<=${range.until}`;
  const criteriaKey = JSON.stringify({
    query: trimmedQuery,
    folderScope,
    startDateTime: range.startDateTime,
    endDateTimeExclusive: range.endDateTimeExclusive
  });
  const params = new URLSearchParams({
    "$top": String(maxResults),
    "$search": `"${escapeSearchValue(kql)}"`,
    "$select": mailSearchSelect
  });

  const initialUrl = options.cursor
    ? decodeMailSearchCursor(options.cursor, criteriaKey).nextUrl
    : `${folderPath}?${params.toString()}`;
  const timeBudgetMs = normalizeTimeBudget(options.timeBudgetMs ?? defaultSearchTimeBudgetMs);
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  let deletedItemsFolderId: string | undefined;
  let nextUrl: string | undefined = initialUrl;
  let partialResult = false;
  let partialReason: MailSearchResult["search"]["partialReason"];
  let messages: GraphMail[] = [];

  try {
    deletedItemsFolderId = folderScope === "all"
      ? await getDeletedItemsFolderId(config, client, {
          totalTimeoutMs: Math.max(1, deadline - now())
        })
      : undefined;
    const page = await client.get<GraphPage<GraphMail>>(initialUrl, {
      totalTimeoutMs: Math.max(1, deadline - now())
    });
    messages = (page.value ?? []).filter(
      (mail) => !deletedItemsFolderId || mail.parentFolderId !== deletedItemsFolderId
    );
    nextUrl = page["@odata.nextLink"];
  } catch (error) {
    if (now() >= deadline || isGraphTimeBudgetError(error)) {
      partialResult = true;
      partialReason = "time-budget-exceeded";
    } else {
      throw error;
    }
  }

  const summaries = messages.map(toMailSummary);
  const fullBodyReturnedCount = summaries.filter((mail) => mail.fullBodyAvailable).length;
  const continuationAvailable = Boolean(nextUrl);
  return {
    search: {
      query: trimmedQuery,
      folderScope,
      range,
      returnedCount: summaries.length,
      maxResults,
      limitReached: continuationAvailable,
      continuationAvailable,
      nextCursor: nextUrl
        ? encodeMailSearchCursor({ version: 1, nextUrl, criteriaKey })
        : undefined,
      partialResult,
      partialReason,
      timeBudgetMs,
      fullBodyReturnedCount,
      fullBodyUnavailableCount: summaries.length - fullBodyReturnedCount
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
  folderScope: MailFolderScope,
  options: MailCountOptions = {},
  client: OutlookGraphClient = defaultOutlookClient(config)
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
  const timeBudgetMs = normalizeTimeBudget(options.timeBudgetMs ?? defaultSearchTimeBudgetMs);
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  let deletedItemsFolderId: string | undefined;
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
  const criteriaKey = JSON.stringify({
    subjectNeedle,
    fromNeedle,
    folderScope,
    startDateTime: range.startDateTime,
    endDateTimeExclusive: range.endDateTimeExclusive,
    dateProperty
  });
  const resumed = options.cursor ? decodeMailCountCursor(options.cursor, criteriaKey) : undefined;
  let nextUrl: string | undefined = resumed?.nextUrl ?? `${folderPath}?${params.toString()}`;
  let scannedCount = resumed?.scannedCount ?? 0;
  let matchedCount = resumed?.matchedCount ?? 0;
  let earliestReceivedDateTime = resumed?.earliestDateTime;
  let latestReceivedDateTime = resumed?.latestDateTime;
  const senderCounts = new Map<string, number>(resumed?.senderCounts ?? []);
  let partialResult = false;
  let partialReason: MailCountResult["count"]["partialReason"];

  try {
    deletedItemsFolderId = folderScope === "all"
      ? await getDeletedItemsFolderId(config, client, {
          totalTimeoutMs: Math.max(1, deadline - now())
        })
      : undefined;
  } catch (error) {
    if (now() >= deadline || isGraphTimeBudgetError(error)) {
      partialResult = true;
      partialReason = "time-budget-exceeded";
    } else {
      throw error;
    }
  }

  while (nextUrl && !partialResult) {
    if (now() >= deadline) {
      partialResult = true;
      partialReason = "time-budget-exceeded";
      break;
    }
    const currentUrl = nextUrl;
    let page: GraphPage<GraphMail>;
    try {
      page = await client.get<GraphPage<GraphMail>>(currentUrl, {
        totalTimeoutMs: Math.max(1, deadline - now())
      });
    } catch (error) {
      if (now() >= deadline || isGraphTimeBudgetError(error)) {
        partialResult = true;
        partialReason = "time-budget-exceeded";
        nextUrl = currentUrl;
        break;
      }
      throw error;
    }
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

  const continuationAvailable = Boolean(nextUrl);

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
      complete: !continuationAvailable,
      continuationAvailable,
      nextCursor: nextUrl
        ? encodeMailCountCursor({
            version: 1,
            nextUrl,
            criteriaKey,
            scannedCount,
            matchedCount,
            earliestDateTime: earliestReceivedDateTime,
            latestDateTime: latestReceivedDateTime,
            senderCounts: [...senderCounts.entries()]
          })
        : undefined,
      partialResult,
      partialReason,
      timeBudgetMs,
      continuationMode: "cumulative"
    },
    breakdownBySender: [...senderCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([sender, count]) => ({ sender, count }))
  };
}

export function toMailSummary(mail: GraphMail): MailSummary {
  const bodyHtml = mail.body?.content;
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
    body: bodyHtml === undefined ? undefined : stripHtml(bodyHtml),
    bodyHtml,
    bodyContentType: mail.body?.contentType,
    fullBodyAvailable: bodyHtml !== undefined,
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

async function getDeletedItemsFolderId(
  config: AppConfig,
  client: OutlookGraphClient = defaultOutlookClient(config),
  options?: GraphRequestOptions
): Promise<string> {
  const folder = await client.get<GraphMailFolder>(
    "/me/mailFolders/deleteditems?$select=id",
    options
  );
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

function normalizePageLimit(requested: number, policyMaximum: number): number {
  if (!Number.isFinite(requested) || requested < 1) {
    throw new Error("limit must be a positive number.");
  }
  return Math.min(Math.floor(requested), policyMaximum, 100);
}

function normalizeTimeBudget(value: number): number {
  if (!Number.isFinite(value) || value < 1) {
    throw new Error("time budget must be a positive number.");
  }
  return Math.floor(value);
}

function defaultOutlookClient(config: AppConfig): OutlookGraphClient {
  return {
    get: <T>(pathOrUrl: string, options?: GraphRequestOptions) =>
      graphGet<T>(config, pathOrUrl, options)
  };
}

function isGraphTimeBudgetError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return error.name === "GraphTimeoutError" || /time budget|timed out/i.test(error.message);
}

function stripHtml(value: string): string {
  return value
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function encodeMailCountCursor(state: MailCountCursorState): string {
  const normalized: MailCountCursorState = {
    ...state,
    nextUrl: decodeGraphContinuation(encodeGraphContinuation(state.nextUrl))
  };
  return Buffer.from(JSON.stringify(normalized), "utf8").toString("base64url");
}

function encodeMailSearchCursor(state: MailSearchCursorState): string {
  const normalized: MailSearchCursorState = {
    ...state,
    nextUrl: decodeGraphContinuation(encodeGraphContinuation(state.nextUrl))
  };
  return Buffer.from(JSON.stringify(normalized), "utf8").toString("base64url");
}

function decodeMailSearchCursor(cursor: string, criteriaKey: string): MailSearchCursorState {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
  } catch {
    throw new Error("search cursor is invalid.");
  }
  if (!isMailSearchCursorState(parsed) || parsed.criteriaKey !== criteriaKey) {
    throw new Error("search cursor does not match the requested search criteria.");
  }
  return {
    ...parsed,
    nextUrl: decodeGraphContinuation(encodeGraphContinuation(parsed.nextUrl))
  };
}

function decodeMailCountCursor(cursor: string, criteriaKey: string): MailCountCursorState {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
  } catch {
    throw new Error("count cursor is invalid.");
  }
  if (!isMailCountCursorState(parsed) || parsed.criteriaKey !== criteriaKey) {
    throw new Error("count cursor does not match the requested search criteria.");
  }
  return {
    ...parsed,
    nextUrl: decodeGraphContinuation(encodeGraphContinuation(parsed.nextUrl))
  };
}

function isMailCountCursorState(value: unknown): value is MailCountCursorState {
  if (!value || typeof value !== "object") return false;
  const state = value as Partial<MailCountCursorState>;
  return state.version === 1 &&
    typeof state.nextUrl === "string" &&
    typeof state.criteriaKey === "string" &&
    Number.isFinite(state.scannedCount) &&
    Number.isFinite(state.matchedCount) &&
    Array.isArray(state.senderCounts) &&
    state.senderCounts.every(
      (entry) => Array.isArray(entry) && typeof entry[0] === "string" && Number.isFinite(entry[1])
    );
}

function isMailSearchCursorState(value: unknown): value is MailSearchCursorState {
  if (!value || typeof value !== "object") return false;
  const state = value as Partial<MailSearchCursorState>;
  return state.version === 1 &&
    typeof state.nextUrl === "string" &&
    typeof state.criteriaKey === "string";
}
