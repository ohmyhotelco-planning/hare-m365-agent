import type { AppConfig } from "./config.js";
import {
  graphGet,
  graphPost,
  type GraphPage,
  type GraphRequestOptions
} from "./graph.js";
import { resolveSearchRange, type SearchRange } from "./search-range.js";

export type TeamSummary = {
  id: string;
  displayName?: string;
  description?: string;
};

export type ChatSummary = {
  id: string;
  topic?: string;
  chatType?: string;
  lastMessageCreatedDateTime?: string;
  lastMessageFrom?: string;
  lastUpdatedDateTime?: string;
};

export type ChatMessageSummary = {
  id: string;
  createdDateTime?: string;
  from?: string;
  importance?: string;
  subject?: string | null;
  body?: string;
  bodyHtml?: string;
  bodyContentType?: string;
  /** @deprecated Use body. Kept as a full-text compatibility alias. */
  bodyPreview?: string;
};

export type ChatMessageSearchSummary = {
  id: string;
  createdDateTime?: string;
  from?: string;
  fromAddress?: string;
  chatId?: string;
  chatTopic?: string;
  channelIdentity?: { teamId?: string; channelId?: string };
  searchSummary?: string;
  body?: string;
  bodyHtml?: string;
  bodyContentType?: string;
  fullBodyAvailable: boolean;
  bodyUnavailableReason?:
    | "missing-message-location"
    | "detail-request-failed"
    | "message-body-missing"
    | "time-budget-exceeded";
  /** @deprecated Use body. Full text when available; search summary otherwise. */
  bodyPreview?: string;
  webLink?: string;
};

export type ChatMessageSearchResult = {
  search: {
    query: string;
    range: SearchRange;
    totalMatchesReported: number;
    returnedCount: number;
    fullBodyReturnedCount: number;
    fullBodyUnavailableCount: number;
    maxResults: number;
    limitReached: boolean;
    offset: number;
    nextOffset?: number;
    continuationAvailable: boolean;
    partialResult: boolean;
    partialReason?: "time-budget-exceeded";
    timeBudgetMs: number;
  };
  messages: ChatMessageSearchSummary[];
};

export type ChatMessageSearchOptions = {
  offset?: number;
  timeBudgetMs?: number;
  now?: () => number;
};

type GraphTeam = {
  id: string;
  displayName?: string;
  description?: string;
};

type GraphChat = {
  id: string;
  topic?: string;
  chatType?: string;
  lastUpdatedDateTime?: string;
  lastMessagePreview?: {
    createdDateTime?: string;
    from?: { user?: { displayName?: string; userIdentityType?: string } };
  };
};

type GraphChatMessage = {
  id: string;
  createdDateTime?: string;
  from?: { user?: { displayName?: string; userIdentityType?: string } };
  importance?: string;
  subject?: string | null;
  body?: { contentType?: string; content?: string };
};

type GraphChatMessageSearchHit = {
  summary?: string;
  resource?: {
    id?: string;
    createdDateTime?: string;
    from?: { emailAddress?: { name?: string; address?: string } };
    chatId?: string;
    channelIdentity?: { teamId?: string; channelId?: string };
    webLink?: string;
  };
};

type GraphSearchResponse = {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: GraphChatMessageSearchHit[];
      total?: number;
      moreResultsAvailable?: boolean;
    }>;
  }>;
};

export type TeamsGraphClient = {
  get<T>(pathOrUrl: string, options?: GraphRequestOptions): Promise<T>;
  post<T>(pathOrUrl: string, body?: unknown, options?: GraphRequestOptions): Promise<T>;
};

const fullBodyConcurrency = 5;
const defaultSearchTimeBudgetMs = 35_000;

export async function listJoinedTeams(
  config: AppConfig,
  client: TeamsGraphClient = defaultTeamsClient(config)
): Promise<TeamSummary[]> {
  const teams = await collectPages<GraphTeam>(client, "/me/joinedTeams");
  return teams.map((team) => ({
    id: team.id,
    displayName: team.displayName,
    description: team.description
  }));
}

export async function listChats(
  config: AppConfig,
  limit: number,
  client: TeamsGraphClient = defaultTeamsClient(config)
): Promise<ChatSummary[]> {
  const top = normalizeSearchLimit(limit, config.policy.maxTeamsFetchLimit);
  const chats = await fetchAllChats(config, client);
  return chats
    .sort((a, b) => (b.lastMessagePreview?.createdDateTime ?? "").localeCompare(a.lastMessagePreview?.createdDateTime ?? ""))
    .slice(0, top)
    .map((chat) => ({
      id: chat.id,
      topic: chat.topic,
      chatType: chat.chatType,
      lastMessageCreatedDateTime: chat.lastMessagePreview?.createdDateTime,
      lastMessageFrom:
        chat.lastMessagePreview?.from?.user?.displayName ??
        chat.lastMessagePreview?.from?.user?.userIdentityType,
      lastUpdatedDateTime: chat.lastUpdatedDateTime
    }));
}

export async function listChatMessages(
  config: AppConfig,
  chatId: string,
  limit: number,
  client: TeamsGraphClient = defaultTeamsClient(config)
): Promise<ChatMessageSummary[]> {
  if (!chatId.trim()) throw new Error("chat-id must not be empty.");
  const top = normalizeSearchLimit(limit, config.policy.maxTeamsFetchLimit);
  const messages = await collectPages<GraphChatMessage>(
    client,
    `/chats/${encodeURIComponent(chatId)}/messages?$top=${Math.min(top, 50)}&$orderby=createdDateTime%20desc`,
    top
  );
  return messages.map(toChatMessageSummary);
}

export async function searchChatMessages(
  config: AppConfig,
  query: string,
  since: string | undefined,
  until: string | undefined,
  requestedLimit: number,
  options: ChatMessageSearchOptions = {},
  client: TeamsGraphClient = defaultTeamsClient(config)
): Promise<ChatMessageSearchResult> {
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
  const initialOffset = normalizeSearchOffset(options.offset ?? 0);
  const timeBudgetMs = normalizeTimeBudget(options.timeBudgetMs ?? defaultSearchTimeBudgetMs);
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  const graphQuery = `${trimmedQuery} AND sent>=${range.since} AND sent<=${range.until}`;
  const hits: GraphChatMessageSearchHit[] = [];
  let totalMatchesReported = 0;
  let moreResultsAvailable = true;
  let offset = initialOffset;
  let partialResult = false;
  let partialReason: ChatMessageSearchResult["search"]["partialReason"];

  const hasTimeRemaining = () => now() < deadline;
  const markTimeBudgetExceeded = () => {
    partialResult = true;
    partialReason = "time-budget-exceeded";
  };
  const graphRequestOptions = (): GraphRequestOptions => ({
    totalTimeoutMs: Math.max(1, deadline - now())
  });

  while (moreResultsAvailable && hits.length < maxResults) {
    if (!hasTimeRemaining()) {
      markTimeBudgetExceeded();
      break;
    }
    const size = Math.min(25, maxResults - hits.length);
    let response: GraphSearchResponse;
    try {
      response = await client.post<GraphSearchResponse>(
        "/search/query",
        {
          requests: [
            {
              entityTypes: ["chatMessage"],
              query: { queryString: graphQuery },
              from: offset,
              size
            }
          ]
        },
        graphRequestOptions()
      );
    } catch (error) {
      if (!hasTimeRemaining() || isGraphTimeBudgetError(error)) {
        markTimeBudgetExceeded();
        break;
      }
      throw error;
    }
    const container = response.value?.[0]?.hitsContainers?.[0];
    const pageHits = container?.hits ?? [];
    totalMatchesReported = container?.total ?? totalMatchesReported;
    hits.push(...pageHits);
    offset += pageHits.length;
    moreResultsAvailable = Boolean(container?.moreResultsAvailable) && pageHits.length > 0;
  }

  const details = await mapWithConcurrency(hits, fullBodyConcurrency, async (hit) => {
    const resource = hit.resource;
    const detailPath = messageDetailPath(resource);
    let detail: GraphChatMessage | undefined;
    let bodyUnavailableReason: ChatMessageSearchSummary["bodyUnavailableReason"];
    if (!hasTimeRemaining()) {
      markTimeBudgetExceeded();
      bodyUnavailableReason = "time-budget-exceeded";
    } else if (!detailPath) {
      bodyUnavailableReason = "missing-message-location";
    } else {
      try {
        detail = await client.get<GraphChatMessage>(detailPath, graphRequestOptions());
        if (detail.body?.content === undefined) {
          bodyUnavailableReason = "message-body-missing";
        }
      } catch (error) {
        if (!hasTimeRemaining() || isGraphTimeBudgetError(error)) {
          markTimeBudgetExceeded();
          bodyUnavailableReason = "time-budget-exceeded";
        } else {
          bodyUnavailableReason = "detail-request-failed";
        }
      }
    }
    return { detail, bodyUnavailableReason };
  });

  const uniqueChatIds = Array.from(new Set(
    hits.map((hit) => hit.resource?.chatId).filter((chatId): chatId is string => Boolean(chatId))
  ));
  const topicEntries = await mapWithConcurrency(uniqueChatIds, fullBodyConcurrency, async (chatId) => {
    if (!hasTimeRemaining()) {
      markTimeBudgetExceeded();
      return [chatId, undefined] as const;
    }
    try {
      const chat = await client.get<GraphChat>(
        `/chats/${encodeURIComponent(chatId)}?$select=id,topic`,
        graphRequestOptions()
      );
      return [chatId, chat.topic] as const;
    } catch (error) {
      if (!hasTimeRemaining() || isGraphTimeBudgetError(error)) markTimeBudgetExceeded();
      return [chatId, undefined] as const;
    }
  });
  const topicByChatId = new Map<string, string | undefined>(topicEntries);

  const messages = hits.map((hit, index) => {
    const resource = hit.resource;
    const { detail, bodyUnavailableReason } = details[index];
    const bodyHtml = detail?.body?.content;
    const body = bodyHtml === undefined ? undefined : stripHtml(bodyHtml);
    const searchSummary = stripHtml(hit.summary ?? "");
    return {
      id: resource?.id ?? "",
      createdDateTime: resource?.createdDateTime,
      from: resource?.from?.emailAddress?.name,
      fromAddress: resource?.from?.emailAddress?.address,
      chatId: resource?.chatId,
      chatTopic: resource?.chatId ? topicByChatId.get(resource.chatId) : undefined,
      channelIdentity: resource?.channelIdentity,
      searchSummary,
      body,
      bodyHtml,
      bodyContentType: detail?.body?.contentType,
      fullBodyAvailable: bodyHtml !== undefined,
      bodyUnavailableReason,
      bodyPreview: body ?? searchSummary,
      webLink: resource?.webLink
    };
  });

  const fullBodyReturnedCount = messages.filter((message) => message.fullBodyAvailable).length;
  const continuationAvailable =
    moreResultsAvailable || totalMatchesReported > initialOffset + hits.length;

  return {
    search: {
      query: trimmedQuery,
      range,
      totalMatchesReported,
      returnedCount: messages.length,
      fullBodyReturnedCount,
      fullBodyUnavailableCount: messages.length - fullBodyReturnedCount,
      maxResults,
      limitReached: continuationAvailable,
      offset: initialOffset,
      nextOffset: continuationAvailable ? initialOffset + hits.length : undefined,
      continuationAvailable,
      partialResult,
      partialReason,
      timeBudgetMs
    },
    messages
  };
}

async function fetchAllChats(
  config: AppConfig,
  client: TeamsGraphClient = defaultTeamsClient(config)
): Promise<GraphChat[]> {
  return collectPages<GraphChat>(client, "/me/chats?$top=50&$expand=lastMessagePreview");
}

async function collectPages<T>(
  client: TeamsGraphClient,
  initialUrl: string,
  maximum = Number.POSITIVE_INFINITY
): Promise<T[]> {
  const values: T[] = [];
  let nextUrl: string | undefined = initialUrl;

  while (nextUrl && values.length < maximum) {
    const page: GraphPage<T> = await client.get<GraphPage<T>>(nextUrl);
    const remaining = maximum - values.length;
    values.push(...(page.value ?? []).slice(0, remaining));
    nextUrl = page["@odata.nextLink"];
  }

  return values;
}

function defaultTeamsClient(config: AppConfig): TeamsGraphClient {
  return {
    get: <T>(pathOrUrl: string, options?: GraphRequestOptions) =>
      graphGet<T>(config, pathOrUrl, options),
    post: <T>(pathOrUrl: string, body?: unknown, options?: GraphRequestOptions) =>
      graphPost<T>(config, pathOrUrl, body, options)
  };
}

function toChatMessageSummary(message: GraphChatMessage): ChatMessageSummary {
  const bodyHtml = message.body?.content ?? "";
  const body = stripHtml(bodyHtml);
  return {
    id: message.id,
    createdDateTime: message.createdDateTime,
    from: message.from?.user?.displayName ?? message.from?.user?.userIdentityType,
    importance: message.importance,
    subject: message.subject,
    body,
    bodyHtml,
    bodyContentType: message.body?.contentType,
    bodyPreview: body
  };
}

function messageDetailPath(
  resource: GraphChatMessageSearchHit["resource"]
): string | undefined {
  if (!resource?.id) return undefined;
  if (resource.chatId) {
    return `/chats/${encodeURIComponent(resource.chatId)}/messages/${encodeURIComponent(resource.id)}`;
  }
  const teamId = resource.channelIdentity?.teamId;
  const channelId = resource.channelIdentity?.channelId;
  if (teamId && channelId) {
    return `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(resource.id)}`;
  }
  return undefined;
}

async function mapWithConcurrency<T, R>(
  values: T[],
  concurrency: number,
  mapper: (value: T, index: number) => Promise<R>
): Promise<R[]> {
  const results = new Array<R>(values.length);
  let nextIndex = 0;

  async function worker(): Promise<void> {
    while (true) {
      const index = nextIndex;
      nextIndex += 1;
      if (index >= values.length) return;
      results[index] = await mapper(values[index], index);
    }
  }

  const workerCount = Math.min(Math.max(1, concurrency), values.length);
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  return results;
}

function stripHtml(value: string): string {
  return value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeSearchLimit(requested: number, policyMaximum: number): number {
  if (!Number.isFinite(requested) || requested < 1) {
    throw new Error("limit must be a positive number.");
  }
  return Math.min(Math.floor(requested), policyMaximum, 1000);
}

function normalizeSearchOffset(requested: number): number {
  if (!Number.isFinite(requested) || requested < 0) {
    throw new Error("offset must be a non-negative number.");
  }
  return Math.floor(requested);
}

function normalizeTimeBudget(requested: number): number {
  if (!Number.isFinite(requested) || requested < 1) {
    throw new Error("time budget must be a positive number.");
  }
  return Math.floor(requested);
}

function isGraphTimeBudgetError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return error.name === "GraphTimeoutError" || /time budget|timed out/i.test(error.message);
}
