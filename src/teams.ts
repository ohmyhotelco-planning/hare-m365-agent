import type { AppConfig } from "./config.js";
import { graphGet, graphPost, type GraphPage } from "./graph.js";
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
  bodyPreview?: string;
};

export type ChatMessageSearchSummary = {
  id: string;
  createdDateTime?: string;
  from?: string;
  fromAddress?: string;
  chatId?: string;
  chatTopic?: string;
  bodyPreview?: string;
  webLink?: string;
};

export type ChatMessageSearchResult = {
  search: {
    query: string;
    range: SearchRange;
    totalMatchesReported: number;
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
  };
  messages: ChatMessageSearchSummary[];
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
  body?: { content?: string };
};

type GraphChatMessageSearchHit = {
  summary?: string;
  resource?: {
    id?: string;
    createdDateTime?: string;
    from?: { emailAddress?: { name?: string; address?: string } };
    chatId?: string;
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

export async function listJoinedTeams(config: AppConfig): Promise<TeamSummary[]> {
  const page = await graphGet<GraphPage<GraphTeam>>(config, "/me/joinedTeams");
  return (page.value ?? []).map((team) => ({
    id: team.id,
    displayName: team.displayName,
    description: team.description
  }));
}

export async function listChats(config: AppConfig, limit: number): Promise<ChatSummary[]> {
  const top = Math.min(limit, config.policy.maxTeamsFetchLimit);
  const page = await graphGet<GraphPage<GraphChat>>(
    config,
    `/me/chats?$top=${top}&$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime%20desc`
  );
  return (page.value ?? [])
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
  limit: number
): Promise<ChatMessageSummary[]> {
  const top = Math.min(limit, config.policy.maxTeamsFetchLimit);
  const page = await graphGet<GraphPage<GraphChatMessage>>(
    config,
    `/chats/${encodeURIComponent(chatId)}/messages?$top=${top}`
  );
  return (page.value ?? []).map((message) => ({
    id: message.id,
    createdDateTime: message.createdDateTime,
    from: message.from?.user?.displayName ?? message.from?.user?.userIdentityType,
    importance: message.importance,
    subject: message.subject,
    bodyPreview: stripHtml(message.body?.content ?? "").slice(0, 500)
  }));
}

export async function searchChatMessages(
  config: AppConfig,
  query: string,
  since: string | undefined,
  until: string | undefined,
  requestedLimit: number
): Promise<ChatMessageSearchResult> {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) {
    throw new Error("query must not be empty.");
  }

  const range = resolveSearchRange(since, until, config.policy.defaultSearchLookbackDays);
  const maxResults = normalizeSearchLimit(requestedLimit, config.policy.maxSearchResults);
  const graphQuery = `${trimmedQuery} AND sent>=${range.since} AND sent<=${range.until}`;
  const hits: GraphChatMessageSearchHit[] = [];
  let totalMatchesReported = 0;
  let moreResultsAvailable = true;
  let offset = 0;

  while (moreResultsAvailable && hits.length < maxResults) {
    const size = Math.min(25, maxResults - hits.length);
    const response = await graphPost<GraphSearchResponse>(config, "/search/query", {
      requests: [
        {
          entityTypes: ["chatMessage"],
          query: { queryString: graphQuery },
          from: offset,
          size
        }
      ]
    });
    const container = response.value?.[0]?.hitsContainers?.[0];
    const pageHits = container?.hits ?? [];
    totalMatchesReported = container?.total ?? totalMatchesReported;
    hits.push(...pageHits);
    offset += pageHits.length;
    moreResultsAvailable = Boolean(container?.moreResultsAvailable) && pageHits.length > 0;
  }

  const chats = await listChats(config, config.policy.maxTeamsFetchLimit);
  const topicByChatId = new Map(chats.map((chat) => [chat.id, chat.topic]));
  const messages = hits.map((hit) => {
    const resource = hit.resource;
    return {
      id: resource?.id ?? "",
      createdDateTime: resource?.createdDateTime,
      from: resource?.from?.emailAddress?.name,
      fromAddress: resource?.from?.emailAddress?.address,
      chatId: resource?.chatId,
      chatTopic: resource?.chatId ? topicByChatId.get(resource.chatId) : undefined,
      bodyPreview: hit.summary,
      webLink: resource?.webLink
    };
  });

  return {
    search: {
      query: trimmedQuery,
      range,
      totalMatchesReported,
      returnedCount: messages.length,
      maxResults,
      limitReached: moreResultsAvailable || totalMatchesReported > messages.length
    },
    messages
  };
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
