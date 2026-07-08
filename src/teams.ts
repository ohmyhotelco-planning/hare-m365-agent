import type { AppConfig } from "./config.js";
import { graphGet, type GraphPage } from "./graph.js";

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

function stripHtml(value: string): string {
  return value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}
