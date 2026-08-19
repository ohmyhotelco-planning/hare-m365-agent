import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { writeStoredText } from "./persistent-storage.js";

const managedClaudeRulesStart = "<!-- HARE_M365_AGENT_RULES_START -->";
const managedClaudeRulesEnd = "<!-- HARE_M365_AGENT_RULES_END -->";

export type SessionRulesOptions = {
  commandPrefix: string;
  repository: string;
  branch: string;
  workDir: string;
  requiredDomains: string[];
};

export function sessionRulesPath(config: AppConfig): string {
  return path.join(config.dataDir, "claude", "hare-m365-agent-rules.md");
}

export function writeSessionRules(
  config: AppConfig,
  options: SessionRulesOptions
): string | undefined {
  if (!config.dataDirPersistent) return undefined;

  const rulesFile = sessionRulesPath(config);
  const contents = buildSessionRules(config, options, rulesFile);
  fs.mkdirSync(path.dirname(rulesFile), { recursive: true });

  if (!fs.existsSync(rulesFile) || fs.readFileSync(rulesFile, "utf8") !== contents) {
    writeStoredText(rulesFile, contents);
  }

  writeClaudeProjectInstructions(config, rulesFile);
  return rulesFile;
}

function writeClaudeProjectInstructions(config: AppConfig, rulesFile: string): void {
  const claudeFile = path.join(config.dataDir, "CLAUDE.md");
  const managedBlock = `${managedClaudeRulesStart}
## Hare M365 Agent

- Before any Outlook, Teams, SharePoint, OneDrive, or Microsoft 365 request, read and follow [the Hare session rules](claude/hare-m365-agent-rules.md).
- Hare CLI is the exclusive tool for these lookups and for Outlook draft creation.
- Do not search for, invoke, or fall back to a Microsoft 365 connector, another connector, Computer Use, browser automation, or an Outlook/Teams/SharePoint user interface.
- If Hare is unavailable, unsupported, or fails, report the failed Hare step and stop instead of using another tool.
- Keep using the exact data directory and command prefix recorded in ${rulesFile}.
${managedClaudeRulesEnd}`;
  const existing = fs.existsSync(claudeFile) ? fs.readFileSync(claudeFile, "utf8") : "";
  const withoutManagedBlocks = existing
    .replace(
      /<!-- HARE_M365_AGENT_RULES_START -->[\s\S]*?<!-- HARE_M365_AGENT_RULES_END -->\s*/g,
      ""
    )
    .trimEnd();
  const nextContents = withoutManagedBlocks
    ? `${withoutManagedBlocks}\n\n${managedBlock}\n`
    : `${managedBlock}\n`;

  if (existing !== nextContents) writeStoredText(claudeFile, nextContents);
}

function buildSessionRules(
  config: AppConfig,
  options: SessionRulesOptions,
  rulesFile: string
): string {
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  const startupCommand = options.commandPrefix;
  const statusCommand = `${options.commandPrefix} auth status`;

  return `# Hare M365 Agent Session Rules

The project folder selected when this Cowork task was opened is Hare's persistent data directory. Keep using that exact project root regardless of its folder name.

## Fixed locations

- Persistent data directory: ${config.dataDir}
- Rules file: ${rulesFile}
- Login cache: ${cacheFile}
- Downloads: ${config.downloadDir}
- Results: ${config.resultsDir}
- Logs: ${config.logsDir}
- Current session app directory: ${options.workDir}

The selected project folder stores Hare data only. Do not clone the repository, run npm ci, or build inside it. Cowork mounts may allow create and overwrite while rejecting shell deletion, so Hare updates its own cache and rule files without requesting folder deletion permission.

The application checkout belongs in the Cowork session runtime shown by the setup command. It can be recreated in a new Cowork task. Never replace the selected project data directory with a temporary or guessed path.

## Code and commands

- Repository: ${options.repository}
- Branch: ${options.branch}
- Verify the latest code with refs/heads/${options.branch} before building.
- Command prefix for this selected project:

    ${options.commandPrefix}

- Startup command:

    ${startupCommand}

- Authentication status command:

    ${statusCommand}

Every command must keep this exact --data-dir. A new Cowork task may recreate the session app, but a usable login in this selected project remains reusable.

## Startup state

1. Run startup and follow only setup.state and setup.nextCommand. A usable login requires both loggedIn=true and tokenUsable=true.
2. READY: Do not start a new login; wait for the user's Microsoft 365 request.
3. LOGIN_START_REQUIRED: run setup.nextCommand once and show the Microsoft URL and user code. Tell the user to sign in with their own company Microsoft account that they will use with Hare. Never name, recommend, or preselect a specific email address from a cache, example, or prior conversation. If authReason is AUTH_APP_CHANGED, explain only that Hare was updated to a new Microsoft application and one Microsoft sign-in is required.
4. LOGIN_COMPLETE_REQUIRED: wait until the user says the browser login is complete, then run setup.nextCommand once.
5. FOLDER_REQUIRED: stop and tell the user to open a new Cowork task with the existing Hare project selected.
6. Do not request deletion permission for the selected project folder. Do not move dataDir to another path.

## Cowork network permission

- Run git, npm, login, and Graph commands in the session sandbox shell where the Cowork domain allowlist applies.
- If HTTP 403 includes X-Proxy-Error: blocked-by-allowlist, report NETWORK_PERMISSION_REQUIRED and stop after identifying the failed domain.
- When the allowlist changes, open a new Cowork task with the same Hare project selected and retry only the failed step once.

## Microsoft 365 operation

- Hare CLI is the exclusive tool for requested Outlook, Teams, SharePoint, OneDrive, and Microsoft 365 lookups and for Outlook draft creation.
- Do not search for, invoke, or fall back to a Microsoft 365 connector, another connector, Computer Use, Outlook/Teams/SharePoint user interfaces, or browser automation for these requests.
- If Hare is unavailable, unsupported, or fails, report the failed Hare step and stop. Do not substitute another tool or data source.
- Default to read-only operation. The only enabled write action is creating an Outlook draft after exact user approval. Sending mail, Teams posting, calendar creation, file upload/delete/share, and permission changes are unavailable.
- Always use the Hare CLI for Outlook draft requests. Never use Computer Use, Outlook desktop/web UI, browser automation, or a Microsoft 365 connector to create a draft or paste its content.
- If a Hare draft command fails, report the failed Hare step and stop. Do not fall back to GUI automation or another connector.
- Drafts support new messages, replies, reply-all, forwards, and file attachments. Run the draft command without an approval token first.
- Show the complete AWAITING_USER_APPROVAL preview, including recipients, subject, body, and attachments, then stop. Only after explicit user approval, rerun the exact same command once with the returned --approval-token.
- Never reuse an approval token after changing content, recipients, or attachments. Hare cannot send a draft.
- Use outlook recent --folder all for general or recent-mail requests. Use outlook inbox only when the user explicitly asks for the Inbox.
- Use outlook flagged --folder all for flagged-mail requests and inspect flagStatus in every mail result.
- When the user asks for an Outlook attachment, use outlook attachments list with the message ID, then outlook attachments download with the selected attachment ID. Downloaded files are saved under Hare downloadsDir and remain subject to the download policy and size limit.
- When the user asks for a Teams chat attachment, use teams attachments list with the chat and message IDs, then teams attachments download with the selected attachment ID. Never expose the original sharing URL. The existing download policy and size approval gate still apply.
- Downloads at or below the default limit run normally. A larger download returns AWAITING_USER_APPROVAL without downloading content. Show the complete source, file name, size, output name, destination, and limits, then stop.
- Only after explicit approval, rerun the exact same download command once with the returned --approval-token. The token expires after 10 minutes, is bound to that exact file and output name, and cannot be reused. Never change the policy or choose another download path to bypass the gate.
- When the user omits a date range, the default lookback is ${config.policy.defaultSearchLookbackDays} days and the actual range must be reported.
- Use outlook count for exact mail counts and sharepoint sites for SharePoint site existence checks.
- Outlook search returns 100 messages per page by default. Use body/bodyHtml as the complete content, check fullBodyUnavailableCount, and pass nextCursor through --cursor while continuationAvailable is true.
- If outlook count returns complete=false, continue with nextCursor. The cursor preserves cumulative state, so only the final matchedCount returned with complete=true is the exact total.
- File search covers accessible SharePoint, Teams, and OneDrive files. Pass nextOffset through --offset while continuationAvailable is true. Use sharepoint sites to check whether a site itself exists.
- Use lastMessageCreatedDateTime, not lastUpdatedDateTime alone, when deciding the latest Teams chat.
- Teams chat-messages body and bodyHtml are the complete untruncated message. bodyPreview is only a compatibility alias for the same full text.
- Teams chat-messages returns 20 messages by default and at most 1,000 unique messages per command. While page.continuationAvailable is true, pass page.nextOffset through --offset.
- Teams search-messages performs detail lookups for full bodies. Use body/bodyHtml, check fullBodyUnavailableCount, and never present searchSummary as the complete message when fullBodyAvailable is false.
- Teams search-messages hydrates at most 100 unique messages per command and Microsoft Search exposes at most a 1,000-result window. Never raise the limit above 100 or continue after searchWindowExhausted or noProgressDetected.
- totalMatchesReported counts messages matched by Microsoft Search across sender, body, and attachments. It is not an exact count of text occurrences in message bodies. Do not exhaustively paginate only to revalidate totalMatchesReported.
- For an exact body-occurrence request, explain the distinction and inspect the first page. If totalMatchesReported exceeds searchWindowLimit, do not paginate; ask for a narrower date range or query because an exact count cannot be guaranteed. Otherwise inspect pages only within the 1,000-result window and stop on duplicate/no-progress metadata.
- When partialResult is true, report partialReason and fullBodyUnavailableCount. Do not present a time-budget-limited result as complete.

## Required domains

${options.requiredDomains.map((domain) => `- ${domain}`).join("\n")}
`;
}
