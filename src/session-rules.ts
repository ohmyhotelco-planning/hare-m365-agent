import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { writeStoredText } from "./persistent-storage.js";

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

  if (fs.existsSync(rulesFile) && fs.readFileSync(rulesFile, "utf8") === contents) {
    return rulesFile;
  }

  writeStoredText(rulesFile, contents);
  return rulesFile;
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
3. LOGIN_START_REQUIRED: run setup.nextCommand once and show the Microsoft URL and user code. If authReason is AUTH_APP_CHANGED, explain only that Hare was updated to a new Microsoft application and one Microsoft sign-in is required.
4. LOGIN_COMPLETE_REQUIRED: wait until the user says the browser login is complete, then run setup.nextCommand once.
5. FOLDER_REQUIRED: stop and tell the user to open a new Cowork task with the existing Hare project selected.
6. Do not request deletion permission for the selected project folder. Do not move dataDir to another path.

## Cowork network permission

- Run git, npm, login, and Graph commands in the session sandbox shell where the Cowork domain allowlist applies.
- If HTTP 403 includes X-Proxy-Error: blocked-by-allowlist, report NETWORK_PERMISSION_REQUIRED and stop after identifying the failed domain.
- When the allowlist changes, open a new Cowork task with the same Hare project selected and retry only the failed step once.

## Microsoft 365 operation

- Use Hare for requested Outlook, Teams, SharePoint, and OneDrive lookups.
- Default to read-only operation. The only enabled write action is creating an Outlook draft after exact user approval. Sending mail, Teams posting, calendar creation, file upload/delete/share, and permission changes are unavailable.
- Drafts support new messages, replies, reply-all, forwards, and file attachments. Run the draft command without an approval token first.
- Show the complete AWAITING_USER_APPROVAL preview, including recipients, subject, body, and attachments, then stop. Only after explicit user approval, rerun the exact same command once with the returned --approval-token.
- Never reuse an approval token after changing content, recipients, or attachments. Hare cannot send a draft.
- Use outlook recent --folder all for general or recent-mail requests. Use outlook inbox only when the user explicitly asks for the Inbox.
- Use outlook flagged --folder all for flagged-mail requests and inspect flagStatus in every mail result.
- When the user omits a date range, the default lookback is ${config.policy.defaultSearchLookbackDays} days and the actual range must be reported.
- Use outlook count for exact mail counts and sharepoint sites for SharePoint site existence checks.
- Use lastMessageCreatedDateTime, not lastUpdatedDateTime alone, when deciding the latest Teams chat.
- Teams chat-messages body and bodyHtml are the complete untruncated message. bodyPreview is only a compatibility alias for the same full text.
- Teams search-messages performs detail lookups for full bodies. Use body/bodyHtml, check fullBodyUnavailableCount, and never present searchSummary as the complete message when fullBodyAvailable is false.

## Required domains

${options.requiredDomains.map((domain) => `- ${domain}`).join("\n")}
`;
}
