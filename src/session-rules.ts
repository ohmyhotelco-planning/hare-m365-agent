import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";

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

  const temporaryPath = `${rulesFile}.${process.pid}.${Date.now()}.tmp`;
  try {
    fs.writeFileSync(temporaryPath, contents, { encoding: "utf8", mode: 0o600 });
    fs.renameSync(temporaryPath, rulesFile);
  } finally {
    fs.rmSync(temporaryPath, { force: true });
  }
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

This file may be copied between the current Cowork project and a Linux session working directory. The project root selected for the current Cowork task is the persistent source of truth.

## Fixed locations

- Rules file: ${rulesFile}
- Data directory (run location): ${config.dataDir}
- Persistent store: the project root selected for the current Cowork task (use the actual selected folder, regardless of its name)
- Login cache: ${cacheFile}
- Downloads: ${config.downloadDir}
- Results: ${config.resultsDir}
- Logs: ${config.logsDir}
- Cowork project folder: the current selected project root

Start every Cowork task with the user's existing Hare project folder selected. Do not search for a folder by the literal name HareM365Agent and do not switch to a similarly named sibling folder. The selected project root is the persistent store; the data directory above is where Hare actually runs.
Cowork Linux defaults such as /root/.local/share are container-local and are never a persistent Hare data directory.
Pick the run location by inspection, not by assumption: if the selected project mount (a path shaped like /sessions/<session>/mnt/<selected-project>) is directly visible and writable in the network-enabled shell, use that mount root as the data directory. If it is not visible there (standard Cowork), use the hybrid model: run Hare in a fixed sandbox path (for example /home/claude/hare) as the data directory, and sync persistent files with the selected project root using the file staging/commit tools.

## Hybrid session start and sync

1. Stage from the current selected project root into the data directory (same relative paths): .hare-app-snapshot.tar.gz, .hare-app-build-head, .cache/msal-cache.json, and the claude/ documents.
2. Run the setup command. An existing snapshot is extracted automatically and skips git clone; when HEAD matches the remote, npm ci and the build are skipped too.
3. After any rebuild, commit the refreshed .hare-app-snapshot.tar.gz and .hare-app-build-head back to that same selected project root. After any login or token-cache change, commit .cache/msal-cache.json back as well. Results the user wants to keep go back under results/.

## Code and commands

- Repository: ${options.repository}
- Branch: ${options.branch}
- Local app directory: ${options.workDir}
- Verify latest code with refs/heads/${options.branch} before building.
- Reuse this app directory in every chat. Run git fetch and git pull --ff-only; build only when HEAD changed or build files are missing.
- Command prefix for this data directory:

\`\`\`text
${options.commandPrefix}
\`\`\`

- Startup command:

\`\`\`text
${startupCommand}
\`\`\`

- Authentication status command:

\`\`\`text
${statusCommand}
\`\`\`

Always keep the exact --data-dir argument shown above. Updating or rebuilding the local app does not require a new Microsoft login when this same data directory contains a usable cache.

## Startup state

1. Run the startup command and follow only setup.state and setup.nextCommand.
2. READY means loggedIn and tokenUsable are both true. Do not start a new login; wait for the user's Microsoft 365 request.
3. LOGIN_START_REQUIRED means no usable login exists in this exact data directory. Run setup.nextCommand once and show the returned Microsoft URL and user code.
4. LOGIN_COMPLETE_REQUIRED means wait until the user says the browser login is complete, then run setup.nextCommand once.
5. FOLDER_REQUIRED means this task was opened without a project folder. Stop and tell the user to open a new Cowork task with their existing Hare project selected.
6. Do not interpret a session-local /sessions or /tmp data directory as the user's persistent login state. Accept /sessions/<session>/mnt/<selected-project> as the connected host folder when it is the current selected project and is directly visible in the running shell.

## Cowork network permission

- Cowork Settings > Features > Domain allowlist applies to the session sandbox shell. The device shell that mounts the connected folder may have all outbound domains blocked regardless of that setting (every domain returns 403 blocked-by-allowlist there).
- Never run git, npm, login, or Graph commands in the device shell. Run them in the sandbox shell where the domain allowlist actually applies.
- If a command returns HTTP 403 with X-Proxy-Error: blocked-by-allowlist, first check which shell ran it. From the device shell, rerun the same command in the sandbox shell instead. From the sandbox shell, report NETWORK_PERMISSION_REQUIRED and ask the user to set the domain preset to None and add the required domains below.
- If the allowlist was changed after the current Cowork task started, the user must open a new Cowork task so the new egress policy is applied. Reconnect the same HareM365Agent folder and retry only the failed command once.
- Do not repeat domain probes, clone, build, or login while the current task still has the old egress policy.

## Microsoft 365 operation

- Use Hare for requested Outlook, Teams, SharePoint, and OneDrive lookups.
- Default to read-only operation. Sending, posting, uploading, deleting, sharing, and permission changes are unavailable.
- When the user omits a date range, the default lookback is ${config.policy.defaultSearchLookbackDays} days and the actual range must be reported.
- Use outlook count for exact mail counts and sharepoint sites for SharePoint site existence checks.
- Use lastMessageCreatedDateTime, not lastUpdatedDateTime alone, when deciding the latest Teams chat.

## Required domains

${options.requiredDomains.map((domain) => `- ${domain}`).join("\n")}
`;
}
