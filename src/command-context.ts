export type RuntimeHost = "cowork" | "claude-code";
export type CommandShell = "powershell" | "posix";

export function quoteCommandArgument(value: string, shell: CommandShell): string {
  if (/[\0\r\n]/.test(value)) throw new Error("Hare command arguments cannot contain null bytes or line breaks.");
  return shell === "powershell"
    ? `'${value.replaceAll("'", "''")}'`
    : `'${value.replaceAll("'", `'"'"'`)}'`;
}

export function commandPath(value: string, shell: CommandShell, platform = process.platform): string {
  // Git Bash passes forward-slash Windows paths to the native Node executable.
  return platform === "win32" && shell === "posix" ? value.replaceAll("\\", "/") : value;
}

export const claudeCodeExecutionGuidance = `Use Claude Code Desktop Local in the user's selected persistent Hare folder. Cloud, remote, WSL, and Git worktree sessions are not this setup mode. Keep the exact --data-dir, --host claude-code, and --command-shell from the command prefix; execute commands in that shell, not another shell's syntax.
Before startup/auth status or the first requested M365 read in a new execution environment, run the same prefix followed by network check --environment claude-code. This public-metadata check does not access the login cache. Do not repeat it for every page.
- REACHABLE verifies connectivity only. Run startup in the same environment; only loggedIn=true and tokenUsable=true permit M365 reads.
- Use Claude Code's normal permission prompts for the exact command, paths and network access needed. Never enable bypass permissions, broadly allow arbitrary node/shell commands, or evade an organization policy. A denied request must stop the operation.
- NETWORK_PERMISSION_REQUIRED is an explicit domain policy block. Request the policy change and stop; do not switch hosts or proxies.
- NETWORK_CHECK_BLOCKED means connectivity is unverified, not token expiry. Report the safe error and stop. Preserve the existing cache; do not automatically retry writes or start login to solve a network failure.
- Read CLAUDE.md and claude/hare-m365-agent-rules.md after setup and in each new session. Never inspect raw authentication/cache files or put them in Git, prompts, logs, or cloud-synced folders. Local execution does not mean that returned Microsoft 365 content stays outside the AI service.`;

export const claudeCodeFolderInstruction = "Report FOLDER_REQUIRED and stop. Open the existing persistent Hare data folder in Claude Code Desktop Local and pass its exact absolute path with --data-dir. Do not guess a replacement folder or use a Git worktree.";
