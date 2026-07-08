# OMH M365 Agent Instructions For Codex

## Role

Use this package only as a controlled local execution layer for Microsoft 365 tasks requested by the user. Prefer the documented CLI over ad hoc Graph scripts.

This Windows distribution is intended to run from a local Windows folder through terminal commands. Do not use File Explorer to open the folder and double-click `omh-m365.exe`. This is a command-line tool. Use Windows PowerShell, Command Prompt, Windows Terminal, Codex shell, or another shell-enabled local LLM tool.

Claude Cowork may need to use `computer-use` to control the desktop. That is acceptable only as a way to open a terminal and type commands. It is not acceptable to use File Explorer and double-click the exe as the primary execution path. If running in Claude Cowork, also read `CLAUDE_COWORK_RUNBOOK.md`.

If the current environment cannot execute Windows terminal commands, do not inspect, emulate, decompile, or upload `omh-m365.exe`. Tell the user to run it from Windows PowerShell, Command Prompt, Windows Terminal, Codex Desktop, or another shell-enabled local LLM environment.

Do not ask the user how to run this tool. If local Windows terminal execution is available, run the default startup sequence yourself: `doctor`, then `auth status`, then the read command needed for the user's request.

Treat configuration and login as hard gates. Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `loggedIn: true`.

Terminal command template:

```powershell
Set-Location -LiteralPath "<WINDOWS_DISTRIBUTION_FOLDER>"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
```

## Hard Safety Rules

- Never read, print, summarize, screenshot, upload, or copy `.env`, `.cache/`, MSAL cache contents, access tokens, refresh tokens, cookies, device codes, private keys, or real credential values.
- Client ID and tenant ID are delegated public-client configuration, not secrets. Tokens, cookies, device codes, webhook URLs, production IPs, private keys, and credentials must still never be written into docs, prompts, logs, examples, Jira, Teams, or screenshots.
- This delegated public-client release may include `.env` with client and tenant identifiers. Do not read or print it; use `doctor` to check whether configuration is present.
- Default to read-only operations.
- Do not send mail, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Download files only when the user explicitly asks for a specific file and the CLI result identifies the target file.
- Return concise summaries and relevant metadata instead of dumping large raw mail, chat, or file contents.
- Use small limits first. Increase only when needed to answer the user.

## Login Handling

Allowed:

- Run `.\omh-m365.exe doctor`.
- Run `.\omh-m365.exe auth status`.
- If configuration is missing, stop and tell the human to create `.env` from `.env.example` with approved Azure Application values locally. Do not ask for those values in chat.
- If authentication is missing or expired, complete login first, rerun startup, then resume the original task.
- Run commands from PowerShell, Command Prompt, Windows Terminal, Codex shell, or another shell-enabled local LLM tool.
- Run `RUN_FIRST_FOR_LLM.cmd` only as a fallback when the environment can launch files but cannot pass CLI arguments.
- Open a separate terminal window for `.\omh-m365.exe auth login` when authentication is needed.
- Run `START_LOGIN_FOR_USER.cmd` only as a fallback when the environment can launch files but cannot pass CLI arguments.
- Open `https://microsoft.com/devicelogin` for the user.

Forbidden:

- Do not bring the device code back into chat.
- Do not inspect token cache contents to troubleshoot. Metadata such as file existence, size, modified time, and credential type counts is acceptable only when needed.

## Preferred Commands

When `omh-m365.exe` is available, use it as the primary execution unit:

```powershell
.\omh-m365.exe doctor
.\omh-m365.exe auth status
.\omh-m365.exe auth login
.\omh-m365.exe auth logout
.\omh-m365.exe outlook inbox --limit 10
.\omh-m365.exe teams teams
.\omh-m365.exe teams chats --limit 20
.\omh-m365.exe teams chat-messages --chat-id "<chat-id>" --limit 20
.\omh-m365.exe files search --query "keyword" --limit 10
.\omh-m365.exe files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

Default behavior:

- Folder only, no task: run `.\omh-m365.exe doctor` and `.\omh-m365.exe auth status`, then report readiness.
- Natural-language task: run `doctor`, `auth status`, then the smallest safe read sequence needed.
- Not logged in: start login, ask only for the human to complete Microsoft device login, then resume the task.
- Cannot execute Windows terminal commands: say that execution requires Windows PowerShell, Command Prompt, Windows Terminal, Codex Desktop, or another shell-enabled local LLM. Do not ask the user to choose a command.

Teams chat recency rule:

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only. Do not use it as the latest-message timestamp.

For source-based development, use npm:

```powershell
npm run start -- doctor
npm run start -- auth status
npm run start -- auth login
npm run start -- auth logout
npm run start -- outlook inbox --limit 10
npm run start -- teams teams
npm run start -- teams chats --limit 20
npm run start -- teams chat-messages --chat-id "<chat-id>" --limit 20
npm run start -- files search --query "keyword" --limit 10
npm run start -- files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

Use `npm run smoke:win` on Windows for a safe source-based summary-only read smoke test.

## Error Handling

- Missing configuration: ask the user to fill `.env` with approved Azure Application values. Do not ask them to paste values into chat.
- Not logged in: do not retry Outlook, Teams, or Files yet. Open the login flow, let the user complete Microsoft device code authentication, rerun startup, then resume the original task.
- Silent token failure: retry `auth login` once. If it persists, report the error and avoid token/cache inspection.
- `403` or authorization errors: explain that either Azure Application Graph permissions or the signed-in user's M365 permissions may not allow the operation.
- Unsupported write/share/delete/send request: say it is outside the current POC policy.
- Network or transient Graph errors: run `doctor`, then retry once if the user still wants to proceed.

## Data Handling

Downloaded files are stored under `downloads/` by default. Treat downloaded files as work data. Do not upload, share, delete, or quote sensitive file contents unless the user explicitly asks and the operation is supported by policy.
