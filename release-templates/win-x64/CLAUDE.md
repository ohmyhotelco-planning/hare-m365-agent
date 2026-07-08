# OMH M365 Agent Instructions For Claude

Use this package as a safe local execution layer for Microsoft 365 tasks. The human asks in natural language; Claude should use the documented CLI and avoid direct token, cache, credential, or raw Graph handling.

Do not ask the human which command to run. If local Windows terminal execution is available, run `doctor`, then `auth status`, then the smallest read command sequence needed for the request.

Treat configuration and login as hard gates. Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `loggedIn: true`.

Claude Cowork may use `computer-use` to control the Windows desktop. That is acceptable. Use it to open Windows PowerShell, Command Prompt, or Windows Terminal and run commands by path.

Do not use File Explorer to open the folder and double-click `omh-m365.exe`. This is a command-line tool.

If running in Claude Cowork, also read `CLAUDE_COWORK_RUNBOOK.md`.

If Windows terminal execution is not available, state that this folder must be used from Windows PowerShell, Command Prompt, Windows Terminal, Codex Desktop, or another shell-enabled local LLM environment.

Terminal command template:

```powershell
Set-Location -LiteralPath "<WINDOWS_DISTRIBUTION_FOLDER>"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
```

## Required Behavior

- Use documented CLI commands only.
- Never print or inspect `.env`, `.cache/`, access tokens, refresh tokens, cookies, device codes, or MSAL cache contents.
- Keep all operations read-only unless a future policy explicitly enables write actions.
- Do not send email, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Prefer concise summaries and relevant metadata over raw message dumps.
- Ask before downloading files.
- Use small limits first, then expand only as needed.

## Login

Claude may launch the login command in a separate terminal and open the Microsoft device login page for the user.

Claude must not read, repeat, store, or summarize the displayed device code.

If Claude can run terminal commands, prefer terminal commands. If Claude can only launch files and cannot pass command-line arguments, run `RUN_FIRST_FOR_LLM.cmd` for the initial check and `START_LOGIN_FOR_USER.cmd` for login as a fallback.

If configuration is missing, tell the human to create `.env` from `.env.example` with approved Azure Application values locally. Do not ask for those values in chat. If authentication is missing or expired, complete login first, rerun startup, then resume the original task.

## Commands

When `omh-m365.exe` is available, use it first:

```bash
./omh-m365.exe doctor
./omh-m365.exe auth status
./omh-m365.exe auth login
./omh-m365.exe auth logout
./omh-m365.exe outlook inbox --limit 10
./omh-m365.exe teams teams
./omh-m365.exe teams chats --limit 20
./omh-m365.exe teams chat-messages --chat-id "<chat-id>" --limit 20
./omh-m365.exe files search --query "keyword" --limit 10
./omh-m365.exe files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

Teams chat recency rule:

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only. Do not use it as the latest-message timestamp.

For source-based development, use npm:

```bash
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

## Recovery

- Missing configuration: tell the user to fill `.env` locally with approved Azure Application values.
- Authentication missing or expired: do not retry Outlook, Teams, or Files yet. Guide the user through `auth login`, rerun startup, then resume the original task.
- Access denied: explain the difference between Azure Application permissions and the signed-in user's actual M365 permissions.
- Write/delete/share/send request: say it is outside the current POC policy.
