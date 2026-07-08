# Start Here For LLM

Use this folder as a Windows-local Microsoft 365 lookup tool.

Do not ask the user how to run this tool. If you can execute local Windows commands, start immediately with the startup checklist below.

This release assumes the user has copied the Windows distribution folder to their own Windows machine. If your current LLM environment can execute local Windows commands, run `omh-m365.exe` from a terminal command line.

Do not use File Explorer to open the folder and double-click `omh-m365.exe`. This is a command-line tool. Use Windows PowerShell, Command Prompt, Windows Terminal, Codex shell, or another shell-enabled local LLM tool.

Terminal command template:

```powershell
Set-Location -LiteralPath "<WINDOWS_DISTRIBUTION_FOLDER>"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
```

If you are controlling the desktop through a GUI/computer-use tool, that is acceptable. Use computer-use to open PowerShell, Command Prompt, or Windows Terminal first, then run the commands above. Do not use computer-use to navigate through File Explorer and double-click `omh-m365.exe` unless the user explicitly asks you to inspect files visually.

For Claude Cowork, read `CLAUDE_COWORK_RUNBOOK.md` and follow it.

If your current environment cannot execute Windows terminal commands, do not inspect, emulate, decompile, or upload the executable. Tell the user that this folder must be used from Windows PowerShell, Command Prompt, Windows Terminal, Codex Desktop, or another shell-enabled local LLM environment. Do not ask whether to run diagnostics, login, or guidance; state the exact blocker and the exact environment needed.

First read `AGENTS.md` and follow it. If there is any conflict, follow the stricter safety rule.

For Microsoft 365 access, use only `omh-m365.exe`. Do not use raw Graph scripts unless the user explicitly asks for development or debugging work.

Startup checklist:

1. Run `.\omh-m365.exe doctor`.
2. Run `.\omh-m365.exe auth status`.
3. Treat configuration and login as hard gates. Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `loggedIn: true`.
4. If configuration is missing, tell the human to create `.env` from `.env.example` with approved Azure Application values locally. Do not ask them to paste those values into chat.
5. If login is missing or expired, open `.\omh-m365.exe auth login` in a separate terminal and let the user complete Microsoft device login.
6. Do not read, repeat, store, or summarize the device code.
7. After login, rerun the startup checklist, then use the relevant read commands to answer the user's natural-language request.

If your tool can launch terminal commands, prefer the terminal commands above. If your tool can only launch files and cannot pass command-line arguments, use `RUN_FIRST_FOR_LLM.cmd` as a fallback. If login is needed, use `START_LOGIN_FOR_USER.cmd` as a fallback.

If the user gives only a folder path and no task, run the startup checklist and report readiness. Do not ask which command to run.

If the user gives a natural-language task, run the startup checklist, then execute the smallest read command sequence needed for the task.

Allowed read commands:

```powershell
.\omh-m365.exe doctor
.\omh-m365.exe auth status
.\omh-m365.exe outlook inbox --limit 10
.\omh-m365.exe teams teams
.\omh-m365.exe teams chats --limit 20
.\omh-m365.exe teams chat-messages --chat-id "<chat-id>" --limit 20
.\omh-m365.exe files search --query "keyword" --limit 10
```

Teams chat recency rule:

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only. Do not use it as the latest-message timestamp.

Allowed only after explicit user request for a specific file:

```powershell
.\omh-m365.exe files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

Do not perform write, send, delete, upload, share, permission-change, or calendar-change actions in this POC.

Never inspect or print `.env`, `.cache/`, token cache files, access tokens, refresh tokens, cookies, device codes, private keys, or real credential values.

This delegated public-client release may include `.env` with client and tenant identifiers. Do not read or print it; use `doctor` to check whether configuration is present.

Prefer concise conclusions over raw data dumps. Use small limits first, expand only when needed, and summarize the evidence behind the answer.

Example user request:

```text
Find my Teams conversation with Mia about the refund and tell me when the refund is expected.
```

Expected LLM behavior:

- Check local tool status.
- Do not ask how to run the CLI.
- Search Teams chats and messages as needed.
- Avoid exposing unnecessary raw message bodies.
- Answer with the relevant dates, people, invoice/reference numbers, and uncertainty.
