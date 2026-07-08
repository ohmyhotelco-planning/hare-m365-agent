# Claude Cowork Runbook

Claude Cowork may use `computer-use` to control the Windows desktop. That is acceptable.

The rule is not "avoid computer-use." The rule is:

- Use `computer-use` to open PowerShell, Command Prompt, or Windows Terminal.
- Do not use `computer-use` to browse with File Explorer and double-click `omh-m365.exe`.
- Run `omh-m365.exe` as a command-line tool with explicit arguments.

## Default Startup

If the user gives this folder path or asks to use this M365 agent, do not ask which command to run.

Open PowerShell or Windows Terminal and run:

```powershell
Set-Location -LiteralPath "C:\Users\OMH\Documents\GitHub\omh-m365-agent\releases\win-x64"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
```

If the folder is different, replace the path with the folder the user provided.

Treat this startup result as a hard gate. Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `loggedIn: true`.

If `configured` is false, stop and tell the human to create `.env` from `.env.example` with approved Azure Application values locally. Do not ask for those values in chat.

## If Not Logged In

Open PowerShell or Windows Terminal and run:

```powershell
Set-Location -LiteralPath "C:\Users\OMH\Documents\GitHub\omh-m365-agent\releases\win-x64"
Start-Process "https://microsoft.com/devicelogin"
.\omh-m365.exe auth login
```

The human user must enter the device code. Claude must not read, repeat, store, or summarize the device code.

After the human says login is complete, rerun:

```powershell
.\omh-m365.exe doctor
.\omh-m365.exe auth status
```

Only then resume the original task.

## After Login

Run the smallest read command sequence needed for the user's request.

Examples:

```powershell
.\omh-m365.exe outlook inbox --limit 10
.\omh-m365.exe teams chats --limit 20
.\omh-m365.exe teams chat-messages --chat-id "<chat-id>" --limit 20
.\omh-m365.exe files search --query "keyword" --limit 10
```

## If Terminal Execution Is Impossible

If Claude Cowork cannot open PowerShell, Command Prompt, or Windows Terminal, stop and say:

```text
I can read the folder instructions, but I cannot execute the M365 agent from this environment. Please run this from Windows PowerShell, Codex Desktop, Claude Code with shell access, or another shell-enabled local LLM environment.
```

Do not upload, decompile, emulate, or inspect `omh-m365.exe`.
