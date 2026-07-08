# OMH M365 Agent POC

Local CLI execution layer for LLM-driven Microsoft 365 work through Microsoft Graph delegated permissions.

This POC is intentionally read-focused. The Azure Enterprise Application controls the maximum Graph permission scope, and this CLI controls the operations an LLM can actually call.

## Current Scope

Supported:

- Microsoft device code login
- Local diagnostics and login status
- Outlook Inbox recent message lookup
- Teams joined team lookup
- Teams chat lookup
- Teams chat message lookup by chat ID
- OneDrive/SharePoint-visible file search
- File download by drive ID and item ID, only when explicitly requested

Not supported:

- Sending email
- Posting Teams messages
- Creating or changing calendar events
- Uploading, editing, deleting, or sharing files
- Changing permissions

## Quick Start For LLM Execution

Give the LLM this folder path and ask it to read `START_HERE_FOR_LLM.md` first.

For Claude Cowork, ask it to read `CLAUDE_COWORK_RUNBOOK.md` as well. Claude Cowork may use `computer-use`, but only to open PowerShell/Command Prompt/Windows Terminal and run command-line instructions. It should not use File Explorer and double-click the exe.

Starter prompt:

```text
Use this folder as a Windows-local Microsoft 365 lookup tool.
First read START_HERE_FOR_LLM.md and AGENTS.md, then follow their rules.
If your environment can execute local Windows commands, use omh-m365.exe only for M365 access.
Open PowerShell, Command Prompt, or Windows Terminal and run commands by path. Do not use File Explorer or double-click the exe.
Do not ask me how to run it. Start by running doctor and auth status.
Do not read or print .env, .cache, tokens, or device codes.
Start with doctor and auth status, then use Outlook, Teams, and Files commands as needed to answer my natural-language request.
If your environment cannot execute Windows binaries, tell me that this folder must be used from Windows PowerShell, Codex Desktop, or another shell-enabled local LLM environment.
```

Preferred Windows execution unit:

```powershell
Set-Location -LiteralPath "<WINDOWS_DISTRIBUTION_FOLDER>"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
.\omh-m365.exe auth login
.\omh-m365.exe outlook inbox --limit 10
.\omh-m365.exe teams chats --limit 20
.\omh-m365.exe teams chat-messages --chat-id "<chat-id>" --limit 20
.\omh-m365.exe files search --query "keyword" --limit 10
```

If an LLM can control the desktop, it should open PowerShell or Windows Terminal and run the commands above. It should not navigate through File Explorer or double-click the exe.

Teams chat lists are ordered by `lastMessageCreatedDateTime`, the actual last-message timestamp from Microsoft Graph `lastMessagePreview/createdDateTime`. LLMs should not use `lastUpdatedDateTime` to decide the newest chat, because that value is chat metadata rather than the latest message time.

If an LLM can only launch files and cannot pass command-line arguments, it may use these fallback launchers:

```text
RUN_FIRST_FOR_LLM.cmd
START_LOGIN_FOR_USER.cmd
```

The npm commands below are for development and source-based testing.

## Source Setup

Requirements:

- Node.js 20 or later
- Approved Azure Application client ID and tenant ID

Windows setup:

```powershell
cd "<repo-or-unzipped-folder>"
npm run setup:win
```

The release may already include `.env` with delegated public-client configuration. If it is missing, fill `.env` with the approved Azure Application client and tenant identifiers.

First login:

```powershell
npm run start -- auth login
```

For LLM-driven testing, the LLM may open a separate PowerShell window and the Microsoft device login page. The LLM must not read, repeat, store, or summarize the device code.

Smoke test:

```powershell
npm run smoke:win
```

## Commands

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

## Packaging

Windows executable:

```powershell
npm run build:exe:win
```

The generated POC executable is written to `releases/win-x64/omh-m365.exe`.

Developer ZIP:

```powershell
npm run package:win
```

The ZIP package is a developer artifact, not the preferred non-developer distribution format. The package script uses an allowlist and excludes `.cache`, `downloads`, `logs`, `node_modules`, and existing release archives.

## Security Notes

- `.env` is delegated public-client configuration and may be included in the release.
- Never commit `.cache/`, `downloads/`, `logs/`, token cache files, access tokens, refresh tokens, device codes, cookies, or production credentials.
- The current POC token cache is stored locally under `.cache/`. A pilot or production build should use safer OS-backed storage such as Windows Credential Manager, DPAPI, or macOS Keychain.
- Keep the default policy read-only. Even if the Azure Application has write-capable permissions, this POC should not expose write commands.
- If a secret is accidentally exposed, deletion is not enough. Revoke or rotate the secret and review exposure scope.

## Verified POC Behavior

Validated in the local POC:

- `doctor`
- `auth status`
- `outlook inbox`
- `teams teams`
- `teams chats`
- `teams chat-messages`
- `files search`

Known issue:

- `npm audit` reports a moderate `uuid` advisory through `@azure/msal-node`. `npm audit fix --force` upgrades `@azure/msal-node` across a breaking boundary, so it is not applied automatically in this POC.
- The Windows executable produced by `build:exe:win` is not code-signed. It is suitable for POC validation, not broad production rollout.

## More Docs

- [LLM starter prompt](START_HERE_FOR_LLM.md)
- [LLM execution rules](AGENTS.md)
- [Claude execution rules](CLAUDE.md)
- [Claude Cowork runbook](CLAUDE_COWORK_RUNBOOK.md)
- [Operation model](docs/operation-model.md)
- [Deployment guide](docs/deployment-guide.md)
