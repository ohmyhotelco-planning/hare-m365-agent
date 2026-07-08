# Hare M365 Agent Instructions For LLMs

## Role

Use Hare M365 Agent as the controlled local execution layer for Microsoft 365 read tasks requested by the user.

Primary model:

```text
LLM -> local shell -> npm exec GitHub Release tarball -> hare-m365 -> Microsoft Graph delegated access
```

This is not primarily a desktop GUI app. Prefer shell commands over Windows GUI control, File Explorer, double-click execution, or `computer-use`.

## Primary Package URL

Use this GitHub Release package URL unless the user provides a newer one:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
```

Because this is an `npm exec --package <url>` flow, do not switch to bare `hare-m365` unless the CLI is globally installed. Repeat the full npm exec prefix for each command.

Command prefix:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365
```

## Connected Runtime Folder

For Claude Cowork or any folder-mounted LLM environment, use a connected runtime folder so the CLI process can reuse Microsoft login state across commands.

Preferred folder layout:

```text
<connected-folder>/
  runtime/
    .cache/
      msal-cache.json
    downloads/
    logs/
```

Run every Hare command from the connected folder with `HARE_M365_DATA_DIR=./runtime`.

Linux/Mac shell:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

Windows PowerShell:

```powershell
$env:HARE_M365_DATA_DIR = "$PWD\runtime"
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

Security boundary:

- The CLI may read and write `runtime/.cache/msal-cache.json` to authenticate Graph calls.
- The LLM must not open, print, summarize, copy, upload, or inspect `runtime/.cache/msal-cache.json`.
- If login is required, tell the human to run `auth login` in the same connected folder with the same `HARE_M365_DATA_DIR=./runtime` setting. Login in a different folder or machine will not help this execution environment.

## Default Behavior

- Do not ask the user which command to run.
- In Claude Cowork or any allow-list environment, confirm the required domains with the user before running npm, npx, curl, downloads, login, or any Microsoft 365 command.
- Run `llm-guide`, then `doctor`, then `auth status` only after the domain gate is satisfied.
- If a connected folder is available, use `HARE_M365_DATA_DIR=./runtime` for every startup, login, and read command.
- Treat configuration and login as hard gates.
- Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `auth status` shows `loggedIn: true`.
- If `loggedIn: false`, stop at the login hard gate. Do not run Outlook, Teams, or Files commands.
- Do not automatically run `auth login` in Claude Cowork, hosted sandboxes, or any environment where command output is primarily visible to the agent rather than directly to the human.
- At the login hard gate, give the exact local login command and ask the human to run it in the same connected folder and runtime path. After the human says "로그인 완료", rerun `doctor` and `auth status`, then resume the original request.
- Only run `auth login` yourself if the user explicitly asks you to start login and the current terminal/browser output is directly visible to the human. Run it once. If it fails before showing a device code, report the error and stop instead of retrying or probing repeatedly.
- If a natural-language Microsoft 365 request is given, run the smallest safe read sequence needed to answer it.

Network/domain gate for Claude Cowork or any allow-list environment:

- GitHub Release execution may need `github.com` and `release-assets.githubusercontent.com`.
- Current v0.1.0 may also need `registry.npmjs.org` while npm installs package dependencies.
- Microsoft 365 calls need `graph.microsoft.com` and `login.microsoftonline.com`.
- If the user has not explicitly confirmed domain setup, stop before the startup checklist and ask them to add these domains first.
- If download or Graph access fails with `403`, `network_error`, proxy, or allow-list symptoms, report the blocked domain/error and stop. Do not switch to Windows GUI control, File Explorer, or double-click execution as a workaround.

Startup checklist:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

Login command:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
```

## Hard Safety Rules

- Never read, print, summarize, screenshot, upload, or copy `.env`, `.cache/`, `runtime/.cache/`, MSAL cache contents, access tokens, refresh tokens, cookies, device codes, private keys, or real credential values.
- Client ID and tenant ID are delegated public-client configuration, not secrets. Do not print them unless the user explicitly asks and it is operationally necessary.
- Default to read-only operations.
- Do not send mail, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Download files only when the user explicitly asks for a specific file and the CLI result identifies the target file.
- Return concise summaries and relevant metadata instead of dumping large raw mail, chat, or file contents.
- Use small limits first. Increase only when needed.

## Preferred Read Commands

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams teams
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chats --limit 20
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files search --query "keyword" --limit 10
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

## Login Handling

Authentication uses Microsoft device-code login and requires human action.

Allowed:

- Provide the exact `auth login` command when `loggedIn: false`.
- Run `auth login` only in a trusted local environment where the human can directly see the code and browser page, and only after the user explicitly asks you to start login.
- Open or guide the Microsoft device login page if the environment supports it.
- If `auth status` returns `loggedIn: false`, do not continue the Microsoft 365 task. Stop and tell the user: `로그인이 필요합니다. 아래 로그인 명령을 같은 연결 폴더에서 실행한 뒤, Microsoft 로그인을 완료하고 "로그인 완료"라고 알려주세요. 코드는 채팅에 붙여넣지 마세요.`

Forbidden:

- Do not auto-run `auth login` in Cowork or hosted sandbox shells.
- Do not paste the device code into chat.
- Do not say that the code cannot be shown to the human. The code can be shown in the user's terminal/browser flow; it just must not be copied back into chat.
- Do not inspect token cache contents to troubleshoot.
- Do not run login in a remote or unclear hosted sandbox where the human cannot complete the browser login safely.

If the environment cannot show the login flow to the human, tell the user:

```text
로그인이 필요합니다. 로컬에서 Hare M365 Agent auth login을 실행하고 Microsoft device-code 로그인을 완료한 뒤 "로그인 완료"라고 알려주세요.
```

## Teams Recency Rule

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only. Do not use it as the latest-message timestamp.

## Error Handling

- Missing configuration: say the package `.env` configuration is missing or invalid. Do not ask the user to paste values into chat.
- Not logged in: do not retry Outlook, Teams, or Files yet. Complete the login flow, rerun startup, then resume the original task.
- `403` or authorization errors: explain that either Azure Application Graph permissions or the signed-in user's M365 permissions may not allow the operation.
- Unsupported write/share/delete/send request: say it is outside the current POC policy.
- Network or transient Graph errors: run `doctor`, then retry once if the user still wants to proceed.
