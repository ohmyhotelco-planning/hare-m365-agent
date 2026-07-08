# Hare M365 Agent Instructions For Claude

Use Hare M365 Agent as a shell-executed Microsoft 365 read-only tool.

Primary workflow:

```text
Claude/Cowork/Claude Code -> local shell -> npm exec GitHub Release tarball -> hare-m365 -> Microsoft Graph delegated access
```

Do not default to `computer-use`, File Explorer, Windows desktop control, or double-click execution. Use shell commands directly when shell access exists.

## Primary Package URL

Use this GitHub Release package URL unless the user gives a newer version:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
```

Command prefix:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365
```

Because this is not a global install, repeat the full npm exec prefix for every command.

## Connected Runtime Folder

For Claude Cowork or any folder-mounted environment, use a connected runtime folder so the CLI can reuse Microsoft login state:

```text
<connected-folder>/runtime/.cache/msal-cache.json
```

Run every Hare command from the connected folder with `HARE_M365_DATA_DIR=./runtime`.

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

The CLI may use `runtime/.cache/msal-cache.json`; Claude must never open, print, summarize, upload, or inspect that file. If login is required, the human must run `auth login` in the same connected folder with the same `HARE_M365_DATA_DIR=./runtime` setting. Login in another local folder or machine does not authenticate this mounted execution environment.

## Required Behavior

- In Claude Cowork or any allow-list environment, confirm required domain allow-list setup before running npm, npx, curl, package download, login, or Microsoft 365 commands.
- Run the startup checklist only after the domain gate is satisfied.
- If a connected folder is available, use `HARE_M365_DATA_DIR=./runtime` for every startup, login, and read command.
- Treat startup as a hard gate.
- Do not run Outlook, Teams, or Files commands until `configured: true` and `loggedIn: true`.
- If `loggedIn: false`, stop at the login hard gate. Do not run Outlook, Teams, or Files commands.
- Do not automatically run `auth login` in Claude Cowork, hosted sandboxes, or any environment where command output is primarily visible to Claude rather than directly to the human.
- At the login hard gate, give the exact local login command and ask the human to run it in the same connected folder and runtime path. After the human says "로그인 완료", rerun startup and resume the original request.
- Only run `auth login` yourself if the user explicitly asks you to start login and the current terminal/browser output is directly visible to the human. Run it once. If it fails before showing a device code, report the error and stop instead of retrying or probing repeatedly.
- Never print or inspect `.env`, `.cache/`, `runtime/.cache/`, access tokens, refresh tokens, cookies, device codes, or MSAL cache contents.
- Keep operations read-only.
- Do not send email, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Prefer concise summaries and relevant metadata over raw message dumps.
- Ask before downloading files.

Startup checklist:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

## Cowork Notes

If Cowork has domain allow-list controls, Microsoft 365 and GitHub Release execution need access to:

- `github.com`
- `release-assets.githubusercontent.com`
- `registry.npmjs.org`
- `graph.microsoft.com`
- `login.microsoftonline.com`

`github.com` alone is not enough for Release assets. GitHub redirects the `.tgz` asset download to `release-assets.githubusercontent.com`.

Current v0.1.0 may also need `registry.npmjs.org` while npm installs package dependencies.

If the user has not explicitly confirmed these domains are allowed, stop before the startup checklist and ask the user to add them first.

If shell access is available, do not ask for Windows desktop control just to run Hare.

If the package download is blocked by network policy, report the blocked domain/error and stop. Do not switch to GUI control unless the user explicitly asks for desktop troubleshooting.

## Login

Login requires human device-code authentication.

When `auth status` shows `loggedIn: false`, stop and tell the human to run this command in a local terminal they can see:

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
```

Do not read, repeat, store, or summarize the displayed device code.

Do not say that the code cannot be shown to the human. The code may appear in the user's terminal or browser flow; the rule is that Claude must not copy it back into chat.

Do not automatically run the login command in Cowork or hosted sandbox shells. Run it only if the user explicitly asks you to start login and the terminal/browser output is directly visible to the human. After starting login, tell the human:

```text
브라우저에 화면의 코드를 직접 입력하고 완료되면 "로그인 완료"라고 알려주세요. 코드는 채팅에 붙여넣지 마세요.
```

After the human says login is complete, rerun startup and resume the original task.

## Commands

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams teams
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chats --limit 20
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files search --query "keyword" --limit 10
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

## Teams Recency Rule

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only.

## Recovery

- Missing configuration: say the package configuration is missing or invalid. Do not ask the user to paste `.env` values into chat.
- Authentication missing or expired: complete login first, rerun startup, then resume.
- Access denied: explain the difference between Azure Application permissions and the signed-in user's actual M365 permissions.
- Write/delete/share/send request: say it is outside the current POC policy.

