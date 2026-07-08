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
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365
```

Because this is not a global install, repeat the full npm exec prefix for every command.

## Required Behavior

- In Claude Cowork or any allow-list environment, confirm required domain allow-list setup before running npm, npx, curl, package download, login, or Microsoft 365 commands.
- Run the startup checklist only after the domain gate is satisfied.
- Treat startup as a hard gate.
- Do not run Outlook, Teams, or Files commands until `configured: true` and `loggedIn: true`.
- If `loggedIn: false`, do not ask which command to run and do not ask whether to log in. Start `auth login` immediately when the current shell output is visible to the human.
- During `auth login`, the terminal may display a Microsoft device code for the human. Do not copy, repeat, summarize, or paste that code into chat. Ask only for the human to complete the browser login and say "로그인 완료".
- If the current environment cannot show the login flow to the human, stop and provide the exact local login command instead of trying desktop or GUI workarounds.
- Never print or inspect `.env`, `.cache/`, access tokens, refresh tokens, cookies, device codes, or MSAL cache contents.
- Keep operations read-only.
- Do not send email, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Prefer concise summaries and relevant metadata over raw message dumps.
- Ask before downloading files.

Startup checklist:

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
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

Use this command only where the human can see the login code and complete the browser step:

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
```

Do not read, repeat, store, or summarize the displayed device code.

Do not say that the code cannot be shown to the human. The code may appear in the user's terminal or browser flow; the rule is that Claude must not copy it back into chat.

When `auth status` shows `loggedIn: false`, run the login command before any Outlook, Teams, or Files command. Do not present a menu. After starting login, tell the human:

```text
브라우저에 화면의 코드를 직접 입력하고 완료되면 "로그인 완료"라고 알려주세요. 코드는 채팅에 붙여넣지 마세요.
```

After the human says login is complete, rerun startup and resume the original task.

## Commands

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams teams
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chats --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files search --query "keyword" --limit 10
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
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

