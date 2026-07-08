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
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365
```

## Default Behavior

- Do not ask the user which command to run.
- Run `llm-guide`, then `doctor`, then `auth status`.
- Treat configuration and login as hard gates.
- Do not run Outlook, Teams, or Files commands until `doctor` shows `configured: true` and `auth status` shows `loggedIn: true`.
- If login is missing, run or guide `auth login` and let the human complete Microsoft device-code login.
- If a natural-language Microsoft 365 request is given, run the smallest safe read sequence needed to answer it.

Startup checklist:

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

Login command:

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
```

## Hard Safety Rules

- Never read, print, summarize, screenshot, upload, or copy `.env`, `.cache/`, MSAL cache contents, access tokens, refresh tokens, cookies, device codes, private keys, or real credential values.
- Client ID and tenant ID are delegated public-client configuration, not secrets. Do not print them unless the user explicitly asks and it is operationally necessary.
- Default to read-only operations.
- Do not send mail, post Teams messages, create calendar events, upload files, delete files, share files, or change permissions in this POC.
- Download files only when the user explicitly asks for a specific file and the CLI result identifies the target file.
- Return concise summaries and relevant metadata instead of dumping large raw mail, chat, or file contents.
- Use small limits first. Increase only when needed.

## Preferred Read Commands

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams teams
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chats --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files search --query "keyword" --limit 10
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "downloaded-file.ext"
```

## Login Handling

Authentication uses Microsoft device-code login and requires human action.

Allowed:

- Run `auth login` in a trusted local environment where the human can see the code and browser page.
- Open or guide the Microsoft device login page if the environment supports it.

Forbidden:

- Do not paste the device code into chat.
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
