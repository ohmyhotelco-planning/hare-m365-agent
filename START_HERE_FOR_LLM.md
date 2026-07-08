# Start Here For LLM

Use this folder as a Linux-local Microsoft 365 lookup tool attached to an LLM.

This project is not primarily a standalone desktop app. The intended use is:

```text
LLM -> local Linux shell/sandbox -> node dist/cli.js -> Microsoft Graph delegated access
```

Do not switch to Windows GUI control, File Explorer, or exe double-click behavior unless the user explicitly asks for Windows-local testing. If you have a Linux shell, use it directly.

If the user provides a Windows path such as `C:\Users\...\releases\linux-llm`, treat it as a location hint only. Do not request Windows desktop control just to open that path.

In Claude Cowork, first check whether the folder is already mounted into the Linux shell. If it is not mounted, use Cowork's folder/project connection tool to connect the local folder to the sandbox, then use the mounted Linux path. Folder connection is allowed; Windows GUI control, File Explorer control, and double-click execution are not.

If the folder cannot be connected or mounted into the shell, stop and say that the folder is not accessible from the shell; do not fall back to `computer-use` or File Explorer.

For Claude Cowork, first read `COWORK_REQUIRED_README.md` and follow it literally.

In Claude Cowork, use `bash run-cowork.sh ...` instead of direct `node omh-m365.cjs ...`. The wrapper sets `NODE_USE_ENV_PROXY=1` so Node/MSAL/Graph traffic uses Cowork's host proxy. Cowork's one-shot shell is suitable for short read commands, but not for long interactive device-code login. For non-developers, ask the human to run `LOGIN_FOR_COWORK.cmd` from the same `linux-llm` folder. WSL is only a technical fallback.

In a packaged `linux-llm` release, prefer the bundled CLI:

```bash
node omh-m365.cjs doctor
node omh-m365.cjs auth status
```

This avoids `npm ci` and does not require network access to npm registries.

## Default Rule

Do not ask the user how to run this tool. Start with the startup checklist.

Before any Outlook, Teams, or Files command, authentication must be confirmed. If configuration is missing or `loggedIn` is false, do not try the read command yet. Resolve the gate first, then resume the user's original request.

## Startup Checklist

Run these commands from this project or release folder:

```bash
pwd
node --version
if [ -f omh-m365.cjs ]; then
  node omh-m365.cjs doctor
  node omh-m365.cjs auth status
else
  test -d node_modules || npm ci
  test -f dist/cli.js || npm run build
  node dist/cli.js doctor
  node dist/cli.js auth status
fi
```

If the user gave only a folder path and no task, run the checklist and report readiness.

If the user gave a natural-language M365 request, run the checklist and then execute the smallest safe read command sequence needed for the request.

If `doctor` shows `configured: false`, stop and tell the human to create `.env` from `.env.example` with approved Azure Application values locally. Do not ask them to paste those values into chat.

If `doctor` or `auth status` shows `loggedIn: false`, stop the read command sequence and use the Login section below first. Do not call `outlook`, `teams`, or `files` commands until login is complete.

## Login

If `auth status` shows the user is not logged in, authentication is required.

Login is a human action because Microsoft device-code authentication displays a one-time code. Do not copy, repeat, summarize, store, screenshot, or paste the device code into chat.

Preferred login flow in a local Linux environment:

```bash
./START_LOGIN_FOR_USER.sh
```

If that script cannot open a user-visible terminal/browser in the local environment, tell the user exactly:

```text
Please run ./START_LOGIN_FOR_USER.sh in your local Linux terminal, complete Microsoft device login, then tell me "login complete".
```

After login, resume the original task without asking which command to run.

## Allowed Read Commands

```bash
node dist/cli.js doctor
node dist/cli.js auth status
node dist/cli.js outlook inbox --limit 10
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20
node dist/cli.js files search --query "keyword" --limit 10
```

In `linux-llm` releases, replace `node dist/cli.js` with `node omh-m365.cjs`.

Teams chat recency rule:

- `teams chats` is ordered by `lastMessageCreatedDateTime`, which comes from Microsoft Graph `lastMessagePreview/createdDateTime`.
- Use `lastMessageCreatedDateTime` to decide the newest chat.
- Treat `lastUpdatedDateTime` as chat metadata only. Do not use it as the latest-message timestamp.

Allowed only after explicit user request for a specific file:

```bash
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

## Forbidden

- Do not perform write, send, delete, upload, share, permission-change, or calendar-change actions in this POC.
- Do not read or print `.env`, `.cache/`, token cache files, access tokens, refresh tokens, cookies, device codes, private keys, or real credential values.
- Do not upload this folder, token cache, or downloaded files to another environment.
- Do not use a remote hosted sandbox for login or token storage. This workflow assumes a local Linux sandbox managed by the user's LLM/Codex/Claude environment.

## Data Handling

This delegated public-client release may include `.env` with client and tenant identifiers. Do not read or print it; use `doctor` to check whether configuration is present.

Prefer concise conclusions over raw data dumps. Use small limits first, expand only when needed, and summarize the evidence behind the answer.

## Example User Request

```text
Find my Teams conversation with Mia about the refund and tell me when the refund is expected.
```

Expected LLM behavior:

- Run the startup checklist.
- Search Teams chats and messages as needed.
- Avoid exposing unnecessary raw message bodies.
- Answer with the relevant dates, people, invoice/reference numbers, and uncertainty.
