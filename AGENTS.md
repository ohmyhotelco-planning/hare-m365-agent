# Hare M365 Agent Instructions

## Purpose

Hare M365 Agent lets an LLM query Microsoft 365 through Microsoft Graph delegated permissions. Default operation is read-only.

## Default Cowork Flow

Use git clone as the default setup path. Do not use GitHub API, GitHub Release asset downloads, or npmjs publish as the primary path.

```bash
git clone https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
cd /tmp/hare-m365-agent
npm ci
npm run build
node dist/cli.js
```

If the repository already exists:

```bash
cd /tmp/hare-m365-agent
git pull
npm ci
npm run build
node dist/cli.js
```

## Domain Checks

Initial required domains:

```text
github.com
login.microsoftonline.com
graph.microsoft.com
ohmylab-my.sharepoint.com
ohmylab.sharepoint.com
```

SharePoint/OneDrive file content downloads may resolve to the tenant SharePoint host, not only Graph. If wildcard domains are supported, `*.sharepoint.com` can cover these hosts.

If `npm ci` fails with an npm registry access error, report that `registry.npmjs.org` is additionally required and stop.

Do not use `api.github.com` as the repo access test. In Cowork, GitHub API may be blocked while `git clone` succeeds. Use `git ls-remote` or `git clone`.

## Storage

Hare uses a fixed OS data folder. Do not create an arbitrary runtime folder.

```text
Windows: %LOCALAPPDATA%\Ohmyhotel\HareM365Agent
Mac: ~/Library/Application Support/Ohmyhotel/HareM365Agent
Linux: ~/.local/share/ohmyhotel/hare-m365-agent
```

If Claude/Cowork mounts the user's fixed Hare folder, set `HARE_M365_DATA_DIR="<mounted Hare folder path>"` for every Hare command so the same cache is used.

## Login Gate

If `loggedIn` is false, do not run Outlook, Teams, or Files lookup.

`loggedIn` is scoped to the printed `dataDir` and `cacheFile`. A false value from a hosted sandbox path is not proof that the user's PC is not logged in.

If a cache file exists in the fixed Hare folder, do not ask the user to login again. Run the requested read command with the same `HARE_M365_DATA_DIR`.

For initial connection, run the printed `humanLoginCommand` in the same shell. Do not ask the user to type a shell command or navigate to the clone folder. The user enters the Microsoft device code in the browser, signs in with the company account, then says "로그인 완료".

## Read Commands

```bash
node dist/cli.js outlook inbox --limit 10 --out latest-mail.json
node dist/cli.js teams teams --out teams.json
node dist/cli.js teams chats --limit 20 --out chats.json
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20 --out chat-messages.json
node dist/cli.js files search --query "keyword" --limit 10 --out files.json
```

If a hosted sandbox can read the cache but Graph calls fail with `fetch failed` or `network_error`, treat it as sandbox egress trouble. Do not loop on diagnostics. Run the same command locally with `--out`, then read the JSON output.

## Safety Boundary

Do not send mail, post Teams messages, create calendar events, upload/delete/share files, or change permissions unless the policy explicitly allows it and the user confirms the exact action.
