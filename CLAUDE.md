# Hare M365 Agent For Claude/Cowork

Use git clone as the default setup path.

```bash
git clone https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
cd /tmp/hare-m365-agent
npm ci
npm run build
node dist/cli.js
```

Before running commands in Cowork, confirm the allow-list contains:

```text
github.com
login.microsoftonline.com
graph.microsoft.com
```

If `npm ci` fails because it cannot reach npm registry, report that `registry.npmjs.org` is additionally required and stop.

Do not test repo access with `api.github.com`. Use `git ls-remote` or `git clone`.

After startup, read the JSON and follow these rules:

- Use only the printed `dataDir` and `cacheFile` state.
- If the user's fixed Hare folder is mounted, set `HARE_M365_DATA_DIR="<mounted Hare folder path>"` for every command.
- If `loggedIn` is false, do not query Outlook/Teams/Files.
- If `cacheFileExists` is true, do not ask the user to login again; continue with the requested read command.
- For initial connection, run `humanLoginCommand` in the same shell. Do not ask the user to type a shell command or navigate to the clone folder. The user enters the Microsoft device code in the browser, signs in, then says "로그인 완료".
- If Graph calls fail in the hosted sandbox but cache exists, run the command on the user's local machine with `--out`, then read the output JSON.

Do not use computer-use or File Explorer when shell commands are available.

Do not perform write/change actions unless explicitly enabled and confirmed.
