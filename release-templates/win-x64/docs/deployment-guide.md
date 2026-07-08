# OMH M365 Agent Deployment Guide

## Recommended Path

Use a staged distribution model.

1. POC distribution: Windows executable plus LLM-readable Markdown rules.
2. Developer artifact: release ZIP with source, setup, and smoke-test scripts.
3. Pilot distribution: signed executable or controlled internal installer.
4. Broad internal distribution: installer or endpoint-management package with bundled prerequisites and update handling.

## Preferred Windows Execution Unit

For LLM use, prefer a direct executable:

```powershell
Set-Location -LiteralPath "<WINDOWS_DISTRIBUTION_FOLDER>"
.\omh-m365.exe doctor
.\omh-m365.exe auth status
.\omh-m365.exe auth login
.\omh-m365.exe outlook inbox --limit 10
.\omh-m365.exe teams chats --limit 20
.\omh-m365.exe files search --query "keyword" --limit 10
```

The executable should be distributed with `README.md`, `START_HERE_FOR_LLM.md`, `AGENTS.md`, `CLAUDE.md`, `policy.json`, `.env`, `.env.example`, and `docs/`.

For Claude Cowork, also distribute `CLAUDE_COWORK_RUNBOOK.md`.

`.env` is delegated public-client configuration and may be included in the release. LLMs must not read or print it; they should use `doctor` to check whether configuration is present.

LLMs should run the executable through PowerShell, Command Prompt, Windows Terminal, Codex shell, or another shell-enabled local tool. They should not use File Explorer and double-click the exe. Claude Cowork may use `computer-use`, but only to open a terminal and type these commands. The `.cmd` files below are fallback launchers for environments that can launch files but cannot pass command-line arguments.

Also include:

- `RUN_FIRST_FOR_LLM.cmd`: runs `doctor` and `auth status` for GUI-style LLM environments that can launch files but cannot pass CLI arguments.
- `START_LOGIN_FOR_USER.cmd`: opens Microsoft device login and starts `auth login`; the human completes the code entry.

POC build:

```powershell
npm run build:exe:win
```

Output:

```text
releases/win-x64/omh-m365.exe
```

The POC executable is not code-signed. Production rollout should use signing and an installer or endpoint-management deployment.

## Developer ZIP Contents

The ZIP is for developer handoff, not non-developer rollout. It should include:

- `dist/`
- `src/`
- `docs/`
- `scripts/`
- `package.json`
- `package-lock.json`
- `tsconfig.json`
- `.env.example`
- `.env`
- `.gitignore`
- `policy.json`
- `config.example.json`
- `README.md`
- `START_HERE_FOR_LLM.md`
- `AGENTS.md`
- `CLAUDE.md`

The ZIP must not include:

- `.cache/`
- `downloads/`
- `logs/`
- `node_modules/`
- real tokens, device codes, cookies, secrets, or credentials

## Build And Package

From the project root:

```powershell
npm install
npm run package:win
```

The package script runs typecheck/build and creates a ZIP under `releases/`.

## Install From ZIP

Unzip the release package, then run:

```powershell
npm run setup:win
```

The release may already include `.env` with delegated public-client configuration. If it is missing, fill `.env` locally with approved Azure Application client and tenant identifiers.

Run login:

```powershell
npm run start -- auth login
```

Run smoke test:

```powershell
npm run smoke:win
```

## LLM Use

LLMs should read `AGENTS.md` or `CLAUDE.md` before calling this tool. These files define the safe command surface and data-handling rules.

For real login testing, an LLM may open a separate terminal and the Microsoft device login page, but must not copy the device code into chat.

## Future Packaging Options

- Private npm package: best next step for versioned pilot rollout.
- `npx` or `npm exec`: useful for controlled execution when registry auth is available.
- Installer: best for non-developer users.
- Node Single Executable Application: possible later, but requires separate Windows/macOS build, signing, update, and endpoint security validation.

Avoid building the long-term plan around deprecated `pkg`-style packaging unless a maintained fork is explicitly selected and approved.
