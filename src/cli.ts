#!/usr/bin/env node
import { Command } from "commander";
import { getAccount, getScopeList, login, logout } from "./auth.js";
import { ensureRuntimeDirs, loadConfig, requireConfigured } from "./config.js";
import { listInbox } from "./outlook.js";
import { downloadDriveItem, searchFiles } from "./sharepoint.js";
import { listChatMessages, listChats, listJoinedTeams } from "./teams.js";

const program = new Command();
const config = loadConfig();
ensureRuntimeDirs(config);

const preferredCommand = "hare-m365";
const packageName = "@ohmyhotel/hare-m365-agent";
const packageVersion = "0.1.0";
const defaultPackageSpec =
  "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz";

const llmGuide = `# Hare M365 Agent LLM Guide

Hare M365 Agent는 사용자의 로컬 shell에서 Microsoft 365를 읽기 전용으로 조회하기 위한 도구입니다.

기본 원칙:

- 사용자가 자연어로 Outlook, Teams, OneDrive, SharePoint 조회를 요청하면 이 CLI를 사용합니다.
- 먼저 doctor와 auth status를 실행합니다.
- Claude Cowork처럼 폴더가 연결된 환경에서는 같은 연결 폴더 안의 runtime 폴더를 인증/다운로드/로그 저장소로 사용합니다.
- 이 경우 모든 Hare 명령에 HARE_M365_DATA_DIR=./runtime을 붙입니다.
- configured가 true이고 loggedIn이 true일 때만 Outlook, Teams, Files 명령을 실행합니다.
- loggedIn이 false이면 로그인 hard gate에서 멈춥니다. Outlook, Teams, Files 조회를 실행하지 않습니다.
- Claude Cowork, hosted sandbox, 명령 출력이 주로 LLM에게만 보이는 환경에서는 auth login을 자동 실행하지 않습니다.
- 로그인 hard gate에서는 사용자가 직접 볼 수 있는 로컬 터미널에서 실행할 정확한 auth login 명령을 안내하고 멈춥니다. 연결 폴더를 쓰는 경우 사용자가 같은 폴더에서 같은 HARE_M365_DATA_DIR=./runtime 설정으로 로그인해야 합니다.
- 사용자가 명시적으로 로그인 시작을 요청했고 현재 터미널/브라우저 출력이 사용자에게 직접 보이는 환경에서만 auth login을 1회 실행할 수 있습니다. 실패하면 반복 재시도하지 말고 오류를 보고하고 멈춥니다.
- auth login 중 터미널에 표시되는 Microsoft device code는 사용자가 직접 브라우저에 입력합니다. LLM은 코드를 채팅으로 복사하거나 반복하지 않습니다.
- .env, .cache, runtime/.cache, token, device code, cookie, credential 값은 읽거나 출력하지 않습니다.

연결형 runtime 원칙:

- CLI 프로세스는 runtime/.cache/msal-cache.json을 사용해 Microsoft Graph 인증을 수행할 수 있습니다.
- LLM은 runtime/.cache/msal-cache.json 파일을 열거나 출력하거나 요약하거나 업로드하지 않습니다.
- 다른 로컬 폴더나 다른 머신에서 로그인한 토큰은 현재 연결된 실행환경에서 재사용되지 않을 수 있습니다.

연결 폴더에서 실행하는 예:

\`\`\`bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "<GITHUB_RELEASE_TGZ_URL>" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "<GITHUB_RELEASE_TGZ_URL>" -- hare-m365 auth status
\`\`\`

Claude/Cowork처럼 도메인 허용 목록이 있는 환경에서는 아래 도메인이 필요할 수 있습니다.

- github.com
- release-assets.githubusercontent.com
- registry.npmjs.org
- graph.microsoft.com
- login.microsoftonline.com

패키지 다운로드나 Graph 호출이 403, network_error, proxy, allowlist 오류로 실패하면 막힌 도메인/오류를 사용자에게 보고하고 중단합니다. Windows GUI, 파일 탐색기, 더블클릭 실행으로 우회하지 않습니다.

도메인 허용 목록이 있는 환경에서 사용자가 아직 도메인 허용 완료를 확인하지 않았다면 npm exec, npx, 다운로드, 로그인, Microsoft 365 조회를 시작하지 않습니다.

도메인 허용 완료 전에는 폴더 연결 요청도 시작하지 않습니다. 도메인 허용 후 먼저 현재 실행환경에서 llm-guide, doctor, auth status를 시도합니다. 그 과정에서 폴더 접근 권한 또는 지속 인증 캐시 위치가 필요하다고 확인될 때만, 가능한 도구로 문서/Documents 안의 Hare M365 Agent 전용 폴더를 먼저 준비한 뒤 Claude/Cowork의 폴더 연결 기능으로 그 폴더 연결/허용을 요청합니다. 폴더 존재 여부를 미리 검사해서 사용자에게 만들라고 요구하지 않습니다. 사용자가 할 일은 연결 허용뿐이라는 전제로 진행합니다. 폴더 선택창이 뜬 경우에만 SharePoint 공용 폴더, Teams 공유 폴더, 다운로드 폴더 전체, 바탕화면 전체는 선택하지 말라고 안내합니다.

GitHub Release tarball로 실행 중이면 모든 명령에 같은 npm exec 접두어를 사용합니다.

\`\`\`bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "<GITHUB_RELEASE_TGZ_URL>" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "<GITHUB_RELEASE_TGZ_URL>" -- hare-m365 auth status
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "<GITHUB_RELEASE_TGZ_URL>" -- hare-m365 auth login
\`\`\`

npm registry에 공개 배포된 경우에는 npx를 사용할 수 있습니다.

\`\`\`bash
npx ${packageName} doctor
npx ${packageName} auth status
\`\`\`

전역 설치된 경우에는 hare-m365를 직접 사용할 수 있습니다.

\`\`\`bash
hare-m365 doctor
hare-m365 auth status
\`\`\`

허용된 읽기 명령:

\`\`\`bash
HARE_M365_DATA_DIR=./runtime hare-m365 doctor
HARE_M365_DATA_DIR=./runtime hare-m365 auth status
HARE_M365_DATA_DIR=./runtime hare-m365 auth login
HARE_M365_DATA_DIR=./runtime hare-m365 outlook inbox --limit 10
HARE_M365_DATA_DIR=./runtime hare-m365 teams teams
HARE_M365_DATA_DIR=./runtime hare-m365 teams chats --limit 20
HARE_M365_DATA_DIR=./runtime hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
HARE_M365_DATA_DIR=./runtime hare-m365 files search --query "keyword" --limit 10
HARE_M365_DATA_DIR=./runtime hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
\`\`\`

안전 규칙:

- 기본 정책은 read-only입니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경은 수행하지 않습니다.
- 파일 다운로드는 사용자가 특정 파일을 명시적으로 요청했을 때만 수행합니다.
- 원문 전체 덤프보다 필요한 메타데이터와 요약을 우선합니다.

Teams 최신성 규칙:

- teams chats 결과의 lastMessageCreatedDateTime으로 최신 채팅을 판단합니다.
- lastUpdatedDateTime은 채팅방 메타데이터 변경 시각이므로 최신 메시지 판단에 사용하지 않습니다.
`;

const llmPrompt = `Hare M365 Agent를 사용해 내 Microsoft 365 요청을 처리해. Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 github.com, release-assets.githubusercontent.com, registry.npmjs.org, graph.microsoft.com, login.microsoftonline.com 허용 여부를 사용자에게 확인해. 도메인 허용 완료 전에는 npm exec, 조회, 로그인, 폴더 연결 요청을 시작하지 마. 도메인 허용 후 먼저 현재 실행환경에서 llm-guide, doctor, auth status를 시도해. 그 과정에서 폴더 접근 권한 또는 지속 인증 캐시 위치가 필요하다고 확인될 때만, 가능한 도구로 문서/Documents 안의 Hare M365 Agent 전용 폴더를 먼저 준비한 뒤 Claude/Cowork의 폴더 연결 기능으로 그 폴더 연결/허용을 요청해. 폴더 존재 여부를 미리 검사해서 사용자에게 만들라고 요구하지 마. 사용자가 할 일은 연결 허용뿐이라는 전제로 진행해. 연결 폴더가 있으면 같은 폴더 안의 ./runtime을 인증 캐시 위치로 쓰고 모든 Hare 명령에 HARE_M365_DATA_DIR=./runtime을 붙여. CLI는 runtime/.cache/msal-cache.json을 사용할 수 있지만, 너는 그 파일을 열거나 출력하거나 요약하지 마. 이후 doctor/auth status로 설정과 로그인을 확인해. loggedIn이 false이면 로그인 hard gate로 멈추고 Outlook/Teams/Files 조회를 실행하지 마. Cowork/샌드박스에서는 auth login을 자동 실행하지 말고, 사용자가 같은 연결 폴더에서 같은 HARE_M365_DATA_DIR=./runtime 설정으로 실행할 로그인 명령을 안내해. 사용자가 "로그인 완료"라고 말하면 doctor/auth status를 다시 확인하고 원래 요청을 이어서 처리해. .env, .cache, runtime/.cache, token, device code는 읽거나 출력하지 마.`;

function getSelfCommand(): string {
  const packageSpec = process.env.HARE_M365_PACKAGE_SPEC ?? process.env.HARE_M365_PACKAGE_URL;
  if (packageSpec) return `npm exec --yes --package "${packageSpec}" -- ${preferredCommand}`;

  const script = process.argv[1] ?? "";
  const executable = process.execPath.toLowerCase();
  if (executable.endsWith("omh-m365.exe")) return ".\\omh-m365.exe";
  if (script.endsWith("omh-m365.cjs")) return "node omh-m365.cjs";
  return `npm exec --yes --package "${defaultPackageSpec}" -- ${preferredCommand}`;
}

function getLoginCommand(): string {
  const packageSpec = process.env.HARE_M365_PACKAGE_SPEC ?? process.env.HARE_M365_PACKAGE_URL;
  if (packageSpec) return withRuntimeEnv(`npm exec --yes --package "${packageSpec}" -- ${preferredCommand} auth login`);

  const script = process.argv[1] ?? "";
  const executable = process.execPath.toLowerCase();
  if (executable.endsWith("omh-m365.exe")) return withRuntimeEnv(".\\START_LOGIN_FOR_USER.cmd");
  if (script.endsWith("omh-m365.cjs")) return withRuntimeEnv("./START_LOGIN_FOR_USER.sh");
  return withRuntimeEnv(`npm exec --yes --package "${defaultPackageSpec}" -- ${preferredCommand} auth login`);
}

function renderLlmGuide(): string {
  const packageSpec = process.env.HARE_M365_PACKAGE_SPEC ?? process.env.HARE_M365_PACKAGE_URL ?? defaultPackageSpec;
  return llmGuide.replaceAll("<GITHUB_RELEASE_TGZ_URL>", packageSpec);
}

function renderLlmPrompt(): string {
  const packageSpec = process.env.HARE_M365_PACKAGE_SPEC ?? process.env.HARE_M365_PACKAGE_URL ?? defaultPackageSpec;
  return `패키지 URL: ${packageSpec}\n\n${llmPrompt}`;
}

function quotePosix(value: string): string {
  return `'${value.replace(/'/g, "'\\''")}'`;
}

function quotePowerShell(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function withRuntimeEnv(command: string): string {
  const runtimeDir = process.env.HARE_M365_DATA_DIR ?? process.env.OMH_M365_DATA_DIR;
  if (!runtimeDir) return command;
  if (process.platform === "win32") return `$env:HARE_M365_DATA_DIR = ${quotePowerShell(runtimeDir)}; ${command}`;
  return `HARE_M365_DATA_DIR=${quotePosix(runtimeDir)} ${command}`;
}

function getRuntimeHint(): string {
  const runtimeDir = process.env.HARE_M365_DATA_DIR ?? process.env.OMH_M365_DATA_DIR;
  if (runtimeDir) return `Use the same runtime data directory for every command: ${runtimeDir}`;
  return "For folder-mounted LLM environments, rerun Hare commands from the connected folder with HARE_M365_DATA_DIR=./runtime.";
}

async function getDoctorStatus() {
  const account = config.clientId && config.tenantId ? await getAccount(config).catch(() => null) : null;
  return {
    configured: Boolean(config.clientId && config.tenantId),
    clientIdPresent: Boolean(config.clientId),
    tenantIdPresent: Boolean(config.tenantId),
    loggedIn: Boolean(account),
    dataDir: config.dataDir,
    policyPath: config.policyPath,
    cacheDir: config.cacheDir,
    downloadDir: config.downloadDir,
    logsDir: config.logsDir
  };
}

program
  .name(preferredCommand)
  .description("Hare M365 Agent CLI for LLM-driven Microsoft 365 Graph access.")
  .version(packageVersion)
  .action(async () => {
    const status = await getDoctorStatus();
    console.log(
      JSON.stringify(
        {
          tool: preferredCommand,
          package: packageName,
          mode: "startup",
          status,
          nextCommand: status.configured
            ? status.loggedIn
              ? `${getSelfCommand()} auth status`
              : "LOGIN_REQUIRED_HARD_GATE"
            : "Create .env from .env.example and fill approved Azure Application values.",
          loginCommand: status.configured && !status.loggedIn ? getLoginCommand() : undefined,
          runtimeHint: getRuntimeHint(),
          llmInstruction:
            "Read llm-guide, confirm configured/loggedIn status, and use the smallest safe read command needed for the user's request. If loggedIn is false, stop at the login hard gate and tell the human to run loginCommand in the same connected folder/runtime they expect the LLM to use. Do not auto-run auth login in Cowork or hosted sandbox shells. The CLI may use runtime/.cache/msal-cache.json, but the LLM must never read or print .env, .cache, runtime/.cache, tokens, or device codes."
        },
        null,
        2
      )
    );
  });

const auth = program.command("auth").description("Authentication commands");

auth.command("login").description("Sign in with Microsoft device code flow").action(async () => {
  requireConfigured(config);
  const result = await login(config);
  console.log(
    JSON.stringify(
      {
        ok: true,
        account: result.account?.username,
        tenantId: result.tenantId,
        scopes: getScopeList()
      },
      null,
      2
    )
  );
});

auth.command("status").description("Show current login and policy status").action(async () => {
  requireConfigured(config);
  const account = await getAccount(config);
  console.log(
    JSON.stringify(
      {
        loggedIn: Boolean(account),
        account: account?.username,
        policy: config.policy,
        dataDir: config.dataDir,
        cacheDir: config.cacheDir,
        downloadDir: config.downloadDir,
        logsDir: config.logsDir
      },
      null,
      2
    )
  );
});

auth.command("logout").description("Delete local token cache").action(() => {
  logout(config);
  console.log(JSON.stringify({ ok: true, message: "Local token cache removed." }, null, 2));
});

program
  .command("doctor")
  .description("Check local configuration without reading token contents")
  .action(async () => {
    console.log(JSON.stringify(await getDoctorStatus(), null, 2));
  });

program
  .command("llm-guide")
  .description("Print the safe LLM usage guide for Hare M365 Agent")
  .action(() => {
    console.log(renderLlmGuide());
  });

program
  .command("llm-prompt")
  .description("Print a short first prompt for an LLM session")
  .action(() => {
    console.log(renderLlmPrompt());
  });

const outlook = program.command("outlook").description("Outlook read commands");

outlook
  .command("inbox")
  .description("List recent Inbox messages")
  .option("--limit <number>", "maximum message count", "10")
  .action(async (options: { limit: string }) => {
    requireConfigured(config);
    const data = await listInbox(config, Number(options.limit));
    console.log(JSON.stringify({ messages: data }, null, 2));
  });

const teams = program.command("teams").description("Teams read commands");

teams.command("teams").description("List joined teams").action(async () => {
  requireConfigured(config);
  const data = await listJoinedTeams(config);
  console.log(JSON.stringify({ teams: data }, null, 2));
});

teams
  .command("chats")
  .description("List recent chats")
  .option("--limit <number>", "maximum chat count", "20")
  .action(async (options: { limit: string }) => {
    requireConfigured(config);
    const data = await listChats(config, Number(options.limit));
    console.log(JSON.stringify({ chats: data }, null, 2));
  });

teams
  .command("chat-messages")
  .description("List messages in one chat")
  .requiredOption("--chat-id <id>", "chat ID returned by teams chats")
  .option("--limit <number>", "maximum message count", "20")
  .action(async (options: { chatId: string; limit: string }) => {
    requireConfigured(config);
    const data = await listChatMessages(config, options.chatId, Number(options.limit));
    console.log(JSON.stringify({ messages: data }, null, 2));
  });

const files = program.command("files").description("SharePoint/OneDrive file commands");

files
  .command("search")
  .description("Search files visible to the signed-in user")
  .requiredOption("--query <text>", "search query")
  .option("--limit <number>", "maximum file count", "10")
  .action(async (options: { query: string; limit: string }) => {
    requireConfigured(config);
    const data = await searchFiles(config, options.query, Number(options.limit));
    console.log(JSON.stringify({ files: data }, null, 2));
  });

files
  .command("download")
  .description("Download one file by driveId and itemId")
  .requiredOption("--drive-id <id>", "drive ID from parentReference.driveId")
  .requiredOption("--item-id <id>", "drive item ID")
  .option("--name <filename>", "output filename")
  .action(async (options: { driveId: string; itemId: string; name?: string }) => {
    requireConfigured(config);
    const outputPath = await downloadDriveItem(config, options.driveId, options.itemId, options.name);
    console.log(JSON.stringify({ ok: true, outputPath }, null, 2));
  });

program.parseAsync(process.argv).catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(JSON.stringify({ ok: false, error: message }, null, 2));
  process.exitCode = 1;
});
