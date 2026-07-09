#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
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
const packageVersion = readPackageVersion();
const repoUrl = "https://github.com/ohmyhotelco-planning/hare-m365-agent.git";
const defaultWorkDir = "/tmp/hare-m365-agent";
const defaultCliCommand = "node dist/cli.js";

const requiredDomains = ["github.com", "graph.microsoft.com", "login.microsoftonline.com"];

function readPackageVersion(): string {
  const packageJsonPath = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "package.json");
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8")) as { version?: string };
  return packageJson.version ?? "0.0.0";
}

const llmGuide = `# Hare M365 Agent LLM Guide

목적: 사용자의 자연어 요청을 Microsoft Graph delegated 권한으로 조회한다. 기본 정책은 read-only다.

시작 순서:
1. Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 3개 도메인 허용 여부를 사용자에게 확인한다.
   - github.com
   - graph.microsoft.com
   - login.microsoftonline.com
2. 도메인 확인은 가능하면 LLM 도구의 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리한다. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 만들지 않는다.
3. 확인 요청 UI가 없는 환경에서만 짧게 yes/no로 묻는다.
4. 도메인 허용 확인 전에는 git, npm, curl, 로그인, Outlook/Teams/Files 조회를 시작하지 않는다.
5. GitHub API(api.github.com)나 GitHub Release asset으로 repo 접근을 판단하지 않는다. Cowork에서는 api.github.com이 막혀도 git clone은 성공할 수 있다.
6. 실제 사용 경로인 git ls-remote 또는 git clone으로 repo 접근을 판단한다.

준비 명령:
\`\`\`bash
git ls-remote ${repoUrl} HEAD
rm -rf ${defaultWorkDir}
git clone ${repoUrl} ${defaultWorkDir}
cd ${defaultWorkDir}
npm ci
npm run build
node dist/cli.js
\`\`\`

npm ci가 npm registry 접근 오류로 실패하면 registry.npmjs.org 허용이 추가로 필요하다고 보고하고 멈춘다.

로그인 하드게이트:
- 시작 명령 또는 auth status 결과에서 loggedIn이 false이면 Outlook/Teams/Files 조회를 실행하지 않는다.
- loggedIn 값은 출력 JSON의 dataDir/cacheFile 기준 상태다. hosted sandbox나 컨테이너의 dataDir에서 나온 false를 사용자 PC 로그인 실패로 해석하지 않는다.
- 사용자가 이미 로그인했다고 말했는데 dataDir가 /sessions, /tmp 같은 임시 샌드박스 경로라면 로그인 실패가 아니라 파일시스템 저장 위치 불일치로 판단한다.
- 그 경우 이미 git clone/build가 끝났다면 소스 분석이나 네트워크 진단으로 빠지지 말고, 출력 JSON의 cacheFile 또는 OS별 고정 Hare 저장 위치 접근 권한을 요청해 로그인 캐시 파일 존재 여부를 확인한다.
- cacheFile이 있으면 로그인 명령을 다시 안내하지 말고 요청된 조회 명령을 진행한다.
- cacheFile이 없을 때만 humanLoginCommand를 사용자에게 안내하고 멈춘다.
- LLM은 humanLoginCommand를 직접 실행하지 않는다.
- 사용자는 humanLoginCommand를 직접 실행한다.
- 사용자에게 다음 로그인 절차를 짧게 안내한다: 명령 실행 -> 브라우저에서 코드 입력 -> 회사 Microsoft 계정 로그인 -> 완료 후 "로그인 완료"라고 말하기.
- 브라우저가 자동으로 열리지 않으면 터미널에 표시된 Microsoft 로그인 주소를 Ctrl+클릭하거나 주소창에 직접 붙여 넣도록 안내한다. Mac에서는 Command+클릭도 가능하다고 안내한다.
- 사용자가 "로그인 완료"라고 말하면 다시 시작 명령 또는 auth status를 실행한 뒤 원래 요청을 이어간다.

저장 위치:
- Hare는 항상 status.dataDir에 표시되는 고정 저장 위치를 사용한다.
- 임의의 별도 저장 폴더를 만들지 않는다.
- LLM 실행환경에서 사용자 PC의 고정 Hare 폴더 접근 권한이 없으면 선택지를 묻지 말고 그 고정 폴더 저장 접근 허용을 요청한다.
- Cowork/hosted sandbox에서 고정 Hare 폴더가 마운트되면 이후 모든 Hare 명령 앞에 HARE_M365_DATA_DIR="<마운트된 Hare 폴더 경로>"를 붙여 같은 캐시를 사용한다.
- 접근 허용 후 같은 HARE_M365_DATA_DIR 값으로 다시 시작 명령 또는 auth status를 실행한다.

hosted sandbox Graph 실패 때:
- cacheFile이 있는데 Outlook/Teams/Files 조회가 fetch failed, network_error 등으로 실패하면 토큰 문제가 아니라 sandbox egress 문제로 판단한다.
- 이때 clone/build를 반복하지 말고, 도메인 사진 확인이나 반복 호출 루프에 들어가지 않는다.
- 사용자의 Windows/Mac 로컬 환경에서 같은 Hare 명령을 실행하고 --out으로 status.dataDir에 결과 JSON을 저장하게 한다.
- 예: node dist/cli.js outlook inbox --limit 1 --out latest-mail.json
- --out 값이 상대 경로이면 Hare 고정 dataDir 아래에 저장된다.
- 결과 파일이 생성되면 그 JSON을 읽어 사용자 요청을 처리한다.

읽기 명령:
\`\`\`bash
node dist/cli.js outlook inbox --limit 10 --out latest-mail.json
node dist/cli.js teams teams --out teams.json
node dist/cli.js teams chats --limit 20 --out chats.json
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20 --out chat-messages.json
node dist/cli.js files search --query "keyword" --limit 10 --out files.json
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
\`\`\`

주의:
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경은 수행하지 않는다.
- Microsoft device code, 토큰 캐시, 쿠키, credential 값은 채팅에 붙여 넣거나 출력하지 않는다.
- Teams 최신 채팅 판단은 lastMessageCreatedDateTime을 우선한다. lastUpdatedDateTime만으로 최신 메시지 방을 판단하지 않는다.
`;

const llmPrompt = `목적: Hare M365 Agent로 내 Microsoft 365 업무 요청을 처리해.

저장소:
${repoUrl}

Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 3개 도메인이 허용되어 있는지 확인해.
github.com
graph.microsoft.com
login.microsoftonline.com

도메인 허용 확인은 가능하면 네가 가진 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리해. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 요구하지 마.
그런 확인 요청 기능이 없는 환경에서만 짧게 yes/no로 물어봐.

도메인 허용 확인 전에는 git, npm, curl, 로그인, Outlook/Teams/Files 조회를 시작하지 마.
GitHub API(api.github.com)나 GitHub Release asset 다운로드로 repo 접근을 판단하지 마.
이 환경에서는 api.github.com이 403이어도 git clone은 성공할 수 있어.
repo 접근은 git ls-remote 또는 git clone으로 판단해.

아래 순서로 준비해.

git ls-remote ${repoUrl} HEAD
rm -rf ${defaultWorkDir}
git clone ${repoUrl} ${defaultWorkDir}
cd ${defaultWorkDir}
npm ci
npm run build
node dist/cli.js

npm ci가 npm registry 접근 오류로 실패하면 registry.npmjs.org 허용이 추가로 필요하다고 보고하고 멈춰.

결과에서 loggedIn이 true이면 바로 내 요청을 처리해.
loggedIn이 false이면 결과 JSON의 dataDir/cacheFile 기준으로 판단해.
단, loggedIn 값은 결과 JSON의 dataDir/cacheFile 기준 상태야. hosted sandbox나 컨테이너 경로에서 나온 false를 사용자 PC 로그인 실패로 해석하지 마.
내가 이미 로그인했다고 말했는데 dataDir가 /sessions, /tmp 같은 임시 샌드박스 경로라면 로그인 실패가 아니라 파일시스템 저장 위치 불일치로 판단해.
그 경우 이미 git clone/build가 끝났다면 추가 소스 분석이나 네트워크 진단으로 빠지지 말고, 결과 JSON의 cacheFile 또는 OS별 고정 Hare 저장 위치 접근 권한을 요청해서 로그인 캐시 파일 존재 여부부터 확인해.
내가 이미 로그인했다고 말했거나 cacheFile이 있으면 로그인 명령을 다시 안내하지 말고 요청된 조회 명령을 진행해.
cacheFile이 없을 때만 humanLoginCommand를 사용자에게 안내하고 즉시 멈춰.
humanLoginCommand는 네가 실행하지 마. 사용자에게 안내하고 즉시 멈춰.
사용자에게 humanLoginCommand를 복사해 직접 실행하라고 안내해.
사용자가 해야 할 로그인 절차도 같이 안내해: 명령 실행 -> 브라우저에서 코드 입력 -> 회사 Microsoft 계정 로그인 -> 완료 후 "로그인 완료"라고 말하기.
브라우저가 자동으로 열리지 않으면 터미널에 표시된 Microsoft 로그인 주소를 Ctrl+클릭하거나 주소창에 직접 붙여 넣으라고 안내해. Mac에서는 Command+클릭도 가능하다고 안내해.
git clone/build 이후에는 소스 재확인, 도메인 사진, network_error 원인 분석 루프를 시작하지 마.
내가 "로그인 완료"라고 말하면 다시 시작 명령 또는 auth status를 확인하고 원래 요청을 이어서 처리해.

Hare는 고정 저장 위치만 사용해. 임의의 별도 저장 폴더를 만들지 마.
네 실행환경에서 사용자 PC의 고정 Hare 폴더 접근 권한이 없으면 선택지를 묻지 말고 그 고정 폴더 접근 허용을 요청해.
Cowork/hosted sandbox에서 고정 Hare 폴더가 마운트되면 이후 모든 Hare 명령 앞에 HARE_M365_DATA_DIR="<마운트된 Hare 폴더 경로>"를 붙여 같은 캐시를 사용해.
허용되면 같은 HARE_M365_DATA_DIR 값으로 다시 시작 명령 또는 auth status를 실행하고 이어서 처리해.

cacheFile이 있는데 Outlook/Teams/Files 조회가 fetch failed 또는 network_error로 실패하면 sandbox 네트워크 문제로 판단해.
clone/build를 반복하지 말고, 도메인 사진확인이나 반복 호출 루프를 시작하지 마.
사용자의 Windows/Mac 로컬 환경에서 같은 Hare 명령을 실행하고 --out으로 status.dataDir에 결과 JSON을 저장하게 해.
예: node dist/cli.js outlook inbox --limit 1 --out latest-mail.json
--out 값이 상대 경로이면 Hare 고정 dataDir 아래에 저장돼.
결과 파일이 생성되면 그 JSON을 읽어 내 요청을 처리해.`;
function getSelfCommand(): string {
  return process.env.HARE_M365_COMMAND ?? defaultCliCommand;
}

function getLoginCommand(): string {
  return `${getSelfCommand()} auth login`;
}

function resolveOutPath(out: string): string {
  return path.resolve(path.isAbsolute(out) ? out : path.join(config.dataDir, out));
}

function emitJson(payload: unknown, out?: string): void {
  if (!out) {
    console.log(JSON.stringify(payload, null, 2));
    return;
  }

  const outputPath = resolveOutPath(out);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, JSON.stringify(payload, null, 2), "utf8");
  console.log(JSON.stringify({ ok: true, outputPath }, null, 2));
}

async function getDoctorStatus() {
  const account = config.clientId && config.tenantId ? await getAccount(config).catch(() => null) : null;
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  return {
    configured: Boolean(config.clientId && config.tenantId),
    clientIdPresent: Boolean(config.clientId),
    tenantIdPresent: Boolean(config.tenantId),
    loggedIn: Boolean(account),
    dataDir: config.dataDir,
    policyPath: config.policyPath,
    cacheDir: config.cacheDir,
    cacheFile,
    cacheFileExists: fs.existsSync(cacheFile),
    downloadDir: config.downloadDir,
    logsDir: config.logsDir
  };
}

function loginGateFields(loggedIn: boolean, cacheFileExists = false) {
  if (loggedIn) return {};
  return {
    nextCommand: "LOGIN_REQUIRED_HARD_GATE",
    llmAction: "STOP_DO_NOT_RUN_LOGIN_COMMAND",
    humanLoginCommand: getLoginCommand(),
    instruction:
      cacheFileExists
        ? "loggedIn is only the state for the printed dataDir/cacheFile. cacheFileExists is true, so do not tell the human to login again. If git clone/build was already completed, continue with the requested Hare read command. Do not start an unbounded source audit or network diagnostics loop."
        : "loggedIn is only the state for the printed dataDir/cacheFile. Do not treat sandbox false as the user's PC login state. If the human already logged in and dataDir is a temporary sandbox path, request access to the user's fixed Hare data folder, set HARE_M365_DATA_DIR to the mounted folder path for every Hare command, and check the cache file first. If no cache file exists, tell the human to copy and run humanLoginCommand, enter the Microsoft device code in the browser, then say login complete. Then stop. Do not run auth login or start an unbounded source audit/network diagnostics loop. After the human confirms login is complete, rerun startup or auth status with the same HARE_M365_DATA_DIR."
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
          version: packageVersion,
          mode: "startup",
          repository: repoUrl,
          setupCommand: `git clone ${repoUrl} ${defaultWorkDir} && cd ${defaultWorkDir} && npm ci && npm run build && node dist/cli.js`,
          requiredDomains,
          conditionalDomain: "registry.npmjs.org is required only if npm ci cannot reach npm registry.",
          status,
          storageRule:
            "Use the fixed Hare data folder only. In Cowork/hosted sandbox, mount the user's fixed Hare folder and set HARE_M365_DATA_DIR to that mounted path for every Hare command.",
          ...(status.configured ? loginGateFields(status.loggedIn, status.cacheFileExists) : {}),
          nextCommand: status.configured
            ? status.loggedIn
              ? `${getSelfCommand()} auth status`
              : "LOGIN_REQUIRED_HARD_GATE"
            : "Create .env from .env.example and fill approved Azure Application values.",
          llmAction: status.configured && !status.loggedIn ? "STOP_DO_NOT_RUN_LOGIN_COMMAND" : undefined,
          humanLoginCommand: status.configured && !status.loggedIn ? getLoginCommand() : undefined
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
  const loggedIn = Boolean(account);
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  console.log(
    JSON.stringify(
      {
        loggedIn,
        account: account?.username,
        policy: config.policy,
        dataDir: config.dataDir,
        cacheDir: config.cacheDir,
        cacheFile,
        cacheFileExists: fs.existsSync(cacheFile),
        downloadDir: config.downloadDir,
        logsDir: config.logsDir,
        ...loginGateFields(loggedIn, fs.existsSync(cacheFile))
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
  .description("Print the LLM usage guide for Hare M365 Agent")
  .action(() => {
    console.log(llmGuide);
  });

program
  .command("llm-prompt")
  .description("Print a short first prompt for an LLM session")
  .action(() => {
    console.log(llmPrompt);
  });

const outlook = program.command("outlook").description("Outlook read commands");

outlook
  .command("inbox")
  .description("List recent Inbox messages")
  .option("--limit <number>", "maximum message count", "10")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare dataDir")
  .action(async (options: { limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listInbox(config, Number(options.limit));
    emitJson({ messages: data }, options.out);
  });

const teams = program.command("teams").description("Teams read commands");

teams
  .command("teams")
  .description("List joined teams")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare dataDir")
  .action(async (options: { out?: string }) => {
    requireConfigured(config);
    const data = await listJoinedTeams(config);
    emitJson({ teams: data }, options.out);
  });

teams
  .command("chats")
  .description("List recent chats")
  .option("--limit <number>", "maximum chat count", "20")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare dataDir")
  .action(async (options: { limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listChats(config, Number(options.limit));
    emitJson({ chats: data }, options.out);
  });

teams
  .command("chat-messages")
  .description("List messages in one chat")
  .requiredOption("--chat-id <id>", "chat ID returned by teams chats")
  .option("--limit <number>", "maximum message count", "20")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare dataDir")
  .action(async (options: { chatId: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listChatMessages(config, options.chatId, Number(options.limit));
    emitJson({ messages: data }, options.out);
  });

const files = program.command("files").description("SharePoint/OneDrive file commands");

files
  .command("search")
  .description("Search files visible to the signed-in user")
  .requiredOption("--query <text>", "search query")
  .option("--limit <number>", "maximum file count", "10")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare dataDir")
  .action(async (options: { query: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await searchFiles(config, options.query, Number(options.limit));
    emitJson({ files: data }, options.out);
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
