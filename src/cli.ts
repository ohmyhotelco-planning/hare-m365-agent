#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Command } from "commander";
import { completeLogin, getAuthStatus, getScopeList, startLogin, logout } from "./auth.js";
import { ensureRuntimeDirs, loadConfig, requireConfigured } from "./config.js";
import { deviceLoginStatePath } from "./device-login.js";
import {
  countMailboxMessages,
  listInbox,
  searchMailbox,
  type MailFolderScope
} from "./outlook.js";
import { downloadDriveItem, searchFiles, searchSites } from "./sharepoint.js";
import { listChatMessages, listChats, listJoinedTeams, searchChatMessages } from "./teams.js";
import { cleanupExpiredResults, resolveResultPath } from "./results.js";

const program = new Command();
const config = loadConfig();
ensureRuntimeDirs(config);
try {
  cleanupExpiredResults(config);
} catch {
  // Result cleanup is best-effort and must not block M365 commands.
}

const preferredCommand = "hare-m365";
const packageName = "@ohmyhotel/hare-m365-agent";
const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageVersion = readPackageVersion();
const repoUrl = "https://github.com/ohmyhotelco-planning/hare-m365-agent.git";
const defaultWorkDir = "/tmp/hare-m365-agent";
const defaultCliCommand = `node "${path.join(packageRoot, "dist", "cli.js")}"`;

const requiredDomains = [
  "github.com",
  "registry.npmjs.org",
  "graph.microsoft.com",
  "login.microsoftonline.com",
  "ohmylab-my.sharepoint.com",
  "ohmylab.sharepoint.com"
];

function readPackageVersion(): string {
  const packageJsonPath = path.join(packageRoot, "package.json");
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8")) as { version?: string };
  return packageJson.version ?? "0.0.0";
}

const llmGuide = `# Hare M365 Agent LLM Guide

목적: 사용자의 자연어 요청을 Microsoft Graph delegated 권한으로 조회한다. 기본 정책은 read-only다.

시작 순서:
1. Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 6개 도메인 허용 여부를 사용자에게 확인한다.
   - github.com
   - registry.npmjs.org
   - graph.microsoft.com
   - login.microsoftonline.com
   - ohmylab-my.sharepoint.com
   - ohmylab.sharepoint.com
2. 도메인 확인은 가능하면 LLM 도구의 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리한다. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 만들지 않는다.
3. 확인 요청 UI가 없는 환경에서만 짧게 yes/no로 묻는다.
4. 도메인 허용 확인 전에는 git, npm, curl, 로그인, Outlook/Teams/Files 조회를 시작하지 않는다.
5. GitHub API(api.github.com)나 GitHub Release asset으로 repo 접근을 판단하지 않는다. Cowork에서는 api.github.com이 막혀도 git clone은 성공할 수 있다.
6. 실제 사용 경로인 git ls-remote 또는 git clone으로 repo 접근을 판단한다.

준비 명령:
\`\`\`bash
REMOTE_HEAD=$(git ls-remote ${repoUrl} refs/heads/master | awk '{print $1}')
test -n "$REMOTE_HEAD"
rm -rf ${defaultWorkDir}
test ! -e ${defaultWorkDir}
git clone --branch master --single-branch --no-tags ${repoUrl} ${defaultWorkDir}
LOCAL_HEAD=$(git -C ${defaultWorkDir} rev-parse HEAD)
test "$LOCAL_HEAD" = "$REMOTE_HEAD"
test ! -d ${defaultWorkDir}/node_modules
cd ${defaultWorkDir} && npm ci --prefer-offline --no-audit --no-fund
cd ${defaultWorkDir} && npm run build
test -f ${defaultWorkDir}/dist/cli.js
test -f ${defaultWorkDir}/dist/proxy.js
test -f ${defaultWorkDir}/dist/msal-network.js
cd ${defaultWorkDir} && node dist/cli.js
\`\`\`

각 셸 호출은 작업 폴더가 초기화될 수 있으므로 모든 npm, build, Hare 명령에 동일한 ${defaultWorkDir} 작업 경로를 명시한다. 준비부터 로그인까지 새로 clone한 이 폴더만 사용한다. 어느 단계든 실패하면 기존 node_modules, dist, 다른 복사본으로 대체하지 말고 실패 단계와 오류 한 줄만 보고한다. npm 설치를 백그라운드로 전환하거나 npm install로 바꾸거나 여러 셸 호출에 나눠 반복하지 않는다.

로그인 하드게이트:
- 시작 명령 또는 auth status 결과에서 loggedIn/tokenUsable이 false이면 Outlook/Teams/Files 조회를 실행하지 않는다.
- dataDirPersistent가 false이면 Microsoft 로그인을 시작하지 않는다. 폴더 연결 도구로 사용자의 Documents 폴더 접근을 요청하고, 마운트된 폴더 안에 "Hare M365 Agent" 하위 폴더를 만든다.
- 이후 모든 Hare 명령에 HARE_M365_DATA_DIR="<마운트된 Documents>/Hare M365 Agent"를 사용한다. 사용자에게 경로를 입력하거나 선택지를 고르라고 요구하지 않는다.
- 같은 dataDir로 다시 시작 상태를 확인한 뒤 humanLoginCommand(auth login-start)를 한 번 실행한다. 이 명령은 로그인 주소와 userCode를 즉시 반환하고 종료된다.
- 주소와 코드를 사용자에게 보여주고 브라우저에서 회사 Microsoft 계정으로 로그인한 뒤 "로그인 완료"라고 말하도록 안내한다.
- 사용자가 완료를 알리면 humanLoginCompleteCommand(auth login-complete)를 새 포그라운드 셸 호출로 실행한다. 이 명령은 최대 25초만 폴링한다.
- poller를 계속 실행하거나 백그라운드, detached, setsid, nohup을 사용하지 않는다.
- login-complete가 authorization_pending이면 login-start를 다시 실행하지 말고 사용자 로그인 완료 여부를 확인한 뒤 login-complete만 재실행한다. 코드가 만료된 경우에만 login-start를 다시 실행한다.
- login-complete 성공 후 같은 dataDir에서 auth status를 실행해 loggedIn/tokenUsable이 모두 true인지 확인한다.
- cacheFile 존재만으로 로그인 성공으로 판단하지 않는다.

저장 위치:
- 로컬 Windows/Mac/Linux에서는 OS 기본 고정 dataDir를 사용한다.
- Cowork/hosted sandbox에서는 연결된 Documents 폴더 아래 "Hare M365 Agent"를 고정 dataDir로 사용한다.
- /sessions 또는 /tmp 기본 경로에서는 login-start가 거부된다.
- 연결 후 모든 Hare 명령에 동일한 HARE_M365_DATA_DIR 값을 사용한다.

hosted sandbox Graph 실패 때:
- cacheFile이 있는데 Outlook/Teams/Files 조회가 fetch failed, network_error 등으로 실패하면 토큰 문제가 아니라 sandbox egress 문제로 판단한다.
- 이때 clone/build를 반복하지 말고, 도메인 사진 확인이나 반복 호출 루프에 들어가지 않는다.
- 사용자의 Windows/Mac 로컬 환경에서 같은 Hare 명령을 실행하고 --out으로 status.resultsDir에 결과 JSON을 저장하게 한다.
- 예: node dist/cli.js outlook inbox --limit 1 --out latest-mail.json
- 일반 조회에서는 --out을 사용하지 않고 화면 출력을 바로 처리한다.
- --out 값이 상대 경로이면 Hare 고정 resultsDir 아래에 저장되며 7일 후 자동 정리된다.
- 결과 파일이 생성되면 그 JSON을 읽어 사용자 요청을 처리한다.

읽기 명령:
\`\`\`bash
node dist/cli.js outlook inbox --limit 10
node dist/cli.js outlook search --query "keyword" --since 2026-04-01 --until 2026-07-10 --folder all
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20
node dist/cli.js teams search-messages --query "keyword" --since 2026-04-01 --until 2026-07-10
node dist/cli.js sharepoint sites --query "Agent Automation"
node dist/cli.js files search --query "keyword" --limit 10
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
\`\`\`

주의:
- 사용자가 기간을 지정한 조회는 inbox/chat-messages의 최근 건수 제한으로 대신하지 말고 search 명령의 --since/--until에 반영한다.
- 사용자가 기간을 지정하지 않은 검색은 최근 90일을 조회한다. 결과 JSON의 search.range.notice를 사용자에게 알려 실제 조회 범위를 명확히 한다.
- search.limitReached가 true이면 일부 결과만 반환된 것이므로 사용자에게 한도 도달 사실을 알린다.
- 메일 건수 질문은 검색 인덱스 결과를 세지 말고 outlook count로 전체 페이지를 검사한다.
- SharePoint 사이트 존재 여부는 sharepoint sites로 확인한다. 개인 OneDrive만 검색하는 files search 결과로 사이트 존재나 접근 권한을 판단하지 않는다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경은 수행하지 않는다.
- login-start의 userCode와 verificationUri는 현재 사용자에게 한 번 보여준다. 내부 device_code, 토큰 캐시, 쿠키, credential 값은 출력하거나 반복하지 않는다.
- Teams 최신 채팅 판단은 lastMessageCreatedDateTime을 우선한다. lastUpdatedDateTime만으로 최신 메시지 방을 판단하지 않는다.
`;

const llmPrompt = `목적: Hare M365 Agent로 내 Microsoft 365 업무 요청을 처리해.

저장소:
${repoUrl}

Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 6개 도메인이 허용되어 있는지 확인해.
github.com
registry.npmjs.org
graph.microsoft.com
login.microsoftonline.com
ohmylab-my.sharepoint.com
ohmylab.sharepoint.com

도메인 허용 확인은 가능하면 네가 가진 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리해. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 요구하지 마.
그런 확인 요청 기능이 없는 환경에서만 짧게 yes/no로 물어봐.

도메인 허용 확인 전에는 git, npm, curl, 로그인, Outlook/Teams/Files 조회를 시작하지 마.
GitHub API(api.github.com)나 GitHub Release asset 다운로드로 repo 접근을 판단하지 마.
이 환경에서는 api.github.com이 403이어도 git clone은 성공할 수 있어.
repo 접근은 git ls-remote 또는 git clone으로 판단해.

아래 순서로 준비해.

REMOTE_HEAD=$(git ls-remote ${repoUrl} refs/heads/master | awk '{print $1}')
test -n "$REMOTE_HEAD"
rm -rf ${defaultWorkDir}
test ! -e ${defaultWorkDir}
git clone --branch master --single-branch --no-tags ${repoUrl} ${defaultWorkDir}
LOCAL_HEAD=$(git -C ${defaultWorkDir} rev-parse HEAD)
test "$LOCAL_HEAD" = "$REMOTE_HEAD"
test ! -d ${defaultWorkDir}/node_modules
cd ${defaultWorkDir} && npm ci --prefer-offline --no-audit --no-fund
cd ${defaultWorkDir} && npm run build
test -f ${defaultWorkDir}/dist/cli.js
test -f ${defaultWorkDir}/dist/proxy.js
test -f ${defaultWorkDir}/dist/msal-network.js
cd ${defaultWorkDir} && node dist/cli.js

각 셸 호출은 작업 폴더가 초기화될 수 있어. 모든 npm, build, Hare 명령에 동일한 ${defaultWorkDir} 작업 경로를 명시해. 준비부터 로그인까지 새로 clone한 이 폴더만 사용해. 어느 단계든 실패하면 기존 node_modules, dist, 다른 복사본으로 대체하지 말고 실패 단계와 오류 한 줄만 보고해. npm 설치를 백그라운드로 돌리거나 npm install로 바꾸거나 여러 셸 호출에 나눠 반복하지 마.

결과에서 loggedIn과 tokenUsable이 모두 true이면 바로 내 요청을 처리해.
둘 중 하나라도 false이고 dataDirPersistent가 false이면 로그인부터 시도하지 마. 폴더 연결 도구로 내 Documents 폴더 접근을 요청하고, 연결된 폴더 안에 "Hare M365 Agent" 하위 폴더를 만들어. 선택지를 묻거나 경로를 직접 입력하라고 하지 마.
이후 모든 Hare 명령에 HARE_M365_DATA_DIR="<마운트된 Documents>/Hare M365 Agent"를 붙이고 같은 값으로 시작 상태를 다시 확인해.

dataDirPersistent가 true인데 로그인이 필요하면 humanLoginCommand(auth login-start)를 네가 실행해. 이 명령은 Microsoft 로그인 주소와 userCode를 즉시 반환하고 끝나야 해.
주소와 코드를 나에게 보여주고 브라우저에서 회사 Microsoft 계정으로 로그인한 뒤 "로그인 완료"라고 말하라고 짧게 안내해.
내가 "로그인 완료"라고 말하면 humanLoginCompleteCommand(auth login-complete)를 새 포그라운드 셸 호출로 실행해. 이 명령은 25초 안에 끝나야 해.
poller를 계속 실행하거나 백그라운드, detached, setsid, nohup을 사용하지 마.
authorization_pending이면 login-start를 다시 실행하지 말고 login-complete만 다시 실행해. 코드가 만료된 경우에만 login-start를 다시 실행해.
성공 후 같은 HARE_M365_DATA_DIR로 auth status를 실행하고 loggedIn과 tokenUsable이 모두 true인지 확인해.
cacheFile 존재만으로 로그인 성공으로 판단하지 마.

cacheFile이 있는데 Outlook/Teams/Files 조회가 fetch failed 또는 network_error로 실패하면 sandbox 네트워크 문제로 판단해.
clone/build를 반복하지 말고, 도메인 사진확인이나 반복 호출 루프를 시작하지 마.
사용자의 Windows/Mac 로컬 환경에서 같은 Hare 명령을 실행하고 --out으로 status.resultsDir에 결과 JSON을 저장하게 해.
예: node dist/cli.js outlook inbox --limit 1 --out latest-mail.json
일반 조회에서는 --out을 사용하지 말고 화면 출력을 바로 처리해.
--out 값이 상대 경로이면 Hare 고정 resultsDir 아래에 저장되고 7일 후 자동 정리돼.
결과 파일이 생성되면 그 JSON을 읽어 내 요청을 처리해.

Outlook 또는 Teams에서 기간·키워드 조회를 요청받으면 inbox/chat-messages의 최근 건수 제한으로 대신하지 말고 outlook search 또는 teams search-messages를 사용해.
메일이 몇 건인지 묻는 정확한 집계 요청은 outlook search 결과를 세지 말고 outlook count를 사용해.
SharePoint 사이트 존재 여부는 sharepoint sites로 확인하고, 개인 OneDrive만 검색하는 files search 결과로 사이트 존재 여부나 권한을 추정하지 마.
사용자가 기간을 말하면 --since/--until에 그대로 반영해.
기간을 말하지 않으면 기본 최근 90일이 적용되며, 결과의 search.range.notice를 답변에 포함해 실제 조회 범위를 알려줘.
search.limitReached가 true이면 검색 결과 한도에 도달했다고 알려줘.`;
function getSelfCommand(): string {
  return process.env.HARE_M365_COMMAND ?? defaultCliCommand;
}

function getLoginCommand(): string {
  return `${getSelfCommand()} auth login-start`;
}

function getLoginCompleteCommand(): string {
  return `${getSelfCommand()} auth login-complete`;
}

function emitJson(payload: unknown, out?: string): void {
  if (!out) {
    console.log(JSON.stringify(payload, null, 2));
    return;
  }

  const outputPath = resolveResultPath(config, out);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, JSON.stringify(payload, null, 2), "utf8");
  console.log(JSON.stringify({ ok: true, outputPath }, null, 2));
}

async function getDoctorStatus() {
  const authStatus = config.clientId && config.tenantId
    ? await getAuthStatus(config)
    : { account: null, loggedIn: false, tokenUsable: false, reason: "NOT_CONFIGURED" };
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  return {
    configured: Boolean(config.clientId && config.tenantId),
    clientIdPresent: Boolean(config.clientId),
    tenantIdPresent: Boolean(config.tenantId),
    loggedIn: authStatus.loggedIn,
    tokenUsable: authStatus.tokenUsable,
    authReason: authStatus.reason,
    dataDir: config.dataDir,
    dataDirSource: config.dataDirSource,
    dataDirPersistent: config.dataDirPersistent,
    policyPath: config.policyPath,
    cacheDir: config.cacheDir,
    cacheFile,
    cacheFileExists: fs.existsSync(cacheFile),
    pendingLoginStateExists: fs.existsSync(deviceLoginStatePath(config)),
    downloadDir: config.downloadDir,
    logsDir: config.logsDir,
    resultsDir: config.resultsDir,
    resultRetentionDays: config.policy.retentionDays
  };
}

function loginGateFields(loggedIn: boolean, dataDirPersistent: boolean) {
  if (loggedIn) return {};
  if (!dataDirPersistent) {
    return {
      nextCommand: "PERSISTENT_DATA_DIR_REQUIRED",
      llmAction: "REQUEST_DOCUMENTS_FOLDER_ACCESS",
      instruction:
        "Request access to the user's Documents folder, create a 'Hare M365 Agent' subfolder in the mounted folder, and set HARE_M365_DATA_DIR to that subfolder for every Hare command. Do not start Microsoft login in a temporary /sessions or /tmp dataDir."
    };
  }
  return {
    nextCommand: "LOGIN_REQUIRED_HARD_GATE",
    llmAction: "RUN_SPLIT_DEVICE_LOGIN",
    humanLoginCommand: getLoginCommand(),
    humanLoginCompleteCommand: getLoginCompleteCommand(),
    instruction:
      "Run humanLoginCommand once; it returns the Microsoft URL and user code immediately. Show them to the user and end the shell call. After the user says login complete, run humanLoginCompleteCommand in a new foreground shell call; it must finish within 25 seconds. Never keep a poller alive, use background/detached processes, or rerun login-start unless the code expired. Then run auth status in the same dataDir."
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
          setupCommand: `REMOTE_HEAD=$(git ls-remote ${repoUrl} refs/heads/master | awk '{print $1}') && test -n "$REMOTE_HEAD" && rm -rf ${defaultWorkDir} && test ! -e ${defaultWorkDir} && git clone --branch master --single-branch --no-tags ${repoUrl} ${defaultWorkDir} && LOCAL_HEAD=$(git -C ${defaultWorkDir} rev-parse HEAD) && test "$LOCAL_HEAD" = "$REMOTE_HEAD" && cd ${defaultWorkDir} && npm ci --prefer-offline --no-audit --no-fund && npm run build && test -f dist/cli.js && test -f dist/proxy.js && test -f dist/msal-network.js && node dist/cli.js`,
          cloneVerification: "LOCAL_HEAD must exactly equal REMOTE_HEAD from refs/heads/master before npm ci.",
          requiredDomains,
          installRule: "registry.npmjs.org is required before npm ci. Do not switch to background or incremental npm install when npm ci fails.",
          searchDefaults: {
            lookbackDays: config.policy.defaultSearchLookbackDays,
            maxResults: config.policy.maxSearchResults,
            instruction:
              "If the user does not specify a date range, search the latest 90 days and report search.range.notice. If limitReached is true, report that results were truncated."
          },
          status,
          storageRule:
            "Use the fixed Hare data folder only. In Cowork/hosted sandbox, mount the user's fixed Hare folder and set HARE_M365_DATA_DIR to that mounted path for every Hare command.",
          ...(status.configured ? loginGateFields(status.loggedIn, status.dataDirPersistent) : {}),
          nextCommand: status.configured
            ? status.loggedIn
              ? `${getSelfCommand()} auth status`
              : status.dataDirPersistent
                ? "LOGIN_REQUIRED_HARD_GATE"
                : "PERSISTENT_DATA_DIR_REQUIRED"
            : "Check hare.config.json or set local environment overrides.",
          llmAction: status.configured && !status.loggedIn
            ? status.dataDirPersistent
              ? "RUN_SPLIT_DEVICE_LOGIN"
              : "REQUEST_DOCUMENTS_FOLDER_ACCESS"
            : undefined,
          humanLoginCommand: status.configured && !status.loggedIn && status.dataDirPersistent
            ? getLoginCommand()
            : undefined,
          humanLoginCompleteCommand: status.configured && !status.loggedIn && status.dataDirPersistent
            ? getLoginCompleteCommand()
            : undefined
        },
        null,
        2
      )
    );
  });

const auth = program.command("auth").description("Authentication commands");

async function runLoginStart(): Promise<void> {
  requireConfigured(config);
  const result = await startLogin(config);
  console.log(
    JSON.stringify(
      {
        ...result,
        nextCommand: getLoginCompleteCommand(),
        instruction:
          "Show verificationUri and userCode to the user. End this shell call. After the user completes Microsoft sign-in, run nextCommand in a new foreground shell call."
      },
      null,
      2
    )
  );
}

auth.command("login").description("Start the split Microsoft device-code login flow").action(runLoginStart);
auth.command("login-start").description("Issue a Microsoft device code and return immediately").action(runLoginStart);

auth.command("login-complete").description("Complete a pending Microsoft device-code login").action(async () => {
  requireConfigured(config);
  const result = await completeLogin(config);
  console.log(
    JSON.stringify(
      {
        ok: true,
        stage: "COMPLETE",
        account: result.account?.username,
        tenantId: result.tenantId,
        scopes: getScopeList(),
        nextCommand: `${getSelfCommand()} auth status`
      },
      null,
      2
    )
  );
});

auth.command("status").description("Show current login and policy status").action(async () => {
  requireConfigured(config);
  const authStatus = await getAuthStatus(config);
  const loggedIn = authStatus.loggedIn;
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  console.log(
    JSON.stringify(
      {
        loggedIn,
        tokenUsable: authStatus.tokenUsable,
        authReason: authStatus.reason,
        account: authStatus.account?.username,
        policy: config.policy,
        dataDir: config.dataDir,
        dataDirSource: config.dataDirSource,
        dataDirPersistent: config.dataDirPersistent,
        cacheDir: config.cacheDir,
        cacheFile,
        cacheFileExists: fs.existsSync(cacheFile),
        downloadDir: config.downloadDir,
        logsDir: config.logsDir,
        resultsDir: config.resultsDir,
        resultRetentionDays: config.policy.retentionDays,
        pendingLoginStateExists: fs.existsSync(deviceLoginStatePath(config)),
        ...loginGateFields(loggedIn, config.dataDirPersistent)
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
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listInbox(config, Number(options.limit));
    emitJson({ messages: data }, options.out);
  });

outlook
  .command("search")
  .description("Search Outlook messages by keyword and date range")
  .requiredOption("--query <text>", "search query or Outlook KQL")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--limit <number>", "maximum matching message count", "1000")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      query: string;
      since?: string;
      until?: string;
      folder: string;
      limit: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const data = await searchMailbox(
        config,
        options.query,
        options.since,
        options.until,
        parseMailFolderScope(options.folder),
        Number(options.limit)
      );
      emitJson(data, options.out);
    }
  );

outlook
  .command("count")
  .description("Count Outlook messages exactly by scanning every page in a date range")
  .option("--subject-contains <text>", "literal text that must appear in the subject")
  .option("--from <text>", "text that must appear in the sender name or address")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      subjectContains?: string;
      from?: string;
      since?: string;
      until?: string;
      folder: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const data = await countMailboxMessages(
        config,
        options.subjectContains,
        options.from,
        options.since,
        options.until,
        parseMailFolderScope(options.folder)
      );
      emitJson(data, options.out);
    }
  );

const teams = program.command("teams").description("Teams read commands");

teams
  .command("teams")
  .description("List joined teams")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { out?: string }) => {
    requireConfigured(config);
    const data = await listJoinedTeams(config);
    emitJson({ teams: data }, options.out);
  });

teams
  .command("chats")
  .description("List recent chats")
  .option("--limit <number>", "maximum chat count", "20")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
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
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { chatId: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listChatMessages(config, options.chatId, Number(options.limit));
    emitJson({ messages: data }, options.out);
  });

teams
  .command("search-messages")
  .description("Search Teams messages by keyword and date range")
  .requiredOption("--query <text>", "search query or Teams KQL")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--limit <number>", "maximum matching message count", "1000")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: { query: string; since?: string; until?: string; limit: string; out?: string }) => {
      requireConfigured(config);
      const data = await searchChatMessages(
        config,
        options.query,
        options.since,
        options.until,
        Number(options.limit)
      );
      emitJson(data, options.out);
    }
  );

const files = program.command("files").description("SharePoint/OneDrive file commands");

const sharepoint = program.command("sharepoint").description("SharePoint site commands");

sharepoint
  .command("sites")
  .description("Search SharePoint sites by name or keyword")
  .requiredOption("--query <text>", "site name or search keyword")
  .option("--limit <number>", "maximum site count", "25")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { query: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await searchSites(config, options.query, Number(options.limit));
    emitJson(data, options.out);
  });

files
  .command("search")
  .description("Search files in the signed-in user's personal OneDrive")
  .requiredOption("--query <text>", "search query")
  .option("--limit <number>", "maximum file count", "10")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
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

function parseMailFolderScope(value: string): MailFolderScope {
  if (value === "all" || value === "inbox" || value === "sent") return value;
  throw new Error("folder must be one of: all, inbox, sent.");
}
