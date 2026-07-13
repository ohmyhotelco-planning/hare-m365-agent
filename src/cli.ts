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
import { writeSessionRules } from "./session-rules.js";
import { buildSetupContract } from "./setup-state.js";
import { buildLocalSetupCommand, getLocalInstallPaths } from "./local-install.js";

const program = new Command();
program.option("--data-dir <path>", "Use this exact Hare data directory for every command in the session");
const config = loadConfig({ dataDir: readDataDirArgument(process.argv.slice(2)) });
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
const localInstallPaths = getLocalInstallPaths(config.dataDir);
const localSetupCommand = buildLocalSetupCommand({
  dataDir: config.dataDir,
  repository: repoUrl,
  branch: "master"
});
const setupCommandForGuide = config.dataDirPersistent
  ? localSetupCommand
  : "# FOLDER_REQUIRED: start Cowork with the HareM365Agent project folder selected, then rerun this guide with its mounted root as --data-dir.";
const defaultCliCommand = `node "${path.join(packageRoot, "dist", "cli.js")}"`;

const requiredDomains = [
  "github.com",
  "registry.npmjs.org",
  "graph.microsoft.com",
  "login.microsoftonline.com",
  "ohmylab-my.sharepoint.com",
  "ohmylab.sharepoint.com"
];

const rulesFile = writeSessionRules(config, {
  commandPrefix: getExplicitSelfCommand(),
  repository: repoUrl,
  branch: "master",
  workDir: localInstallPaths.appDir,
  requiredDomains
});

function readPackageVersion(): string {
  const packageJsonPath = path.join(packageRoot, "package.json");
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8")) as { version?: string };
  return packageJson.version ?? "0.0.0";
}

function readDataDirArgument(args: string[]): string | undefined {
  for (let index = 0; index < args.length; index += 1) {
    const value = args[index];
    if (value === "--data-dir") return args[index + 1];
    if (value.startsWith("--data-dir=")) return value.slice("--data-dir=".length);
  }
  return undefined;
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
7. Cowork에는 실행 셸이 두 곳 있을 수 있다. 세션 샌드박스 셸에는 Cowork 설정의 기능 > 도메인 허용 목록이 적용되고, 연결 폴더가 마운트된 디바이스 셸은 설정과 무관하게 모든 외부 도메인이 차단될 수 있다(전 도메인 403 blocked-by-allowlist). git, npm, 로그인, Graph 호출은 반드시 도메인 허용 목록이 실제로 적용되는 셸에서 실행한다.
8. HARE_ROOT 판별: 네트워크가 되는 실행 셸에서 연결 폴더 마운트(/sessions/<session>/mnt/HareM365Agent 형태)가 직접 보이고 쓰기 가능하면 그 마운트 루트를 HARE_ROOT로 사용한다(단일 환경). 보이지 않으면 하이브리드 모델을 사용한다: 샌드박스 고정 경로(예: /home/claude/hare)를 HARE_ROOT와 --data-dir로 쓰고, 연결 폴더는 영속 저장소로만 사용해 파일 스테이징/커밋 도구로 동기화한다.
9. 도메인 설정을 현재 Cowork 작업이 시작된 뒤 변경했다면 새 Cowork 채팅을 열어야 변경된 네트워크 정책이 적용된다.
10. 사용자는 Cowork 작업을 만들 때 HareM365Agent 프로젝트 폴더를 먼저 선택한다. 연결된 HareM365Agent 프로젝트가 없으면 FOLDER_REQUIRED로 멈추고 해당 프로젝트를 선택한 새 Cowork 작업에서 다시 시작하도록 안내한다.

하이브리드 세션 시작(연결 폴더가 실행 셸에 보이지 않는 표준 Cowork):
1. 연결 폴더에서 .hare-app-snapshot.tar.gz, .hare-app-build-head, .cache/msal-cache.json, claude/ 문서를 가져와 HARE_ROOT 아래 같은 상대 경로에 배치한다.
2. 준비 명령을 실행한다. 스냅샷이 있으면 자동으로 풀려 clone을 건너뛰고, HEAD가 원격과 같으면 npm ci와 build도 건너뛴다.
3. 재빌드가 일어났거나 로그인 캐시가 갱신되면 .hare-app-snapshot.tar.gz, .hare-app-build-head, .cache/msal-cache.json을 연결 폴더로 다시 커밋해 다음 채팅에서 재사용한다.

준비 명령:
\`\`\`bash
${setupCommandForGuide}
\`\`\`

HARE_ROOT 하나만 모든 명령의 --data-dir로 사용한다. 프로그램은 그 아래 app, 인증 캐시는 .cache, 규칙은 claude에 유지한다. 스냅샷과 app이 모두 없을 때만 clone하고, 이후에는 같은 app을 git pull --ff-only로 갱신한다. HEAD가 바뀌었거나 빌드가 없을 때만 npm ci와 build를 실행한다. 어느 단계든 실패하면 다른 복사본이나 임시 폴더로 대체하지 말고 실패 단계와 오류 한 줄만 보고한다. npm ci를 npm install로 바꾸거나 여러 셸 호출에 나눠 반복하지 않는다.

로그인 하드게이트:
- startup, doctor, auth status의 setup.state만 확인하고 해당 상태의 setup.nextAction 하나만 수행한다.
- SETUP_REQUIRED: 설정 누락을 한 문장으로 보고하고 멈춘다.
- FOLDER_REQUIRED: 현재 작업이 HareM365Agent 프로젝트 폴더 없이 시작된 상태다. 사용자가 해당 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 열고 같은 프롬프트를 붙여넣도록 한 문장으로 안내한 뒤 멈춘다.
- Cowork Linux의 /root/.local/share 또는 ~/.local/share OS 기본 경로는 login-start가 거부된다. 하이브리드 모델에서는 --data-dir로 명시한 샌드박스 HARE_ROOT(예: /home/claude/hare)에 로그인하고, 캐시를 연결 폴더로 동기화해 영속성을 확보한다. /sessions/<session>/mnt/HareM365Agent 형태는 연결 폴더의 정상 마운트 경로이므로 실행 셸에서 직접 보이는 환경이면 그대로 --data-dir로 사용한다.
- LOGIN_START_REQUIRED: setup.nextCommand를 수정하지 않고 한 번 실행하고 Microsoft 로그인 주소와 userCode를 보여준 뒤 멈춘다.
- LOGIN_COMPLETE_REQUIRED: 사용자가 "로그인 완료"라고 말할 때까지 기다린다. 완료 후 setup.nextCommand를 수정하지 않고 한 번 실행한다. COMPLETE는 저장된 캐시 재검증까지 성공했다는 뜻이다. 하이브리드 모델에서는 갱신된 .cache/msal-cache.json을 즉시 연결 폴더로 커밋한다.
- READY: 사용 준비 완료를 보고하고 사용자의 업무 조회 요청을 기다린다.
- HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 함께 나오면 먼저 실행 셸을 확인한다. 연결 폴더가 마운트된 디바이스 셸이었다면 실행 위치 문제이므로 같은 명령을 샌드박스 셸에서 한 번 실행한다. 샌드박스 셸이었다면 NETWORK_PERMISSION_REQUIRED로 보고한다. Cowork 설정의 도메인 허용 목록을 확인하고, 설정을 변경했다면 같은 HareM365Agent 프로젝트에서 새 Cowork 작업을 열어 실패한 명령만 한 번 재시도하도록 안내한 뒤 멈춘다.
- 그 밖의 명령 실패는 BLOCKED로 취급해 실패 단계와 오류 한 줄만 보고한다. 다른 경로를 추측하거나 반복 실행하지 않는다.

저장 위치:
- 로컬 Windows/Mac/Linux에서는 OS 기본 고정 dataDir를 사용한다.
- Cowork/hosted sandbox에서는 샌드박스 고정 HARE_ROOT를 --data-dir로 명시해 실행하고, 영속 원본은 작업을 시작할 때 선택한 HareM365Agent 프로젝트 폴더에 둔다. 스냅샷, 빌드 마커, 로그인 캐시, 규칙, 보존할 결과를 연결 폴더와 동기화한다.
- 일반 /sessions 또는 /tmp 기본 경로에서는 login-start가 거부된다. /sessions/<session>/mnt/HareM365Agent는 연결된 호스트 폴더이므로 허용한다.
- Cowork Linux OS 기본 dataDir(/root/.local/share 등)도 임시 컨테이너 경로로 취급하고 login-start를 거부한다.
- startup이 반환한 setup.nextCommand에는 동일한 --data-dir가 포함된다. 환경변수에 의존하거나 명령을 다시 조합하지 않고 그대로 실행한다.

Cowork 네트워크 실패 때:
- 먼저 실패한 명령을 어느 셸에서 실행했는지 확인한다. 연결 폴더가 마운트된 디바이스 셸은 설정과 무관하게 모든 도메인이 403 blocked-by-allowlist로 차단될 수 있다. 이 경우 설정 문제가 아니라 실행 위치 문제이므로 같은 명령을 샌드박스 셸에서 실행한다.
- 샌드박스 셸에서 HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 나오면 현재 Cowork 작업의 네트워크 허용 정책 문제다. Cowork 설정의 기능 > 도메인 허용 목록을 확인한다.
- 설정을 변경한 뒤에는 새 Cowork 채팅을 열고 같은 HareM365Agent 폴더를 다시 연결해야 새 정책이 적용될 수 있다.
- 이전 작업에서 반복 진단하지 말고, 새 작업에서 실패했던 명령만 한 번 재시도한다. clone/build/login 전체를 반복하지 않는다.
- 일반적인 fetch failed 또는 network_error에는 실제 오류 한 줄만 보고하고, 도메인 반복 호출이나 임의 프록시 우회를 시작하지 않는다.

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
Cowork에는 실행 셸이 두 곳 있을 수 있어. 세션 샌드박스 셸에는 설정의 기능 > 도메인 허용 목록이 적용되고, 연결 폴더가 마운트된 디바이스 셸은 설정과 무관하게 모든 도메인이 403 blocked-by-allowlist로 차단될 수 있어.
git, npm, 로그인, Graph 호출은 반드시 도메인 허용 목록이 실제로 적용되는 셸(샌드박스)에서 실행해.
도메인 설정을 이 Cowork 작업을 연 뒤 변경했다면 새 Cowork 채팅을 열어야 변경된 정책이 적용돼.
GitHub API(api.github.com)나 GitHub Release asset 다운로드로 repo 접근을 판단하지 마.
이 환경에서는 api.github.com이 403이어도 git clone은 성공할 수 있어.
repo 접근은 git ls-remote 또는 git clone으로 판단해.

이 Cowork 작업을 만들 때 사용자가 미리 선택한 HareM365Agent 프로젝트 폴더를 찾아.
연결된 HareM365Agent 프로젝트가 없으면 FOLDER_REQUIRED라고 말하고, HareM365Agent 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 연 뒤 같은 프롬프트를 붙여넣으라고 한 문장으로 안내하고 멈춰.

HARE_ROOT 판별: 네트워크가 되는 실행 셸에서 연결 폴더 마운트(/sessions/<session>/mnt/HareM365Agent 형태)가 직접 보이고 쓰기 가능하면 그 마운트 루트를 HARE_ROOT와 --data-dir로 사용해.
보이지 않으면 하이브리드 모델을 사용해: 샌드박스 고정 경로(예: /home/claude/hare)를 HARE_ROOT로 쓰고, 연결 폴더는 영속 저장소로만 사용해.
하이브리드 모델에서는 준비 전에 연결 폴더의 .hare-app-snapshot.tar.gz, .hare-app-build-head, .cache/msal-cache.json, claude/ 문서를 가져와 HARE_ROOT 아래 같은 상대 경로에 배치하고, 재빌드나 로그인 캐시 갱신이 생기면 같은 파일들을 연결 폴더로 다시 커밋해.

아래 순서로 준비해.

${setupCommandForGuide}

HARE_ROOT 하나만 모든 명령의 --data-dir로 사용해. 프로그램은 그 아래 app, 인증 캐시는 .cache, 규칙은 claude에 유지해. 스냅샷이 있으면 자동으로 풀려 clone을 건너뛰고, 스냅샷과 app이 모두 없을 때만 clone해. 이후 채팅에서는 같은 app을 git pull --ff-only로 갱신해. HEAD가 바뀌었거나 빌드가 없을 때만 npm ci와 build를 실행해. 실패하면 다른 복사본이나 임시 폴더로 대체하지 말고 실패 단계와 오류 한 줄만 알려줘.

startup JSON의 setup.state만 확인하고 아래에서 일치하는 한 단계만 수행해. setup.state와 다른 경로를 추측하지 마.

- SETUP_REQUIRED: 설정 누락을 한 문장으로 알려주고 멈춰.
- FOLDER_REQUIRED: HareM365Agent 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 열고 같은 프롬프트를 붙여넣으라고 한 문장으로 안내한 뒤 멈춰.
- Cowork Linux의 /root/.local/share 또는 ~/.local/share OS 기본 경로에는 로그인하지 마. 하이브리드 모델에서는 --data-dir로 명시한 샌드박스 HARE_ROOT에 로그인하고 캐시를 연결 폴더로 동기화해. /sessions/<session>/mnt/HareM365Agent 형태는 내 컴퓨터 폴더의 정상 마운트 경로야. 실행 셸에서 직접 보이는 환경이면 그대로 --data-dir로 사용해.
- LOGIN_START_REQUIRED: setup.nextCommand를 수정하지 않고 한 번 실행해. 반환된 Microsoft 로그인 주소와 userCode를 보여주고 "로그인 완료"라고 말해달라고 안내한 뒤 멈춰.
- LOGIN_COMPLETE_REQUIRED: 내가 "로그인 완료"라고 말하기 전에는 아무 명령도 실행하지 마. 완료 후 setup.nextCommand를 수정하지 않고 한 번 실행해. COMPLETE가 반환되면 저장된 캐시 재검증까지 성공한 상태야. 하이브리드 모델이면 갱신된 .cache/msal-cache.json을 즉시 연결 폴더로 커밋해.
- READY: "Hare M365 Agent 사용 준비 완료. 이제 자연어로 업무 조회를 요청하세요."라고 말하고 멈춰.
- HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 함께 나오면 먼저 실행 셸을 확인해. 디바이스 셸이었다면 같은 명령을 샌드박스 셸에서 한 번 실행해. 샌드박스 셸이었다면 NETWORK_PERMISSION_REQUIRED라고 알려주고, Cowork 설정의 도메인 허용 목록을 확인한 뒤 설정을 바꿨다면 같은 HareM365Agent 프로젝트에서 새 Cowork 작업을 열어 실패했던 명령만 한 번 재시도하라고 안내하고 멈춰.
- 그 밖의 명령 실패: BLOCKED로 취급해 실패 단계와 오류 한 줄만 알려주고 멈춰. 다른 경로를 시도하거나 같은 명령을 반복하지 마.

clone/build 전체 반복, 도메인 반복 진단, 임의 프록시 우회는 시작하지 마.

Outlook 또는 Teams에서 기간·키워드 조회를 요청받으면 inbox/chat-messages의 최근 건수 제한으로 대신하지 말고 outlook search 또는 teams search-messages를 사용해.
메일이 몇 건인지 묻는 정확한 집계 요청은 outlook search 결과를 세지 말고 outlook count를 사용해.
SharePoint 사이트 존재 여부는 sharepoint sites로 확인하고, 개인 OneDrive만 검색하는 files search 결과로 사이트 존재 여부나 권한을 추정하지 마.
사용자가 기간을 말하면 --since/--until에 그대로 반영해.
기간을 말하지 않으면 기본 최근 90일이 적용되며, 결과의 search.range.notice를 답변에 포함해 실제 조회 범위를 알려줘.
search.limitReached가 true이면 검색 결과 한도에 도달했다고 알려줘.`;
function getSelfCommand(): string {
  const command = process.env.HARE_M365_COMMAND ?? defaultCliCommand;
  if (config.dataDirSource === "os-default") return command;
  return `${command} --data-dir ${quoteCommandArgument(config.dataDir)}`;
}

function getExplicitSelfCommand(): string {
  const command = process.env.HARE_M365_COMMAND ?? defaultCliCommand;
  return `${command} --data-dir ${quoteCommandArgument(config.dataDir)}`;
}

function quoteCommandArgument(value: string): string {
  if (value.includes('"')) throw new Error("Hare data directory cannot contain a double quote.");
  return `"${value}"`;
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
    resultRetentionDays: config.policy.retentionDays,
    rulesFile,
    rulesFileExists: Boolean(rulesFile && fs.existsSync(rulesFile))
  };
}

program
  .name(preferredCommand)
  .description("Hare M365 Agent CLI for LLM-driven Microsoft 365 Graph access.")
  .version(packageVersion)
  .action(async () => {
    const status = await getDoctorStatus();
    const setup = buildSetupContract(status, getSelfCommand());
    console.log(
      JSON.stringify(
        {
          tool: preferredCommand,
          package: packageName,
          version: packageVersion,
          mode: "startup",
          repository: repoUrl,
          appDir: config.dataDirPersistent ? localInstallPaths.appDir : undefined,
          setupCommand: config.dataDirPersistent ? localSetupCommand : undefined,
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
          setup,
          sessionRules: {
            path: rulesFile,
            exists: Boolean(rulesFile && fs.existsSync(rulesFile)),
            instruction: rulesFile
              ? "Read this file before login or Microsoft 365 lookup, and keep using the exact data directory and command prefix recorded there."
              : "Open a Cowork task with the HareM365Agent project folder selected; the session rules file is created only in that persistent project."
          },
          storageRule:
            "Use setup.state and perform setup.nextAction only. Do not infer a different setup path."
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
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  console.log(
    JSON.stringify(
      {
        ok: true,
        stage: "COMPLETE",
        cacheVerified: true,
        account: result.account?.username,
        tenantId: result.tenantId,
        dataDir: config.dataDir,
        cacheFile,
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
  const pendingLoginStateExists = fs.existsSync(deviceLoginStatePath(config));
  const setup = buildSetupContract(
    {
      configured: true,
      loggedIn,
      tokenUsable: authStatus.tokenUsable,
      dataDirPersistent: config.dataDirPersistent,
      pendingLoginStateExists
    },
    getSelfCommand()
  );
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
        rulesFile,
        rulesFileExists: Boolean(rulesFile && fs.existsSync(rulesFile)),
        pendingLoginStateExists,
        setup
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
    const status = await getDoctorStatus();
    console.log(
      JSON.stringify(
        {
          ...status,
          setup: buildSetupContract(status, getSelfCommand())
        },
        null,
        2
      )
    );
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
