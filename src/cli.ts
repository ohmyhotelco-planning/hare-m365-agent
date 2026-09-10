#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Command, Option } from "commander";
import { completeLogin, getAuthStatus, getScopeList, startLogin, logout, type AuthStatus } from "./auth.js";
import { ensureRuntimeDirs, loadConfig, requireConfigured } from "./config.js";
import { hasPendingDeviceLoginState } from "./device-login.js";
import {
  countMailboxMessages,
  listFlaggedMessages,
  listInbox,
  listRecentMailbox,
  searchMailbox,
  type MailFolderScope
} from "./outlook.js";
import {
  createApprovedDraft,
  prepareDraft,
  type DraftContentType,
  type DraftInput,
  type DraftKind
} from "./outlook-drafts.js";
import {
  downloadMessageAttachment,
  listMessageAttachments
} from "./outlook-attachments.js";
import { resolveMailboxTarget, runMailboxRead } from "./outlook-mailbox.js";
import { downloadDriveItem, searchFiles, searchSites } from "./sharepoint.js";
import { exportSharePointFiles } from "./sharepoint-export.js";
import { listChatMessagesPage, listChats, listJoinedTeams, searchChatMessages } from "./teams.js";
import {
  downloadTeamsMessageAttachment,
  listTeamsMessageAttachments
} from "./teams-attachments.js";
import {
  downloadTeamsInlineImage,
  listTeamsInlineImages
} from "./teams-inline-images.js";
import { cleanupExpiredResults, resolveResultPath } from "./results.js";
import { writeSessionRules } from "./session-rules.js";
import { buildSetupContract } from "./setup-state.js";
import { buildLocalSetupCommand } from "./local-install.js";
import { checkMicrosoftConnectivity, networkExecutionGuidance, type ExecutionEnvironment } from "./network-check.js";

const program = new Command();
program.option("--data-dir <path>", "Use this exact Hare data directory for every command in the session");
const config = loadConfig({ dataDir: readDataDirArgument(process.argv.slice(2)) });

const preferredCommand = "hare-m365";
const packageName = "@ohmyhotel/hare-m365-agent";
const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageVersion = readPackageVersion();
const repoUrl = "https://github.com/ohmyhotelco-planning/hare-m365-agent.git";
const localSetupCommand = buildLocalSetupCommand({
  dataDir: config.dataDir,
  repository: repoUrl,
  branch: "master"
});
const setupCommandForGuide = config.dataDirPersistent
  ? localSetupCommand
  : "# FOLDER_REQUIRED: start Cowork with the user's existing Hare project folder selected, then rerun this guide with that selected project root as the persistent store.";
const defaultCliCommand = `${process.platform === "win32" ? "& " : ""}${quoteCommandArgument(process.execPath)} ${quoteCommandArgument(path.join(packageRoot, "dist", "cli.js"))}`;

const requiredDomains = [
  "github.com",
  "registry.npmjs.org",
  "graph.microsoft.com",
  "login.microsoftonline.com",
  "outlook.office.com",
  "ohmylab-my.sharepoint.com",
  "ohmylab.sharepoint.com"
];

let rulesFile: string | undefined;
program.hook("preAction", (_command, actionCommand) => {
  if (actionCommand.name() === "check" && actionCommand.parent?.name() === "network") return;
  ensureRuntimeDirs(config);
  try {
    cleanupExpiredResults(config);
  } catch {
    // Result cleanup is best-effort and must not block M365 commands.
  }
  rulesFile = writeSessionRules(config, {
    commandPrefix: getExplicitSelfCommand(),
    repository: repoUrl,
    branch: "master",
    workDir: packageRoot,
    requiredDomains
  });
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

실행 환경 사전 확인 (startup/auth status 또는 첫 조회 전에 적용):
${networkExecutionGuidance}

시작 순서:
1. Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 7개 도메인 허용 여부를 사용자에게 확인한다.
   - github.com
   - registry.npmjs.org
   - graph.microsoft.com
   - login.microsoftonline.com
   - outlook.office.com
   - ohmylab-my.sharepoint.com
   - ohmylab.sharepoint.com
2. 도메인 확인은 가능하면 LLM 도구의 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리한다. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 만들지 않는다.
3. 확인 요청 UI가 없는 환경에서만 짧게 yes/no로 묻는다.
4. 도메인 허용 확인 전에는 git, npm, curl, 로그인, Outlook/Teams/Files 조회를 시작하지 않는다.
5. GitHub API(api.github.com)나 GitHub Release asset으로 repo 접근을 판단하지 않는다. Cowork에서는 api.github.com이 막혀도 git clone은 성공할 수 있다.
6. 실제 사용 경로인 git ls-remote 또는 git clone으로 repo 접근을 판단한다.
7. git, npm, 로그인, Graph 호출은 Cowork 도메인 허용 목록이 적용되는 세션 셸에서 실행한다.
8. 현재 Cowork 작업에 선택된 프로젝트 마운트만 영구 dataDir로 사용한다. 폴더 이름으로 다른 경로를 찾지 않는다.
9. 선택 프로젝트에는 인증 캐시, 규칙, 다운로드와 결과만 저장한다. git clone, npm ci와 build는 준비 명령이 지정한 세션 런타임에서 실행한다.
10. 도메인 설정을 현재 Cowork 작업이 시작된 뒤 변경했다면 새 Cowork 채팅을 열어야 변경된 네트워크 정책이 적용된다.
11. 사용자는 Cowork 작업을 만들 때 기존 Hare 프로젝트 폴더를 먼저 선택한다. 선택된 프로젝트가 없으면 FOLDER_REQUIRED로 멈추고 기존 Hare 프로젝트를 선택한 새 Cowork 작업에서 다시 시작하도록 안내한다.

준비 명령:
\`\`\`bash
${setupCommandForGuide}
\`\`\`

준비 명령은 세션 런타임의 app을 최신 master로 맞추고 모든 Hare 명령에 선택 프로젝트의 정확한 --data-dir를 붙인다. 프로젝트 폴더에서 git/npm/build를 실행하거나 삭제 권한을 요청하지 않는다. /tmp, /dev/shm, /home/claude, /root/.local/share 또는 다른 추측 경로를 dataDir로 사용하지 않는다.

로그인 하드게이트:
- startup, doctor, auth status의 setup.state만 확인하고 해당 상태의 setup.nextAction 하나만 수행한다.
- SETUP_REQUIRED: 설정 누락을 한 문장으로 보고하고 멈춘다.
- FOLDER_REQUIRED: 현재 작업이 프로젝트 폴더 없이 시작된 상태다. 사용자가 기존 Hare 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 열고 같은 프롬프트를 붙여넣도록 한 문장으로 안내한 뒤 멈춘다.
- LOGIN_START_REQUIRED: setup.nextCommand를 수정하지 않고 한 번 실행하고 Microsoft 로그인 주소와 userCode를 보여준 뒤 멈춘다.
- LOGIN_COMPLETE_REQUIRED: 사용자가 "로그인 완료"라고 말할 때까지 기다린다. 완료 후 setup.nextCommand를 수정하지 않고 한 번 실행한다. COMPLETE는 선택 프로젝트에 저장된 캐시 재검증까지 성공했다는 뜻이다.
- READY: 사용 준비 완료를 보고하고 사용자의 업무 조회 요청을 기다린다.
- BLOCKED에 AUTH_CHECK_BLOCKED가 있으면 loggedIn/tokenUsable의 null은 미로그인이 아니라 확인 불가를 뜻한다. 위 실행 환경 사전 확인 절차를 한 번 적용하며 재로그인이나 캐시 초기화는 하지 않는다.
- BLOCKED에 TOKEN_ACQUISITION_FAILED 또는 네트워크 오류가 있으면 기존 캐시를 유지하고 오류만 보고한다. login-start를 실행하지 않는다.
- HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 함께 나오면 NETWORK_PERMISSION_REQUIRED로 보고하고 막힌 도메인만 알려준 뒤 멈춘다.
- 그 밖의 명령 실패는 BLOCKED로 취급해 실패 단계와 오류 한 줄만 보고한다. 다른 경로를 추측하거나 반복 실행하지 않는다.

저장 위치:
- 로컬 Windows/Mac/Linux에서는 OS 기본 고정 dataDir를 사용한다.
- Cowork에서는 /sessions/<session>/mnt/<selected-project> 형태의 현재 선택 프로젝트를 dataDir로 사용한다.
- 앱 코드는 세션 런타임에 두고 로그인 캐시와 사용자 결과만 선택 프로젝트에 유지한다.
- startup이 반환한 setup.nextCommand에는 동일한 --data-dir가 포함된다. 환경변수에 의존하거나 명령을 다시 조합하지 않고 그대로 실행한다.

Cowork 네트워크 실패 때:
- HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 나오면 Cowork 설정의 기능 > 도메인 허용 목록을 확인한다.
- 설정을 변경한 뒤에는 새 Cowork 채팅을 열고 같은 HareM365Agent 폴더를 다시 연결해야 새 정책이 적용될 수 있다.
- 새 작업에서 실패했던 명령만 한 번 재시도하고, 일반적인 fetch failed 또는 network_error에는 실제 오류 한 줄만 보고한다.

읽기 명령:
\`\`\`bash
node dist/cli.js outlook recent --folder all --limit 10
node dist/cli.js outlook inbox --limit 10
node dist/cli.js outlook flagged --folder all --limit 1000
node dist/cli.js outlook search --query "keyword" --since 2026-04-01 --until 2026-07-10 --folder all
node dist/cli.js outlook search --mailbox "CTO" --query "keyword" --folder all
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all
node dist/cli.js outlook attachments list --mailbox "<shared-mailbox-address>" --message-id "<message-id>"
node dist/cli.js outlook attachments download --mailbox "<shared-mailbox-address>" --message-id "<message-id>" --attachment-id "<attachment-id>"
node dist/cli.js outlook draft new --to "user@example.com" --subject "Subject" --body "Body"
node dist/cli.js outlook draft reply --message-id "<message-id>" --body "Reply body"
node dist/cli.js outlook draft reply --message-id "<message-id>" --reply-all --body "Reply-all body"
node dist/cli.js outlook draft forward --message-id "<message-id>" --to "user@example.com" --body "Forward note" --attachment "<file-path>"
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 100 --offset 0
node dist/cli.js teams search-messages --query "keyword" --since 2026-04-01 --until 2026-07-10
node dist/cli.js teams attachments list --chat-id "<chat-id>" --message-id "<message-id>"
node dist/cli.js teams attachments download --chat-id "<chat-id>" --message-id "<message-id>" --attachment-id "<attachment-id>"
node dist/cli.js teams inline-images list --chat-id "<chat-id>" --message-id "<message-id>"
node dist/cli.js teams inline-images download --chat-id "<chat-id>" --message-id "<message-id>" --hosted-content-id "<hosted-content-id>"
node dist/cli.js sharepoint sites --query "Agent Automation"
node dist/cli.js sharepoint export-files --site-url "https://ohmylab.sharepoint.com/sites/<site>" --destination "<absolute-path>"
node dist/cli.js files search --query "keyword" --limit 10
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
\`\`\`

주의:
- Outlook, Teams, SharePoint, OneDrive, Microsoft 365 조회와 Outlook 초안 작성에는 Hare CLI만 사용한다. Microsoft 365 커넥터, 다른 커넥터, Computer Use, Outlook/Teams/SharePoint UI 또는 브라우저 자동화를 검색하거나 호출하거나 대체 수단으로 사용하지 않는다.
- 연결 오류는 위 실행 환경 사전 확인 절차를 적용한다. 그 외에 Hare가 지원하지 않거나 실행에 실패하면 실패한 Hare 단계와 오류만 보고하고 멈춘다. 다른 도구나 데이터 소스로 우회하지 않는다.
- 일반적인 메일 조회와 최근 메일 요청은 outlook recent --folder all을 사용해 삭제된 항목을 제외한 전체 메일함을 대상으로 한다. outlook inbox는 사용자가 받은편지함을 명시한 경우에만 사용한다.
- 플래그된 메일 요청은 outlook flagged --folder all을 사용한다. 모든 메일 조회 결과의 flagStatus를 함께 확인한다.
- --mailbox가 없는 Outlook 명령은 로그인한 사용자의 사서함만 조회한다. 사용자가 공유 사서함 이름이나 주소를 명시하면 recent, flagged, search, count에 --mailbox를 사용한다. 이름 후보가 여러 개면 정확한 주소를 확인하고, 공유 사서함 조회가 실패해도 본인 사서함으로 대체하지 않는다.
- 공유 사서함 메일의 첨부파일은 조회에 사용한 동일한 --mailbox를 attachments list와 attachments download에도 전달한다.
- 사용자가 기간을 지정한 조회는 inbox/chat-messages의 최근 건수 제한으로 대신하지 말고 search 명령의 --since/--until에 반영한다.
- 사용자가 기간을 지정하지 않은 검색은 최근 90일을 조회한다. 결과 JSON의 search.range.notice를 사용자에게 알려 실제 조회 범위를 명확히 한다.
- search.limitReached가 true이면 일부 결과만 반환된 것이므로 사용자에게 한도 도달 사실을 알린다.
- outlook search는 기본 100건씩 반환한다. search.continuationAvailable이 true이면 search.nextCursor를 --cursor에 그대로 전달해 다음 결과를 이어서 조회한다.
- outlook search의 body와 bodyHtml은 전체 메일 본문이다. fullBodyUnavailableCount가 0이 아니면 본문 누락 사실을 알리고 bodyPreview를 전체 본문으로 간주하지 않는다.
- teams search-messages는 명령당 최대 100개의 고유 메시지만 전체 본문으로 조회하며 Microsoft Search 검색 창은 최대 1,000건이다. limit를 100보다 높이지 말고 searchWindowExhausted 또는 noProgressDetected가 true이면 반복을 중단한다.
- totalMatchesReported는 발신자, 본문, 첨부파일 중 검색어가 일치한 메시지 수이며 본문 안의 정확한 단어 등장 횟수가 아니다. 이 값을 재검증하려고 모든 페이지를 반복 조회하지 않는다.
- 본문 안의 정확한 등장 횟수를 요청받으면 차이를 먼저 설명하고 첫 페이지만 조회한다. totalMatchesReported가 searchWindowLimit보다 크면 정확한 전체 집계가 불가능하므로 더 진행하지 말고 기간이나 검색어를 좁혀 달라고 요청한다. 그 이하일 때만 1,000건 검색 창 안에서 본문을 확인하며, 중복 또는 무진행 상태가 나오면 같은 페이지를 반복하지 않는다.
- teams search-messages의 search.partialResult가 true이면 시간 예산 안에 처리한 부분 결과다. partialReason과 fullBodyUnavailableCount를 알리고, 필요한 경우 같은 범위를 더 작은 limit으로 다시 조회한다.
- 메일 건수 질문은 검색 인덱스 결과를 세지 말고 outlook count로 전체 페이지를 검사한다. count.complete가 false이면 nextCursor를 --cursor에 전달해 계속한다. 커서가 누적 집계를 보존하므로 complete가 true인 마지막 matchedCount만 정확한 전체 건수로 답한다.
- files search는 접근 가능한 SharePoint, Teams, OneDrive 파일 전체를 검색한다. search.continuationAvailable이 true이면 search.nextOffset을 --offset에 전달한다. SharePoint 사이트 자체의 존재 여부는 sharepoint sites로 확인한다.
- Teams 채팅 첨부파일은 teams attachments list로 메타데이터를 확인한 뒤 teams attachments download로 내려받는다. 원본 공유 URL은 출력하지 않으며 기존 다운로드 크기 상한과 승인 절차를 그대로 적용한다.
- Teams 메시지 본문이 비어 있고 bodyHtml에 hostedContents 이미지가 있으면 teams inline-images list로 ID를 확인한 뒤 teams inline-images download로 내려받아 판독한다. 다운로드 응답이 허용된 래스터 이미지 형식인지 검증하며 Graph 원본 URL이나 바이너리를 출력하지 않는다. 크기 미상 이미지는 기본 다운로드 상한까지만 허용한다.
- 다운로드가 기본 상한을 초과해 AWAITING_USER_APPROVAL을 반환하면 출처, 파일명, 크기, 출력명, 저장 위치와 상한을 모두 사용자에게 보여주고 멈춘다. 사용자가 명시적으로 동의한 뒤에만 동일한 명령에 반환된 --approval-token을 추가해 한 번 실행한다. 토큰은 10분 동안 정확히 같은 파일과 출력명에만 유효하며 재사용할 수 없다.
- SharePoint 사이트 파일 일괄 복사는 sharepoint export-files를 approval-token 없이 먼저 실행한다. ENUMERATION_IN_PROGRESS가 반환되면 저장된 Graph 페이지 체크포인트를 이어가도록 approval-token 없이 같은 명령을 다시 실행한다. 전체 파일 수, 총용량, 기존 파일, 10GiB 초과 제외 파일, 대상 경로와 디스크 여유가 포함된 전체 계획이 나온 뒤에만 멈춰 사용자 승인을 받는다. 사용자가 전체 계획을 명시적으로 승인한 뒤에만 같은 명령에 반환된 --approval-token을 추가한다. 같은 계정, 사이트, 파일 목록, 대상과 정책에 묶인 승인 작업은 24시간 동안 파일별 추가 승인 없이 같은 명령으로 재개하며 기존 파일을 덮어쓰지 않는다.
- Outlook 초안 작성 요청은 반드시 Hare CLI로 처리한다. 초안을 만들거나 본문을 붙여넣기 위해 Computer Use, Outlook 데스크톱/웹 UI, 브라우저 자동화 또는 Microsoft 365 커넥터를 사용하지 않는다.
- Hare 초안 명령이 실패하면 실패 단계와 오류만 보고하고 멈춘다. GUI 자동화나 다른 커넥터로 우회하지 않는다.
- Outlook 초안은 신규, 답장, 전체답장, 전달과 첨부파일을 지원한다. 먼저 approval-token 없이 명령을 실행해 AWAITING_USER_APPROVAL 미리보기를 만들고 수신자, 제목, 본문, 첨부파일 전체를 사용자에게 보여준 뒤 멈춘다.
- 사용자가 명시적으로 동의한 경우에만 동일한 명령에 반환된 --approval-token을 추가해 한 번 실행한다. 내용이나 첨부가 바뀌면 기존 토큰을 재사용하지 말고 새 미리보기와 동의를 받는다.
- 메일 발송은 지원하지 않는다. Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경도 수행하지 않는다.
- login-start의 userCode와 verificationUri는 현재 사용자에게 한 번 보여준다. 내부 device_code, 토큰 캐시, 쿠키, credential 값은 출력하거나 반복하지 않는다.
- 로그인 계정은 "Hare를 실제로 사용할 사용자 본인의 회사 Microsoft 계정"으로만 안내한다. 기존 캐시, 예시, 대화 문맥에서 발견한 특정 이메일 주소를 로그인 대상으로 표시하거나 추천하지 않는다.
- Teams 최신 채팅 판단은 lastMessageCreatedDateTime을 우선한다. lastUpdatedDateTime만으로 최신 메시지 방을 판단하지 않는다.
- teams chat-messages의 body와 bodyHtml은 잘리지 않은 전체 본문이다. bodyPreview는 호환용 전체 텍스트 별칭이며 요약 필드로 취급하지 않는다.
- teams chat-messages는 기본 20건, 명령당 최대 1,000개의 고유 메시지를 반환한다. page.continuationAvailable이 true이면 page.nextOffset을 --offset에 전달해 이어서 조회한다.
- teams search-messages는 검색 스니펫과 별도로 메시지 상세를 조회해 body와 bodyHtml에 전체 본문을 반환한다. fullBodyUnavailableCount가 0이 아니면 누락된 전체 본문이 있다고 사용자에게 알리고, searchSummary를 전체 본문으로 간주하지 않는다.
`;

const llmPrompt = `목적: Hare M365 Agent 연결 상태를 준비하고, 이후 내 Microsoft 365 업무 요청을 처리해.

저장소:
${repoUrl}

Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 먼저 아래 7개 도메인이 허용되어 있는지 확인해.
github.com
registry.npmjs.org
graph.microsoft.com
login.microsoftonline.com
outlook.office.com
ohmylab-my.sharepoint.com
ohmylab.sharepoint.com

도메인 허용 확인은 가능하면 네가 가진 사용자 확인 요청, 선택형 질문, 승인 요청 UI로 처리해. 사용자가 직접 "도메인 허용 완료" 같은 문구를 타이핑하게 요구하지 마.
그런 확인 요청 기능이 없는 환경에서만 짧게 yes/no로 물어봐.

도메인 허용 확인 전에는 git, npm, 로그인, Outlook/Teams/Files 조회를 시작하지 마.
git, npm, 로그인, Graph 호출은 도메인 허용 목록이 적용되는 Cowork 세션 셸에서 실행해.
도메인 설정을 이 Cowork 작업을 연 뒤 변경했다면 새 Cowork 채팅을 열어야 변경된 정책이 적용돼.
GitHub API(api.github.com)나 GitHub Release asset 다운로드로 repo 접근을 판단하지 마.
이 환경에서는 api.github.com이 403이어도 git clone은 성공할 수 있어.
repo 접근은 git ls-remote 또는 git clone으로 판단해.

이 Cowork 작업을 만들 때 사용자가 미리 선택한 프로젝트 루트를 확인해. 이름이 무엇이든 현재 선택된 프로젝트 루트 하나가 Hare의 영구 dataDir야.
선택된 프로젝트가 없으면 FOLDER_REQUIRED라고 말하고, 기존 Hare 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 연 뒤 같은 프롬프트를 붙여넣으라고 한 문장으로 안내하고 멈춰.

Cowork 세션 셸이 시작되지 않고 workspace unavailable, still starting 또는 유사한 작업공간 시작 오류가 나오면 WORKSPACE_FOLDER_RETRY라고 알려줘. 가상화 진단이나 다른 명령을 반복하지 말고, OneDrive가 아닌 PC 로컬 바탕화면에 HareM365Agent 새 폴더를 만든 뒤 그 폴더 하나만 프로젝트로 선택해 새 Cowork 작업을 열고 같은 프롬프트를 붙여넣으라고 안내하고 멈춰.

아래 순서로 준비해.

${setupCommandForGuide}

준비 명령은 git/npm/build를 세션 런타임에서 수행하고 --data-dir에는 현재 선택 프로젝트만 사용해. 프로젝트 폴더 안에서 git clone, npm ci 또는 build를 실행하지 마. 프로젝트 폴더의 삭제 권한을 요청하지 마. /tmp, /dev/shm, /home/claude, /root/.local/share 또는 다른 경로를 dataDir로 시도하지 마.

startup JSON의 setup.state만 확인하고 아래에서 일치하는 한 단계만 수행해. setup.state와 다른 경로를 추측하지 마.

- SETUP_REQUIRED: 설정 누락을 한 문장으로 알려주고 멈춰.
- FOLDER_REQUIRED: 기존 Hare 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 열고 같은 프롬프트를 붙여넣으라고 한 문장으로 안내한 뒤 멈춰.
- LOGIN_START_REQUIRED: setup.nextCommand를 수정하지 않고 한 번 실행해. 반환된 Microsoft 로그인 주소와 userCode를 보여주고 "로그인 완료"라고 말해달라고 안내한 뒤 멈춰.
- 로그인 계정은 "Hare를 실제로 사용할 사용자 본인의 회사 Microsoft 계정"으로만 안내해. 특정 이메일 주소를 예시로 들거나 로그인 대상으로 지정하지 마.
- LOGIN_COMPLETE_REQUIRED: 내가 "로그인 완료"라고 말하기 전에는 아무 명령도 실행하지 마. 완료 후 setup.nextCommand를 수정하지 않고 한 번 실행해. COMPLETE가 반환되면 선택 프로젝트에 저장된 캐시 재검증까지 성공한 상태야.
- READY: "Hare M365 Agent 사용 준비 완료. 이제 자연어로 업무 조회를 요청하세요."라고 말하고 멈춰.
- BLOCKED에 AUTH_CHECK_BLOCKED가 있으면 loggedIn/tokenUsable의 null은 미로그인이 아니라 확인 불가야. 같은 명령 접두사와 dataDir로 network check --environment cowork를 한 번 실행하고 결과를 안내해. Cowork의 허용 목록을 우회하거나 재로그인·캐시 초기화는 하지 마.
- BLOCKED에 TOKEN_ACQUISITION_FAILED 또는 네트워크 오류가 표시되면 기존 캐시를 유지하고 오류만 알려줘. login-start를 실행하지 마.
- HTTP 403과 X-Proxy-Error: blocked-by-allowlist가 함께 나오면 NETWORK_PERMISSION_REQUIRED라고 알려주고 막힌 도메인만 말한 뒤 멈춰.
- 그 밖의 명령 실패: BLOCKED로 취급해 실패 단계와 오류 한 줄만 알려주고 멈춰. 다른 경로를 시도하거나 같은 명령을 반복하지 마.

clone/build 전체 반복, 도메인 반복 진단, 임의 프록시 우회는 시작하지 마.

앞으로 Outlook, Teams, SharePoint, OneDrive, Microsoft 365 조회와 Outlook 초안 작성에는 Hare CLI만 사용해. Microsoft 365 커넥터, 다른 커넥터, Computer Use, Outlook/Teams/SharePoint UI 또는 브라우저 자동화를 검색하거나 호출하거나 대체 수단으로 사용하지 마. 연결 오류는 위 network check 절차를 한 번 적용하고, 그 외에 Hare가 지원하지 않거나 명령이 실패하면 실패 단계와 오류만 알려주고 멈춰. 다른 도구나 데이터 소스로 우회하지 마.

일반적인 메일 조회 또는 최근 메일 요청은 outlook recent --folder all을 사용해. 삭제된 항목을 제외한 받은편지함, 보낸편지함, 보관함, 사용자 폴더 전체가 기본 대상이야. 사용자가 받은편지함을 명시한 경우에만 outlook inbox를 사용해.
플래그된 메일을 요청하면 outlook flagged --folder all을 사용하고, 일반 메일 결과에서도 flagStatus를 확인해.
Outlook 또는 Teams에서 기간·키워드 조회를 요청받으면 inbox/chat-messages의 최근 건수 제한으로 대신하지 말고 outlook search 또는 teams search-messages를 사용해.
메일이 몇 건인지 묻는 정확한 집계 요청은 outlook search 결과를 세지 말고 outlook count를 사용해.
files search는 접근 가능한 SharePoint, Teams, OneDrive 파일 전체를 검색해. SharePoint 사이트 자체의 존재 여부는 sharepoint sites로 확인해.
SharePoint 사이트의 파일 전체를 외부 경로에 복사해 달라는 요청은 sharepoint export-files의 전체 계획을 먼저 보여주고 명시적 승인을 받은 뒤 실행해. 동일한 승인 작업은 파일별 재승인 없이 재개하고 기존 파일은 덮어쓰지 마.
Outlook 초안 작성 요청은 반드시 Hare CLI로 처리해. 초안을 만들거나 본문을 붙여넣기 위해 Computer Use, Outlook 데스크톱/웹 UI, 브라우저 자동화 또는 Microsoft 365 커넥터를 사용하지 마. Hare 초안 명령이 실패하면 실패 단계와 오류만 알려주고 멈춰. 다른 방식으로 우회하지 마.
사용자가 기간을 말하면 --since/--until에 그대로 반영해.
기간을 말하지 않으면 기본 최근 90일이 적용되며, 결과의 search.range.notice를 답변에 포함해 실제 조회 범위를 알려줘.
search.limitReached가 true이면 검색 결과 한도에 도달했다고 알려줘.
Outlook 검색의 search.continuationAvailable이 true이면 search.nextCursor를 --cursor에 전달해 다음 100건을 이어서 조회해. body와 bodyHtml을 전체 본문으로 사용하고 fullBodyUnavailableCount를 확인해.
Teams 검색은 limit를 100보다 높이지 마. totalMatchesReported는 발신자·본문·첨부파일을 포함한 검색 일치 메시지 수이지 본문 등장 횟수가 아니야. 정확한 본문 등장 횟수가 필요하면 첫 페이지만 확인하고, totalMatchesReported가 searchWindowLimit보다 크면 이어서 조회하지 말고 기간이나 검색어를 좁혀 달라고 요청해. 그 이하일 때만 search.nextOffset으로 이어가며 searchWindowExhausted 또는 noProgressDetected가 true이면 즉시 중단해.
파일 검색의 search.continuationAvailable이 true이면 search.nextOffset을 --offset에 전달해 다음 결과를 이어서 조회해.
outlook count의 count.complete가 false이면 nextCursor를 --cursor에 전달해 계속 조회해. 커서가 누적값을 보존하므로 complete가 true인 마지막 matchedCount만 정확한 전체 건수로 답해.
search.partialResult가 true이면 시간 예산 안에 처리한 부분 결과임을 알리고 partialReason과 fullBodyUnavailableCount를 함께 설명해.`;
function getSelfCommand(): string {
  return getExplicitSelfCommand();
}

function getExplicitSelfCommand(): string {
  const command = process.env.HARE_M365_COMMAND ?? defaultCliCommand;
  return `${command} --data-dir ${quoteCommandArgument(config.dataDir)}`;
}

function quoteCommandArgument(value: string): string {
  if (/[\0\r\n]/.test(value)) throw new Error("Hare command arguments cannot contain null bytes or line breaks.");
  return process.platform === "win32"
    ? `'${value.replaceAll("'", "''")}'`
    : `'${value.replaceAll("'", `'"'"'`)}'`;
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
  const authStatus: AuthStatus = config.clientId && config.tenantId
    ? await getAuthStatus(config)
    : {
        account: null,
        loggedIn: false,
        tokenUsable: false,
        migrationRequired: false,
        reason: "NOT_CONFIGURED"
      };
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  return {
    configured: Boolean(config.clientId && config.tenantId),
    clientIdPresent: Boolean(config.clientId),
    tenantIdPresent: Boolean(config.tenantId),
    loggedIn: authStatus.loggedIn,
    tokenUsable: authStatus.tokenUsable,
    authMigrationRequired: authStatus.migrationRequired,
    authReason: authStatus.reason,
    authNetworkFailure: authStatus.networkFailure,
    dataDir: config.dataDir,
    dataDirSource: config.dataDirSource,
    dataDirPersistent: config.dataDirPersistent,
    policyPath: config.policyPath,
    cacheDir: config.cacheDir,
    cacheFile,
    cacheFileExists: fs.existsSync(cacheFile),
    pendingLoginStateExists: hasPendingDeviceLoginState(config),
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
          appDir: config.dataDirPersistent ? packageRoot : undefined,
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
              : "Open a Cowork task with the user's existing Hare project selected; the session rules file is created only in that selected persistent project."
          },
          storageRule:
            "Use setup.state and perform setup.nextAction only. Do not infer a different setup path."
        },
        null,
        2
      )
    );
  });

program.command("network").description("Unauthenticated connection diagnostics")
  .command("check")
  .description("Check Microsoft connectivity without reading or changing the data directory")
  .addOption(new Option("--environment <host>", "Actual host; this option does not grant execution permission")
    .choices(["codex", "cowork", "unknown"]).default("unknown"))
  .action(async (options: { environment: ExecutionEnvironment }) => {
    const result = await checkMicrosoftConnectivity(options.environment);
    console.log(JSON.stringify(result, null, 2));
    if (!result.ok) process.exitCode = 1;
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
          "Show verificationUri and userCode to the user. Tell the user to sign in with their own company Microsoft account that they will use with Hare. Never name, recommend, or preselect a specific email address. End this shell call. After the user completes Microsoft sign-in, run nextCommand in a new foreground shell call."
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
  const pendingLoginStateExists = hasPendingDeviceLoginState(config);
  const setup = buildSetupContract(
    {
      configured: true,
      loggedIn,
      tokenUsable: authStatus.tokenUsable,
      authMigrationRequired: authStatus.migrationRequired,
      authReason: authStatus.reason,
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
        authMigrationRequired: authStatus.migrationRequired,
        authReason: authStatus.reason,
        authNetworkFailure: authStatus.networkFailure,
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
  .command("recent")
  .description("List recent messages across the mailbox; deleted items are excluded for all scope")
  .option("--mailbox <name-or-address>", "shared mailbox display name or exact email address")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--limit <number>", "maximum message count", "10")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { mailbox?: string; folder: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const mailbox = await resolveMailboxTarget(config, options.mailbox);
    const data = await runMailboxRead(
      mailbox,
      () => listRecentMailbox(
        config,
        parseMailFolderScope(options.folder),
        Number(options.limit),
        mailbox
      )
    );
    emitJson(data, options.out);
  });

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
  .command("flagged")
  .description("List flagged Outlook messages by date range")
  .option("--mailbox <name-or-address>", "shared mailbox display name or exact email address")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--limit <number>", "maximum flagged message count", "1000")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      since?: string;
      until?: string;
      mailbox?: string;
      folder: string;
      limit: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const mailbox = await resolveMailboxTarget(config, options.mailbox);
      const data = await runMailboxRead(
        mailbox,
        () => listFlaggedMessages(
          config,
          options.since,
          options.until,
          parseMailFolderScope(options.folder),
          Number(options.limit),
          mailbox
        )
      );
      emitJson(data, options.out);
    }
  );

outlook
  .command("search")
  .description("Search Outlook messages by keyword and date range")
  .requiredOption("--query <text>", "search query or Outlook KQL")
  .option("--mailbox <name-or-address>", "shared mailbox display name or exact email address")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--limit <number>", "maximum matching message count per page", "100")
  .option("--cursor <cursor>", "opaque continuation cursor from search.nextCursor")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      query: string;
      mailbox?: string;
      since?: string;
      until?: string;
      folder: string;
      limit: string;
      cursor?: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const mailbox = await resolveMailboxTarget(config, options.mailbox);
      const data = await runMailboxRead(
        mailbox,
        () => searchMailbox(
          config,
          options.query,
          options.since,
          options.until,
          parseMailFolderScope(options.folder),
          Number(options.limit),
          { cursor: options.cursor, mailbox }
        )
      );
      emitJson(data, options.out);
    }
  );

outlook
  .command("count")
  .description("Count Outlook messages exactly by scanning every page in a date range")
  .option("--mailbox <name-or-address>", "shared mailbox display name or exact email address")
  .option("--subject-contains <text>", "literal text that must appear in the subject")
  .option("--from <text>", "text that must appear in the sender name or address")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--folder <scope>", "mailbox scope: all, inbox, or sent", "all")
  .option("--cursor <cursor>", "opaque continuation cursor from count.nextCursor")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      subjectContains?: string;
      from?: string;
      mailbox?: string;
      since?: string;
      until?: string;
      folder: string;
      cursor?: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const mailbox = await resolveMailboxTarget(config, options.mailbox);
      const data = await runMailboxRead(
        mailbox,
        () => countMailboxMessages(
          config,
          options.subjectContains,
          options.from,
          options.since,
          options.until,
          parseMailFolderScope(options.folder),
          { cursor: options.cursor, mailbox }
        )
      );
      emitJson(data, options.out);
    }
  );

const outlookAttachments = outlook
  .command("attachments")
  .description("List and download attachments from an Outlook message");

outlookAttachments
  .command("list")
  .description("List attachment metadata for one Outlook message")
  .requiredOption("--message-id <id>", "message ID returned by an Outlook read command")
  .option("--mailbox <name-or-address>", "shared mailbox used to retrieve the message")
  .option("--limit <number>", "maximum attachment count", "20")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: {
    messageId: string;
    mailbox?: string;
    limit: string;
    out?: string;
  }) => {
    requireConfigured(config);
    const mailbox = await resolveMailboxTarget(config, options.mailbox);
    const data = await runMailboxRead(
      mailbox,
      () => listMessageAttachments(
        config,
        options.messageId,
        Number(options.limit),
        undefined,
        mailbox
      )
    );
    emitJson(data, options.out);
  });

outlookAttachments
  .command("download")
  .description("Download one Outlook message attachment")
  .requiredOption("--message-id <id>", "message ID returned by an Outlook read command")
  .requiredOption("--attachment-id <id>", "attachment ID returned by outlook attachments list")
  .option("--mailbox <name-or-address>", "shared mailbox used to retrieve the message")
  .option("--name <filename>", "output filename; defaults to the attachment name")
  .option(
    "--approval-token <token>",
    "one-time token returned after previewing a download above the default size limit"
  )
  .action(async (options: {
    messageId: string;
    attachmentId: string;
    mailbox?: string;
    name?: string;
    approvalToken?: string;
  }) => {
    requireConfigured(config);
    const mailbox = await resolveMailboxTarget(config, options.mailbox);
    const data = await runMailboxRead(
      mailbox,
      () => downloadMessageAttachment(
        config,
        options.messageId,
        options.attachmentId,
        options.name,
        options.approvalToken,
        undefined,
        mailbox
      )
    );
    emitJson({ ok: true, ...data });
  });

type DraftCliOptions = {
  messageId?: string;
  replyAll?: boolean;
  to?: string;
  cc?: string;
  bcc?: string;
  subject?: string;
  body?: string;
  bodyFile?: string;
  format: string;
  attachment: string[];
  approvalToken?: string;
};

const outlookDraft = outlook
  .command("draft")
  .description("Preview and create Outlook drafts; sending is not available");

addDraftCommonOptions(
  outlookDraft
    .command("new")
    .description("Preview or create a new Outlook draft")
    .requiredOption("--to <addresses>", "comma-separated To recipients")
    .requiredOption("--subject <text>", "draft subject")
).action(async (options: DraftCliOptions) => {
  await runDraftCommand("new", options);
});

addDraftCommonOptions(
  outlookDraft
    .command("reply")
    .description("Preview or create a reply or reply-all draft")
    .requiredOption("--message-id <id>", "source message ID returned by an Outlook read command")
    .option("--reply-all", "reply to the sender and all original recipients")
    .option("--to <addresses>", "comma-separated additional To recipients")
).action(async (options: DraftCliOptions) => {
  await runDraftCommand(options.replyAll ? "replyAll" : "reply", options);
});

addDraftCommonOptions(
  outlookDraft
    .command("forward")
    .description("Preview or create a forward draft")
    .requiredOption("--message-id <id>", "source message ID returned by an Outlook read command")
    .requiredOption("--to <addresses>", "comma-separated To recipients")
).action(async (options: DraftCliOptions) => {
  await runDraftCommand("forward", options);
});

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
  .option("--limit <number>", "maximum message count (hard maximum: 1000)", "20")
  .option("--offset <number>", "unique message offset for the next page", "0")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: {
    chatId: string;
    limit: string;
    offset: string;
    out?: string;
  }) => {
    requireConfigured(config);
    const data = await listChatMessagesPage(
      config,
      options.chatId,
      Number(options.limit),
      Number(options.offset)
    );
    emitJson(data, options.out);
  });

teams
  .command("search-messages")
  .description("Search Teams messages by keyword and date range")
  .requiredOption("--query <text>", "search query or Teams KQL")
  .option("--since <YYYY-MM-DD>", "inclusive start date; defaults to the last 90 days")
  .option("--until <YYYY-MM-DD>", "inclusive end date; defaults to today")
  .option("--limit <number>", "maximum matching message count per command (hard maximum: 100)", "100")
  .option("--offset <number>", "result offset for the next page", "0")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(
    async (options: {
      query: string;
      since?: string;
      until?: string;
      limit: string;
      offset: string;
      out?: string;
    }) => {
      requireConfigured(config);
      const data = await searchChatMessages(
        config,
        options.query,
        options.since,
        options.until,
        Number(options.limit),
        { offset: Number(options.offset) }
      );
      emitJson(data, options.out);
    }
  );

const teamsAttachments = teams
  .command("attachments")
  .description("List and download SharePoint files attached to a Teams chat message");

teamsAttachments
  .command("list")
  .description("List attachment metadata for one Teams chat message")
  .requiredOption("--chat-id <id>", "chat ID returned by a Teams read command")
  .requiredOption("--message-id <id>", "message ID returned by a Teams read command")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { chatId: string; messageId: string; out?: string }) => {
    requireConfigured(config);
    const data = await listTeamsMessageAttachments(config, options.chatId, options.messageId);
    emitJson(data, options.out);
  });

teamsAttachments
  .command("download")
  .description("Download one SharePoint file attached to a Teams chat message")
  .requiredOption("--chat-id <id>", "chat ID returned by a Teams read command")
  .requiredOption("--message-id <id>", "message ID returned by a Teams read command")
  .requiredOption("--attachment-id <id>", "attachment ID returned by teams attachments list")
  .option("--name <filename>", "output filename; defaults to the attachment name")
  .option(
    "--approval-token <token>",
    "one-time token returned after previewing a download above the default size limit"
  )
  .action(async (options: {
    chatId: string;
    messageId: string;
    attachmentId: string;
    name?: string;
    approvalToken?: string;
  }) => {
    requireConfigured(config);
    const data = await downloadTeamsMessageAttachment(
      config,
      options.chatId,
      options.messageId,
      options.attachmentId,
      options.name,
      options.approvalToken
    );
    emitJson({ ok: true, ...data });
  });

const teamsInlineImages = teams
  .command("inline-images")
  .description("List and download images embedded in a Teams chat message");

teamsInlineImages
  .command("list")
  .description("List inline image metadata for one Teams chat message")
  .requiredOption("--chat-id <id>", "chat ID returned by a Teams read command")
  .requiredOption("--message-id <id>", "message ID returned by a Teams read command")
  .option("--limit <number>", "maximum inline image count", "20")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { chatId: string; messageId: string; limit: string; out?: string }) => {
    requireConfigured(config);
    const data = await listTeamsInlineImages(
      config,
      options.chatId,
      options.messageId,
      Number(options.limit)
    );
    emitJson(data, options.out);
  });

teamsInlineImages
  .command("download")
  .description("Download one image embedded in a Teams chat message")
  .requiredOption("--chat-id <id>", "chat ID returned by a Teams read command")
  .requiredOption("--message-id <id>", "message ID returned by a Teams read command")
  .requiredOption("--hosted-content-id <id>", "hosted content ID returned by teams inline-images list")
  .option("--name <filename>", "output filename; defaults to a safe image filename")
  .action(async (options: {
    chatId: string;
    messageId: string;
    hostedContentId: string;
    name?: string;
  }) => {
    requireConfigured(config);
    const data = await downloadTeamsInlineImage(
      config,
      options.chatId,
      options.messageId,
      options.hostedContentId,
      options.name
    );
    emitJson({ ok: true, ...data });
  });

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

sharepoint
  .command("export-files")
  .description("Export the default document library while preserving its folder structure")
  .requiredOption("--site-url <url>", "SharePoint site URL")
  .requiredOption("--destination <path>", "absolute destination directory")
  .option(
    "--approval-token <token>",
    "one-time token returned after previewing the complete export manifest"
  )
  .option(
    "--time-budget-ms <number>",
    "stop between files after this time and return a resumable result",
    "240000"
  )
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: {
    siteUrl: string;
    destination: string;
    approvalToken?: string;
    timeBudgetMs: string;
    out?: string;
  }) => {
    requireConfigured(config);
    const auth = await getAuthStatus(config);
    if (!auth.tokenUsable || !auth.account?.homeAccountId) {
      throw new Error("A usable logged-in account is required for SharePoint export.");
    }
    const result = await exportSharePointFiles(
      config,
      options.siteUrl,
      options.destination,
      {
        accountId: auth.account.homeAccountId,
        approvalToken: options.approvalToken,
        timeBudgetMs: Number(options.timeBudgetMs)
      }
    );
    emitJson({ ok: true, ...result }, options.out);
  });


files
  .command("search")
  .description("Search accessible files across SharePoint, Teams, and OneDrive")
  .requiredOption("--query <text>", "search query")
  .option("--limit <number>", "maximum file count", "10")
  .option("--offset <number>", "result offset for the next page", "0")
  .option("--out <path>", "write JSON result to a file; relative paths are saved under Hare resultsDir")
  .action(async (options: { query: string; limit: string; offset: string; out?: string }) => {
    requireConfigured(config);
    const data = await searchFiles(
      config,
      options.query,
      Number(options.limit),
      { offset: Number(options.offset) }
    );
    emitJson(data, options.out);
  });

files
  .command("download")
  .description("Download one file by driveId and itemId")
  .requiredOption("--drive-id <id>", "drive ID from parentReference.driveId")
  .requiredOption("--item-id <id>", "drive item ID")
  .option("--name <filename>", "output filename")
  .option(
    "--approval-token <token>",
    "one-time token returned after previewing a download above the default size limit"
  )
  .action(async (options: {
    driveId: string;
    itemId: string;
    name?: string;
    approvalToken?: string;
  }) => {
    requireConfigured(config);
    const result = await downloadDriveItem(
      config,
      options.driveId,
      options.itemId,
      options.name,
      options.approvalToken
    );
    if (typeof result === "string") {
      emitJson({ ok: true, stage: "DOWNLOADED", outputPath: result });
      return;
    }
    emitJson({ ok: true, ...result });
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

function addDraftCommonOptions(command: Command): Command {
  return command
    .option("--cc <addresses>", "comma-separated Cc recipients")
    .option("--bcc <addresses>", "comma-separated Bcc recipients")
    .option("--body <text>", "draft body text or HTML")
    .option("--body-file <path>", "read the draft body from a UTF-8 file")
    .option("--format <type>", "body format: text or html", "text")
    .option(
      "--attachment <path>",
      "file to attach; repeat this option for multiple files",
      collectOption,
      []
    )
    .option(
      "--approval-token <token>",
      "token returned by the preview after the user explicitly approves the exact draft"
    );
}

function collectOption(value: string, previous: string[]): string[] {
  return [...previous, value];
}

async function runDraftCommand(kind: DraftKind, options: DraftCliOptions): Promise<void> {
  requireConfigured(config);
  const input: DraftInput = {
    kind,
    sourceMessageId: options.messageId,
    to: options.to ? [options.to] : [],
    cc: options.cc ? [options.cc] : [],
    bcc: options.bcc ? [options.bcc] : [],
    subject: options.subject,
    body: readDraftBody(options),
    contentType: parseDraftContentType(options.format),
    attachmentPaths: options.attachment
  };

  if (!options.approvalToken) {
    const prepared = await prepareDraft(config, input);
    emitJson({
      ok: true,
      stage: "AWAITING_USER_APPROVAL",
      preview: prepared.preview,
      approval: {
        token: prepared.approvalToken,
        instruction:
          "Show the complete preview to the user and stop. Only after explicit approval, rerun the exact same command with --approval-token set to this token."
      }
    });
    return;
  }

  emitJson(await createApprovedDraft(config, input, options.approvalToken));
}

function readDraftBody(options: DraftCliOptions): string {
  if (options.body !== undefined && options.bodyFile !== undefined) {
    throw new Error("Use either --body or --body-file, not both.");
  }
  if (options.bodyFile !== undefined) {
    return fs.readFileSync(path.resolve(options.bodyFile), "utf8");
  }
  if (options.body !== undefined) return options.body;
  throw new Error("Draft body is required. Use --body or --body-file.");
}

function parseDraftContentType(value: string): DraftContentType {
  const normalized = value.toLowerCase();
  if (normalized === "text" || normalized === "html") return normalized;
  throw new Error("format must be text or html.");
}
