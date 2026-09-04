# Hare M365 Agent

Hare M365 Agent는 LLM이 Microsoft Graph delegated 권한으로 Outlook, Teams, OneDrive, SharePoint를 조회하고, 사용자 승인 후 Outlook 초안을 작성할 수 있도록 만든 Node/TypeScript CLI입니다.

Cowork에서는 작업을 열 때 선택한 프로젝트 폴더를 Hare의 영구 `dataDir`로 사용합니다. 앱 코드는 Cowork 세션 런타임에서 clone·build하고, 선택 프로젝트에는 인증 캐시와 사용자 결과만 저장합니다. GitHub API나 GitHub Release asset은 사용하지 않습니다.

사람이 읽는 안내는 아래 한 장만 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드.html
```

## 실행 개요

저장소:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent.git
```

선택 프로젝트(dataDir) 폴더 구조:

```text
selected-project/
├─ .cache/                   Microsoft 로그인 캐시
├─ claude/                   새 채팅에서 읽을 Hare 운영 규칙
├─ downloads/
├─ results/
└─ logs/
```

`startup.setupCommand`는 앱을 `${XDG_CACHE_HOME:-$HOME/.cache}/hare-m365-agent-runtime/app`에 준비하고, 선택 프로젝트를 정확한 `--data-dir`로 전달합니다. 세션 런타임은 다시 만들 수 있지만 선택 프로젝트의 로그인 캐시는 다음 Cowork 작업에서도 재사용합니다.

최초 설치에는 `registry.npmjs.org`가 필수입니다. `npm ci`가 실패하면 백그라운드 실행이나 `npm install` 증분 설치로 우회하지 않고 원인을 보고합니다.

Cowork의 도메인 허용 기준은 `설정 > 기능 > 도메인 허용 목록`입니다. 프리셋을 `없음`으로 선택하고 필수 도메인을 추가한 뒤, 이미 열려 있던 Cowork 작업이 있다면 새 채팅을 열어 변경된 네트워크 정책을 적용합니다.

## 설정

기본 Azure Application 설정은 `hare.config.json`에 포함됩니다. 일반 사용자는 `.env`를 만들거나 수정하지 않습니다. 과거 POC에서 만든 `.env`가 남아 있어도 앱 설정에는 사용되지 않습니다.

Azure Application 또는 요청 권한이 변경되면 Hare는 기존 앱의 인증 캐시와 진행 중인 로그인 상태만 자동 초기화합니다. 다운로드, 조회 결과, 로그와 Claude 운영 규칙은 유지됩니다. startup의 `authReason`이 `AUTH_APP_CHANGED`이면 변경된 인증 권한 또는 앱으로 Microsoft 로그인을 한 번 완료한 뒤 기존과 같이 사용합니다.

개발자용 로컬 override가 필요하면 실행 프로세스의 `OMH_M365_CLIENT_ID`, `OMH_M365_TENANT_ID` 환경 변수를 명시적으로 설정합니다.

## 로그인과 저장 위치

일반 로컬 실행은 OS별 기본 `dataDir`를 사용합니다.

```text
Windows: %LOCALAPPDATA%\Ohmyhotel\HareM365Agent
Mac: ~/Library/Application Support/Ohmyhotel/HareM365Agent
Linux local default: ~/.local/share/ohmyhotel/hare-m365-agent
```

Cowork에서는 `/sessions/<session>/mnt/<selected-project>` 형태의 현재 선택 프로젝트를 `--data-dir`로 명시합니다. `/tmp`, `/dev/shm`, 일반 `/sessions`, `/home/claude`, Linux OS 기본 저장소는 Cowork 로그인 위치로 사용하지 않습니다.

`loggedIn`과 `tokenUsable`이 모두 `true`일 때만 조회 가능한 상태입니다. 캐시 파일이 존재하더라도 토큰을 획득할 수 없으면 로그인 완료로 판단하지 않습니다. `auth login-complete`도 저장된 캐시를 다시 열어 검증한 뒤에만 `COMPLETE`를 반환합니다.

Cowork에서는 작업을 만들 때 기존 Hare 프로젝트 또는 폴더를 먼저 선택합니다. startup JSON의 `setup.state`와 `setup.nextAction` 하나만 따르며, CLI가 반환한 `setup.nextCommand`를 수정하지 않고 실행합니다. 이 명령에는 동일한 `--data-dir`가 포함되므로 셸과 채팅이 바뀌어도 같은 캐시를 사용합니다. 일반 `/sessions`, `/tmp`, Cowork Linux OS 기본 저장소에서는 로그인이 거부되지만 `/sessions/<session>/mnt/<selected-project>` 형태의 현재 프로젝트 마운트는 허용됩니다.

프로젝트 폴더 없이 작업을 시작하면 Hare는 `FOLDER_REQUIRED`로 중단합니다. 사용자는 `HareM365Agent` 프로젝트 또는 폴더를 선택해 새 Cowork 작업을 연 뒤 같은 프롬프트를 붙여넣습니다. 실행 중인 작업에서 AI가 폴더를 생성하거나 추가 접근 요청을 반복하지 않습니다.

명령에서 `HTTP 403`과 `X-Proxy-Error: blocked-by-allowlist`가 함께 나오면 Cowork 도메인 정책을 확인합니다. 설정을 변경한 뒤 새 Cowork 채팅에서 같은 프로젝트를 선택하고 실패한 명령만 한 번 재시도합니다.

초기 로그인은 45초 셸 제한에 맞춘 두 단계입니다.

```bash
node dist/cli.js auth login-start
# 사용자가 브라우저에서 Hare를 사용할 본인의 회사 Microsoft 계정으로 로그인
node dist/cli.js auth login-complete
node dist/cli.js auth status
```

로그인 안내에는 특정 이메일 주소를 예시 또는 권장 계정으로 표시하지 않습니다. 화면에 다른 계정명이 보이더라도 사용자가 Hare에서 실제로 사용할 본인의 회사 Microsoft 계정을 직접 선택합니다.

`login-start`는 주소와 user code를 즉시 반환하고 종료합니다. `login-complete`는 최대 25초만 토큰을 확인하므로 장기 poller나 백그라운드 프로세스가 필요하지 않습니다.

## 주요 명령

```bash
node dist/cli.js
node dist/cli.js auth status
node dist/cli.js outlook recent --folder all --limit 10
node dist/cli.js outlook inbox --limit 10
node dist/cli.js outlook flagged --folder all --limit 1000
node dist/cli.js outlook search --query "나이스페이 OR nicepay" --since 2026-06-26 --until 2026-07-10 --folder all
node dist/cli.js outlook search --mailbox "CTO" --query "keyword" --folder all
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all
node dist/cli.js outlook attachments list --mailbox "<shared-mailbox-address>" --message-id "<message-id>"
node dist/cli.js outlook attachments download --mailbox "<shared-mailbox-address>" --message-id "<message-id>" --attachment-id "<attachment-id>"
node dist/cli.js outlook draft new --to "user@example.com" --subject "제목" --body "본문"
node dist/cli.js outlook draft reply --message-id "<message-id>" --body "답장 본문"
node dist/cli.js outlook draft reply --message-id "<message-id>" --reply-all --body "전체답장 본문"
node dist/cli.js outlook draft forward --message-id "<message-id>" --to "user@example.com" --body "전달 메모" --attachment "<file-path>"
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 100 --offset 0
node dist/cli.js teams search-messages --query "와플" --since 2026-04-01 --until 2026-07-10
node dist/cli.js teams attachments list --chat-id "<chat-id>" --message-id "<message-id>"
node dist/cli.js teams attachments download --chat-id "<chat-id>" --message-id "<message-id>" --attachment-id "<attachment-id>"
node dist/cli.js teams inline-images list --chat-id "<chat-id>" --message-id "<message-id>"
node dist/cli.js teams inline-images download --chat-id "<chat-id>" --message-id "<message-id>" --hosted-content-id "<hosted-content-id>"
node dist/cli.js sharepoint sites --query "Agent Automation"
node dist/cli.js sharepoint export-files --site-url "https://ohmylab.sharepoint.com/sites/<site>" --destination "<absolute-path>"
node dist/cli.js files search --query "keyword" --limit 10
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

일반 조회 결과는 화면 출력을 바로 사용합니다. 별도 파일이 필요한 경우에만 `--out <path>`를 사용하며, 상대 경로는 Hare 고정 `resultsDir` 아래에 저장되고 7일 후 자동 정리됩니다.

Outlook, Teams, SharePoint, OneDrive, Microsoft 365 조회와 Outlook 초안 작성에는 Hare CLI만 사용합니다. Microsoft 365 커넥터, 다른 커넥터, Computer Use, Outlook/Teams/SharePoint UI, 브라우저 자동화를 검색하거나 대체 수단으로 사용하지 않습니다. Hare가 지원하지 않거나 명령이 실패하면 다른 도구로 우회하지 않고 실패한 Hare 단계와 오류를 보고합니다. 선택 프로젝트 루트의 `CLAUDE.md`에는 이 규칙을 자동으로 읽을 수 있는 Hare 관리 구역이 생성되며, 기존 사용자 작성 내용은 유지됩니다.

일반적인 메일 조회와 최근 메일 조회는 `outlook recent --folder all`을 사용합니다. 삭제된 항목을 제외한 받은편지함, 보낸편지함, 보관함, 사용자 폴더 전체가 기본 대상입니다. `outlook inbox`는 받은편지함이 명시된 요청에만 사용합니다. 플래그된 메일은 `outlook flagged --folder all`로 조회하며 모든 메일 결과에는 `flagStatus`가 포함됩니다.

`--mailbox <name-or-address>`가 없으면 Outlook 명령은 로그인한 사용자의 사서함만 조회합니다. 사용자가 `CTO 공유 사서함`처럼 대상을 명시하면 `recent`, `flagged`, `search`, `count`, 첨부파일 목록·다운로드에 `--mailbox`를 사용합니다. 이름은 기본 디렉터리 정보로 주소를 해석하며, 후보가 여러 개면 정확한 주소를 확인하기 전까지 중단합니다. 공유 사서함 오류가 발생해도 본인 사서함으로 자동 대체하지 않습니다.

메일 첨부파일은 메일 조회 결과의 `id`를 `outlook attachments list --message-id`에 전달해 목록을 확인하고, 반환된 첨부파일 `id`를 `outlook attachments download --attachment-id`에 전달해 내려받습니다. 공유 사서함 메일이면 조회에 사용한 동일한 `--mailbox`를 두 첨부파일 명령에도 전달합니다. 파일은 Hare의 `downloadsDir`에 저장되며 SharePoint/Teams/OneDrive 파일과 동일한 다운로드 정책이 적용됩니다. Excel 등 내려받은 파일의 내용 분석은 다운로드 완료 후 로컬 파일 처리 도구로 수행합니다.

Teams 채팅 첨부파일은 메시지의 `chatId`와 `id`를 `teams attachments list`에 전달해 안전한 메타데이터를 확인하고, 반환된 첨부파일 `id`를 `teams attachments download`에 전달해 내려받습니다. Microsoft Search 색인에서 파일을 찾지 못해도 Teams 메시지에 포함된 회사 SharePoint 직접 경로를 해석하며, 원본 공유 URL은 결과에 출력하지 않습니다. 기존 다운로드 크기 상한과 승인 절차를 그대로 적용합니다.

Teams 본문에 붙여넣은 이미지는 일반 첨부파일과 달리 `hostedContents`로 저장됩니다. Graph 목록 API는 형식을 반환하지 않으므로 `teams inline-images list`에서는 hosted content ID만 확인하고, `teams inline-images download`가 실제 다운로드 응답이 PNG/JPEG/GIF/WebP/BMP/TIFF 래스터 이미지인지 검증한 뒤 Hare의 `downloadsDir`에 저장합니다. Graph 원본 URL과 이미지 바이트는 목록 결과에 노출하지 않으며, 크기를 미리 알 수 없는 인라인 이미지는 기본 다운로드 상한까지만 스트리밍합니다.

기본 다운로드 상한은 `maxDownloadBytes`의 100MiB입니다. 이를 초과하고 `maxApprovedDownloadBytes`의 1GiB 이하인 파일은 다운로드를 시작하지 않고 `AWAITING_USER_APPROVAL` 미리보기를 반환합니다. LLM은 출처, 파일명, 크기와 저장 위치를 모두 보여주고 명시적 동의를 받은 뒤 동일한 명령에 반환된 `--approval-token`을 추가해 한 번만 실행합니다. 토큰은 10분 동안 유효하고 정확히 같은 파일과 출력명에만 사용할 수 있으며 사용 즉시 무효화됩니다. 1GiB를 초과하는 파일은 승인 여부와 관계없이 차단됩니다.

`sharepoint export-files`는 지정한 SharePoint 사이트의 기본 문서 라이브러리를 검색 색인 없이 재귀 열거하고, 파일 폴더 구조를 지정한 절대 경로에 보존합니다. 파일 목록 열거가 시간 예산을 넘으면 `ENUMERATION_IN_PROGRESS`와 체크포인트를 반환하므로 승인 토큰 없이 같은 명령을 다시 실행합니다. 열거가 끝난 첫 실행은 전체 파일 수·총용량·기존 파일·10GiB 초과 제외 파일·디스크 여유를 포함한 `AWAITING_USER_APPROVAL` 계획만 반환하며 다운로드하지 않습니다. 한 번 승인하면 같은 계정·사이트·파일 목록·대상·정책에 묶인 작업을 24시간 동안 파일별 추가 승인 없이 재개합니다. 기존 파일은 덮어쓰지 않고 건너뛰며, 파일당 상한은 `maxSharePointExportFileBytes`의 10GiB입니다. 부분 파일은 HTTP Range로 이어받고, 서버가 Range를 무시하면 해당 부분 파일을 안전하게 처음부터 다시 씁니다. 빈 폴더, 버전 기록, 권한, 메타데이터와 10GiB 초과 파일은 복사하지 않습니다.
파일시스템이 hard-link를 지원하지 않을 때도 기존 파일을 덮어쓰지 않고 배타적 복사로 마무리할 수 있도록, 계획의 `requiredDiskBytes`에는 전체 대기 용량에 가장 큰 대상 파일 1개분의 임시 여유가 보수적으로 추가됩니다.

기간을 지정하지 않은 `outlook search`와 `teams search-messages`는 `Asia/Seoul` 기준 최근 90일을 조회합니다. 두 검색 모두 기본 100건씩 반환합니다. Outlook은 `search.nextCursor`를 `--cursor`로, Teams는 `search.nextOffset`을 `--offset`으로 전달해 이어서 조회합니다. Teams는 명령당 최대 100개의 고유 메시지만 전체 본문으로 조회하고, Microsoft Search의 최대 1,000건 검색 창 안에서만 이어봅니다. `duplicateHitCount`, `noProgressDetected`, `searchWindowExhausted`가 중복 및 중단 사유를 보여줍니다. Outlook 검색 결과의 `body`와 `bodyHtml`은 전체 본문이며 `fullBodyUnavailableCount`로 누락 여부를 확인합니다. `search.partialResult`가 `true`이면 35초 시간 예산 안에 처리한 부분 결과입니다. 결과 JSON의 `search.range.notice`에는 실제 조회 기간이 표시됩니다. 기간이 명확한 요청은 `--since`와 `--until`에 `YYYY-MM-DD` 형식으로 지정합니다.

정확한 메일 건수 집계는 검색 인덱스를 사용하는 `outlook search` 대신 `outlook count`를 사용합니다. `outlook count`는 지정 기간의 모든 메일 페이지를 순회하고 `--subject-contains`와 `--from` 조건을 직접 대조합니다. 35초 안에 끝나지 않으면 누적 상태가 포함된 `nextCursor`를 반환합니다. 다음 실행에서 `--cursor`로 이어가며, `complete:true`인 마지막 `matchedCount`가 정확한 전체 건수입니다. `--folder all`은 삭제된 메일을 제외하며, 보낸편지함은 `sentDateTime`, 나머지는 `receivedDateTime`을 기준으로 집계합니다.

`files search`는 Microsoft Search API를 통해 사용자가 접근할 수 있는 SharePoint, Teams, OneDrive 파일 전체를 검색합니다. 결과의 `nextOffset`을 `--offset`으로 전달해 다음 페이지를 조회할 수 있습니다. SharePoint 사이트 자체의 존재 여부는 `sharepoint sites`로 확인합니다.

Teams `chat-messages`는 `body`에 전체 일반 텍스트, `bodyHtml`에 Graph 원본 HTML을 반환합니다. `search-messages`도 검색 결과마다 채팅 또는 채널 메시지 상세를 추가 조회해 같은 전체 본문 필드를 반환합니다. 일부 상세 조회가 불가능하면 `fullBodyUnavailableCount`와 항목별 `bodyUnavailableReason`으로 명시하며 검색 스니펫을 전체 본문으로 취급하지 않습니다.

한 채팅방의 메시지는 기본 20건, 명령당 최대 1,000개의 고유 메시지까지 조회합니다. 결과의 `page.continuationAvailable`이 `true`이면 `page.nextOffset`을 다음 명령의 `--offset`으로 전달해 이어서 조회합니다. Graph 페이지는 내부적으로 50건씩 처리하며 중복 메시지 ID는 제거합니다.

Teams의 `totalMatchesReported`는 Microsoft Search가 발신자, 본문, 첨부파일을 대상으로 찾은 메시지 수입니다. 따라서 “검색어가 포함된 메시지가 몇 개인가”에는 참고할 수 있지만 “메시지 본문에 단어가 총 몇 번 등장했는가”와는 다릅니다. 정확한 본문 등장 횟수는 첫 페이지에서 전체 후보 수를 확인합니다. 후보가 `searchWindowLimit`인 1,000건을 초과하면 더 조회하지 않고 기간이나 검색어를 좁혀야 합니다. 그 이하일 때만 고유 메시지 본문을 이어서 확인하며, 중복 페이지나 무진행 상태가 감지되면 반복을 중단하고 한계를 보고합니다.

## Outlook 초안

신규·답장·전체답장·전달 초안과 첨부파일을 지원합니다. 초안 명령을 승인 토큰 없이 먼저 실행하면 `AWAITING_USER_APPROVAL` 미리보기가 반환됩니다. LLM은 수신자, 제목, 본문, 첨부파일을 사용자에게 모두 보여주고 명시적 동의를 받은 뒤, 동일한 명령에 반환된 `--approval-token`을 추가해 실행합니다. 내용이나 첨부파일이 바뀌면 승인 토큰이 무효화됩니다.

3MB 미만 파일은 Graph에 직접 첨부하고 3~150MB 파일은 Outlook 업로드 세션을 사용합니다. 조직의 Exchange 메시지 크기 제한이 더 작으면 Microsoft 365 정책에 따라 거부될 수 있습니다. 첨부 중 실패하면 Hare가 불완전한 초안을 자동 삭제합니다.

Hare에는 메일 발송 명령이 없으며 새 Azure Application에도 `Mail.Send` 권한이 없습니다. 생성된 초안은 사용자가 Outlook에서 다시 확인하고 직접 발송합니다.

## 개발 검증

```bash
npm install
npm run typecheck
npm run build
npm test
npm run verify
npm start
```

## 배포 판단

- 기본: 세션 런타임에서 `git clone`/`git pull --ff-only`와 빌드 + 선택 프로젝트를 고정 `--data-dir`로 사용
- 보류: npmjs publish
- 보조/비권장: GitHub Release asset 직접 다운로드
- 폐기: exe/pkg/SharePoint zip 중심 배포
