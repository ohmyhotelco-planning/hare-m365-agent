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

Azure Application 또는 요청 권한이 변경되면 Hare는 기존 앱의 인증 캐시와 진행 중인 로그인 상태만 자동 초기화합니다. 다운로드, 조회 결과, 로그와 Claude 운영 규칙은 유지됩니다. startup의 `authReason`이 `AUTH_APP_CHANGED`이면 새 앱으로 Microsoft 로그인을 한 번 완료한 뒤 기존과 같이 사용합니다.

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
# 사용자가 브라우저에서 회사 Microsoft 계정 로그인
node dist/cli.js auth login-complete
node dist/cli.js auth status
```

`login-start`는 주소와 user code를 즉시 반환하고 종료합니다. `login-complete`는 최대 25초만 토큰을 확인하므로 장기 poller나 백그라운드 프로세스가 필요하지 않습니다.

## 주요 명령

```bash
node dist/cli.js
node dist/cli.js auth status
node dist/cli.js outlook recent --folder all --limit 10
node dist/cli.js outlook inbox --limit 10
node dist/cli.js outlook flagged --folder all --limit 1000
node dist/cli.js outlook search --query "나이스페이 OR nicepay" --since 2026-06-26 --until 2026-07-10 --folder all
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all
node dist/cli.js outlook draft new --to "user@example.com" --subject "제목" --body "본문"
node dist/cli.js outlook draft reply --message-id "<message-id>" --body "답장 본문"
node dist/cli.js outlook draft reply --message-id "<message-id>" --reply-all --body "전체답장 본문"
node dist/cli.js outlook draft forward --message-id "<message-id>" --to "user@example.com" --body "전달 메모" --attachment "<file-path>"
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20
node dist/cli.js teams search-messages --query "와플" --since 2026-04-01 --until 2026-07-10
node dist/cli.js sharepoint sites --query "Agent Automation"
node dist/cli.js files search --query "keyword" --limit 10
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

일반 조회 결과는 화면 출력을 바로 사용합니다. 별도 파일이 필요한 경우에만 `--out <path>`를 사용하며, 상대 경로는 Hare 고정 `resultsDir` 아래에 저장되고 7일 후 자동 정리됩니다.

일반적인 메일 조회와 최근 메일 조회는 `outlook recent --folder all`을 사용합니다. 삭제된 항목을 제외한 받은편지함, 보낸편지함, 보관함, 사용자 폴더 전체가 기본 대상입니다. `outlook inbox`는 받은편지함이 명시된 요청에만 사용합니다. 플래그된 메일은 `outlook flagged --folder all`로 조회하며 모든 메일 결과에는 `flagStatus`가 포함됩니다.

기간을 지정하지 않은 `outlook search`와 `teams search-messages`는 `Asia/Seoul` 기준 최근 90일을 조회하며 최대 1,000건을 반환합니다. 결과 JSON의 `search.range.notice`에 실제 조회 기간이 표시되고, `search.limitReached`로 결과 한도 도달 여부를 확인할 수 있습니다. 기간이 명확한 요청은 `--since`와 `--until`에 `YYYY-MM-DD` 형식으로 지정합니다.

정확한 메일 건수 집계는 검색 인덱스를 사용하는 `outlook search` 대신 `outlook count`를 사용합니다. `outlook count`는 지정 기간의 모든 메일 페이지를 순회하고 `--subject-contains`와 `--from` 조건을 직접 대조합니다. `--folder all`은 삭제된 메일을 제외하며, 보낸편지함은 `sentDateTime`, 나머지는 `receivedDateTime`을 기준으로 집계합니다.

SharePoint 사이트의 존재 여부는 `sharepoint sites`로 확인합니다. `files search`는 현재 개인 OneDrive 범위이므로, 해당 결과만으로 SharePoint 사이트 존재 여부나 접근 권한을 판단하지 않습니다.

Teams `chat-messages`는 `body`에 전체 일반 텍스트, `bodyHtml`에 Graph 원본 HTML을 반환합니다. `search-messages`도 검색 결과마다 채팅 또는 채널 메시지 상세를 추가 조회해 같은 전체 본문 필드를 반환합니다. 일부 상세 조회가 불가능하면 `fullBodyUnavailableCount`와 항목별 `bodyUnavailableReason`으로 명시하며 검색 스니펫을 전체 본문으로 취급하지 않습니다.

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
