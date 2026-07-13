# Hare M365 Agent

Hare M365 Agent는 LLM이 Microsoft Graph delegated 권한으로 Outlook, Teams, OneDrive, SharePoint를 조회할 수 있도록 만든 Node/TypeScript CLI입니다.

현재 기본 실행 방식은 Windows `%USERPROFILE%\HareM365Agent` 또는 Mac `~/HareM365Agent` 폴더를 Cowork에 연결하고, 프로그램과 로그인 상태를 모두 그 로컬 폴더에서 유지하는 방식입니다. GitHub API나 GitHub Release asset 다운로드는 Cowork 프록시 정책에 막힐 수 있으므로 사용하지 않습니다.

사람이 읽는 안내는 아래 한 장만 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드_fixed_final_real.html
```

## 실행 개요

저장소:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent.git
```

로컬 폴더 구조:

```text
HareM365Agent/
├─ app/          Git 저장소, node_modules, dist
├─ .cache/       Microsoft 로그인 캐시
├─ claude/       새 채팅에서 읽을 Hare 운영 규칙
├─ downloads/
├─ results/
└─ logs/
```

최초 연결에서는 로컬 폴더의 `app/`에 저장소를 clone합니다. 이후 채팅에서는 같은 폴더를 다시 연결하고 `git fetch`와 `git pull --ff-only`로 최신 `master`를 확인합니다. HEAD가 바뀌었거나 빌드 산출물이 없을 때만 `npm ci`와 빌드를 실행합니다. `startup.setupCommand`가 이 절차와 정확한 `--data-dir`를 함께 반환합니다.

최초 설치에는 `registry.npmjs.org`가 필수입니다. `npm ci`가 실패하면 백그라운드 실행이나 `npm install` 증분 설치로 우회하지 않고 원인을 보고합니다.

## 설정

기본 Azure Application 설정은 `hare.config.json`에 포함됩니다. 일반 사용자는 `.env`를 만들거나 수정하지 않습니다.

`.env`는 개발자용 로컬 override가 필요할 때만 사용합니다.

## 로그인과 저장 위치

일반 로컬 실행은 OS별 기본 `dataDir`를 사용합니다.

```text
Windows: %LOCALAPPDATA%\Ohmyhotel\HareM365Agent
Mac: ~/Library/Application Support/Ohmyhotel/HareM365Agent
Linux local default: ~/.local/share/ohmyhotel/hare-m365-agent
```

Cowork에서는 위 OS 기본값과 별개로 Windows `%USERPROFILE%\HareM365Agent` 또는 Mac `~/HareM365Agent`를 호스트 고정 폴더로 사용합니다. 연결 도구가 반환한 마운트 루트를 `--data-dir`로 명시합니다. 프로그램도 이 루트의 `app/`에서 실행하므로 코드와 로그인 상태가 임시 컨테이너에 남지 않습니다.

`loggedIn`과 `tokenUsable`이 모두 `true`일 때만 조회 가능한 상태입니다. 캐시 파일이 존재하더라도 토큰을 획득할 수 없으면 로그인 완료로 판단하지 않습니다. `auth login-complete`도 저장된 캐시를 다시 열어 검증한 뒤에만 `COMPLETE`를 반환합니다.

Cowork에서는 작업 전에 고정 호스트 폴더를 연결합니다. startup JSON의 `setup.state`와 `setup.nextAction` 하나만 따르며, CLI가 반환한 `setup.nextCommand`를 수정하지 않고 실행합니다. 이 명령에는 동일한 `--data-dir`가 포함되므로 셸과 채팅이 바뀌어도 같은 캐시를 사용합니다. `/sessions`, `/tmp`, Cowork Linux OS 기본 저장소에서는 로그인이 거부됩니다.

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
node dist/cli.js outlook inbox --limit 10
node dist/cli.js outlook search --query "나이스페이 OR nicepay" --since 2026-06-26 --until 2026-07-10 --folder all
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all
node dist/cli.js teams teams
node dist/cli.js teams chats --limit 20
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20
node dist/cli.js teams search-messages --query "와플" --since 2026-04-01 --until 2026-07-10
node dist/cli.js sharepoint sites --query "Agent Automation"
node dist/cli.js files search --query "keyword" --limit 10
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

일반 조회 결과는 화면 출력을 바로 사용합니다. 별도 파일이 필요한 경우에만 `--out <path>`를 사용하며, 상대 경로는 Hare 고정 `resultsDir` 아래에 저장되고 7일 후 자동 정리됩니다.

기간을 지정하지 않은 `outlook search`와 `teams search-messages`는 `Asia/Seoul` 기준 최근 90일을 조회하며 최대 1,000건을 반환합니다. 결과 JSON의 `search.range.notice`에 실제 조회 기간이 표시되고, `search.limitReached`로 결과 한도 도달 여부를 확인할 수 있습니다. 기간이 명확한 요청은 `--since`와 `--until`에 `YYYY-MM-DD` 형식으로 지정합니다.

정확한 메일 건수 집계는 검색 인덱스를 사용하는 `outlook search` 대신 `outlook count`를 사용합니다. `outlook count`는 지정 기간의 모든 메일 페이지를 순회하고 `--subject-contains`와 `--from` 조건을 직접 대조합니다. `--folder all`은 삭제된 메일을 제외하며, 보낸편지함은 `sentDateTime`, 나머지는 `receivedDateTime`을 기준으로 집계합니다.

SharePoint 사이트의 존재 여부는 `sharepoint sites`로 확인합니다. `files search`는 현재 개인 OneDrive 범위이므로, 해당 결과만으로 SharePoint 사이트 존재 여부나 접근 권한을 판단하지 않습니다.

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

- 기본: 연결된 로컬 `HareM365Agent/app`에서 `git clone` 또는 `git pull --ff-only` + 필요할 때만 빌드 + 고정 `--data-dir` 실행
- 보류: npmjs publish
- 보조/비권장: GitHub Release asset 직접 다운로드
- 폐기: exe/pkg/SharePoint zip 중심 배포
