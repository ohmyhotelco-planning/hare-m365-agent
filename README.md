# Hare M365 Agent

Hare M365 Agent는 LLM이 Microsoft Graph delegated 권한으로 Outlook, Teams, OneDrive, SharePoint를 조회할 수 있도록 만든 Node/TypeScript CLI입니다.

현재 기본 실행 방식은 Claude/Cowork 샌드박스에서 GitHub 저장소를 `git clone`한 뒤 빌드해 실행하는 방식입니다. GitHub API나 GitHub Release asset 다운로드는 Cowork 프록시 정책에 막힐 수 있으므로 기본 경로로 사용하지 않습니다.

사람이 읽는 안내는 아래 한 장만 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드_fixed.html
```

## 실행 개요

저장소:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent.git
```

초기 준비:

```bash
REMOTE_HEAD=$(git ls-remote https://github.com/ohmyhotelco-planning/hare-m365-agent.git refs/heads/master | awk '{print $1}')
test -n "$REMOTE_HEAD"
rm -rf /tmp/hare-m365-agent
test ! -e /tmp/hare-m365-agent
git clone --branch master --single-branch --no-tags https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
LOCAL_HEAD=$(git -C /tmp/hare-m365-agent rev-parse HEAD)
test "$LOCAL_HEAD" = "$REMOTE_HEAD"
test ! -d /tmp/hare-m365-agent/node_modules
cd /tmp/hare-m365-agent && npm ci --prefer-offline --no-audit --no-fund
cd /tmp/hare-m365-agent && npm run build
test -f /tmp/hare-m365-agent/dist/cli.js
test -f /tmp/hare-m365-agent/dist/proxy.js
test -f /tmp/hare-m365-agent/dist/msal-network.js
cd /tmp/hare-m365-agent && node dist/cli.js
```

Cowork에서는 기존 clone을 업데이트해 재사용하지 않습니다. 매번 임시 작업 폴더를 지우고 `master`를 새로 clone한 뒤 `LOCAL_HEAD`와 `REMOTE_HEAD`가 일치할 때만 빌드합니다. 셸 호출마다 작업 폴더가 초기화될 수 있으므로 각 명령에 `/tmp/hare-m365-agent`를 명시합니다.

최초 설치에는 `registry.npmjs.org`가 필수입니다. `npm ci`가 실패하면 백그라운드 실행이나 `npm install` 증분 설치로 우회하지 않고 원인을 보고합니다.

## 설정

기본 Azure Application 설정은 `hare.config.json`에 포함됩니다. 일반 사용자는 `.env`를 만들거나 수정하지 않습니다.

`.env`는 개발자용 로컬 override가 필요할 때만 사용합니다.

## 로그인과 저장 위치

Hare는 OS별 고정 `dataDir`에 로그인 캐시와 다운로드를 저장합니다.

```text
Windows: %LOCALAPPDATA%\Ohmyhotel\HareM365Agent
Mac: ~/Library/Application Support/Ohmyhotel/HareM365Agent
Linux: ~/.local/share/ohmyhotel/hare-m365-agent
```

`loggedIn`과 `tokenUsable`이 모두 `true`일 때만 조회 가능한 상태입니다. 캐시 파일이 존재하더라도 토큰을 획득할 수 없으면 로그인 완료로 판단하지 않습니다.

Cowork에서 `dataDirPersistent: false`가 나오면 로그인 전에 사용자의 Documents 폴더를 연결하고, 마운트된 폴더 아래 `Hare M365 Agent`를 `HARE_M365_DATA_DIR`로 지정합니다. `/sessions` 또는 `/tmp` 기본 저장소에서는 로그인이 거부됩니다.

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

- 기본: Cowork `git clone` + `npm ci --prefer-offline --no-audit --no-fund` + `npm run build` + `node dist/cli.js`
- 보류: npmjs publish
- 보조/비권장: GitHub Release asset 직접 다운로드
- 폐기: exe/pkg/SharePoint zip 중심 배포
