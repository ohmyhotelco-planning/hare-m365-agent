# Hare M365 Agent

Hare M365 Agent는 LLM이 Microsoft Graph delegated 권한으로 Outlook, Teams, OneDrive, SharePoint를 조회할 수 있도록 만든 Node/TypeScript CLI입니다.

현재 기본 실행 방식은 Claude/Cowork 샌드박스에서 GitHub 저장소를 `git clone`한 뒤 빌드해 실행하는 방식입니다. GitHub API나 GitHub Release asset 다운로드는 Cowork 프록시 정책에 막힐 수 있으므로 기본 경로로 사용하지 않습니다.

사람이 읽는 안내는 아래 한 장만 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드.html
```

## 실행 개요

저장소:

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent.git
```

초기 준비:

```bash
git clone https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
cd /tmp/hare-m365-agent
npm ci
npm run build
node dist/cli.js
```

이미 clone되어 있으면:

```bash
cd /tmp/hare-m365-agent
git pull
npm ci
npm run build
node dist/cli.js
```

`npm ci` 단계에서 npm registry 접근 오류가 발생할 때만 `registry.npmjs.org` 도메인 허용이 필요합니다.

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

`loggedIn: false`는 출력 JSON의 `dataDir`/`cacheFile` 기준 상태입니다. hosted sandbox나 컨테이너 경로에서 나온 `false`를 사용자 PC 로그인 실패로 해석하지 않습니다. Cowork처럼 사용자 폴더를 마운트할 수 있는 환경에서는 사용자 PC의 고정 Hare 폴더를 연결하고, 이후 모든 Hare 명령에 `HARE_M365_DATA_DIR="<마운트된 Hare 폴더 경로>"`를 붙여 같은 캐시를 사용합니다.

## 주요 명령

```bash
node dist/cli.js
node dist/cli.js auth status
node dist/cli.js outlook inbox --limit 10 --out latest-mail.json
node dist/cli.js outlook search --query "나이스페이 OR nicepay" --since 2026-06-26 --until 2026-07-10 --folder all --out mail-search.json
node dist/cli.js outlook count --subject-contains "[RPA]" --since 2024-07-10 --until 2026-07-10 --folder all --out mail-count.json
node dist/cli.js teams teams --out teams.json
node dist/cli.js teams chats --limit 20 --out chats.json
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20 --out chat-messages.json
node dist/cli.js teams search-messages --query "와플" --since 2026-04-01 --until 2026-07-10 --out teams-search.json
node dist/cli.js sharepoint sites --query "Agent Automation" --out sharepoint-sites.json
node dist/cli.js files search --query "keyword" --limit 10 --out files.json
node dist/cli.js files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

조회 명령은 `--out <path>`를 지원합니다. 상대 경로를 주면 Hare 고정 `dataDir` 아래에 JSON을 저장합니다.

기간을 지정하지 않은 `outlook search`와 `teams search-messages`는 최근 90일을 조회하며 최대 1,000건을 반환합니다. 결과 JSON의 `search.range.notice`에 실제 조회 기간이 표시되고, `search.limitReached`로 결과 한도 도달 여부를 확인할 수 있습니다. 기간이 명확한 요청은 `--since`와 `--until`에 `YYYY-MM-DD` 형식으로 지정합니다.

정확한 메일 건수 집계는 검색 인덱스를 사용하는 `outlook search` 대신 `outlook count`를 사용합니다. `outlook count`는 지정 기간의 모든 메일 페이지를 순회하고 `--subject-contains`와 `--from` 조건을 직접 대조합니다.

SharePoint 사이트의 존재 여부는 `sharepoint sites`로 확인합니다. `files search`는 현재 개인 OneDrive 범위이므로, 해당 결과만으로 SharePoint 사이트 존재 여부나 접근 권한을 판단하지 않습니다.

## 개발 검증

```bash
npm install
npm run typecheck
npm run build
npm start
```

## 배포 판단

- 기본: Cowork `git clone` + `npm ci` + `npm run build` + `node dist/cli.js`
- 보류: npmjs publish
- 보조/비권장: GitHub Release asset 직접 다운로드
- 폐기: exe/pkg/SharePoint zip 중심 배포
