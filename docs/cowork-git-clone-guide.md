# Hare M365 Agent Cowork Git Clone 가이드

## 현재 기준

Hare M365 Agent는 Claude/Cowork 샌드박스에서 GitHub 저장소를 `git clone`하고, `npm ci`와 TypeScript build 후 `node dist/cli.js`로 실행합니다.

```text
LLM -> Cowork shell -> git clone -> npm ci -> npm run build -> node dist/cli.js -> Microsoft Graph delegated access
```

Cowork에서는 GitHub API(`api.github.com`)나 GitHub Release asset 다운로드가 프록시 정책에 막힐 수 있습니다. repo 접근 확인은 실제 사용 경로와 같은 `git ls-remote` 또는 `git clone`으로 판단합니다.

## 저장소

```text
https://github.com/ohmyhotelco-planning/hare-m365-agent.git
```

## 초기 준비 명령

```bash
git ls-remote https://github.com/ohmyhotelco-planning/hare-m365-agent.git HEAD
rm -rf /tmp/hare-m365-agent
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

## Claude/Cowork 도메인 허용

처음 필요한 도메인은 아래 3개입니다.

```text
github.com
graph.microsoft.com
login.microsoftonline.com
ohmylab-my.sharepoint.com
ohmylab.sharepoint.com
```

SharePoint/OneDrive 파일 본문 다운로드는 Graph 조회 후 SharePoint 테넌트 도메인으로 연결될 수 있습니다. 와일드카드를 지원하는 환경이면 `*.sharepoint.com`으로 대체할 수 있습니다.

`npm ci`가 npm registry 접근 오류로 실패할 때만 아래 도메인이 추가로 필요합니다.

```text
registry.npmjs.org
```

## 로그인과 캐시

Hare는 OS별 고정 `dataDir`를 사용합니다. `loggedIn: false`는 출력 JSON의 `dataDir`/`cacheFile` 기준 상태입니다. hosted sandbox나 컨테이너 경로에서 나온 `false`를 사용자 PC 로그인 실패로 해석하지 않습니다.

Cowork에서 사용자 PC의 고정 Hare 폴더가 마운트되면 이후 모든 Hare 명령에 아래 환경변수를 붙여 같은 캐시를 사용합니다.

```bash
HARE_M365_DATA_DIR="<mounted Hare folder path>" node dist/cli.js auth status
```

## 조회 명령

```bash
node dist/cli.js outlook inbox --limit 10 --out latest-mail.json
node dist/cli.js teams chats --limit 20 --out chats.json
node dist/cli.js teams chat-messages --chat-id "<chat-id>" --limit 20 --out chat-messages.json
node dist/cli.js files search --query "keyword" --limit 10 --out files.json
```

## 검증

```bash
npm run typecheck
npm run build
node dist/cli.js --version
node dist/cli.js
```
