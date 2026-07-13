# Hare M365 Agent Cowork 실행 기준

사람이 따라 하는 배포 안내는 아래 HTML 한 장을 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드_fixed.html
```

LLM 실행 규칙은 루트 `AGENTS.md`와 CLI의 `llm-guide` 출력이 기준입니다.

## 검증 순서

```bash
git ls-remote https://github.com/ohmyhotelco-planning/hare-m365-agent.git HEAD
rm -rf /tmp/hare-m365-agent
git clone https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
cd /tmp/hare-m365-agent && npm ci --prefer-offline --no-audit --no-fund && npm run build
test -f /tmp/hare-m365-agent/dist/cli.js
test -f /tmp/hare-m365-agent/dist/proxy.js
test -f /tmp/hare-m365-agent/dist/msal-network.js
cd /tmp/hare-m365-agent && node dist/cli.js
```

로그인은 같은 셸 호출에서 `node dist/cli.js auth login`을 포그라운드로 실행합니다. 브라우저 인증이 끝나도 명령이 `ok: true`를 반환할 때까지 호출을 유지한 다음, 같은 `dataDir`에서 `auth status`를 확인합니다.

`loggedIn: true`와 `tokenUsable: true`가 모두 확인되기 전에는 Outlook, Teams, SharePoint/OneDrive 조회를 실행하지 않습니다.
