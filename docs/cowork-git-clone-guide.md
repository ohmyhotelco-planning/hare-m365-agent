# Hare M365 Agent Cowork 실행 기준

사람이 따라 하는 배포 안내는 아래 HTML 한 장을 사용합니다.

```text
release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드_fixed_final.html
```

LLM 실행 규칙은 루트 `AGENTS.md`와 CLI의 `llm-guide` 출력이 기준입니다.

## 검증 순서

```bash
REMOTE_HEAD=$(git ls-remote https://github.com/ohmyhotelco-planning/hare-m365-agent.git refs/heads/master | awk '{print $1}')
test -n "$REMOTE_HEAD"
rm -rf /tmp/hare-m365-agent
git clone --branch master --single-branch --no-tags https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
LOCAL_HEAD=$(git -C /tmp/hare-m365-agent rev-parse HEAD)
test "$LOCAL_HEAD" = "$REMOTE_HEAD"
cd /tmp/hare-m365-agent && npm ci --prefer-offline --no-audit --no-fund && npm run build
test -f /tmp/hare-m365-agent/dist/cli.js
test -f /tmp/hare-m365-agent/dist/proxy.js
test -f /tmp/hare-m365-agent/dist/msal-network.js
cd /tmp/hare-m365-agent && node dist/cli.js
```

Cowork에서는 작업을 만들 때 `HareM365Agent` 프로젝트 폴더를 선택합니다. 네트워크가 되는 실행 셸에서 이 폴더의 마운트가 직접 보이면 그 마운트 루트를 `--data-dir`로 사용하고, 보이지 않으면(표준 Cowork) 샌드박스 고정 경로를 `--data-dir`로 사용하면서 스냅샷과 로그인 캐시를 연결 폴더와 동기화합니다. 연결 폴더가 마운트된 디바이스 셸에서는 `git`, `npm`, 로그인, Graph 호출을 실행하지 않습니다(모든 도메인이 차단될 수 있음).

`auth login-start`로 로그인 주소와 코드를 즉시 받은 다음 사용자가 브라우저 로그인을 완료하면, 별도 호출에서 `auth login-complete`를 실행합니다. 장기 실행 poller나 백그라운드 프로세스는 사용하지 않습니다.

`loggedIn: true`와 `tokenUsable: true`가 모두 확인되기 전에는 Outlook, Teams, SharePoint/OneDrive 조회를 실행하지 않습니다.
