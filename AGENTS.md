# Hare M365 Agent LLM 지침

## 목적

Hare는 Microsoft Graph delegated 권한으로 Outlook, Teams, SharePoint/OneDrive를 조회하는 읽기 전용 CLI입니다.

## 준비

기본 경로는 GitHub API나 Release asset이 아닌 `git clone`입니다.

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

필수 도메인은 `github.com`, `registry.npmjs.org`, `login.microsoftonline.com`, `graph.microsoft.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com`입니다.

`npm ci`가 실패하면 백그라운드 실행, `npm install` 전환, 셸 호출별 증분 설치로 우회하지 않고 실패 단계와 오류만 보고합니다.

## 로그인 하드게이트

- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365 조회를 실행합니다.
- 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
- `dataDirPersistent: false`이면 Windows `%USERPROFILE%\HareM365Agent` 또는 Mac `~/HareM365Agent` 고정 폴더 접근을 요청하고, 연결된 마운트 루트를 모든 Hare 명령의 동일한 `--data-dir`로 사용합니다.
- Cowork Linux의 OS 기본 경로(`/root/.local/share/...`, `~/.local/share/...`)는 컨테이너 내부이므로 로그인이나 규칙 파일 저장에 사용하지 않습니다.
- `auth login-start`는 Microsoft 로그인 주소와 코드를 즉시 반환하고 종료합니다.
- 사용자가 브라우저 로그인을 완료하면 별도 셸 호출에서 `auth login-complete`를 실행합니다. 이 명령은 최대 25초만 실행됩니다.
- 장기 poller, 백그라운드, detached, `setsid`, `nohup`을 사용하지 않습니다.
- `LOGIN_PENDING`이면 기존 코드로 `auth login-complete`만 다시 실행하고, 코드가 만료된 경우에만 `auth login-start`를 다시 실행합니다.
- 이후 같은 `dataDir`에서 `auth status`를 실행해 `loggedIn: true`, `tokenUsable: true`를 확인합니다.
- 사용자는 Microsoft 로그인 주소와 user code를 보고 회사 계정으로 로그인하는 과정만 수행합니다.

## 조회 기준

- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일입니다. 결과의 `search.range.notice`를 답변에 포함합니다.
- 정확한 메일 건수는 `outlook count`를 사용합니다. 기본 `all` 범위는 삭제된 메일을 제외합니다.
- 최신 Teams 채팅은 실제 마지막 메시지 생성 시각으로 판단합니다.
- SharePoint 사이트 존재 여부는 `sharepoint sites`로 확인합니다.
- 일반 조회는 화면 출력을 사용하고, 파일 결과가 필요한 경우에만 `--out`을 사용합니다.

메일 발송, Teams 게시, 일정 생성, 파일 업로드·삭제·공유, 권한 변경은 서비스하지 않습니다.
