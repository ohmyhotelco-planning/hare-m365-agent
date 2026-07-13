# Hare M365 Agent LLM 지침

## 목적

Hare는 Microsoft Graph delegated 권한으로 Outlook, Teams, SharePoint/OneDrive를 조회하는 읽기 전용 CLI입니다.

## 준비

기본 경로는 GitHub API나 Release asset이 아닌 `git clone`입니다.

```bash
git ls-remote https://github.com/ohmyhotelco-planning/hare-m365-agent.git HEAD
rm -rf /tmp/hare-m365-agent
git clone https://github.com/ohmyhotelco-planning/hare-m365-agent.git /tmp/hare-m365-agent
cd /tmp/hare-m365-agent && npm ci && npm run build
test -f /tmp/hare-m365-agent/dist/cli.js
test -f /tmp/hare-m365-agent/dist/proxy.js
test -f /tmp/hare-m365-agent/dist/msal-network.js
cd /tmp/hare-m365-agent && node dist/cli.js
```

필수 도메인은 `github.com`, `login.microsoftonline.com`, `graph.microsoft.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com`입니다. `npm ci`가 registry 연결 오류로 실패할 때만 `registry.npmjs.org`를 추가로 요청합니다.

## 로그인 하드게이트

- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365 조회를 실행합니다.
- 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
- 초기 연결에서는 출력된 `humanLoginCommand`를 동일한 셸 호출에서 **포그라운드로 직접 실행**합니다.
- 백그라운드, detached, `setsid`, `nohup` 또는 별도 프로세스로 로그인하지 않습니다.
- Microsoft 브라우저 로그인이 끝나도 로그인 명령을 종료하지 않습니다. 명령이 `ok: true`를 반환하고 캐시 저장이 끝날 때까지 같은 호출을 유지합니다.
- 이후 같은 `dataDir`에서 `auth status`를 실행해 `loggedIn: true`, `tokenUsable: true`를 확인합니다.
- 사용자는 Microsoft 로그인 주소와 device code를 보고 회사 계정으로 로그인하는 과정만 수행합니다.

## 조회 기준

- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일입니다. 결과의 `search.range.notice`를 답변에 포함합니다.
- 정확한 메일 건수는 `outlook count`를 사용합니다. 기본 `all` 범위는 삭제된 메일을 제외합니다.
- 최신 Teams 채팅은 실제 마지막 메시지 생성 시각으로 판단합니다.
- SharePoint 사이트 존재 여부는 `sharepoint sites`로 확인합니다.
- 일반 조회는 화면 출력을 사용하고, 파일 결과가 필요한 경우에만 `--out`을 사용합니다.

메일 발송, Teams 게시, 일정 생성, 파일 업로드·삭제·공유, 권한 변경은 서비스하지 않습니다.
