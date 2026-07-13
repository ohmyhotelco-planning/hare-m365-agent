# Hare M365 Agent LLM 지침

## 목적

Hare는 Microsoft Graph delegated 권한으로 Outlook, Teams, SharePoint/OneDrive를 조회하는 읽기 전용 CLI입니다.

## 준비

기본 경로는 GitHub API나 Release asset이 아닌 `git clone`입니다. 다만 clone 위치는 임시 컨테이너가 아니라 연결된 호스트 `HareM365Agent/app`입니다. 최초에는 clone하고 이후에는 기존 app에서 `git fetch`와 `git pull --ff-only`를 실행합니다. HEAD가 바뀌었거나 빌드가 없을 때만 `npm ci`와 `npm run build`를 실행합니다.

필수 도메인은 `github.com`, `registry.npmjs.org`, `login.microsoftonline.com`, `graph.microsoft.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com`입니다.

Cowork `설정 > 기능 > 도메인 허용 목록`이 유일한 네트워크 허용 목록입니다. 연결 폴더용 별도 허용 목록은 없습니다. 설정을 현재 작업을 연 뒤 변경했다면 새 Cowork 채팅에서 진행합니다.

`npm ci`가 실패하면 백그라운드 실행, `npm install` 전환, 셸 호출별 증분 설치로 우회하지 않고 실패 단계와 오류만 보고합니다.

## 로그인 하드게이트

- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365 조회를 실행합니다.
- 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
- 사용자는 Cowork 작업을 만들 때 `HareM365Agent` 프로젝트 또는 폴더를 먼저 선택합니다. 이미 연결된 프로젝트의 마운트 루트를 모든 Hare 명령의 동일한 `--data-dir`로 사용합니다.
- 프로젝트가 연결되지 않았으면 `FOLDER_REQUIRED`로 중단하고, 해당 프로젝트를 선택해 새 Cowork 작업을 열도록 안내합니다.
- 프로그램은 해당 루트의 `app`, 인증 캐시는 `.cache`, 세션 규칙은 `claude`에 둡니다. 새 작업도 같은 프로젝트에서 시작하고 `claude/hare-m365-agent-rules.md`를 읽습니다.
- Cowork Linux의 OS 기본 경로(`/root/.local/share/...`, `~/.local/share/...`)는 컨테이너 내부이므로 로그인이나 규칙 파일 저장에 사용하지 않습니다.
- 코드, node_modules, dist도 session-local `/sessions`, `/tmp`, `/root` 아래에 설치하거나 실행하지 않습니다. Cowork 프로젝트의 `/sessions/<session>/mnt/HareM365Agent`는 로컬 프로젝트 폴더의 정상 마운트 경로이므로 사용합니다.
- `HTTP 403`과 `X-Proxy-Error: blocked-by-allowlist`가 함께 나오면 `NETWORK_PERMISSION_REQUIRED`로 보고합니다. 설정을 변경했다면 같은 프로젝트에서 새 Cowork 작업을 열고 실패한 명령만 한 번 재시도합니다.
- 마운트 폴더와 네트워크를 서로 다른 로컬/클라우드 실행 환경으로 나누어 설명하지 않습니다. 둘 다 같은 Cowork Linux 셸에서 사용합니다.
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
