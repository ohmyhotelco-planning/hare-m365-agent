# Hare M365 Agent for Claude/Cowork

이 저장소에서는 [AGENTS.md](AGENTS.md)를 Hare 실행의 단일 기준으로 사용합니다.

핵심 규칙:

1. GitHub API가 아닌 `git clone` 경로를 사용하고, 연결된 호스트 `HareM365Agent/app`에 설치합니다.
2. 최초 설치 전 `registry.npmjs.org`를 포함한 필수 도메인 6개를 허용합니다.
3. `dist/cli.js`, `dist/proxy.js`, `dist/msal-network.js`를 모두 확인합니다.
4. `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365를 조회합니다.
5. session-local `/sessions`, `/tmp`, `/root`에는 코드나 로그인 상태를 두지 않습니다. 단, `/sessions/<session>/mnt/HareM365Agent`는 선택된 프로젝트의 마운트이므로 사용합니다. 프로그램은 프로젝트의 `app`, 로그인 상태는 `.cache`에 유지합니다.
6. 프로젝트가 연결되지 않았으면 `FOLDER_REQUIRED`로 중단하고 같은 프로젝트를 선택한 새 Cowork 작업에서 다시 시작합니다.
7. 새 작업은 같은 `HareM365Agent` 프로젝트에서 시작하고 기존 app을 `git pull --ff-only`로 갱신한 뒤 `claude/hare-m365-agent-rules.md`를 읽습니다.
8. `auth login-start`로 코드를 즉시 받은 뒤, 사용자 로그인 완료 후 별도 호출에서 `auth login-complete`를 실행합니다.
9. 장기 poller, 백그라운드, detached, `setsid`, `nohup` 로그인을 사용하지 않습니다.
10. 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
11. 기본 동작은 읽기 전용입니다.
12. Cowork `설정 > 기능 > 도메인 허용 목록`이 유일한 네트워크 허용 목록입니다. 설정 변경 후에는 같은 프로젝트에서 새 Cowork 작업을 엽니다.
13. `HTTP 403`과 `X-Proxy-Error: blocked-by-allowlist`는 `NETWORK_PERMISSION_REQUIRED`로 처리하고, 로컬/클라우드 환경 분리로 오판하지 않습니다.
