# Hare M365 Agent for Claude/Cowork

이 저장소에서는 [AGENTS.md](AGENTS.md)를 Hare 실행의 단일 기준으로 사용합니다.

핵심 규칙:

1. GitHub API가 아닌 `git clone` 경로를 사용합니다. 최초 설치 전 `registry.npmjs.org`를 포함한 필수 도메인 6개를 허용합니다.
2. Cowork에는 실행 셸이 두 곳 있을 수 있습니다. 도메인 허용 목록은 세션 샌드박스 셸에 적용되고, 연결 폴더가 마운트된 디바이스 셸은 설정과 무관하게 모든 도메인이 차단될 수 있습니다. `git`, `npm`, 로그인, Graph 호출은 반드시 샌드박스 셸에서 실행합니다.
3. 영속 저장소는 폴더 이름으로 검색하지 않고 현재 Cowork 작업에 선택된 프로젝트 루트로 확정합니다. 네트워크가 되는 셸에서 현재 프로젝트 마운트(`/sessions/<session>/mnt/<selected-project>`)가 직접 보이면 그 마운트를 HARE_ROOT로 사용합니다. 보이지 않으면 하이브리드 모델로, 샌드박스 고정 경로(예: `/home/claude/hare`)를 HARE_ROOT로 쓰고 현재 선택된 프로젝트 루트를 영속 저장소로 동기화합니다. 비슷한 이름의 다른 로컬 폴더로 전환하지 않습니다.
4. 하이브리드 모델에서는 세션 시작 시 연결 폴더의 `.hare-app-snapshot.tar.gz`, `.hare-app-build-head`, `.cache/msal-cache.json`, `claude/` 문서를 HARE_ROOT로 가져오고, 재빌드나 로그인 캐시 갱신 후 같은 파일들을 연결 폴더로 커밋합니다.
5. 스냅샷이 있으면 setup 명령이 자동으로 풀어 clone을 건너뛰고, HEAD가 원격과 같으면 빌드도 건너뜁니다. `dist/cli.js`, `dist/proxy.js`, `dist/msal-network.js`를 모두 확인합니다.
6. `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365를 조회합니다. 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
7. 일반 `/sessions`, `/tmp`, `/root` 경로와 Cowork Linux OS 기본 경로에는 로그인 상태를 남기지 않습니다. 프로그램은 HARE_ROOT의 `app`, 로그인 상태는 `.cache`에 유지합니다.
8. 프로젝트가 연결되지 않았으면 `FOLDER_REQUIRED`로 중단하고 같은 프로젝트를 선택한 새 Cowork 작업에서 다시 시작합니다.
9. 새 작업은 같은 `HareM365Agent` 프로젝트에서 시작하고 `claude/hare-m365-agent-rules.md`를 읽습니다.
10. `auth login-start`로 코드를 즉시 받은 뒤, 사용자 로그인 완료 후 별도 호출에서 `auth login-complete`를 실행합니다. 장기 poller, 백그라운드, detached, `setsid`, `nohup` 로그인을 사용하지 않습니다.
11. 기본 동작은 읽기 전용입니다.
12. `HTTP 403`과 `X-Proxy-Error: blocked-by-allowlist`가 나오면 먼저 실행 셸을 확인합니다. 디바이스 셸이었다면 샌드박스 셸에서 재실행하고, 샌드박스 셸이었다면 `NETWORK_PERMISSION_REQUIRED`로 처리합니다. 설정 변경 후에는 같은 프로젝트에서 새 Cowork 작업을 엽니다.
