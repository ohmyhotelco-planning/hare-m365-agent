# Hare M365 Agent for Claude/Cowork

이 저장소에서는 [AGENTS.md](AGENTS.md)를 Hare 실행의 단일 기준으로 사용합니다.

핵심 규칙:

1. GitHub API가 아닌 `git clone` 경로를 사용합니다.
2. 최초 설치 전 `registry.npmjs.org`를 포함한 필수 도메인 6개를 허용합니다.
3. `dist/cli.js`, `dist/proxy.js`, `dist/msal-network.js`를 모두 확인합니다.
4. `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365를 조회합니다.
5. `auth login`은 같은 셸 호출의 포그라운드에서 실행하며 성공 JSON이 나올 때까지 호출을 유지합니다.
6. 백그라운드, detached, `setsid`, `nohup` 로그인을 사용하지 않습니다.
7. 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
8. 기본 동작은 읽기 전용입니다.
