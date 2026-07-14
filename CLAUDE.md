# Hare M365 Agent for Claude/Cowork

이 저장소에서는 [AGENTS.md](AGENTS.md)를 실행 기준으로 사용합니다.

1. Cowork 작업을 열 때 선택한 프로젝트 마운트를 영구 `dataDir`로 사용합니다.
2. 앱 코드는 Cowork 세션 런타임에서 clone·build하고, 선택 프로젝트에는 인증 캐시와 사용자 결과만 저장합니다.
3. 선택 프로젝트에서 git/npm/build를 실행하거나 삭제 권한을 요청하지 않습니다.
4. `/tmp`, `/dev/shm`, `/home/claude`, `/root/.local/share` 또는 다른 추측 경로로 우회하지 않습니다.
5. GitHub API가 아닌 `git ls-remote` 또는 `git clone`으로 저장소 접근을 확인합니다.
6. `startup.setup.state`와 `setup.nextCommand`만 따릅니다.
7. `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365를 조회합니다.
8. 로그인은 `auth login-start`와 `auth login-complete` 두 단계로 수행합니다.
9. 기본 동작은 읽기 전용입니다.
