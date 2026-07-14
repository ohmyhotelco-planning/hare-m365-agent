# Hare M365 Agent LLM 지침

## 목적

Hare는 Microsoft Graph delegated 권한으로 Outlook, Teams, SharePoint/OneDrive를 조회하는 읽기 전용 CLI입니다.

## Cowork 실행 모델

- 사용자가 Cowork 작업을 열 때 선택한 프로젝트 마운트만 Hare의 영구 `dataDir`로 사용합니다.
- 선택 프로젝트에는 `.cache`, `claude`, `downloads`, `results`, `logs`만 저장합니다.
- `git clone`, `npm ci`, 빌드는 준비 명령이 지정한 Cowork 세션 런타임에서만 수행합니다.
- 선택 프로젝트 안에서 저장소를 clone하거나 빌드하지 않습니다.
- 선택 프로젝트의 삭제 권한을 요청하지 않습니다. Hare 상태 파일은 삭제 없이 덮어씁니다.
- `/tmp`, `/dev/shm`, `/home/claude`, `/root/.local/share` 또는 이름이 비슷한 다른 폴더를 `dataDir`로 사용하지 않습니다.
- 프로젝트가 선택되지 않았으면 `FOLDER_REQUIRED`로 중단합니다.

필수 도메인은 `github.com`, `registry.npmjs.org`, `login.microsoftonline.com`, `graph.microsoft.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com`입니다. GitHub API나 Release asset이 아닌 `git ls-remote`와 `git clone` 경로를 사용합니다.

## 로그인 하드게이트

- `startup.setup.state`와 `setup.nextCommand`만 따릅니다.
- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365 조회를 실행합니다.
- 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
- `LOGIN_START_REQUIRED`이면 `auth login-start`를 한 번 실행하고 사용자에게 Microsoft 주소와 코드를 보여줍니다.
- 사용자가 로그인을 마쳤다고 말하면 `LOGIN_COMPLETE_REQUIRED`의 명령을 한 번 실행합니다.
- 장기 poller, 백그라운드, `setsid`, `nohup`을 사용하지 않습니다.
- 실패 시 다른 경로로 이동하거나 삭제 권한을 요청하지 않고 실패 단계와 오류 한 줄만 보고합니다.

## 조회 기준

- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일이며 실제 범위를 답변에 포함합니다.
- 정확한 메일 건수는 `outlook count`를 사용합니다.
- 최신 Teams 채팅은 실제 마지막 메시지 생성 시각으로 판단합니다.
- SharePoint 사이트 존재 여부는 `sharepoint sites`로 확인합니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드·삭제·공유, 권한 변경은 서비스하지 않습니다.
