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

- 새 실행 환경에서 startup/auth status 또는 첫 업무 조회 전, 같은 명령 접두사와 `--data-dir`로 `network check --environment <codex|cowork|unknown>`을 실행합니다. 실제 호스트를 지정하며 이 옵션 자체가 권한을 부여하지는 않습니다. 인증 캐시 접근·파일 변경 없이 최대 3초 연결 검사만 수행합니다.
- Codex에서 `EXECUTION_PERMISSION_REQUIRED`이면 일반적인 실패 안내만 하고 끝내지 말고 호스트의 표준 실행 권한 요청 도구를 사용합니다. 승인된 경우에만 같은 검사를 1회 재실행하고, 같은 실행 파일·dataDir로 인증 확인 후 원래 요청한 읽기를 진행합니다. 거부·도구 없음·재실패 시 중단합니다. 쓰기 작업은 자동 재시도하지 않습니다.
- `NETWORK_PERMISSION_REQUIRED`인 허용 목록 차단은 Codex에서도 실행 권한 요청으로 우회하지 않습니다. Cowork의 `EACCES`도 로컬 셸로 우회하지 않으며 정책이 바뀌면 같은 프로젝트로 새 Cowork 작업을 엽니다. 그 외 오류는 `NETWORK_CHECK_BLOCKED`로 보고합니다.
- `REACHABLE`은 로그인 서버의 공개 주소 연결만 확인한 것입니다. 인증이나 Graph/SharePoint 권한 성공을 뜻하지 않습니다. 승인 유효 범위 안에서 같은 실행 환경을 유지하고 매 페이지마다 사전 검사를 반복하지 않습니다.
- `startup.setup.state`와 `setup.nextCommand`만 따릅니다.
- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 M365 조회를 실행합니다.
- 캐시 파일 존재만으로 로그인 성공으로 판단하지 않습니다.
- `AUTH_CHECK_BLOCKED`의 `loggedIn=null`, `tokenUsable=null`은 네트워크 문제로 확인 불가이며 만료나 로그아웃이 아닙니다. 위 사전 검사를 한 번 적용하고 기존 캐시를 유지합니다. 재로그인, 캐시 초기화, 정책 우회 또는 인증 확인 전 M365 조회는 진행하지 않습니다.
- `LOGIN_START_REQUIRED`이면 `auth login-start`를 한 번 실행하고 사용자에게 Microsoft 주소와 코드를 보여줍니다.
- 사용자가 로그인을 마쳤다고 말하면 `LOGIN_COMPLETE_REQUIRED`의 명령을 한 번 실행합니다.
- 장기 poller, 백그라운드, `setsid`, `nohup`을 사용하지 않습니다.
- 네트워크 실패는 위 사전 확인 절차를 적용합니다. 그 밖의 실패 시 다른 경로로 이동하거나 삭제 권한을 요청하지 않고 실패 단계와 오류 한 줄만 보고합니다.

## 조회 기준

- 일반·최근 메일은 `outlook recent --folder all`을 사용하고, `outlook inbox`는 받은편지함이 명시된 경우에만 사용합니다.
- 플래그된 메일은 `outlook flagged --folder all`을 사용하며 모든 메일 결과의 `flagStatus`를 확인합니다.
- 공유 사서함이 명시된 요청은 `--mailbox <name-or-address>`를 사용합니다. 이름 후보가 모호하거나 접근이 거부되면 중단하며 본인 사서함으로 대체하지 않습니다. 공유 사서함 첨부파일 명령에도 같은 `--mailbox`를 유지합니다.
- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일이며 실제 범위를 답변에 포함합니다.
- 정확한 메일 건수는 `outlook count`를 사용합니다.
- 최신 Teams 채팅은 실제 마지막 메시지 생성 시각으로 판단합니다.
- SharePoint 사이트 존재 여부는 `sharepoint sites`로 확인합니다.
- SharePoint 사이트 파일을 외부 경로에 복사할 때는 `sharepoint export-files`의 전체 계획을 먼저 보여주고 명시적 승인을 받은 뒤 실행합니다. 동일한 승인 작업은 파일별 재승인 없이 재개하며 기존 파일을 덮어쓰지 않습니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드·삭제·공유, 권한 변경은 서비스하지 않습니다.
