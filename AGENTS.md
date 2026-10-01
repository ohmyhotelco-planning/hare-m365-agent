# Hare M365 Agent LLM 지침

## 목적

Hare는 Microsoft Graph delegated 권한으로 Outlook, Teams, SharePoint/OneDrive를 조회하는 읽기 전용 CLI입니다.

## Claude Code Desktop Local 실행 모델

- `--host claude-code`는 Windows·Mac의 Desktop Local 실행용입니다. Cloud/원격/WSL 실행이나 권한 우회를 뜻하지 않습니다.
- 사용자가 선택한 기존 Hare 데이터 폴더를 작업 폴더로 열고, 모든 명령에 그 절대 경로를 `--data-dir`로 유지합니다. 새 세션에서도 같은 폴더의 `CLAUDE.md`와 `claude/hare-m365-agent-rules.md`를 읽습니다. Git worktree나 비슷한 이름의 새 폴더로 바꾸지 않습니다.
- 신규 사용자는 Git 저장소·OneDrive 등 동기화 폴더 밖의 로컬 폴더를 선택합니다. 기존 인증이 있다는 이유로 캐시 원문을 읽거나 출력하지 않습니다.
- 앱 clone·build는 데이터 폴더 밖의 로컬 런타임에서 수행합니다. Windows는 PowerShell, Mac은 POSIX 셸 기준이며 `--command-shell powershell|posix`와 반환된 명령 구문을 일치시킵니다.
- Git과 Node.js 20.18.1 이상이 실제 Claude Code 세션의 PATH에서 실행 가능해야 합니다. 설치·업데이트 실패 시 중단하고 기존 인증 데이터를 유지합니다.
- 연결 확인은 같은 접두사로 `network check --environment claude-code`를 사용합니다. Claude Code 표준 권한 UI로 필요한 범위만 승인받으며 전체 권한 우회 모드를 사용하지 않습니다. Cowork 허용 목록이 Local 실행 권한을 부여하지 않습니다.
- 호스트 변경만으로 재로그인하지 않습니다. startup의 상태를 따르며 앱·테넌트·계정 불일치나 네트워크 실패는 기존 캐시를 보존하고 중단합니다.

## Cowork 실행 모델 (Cowork에서만 적용)

- 사용자가 Cowork 작업을 열 때 선택한 프로젝트 마운트만 Hare의 영구 `dataDir`로 사용합니다.
- 선택 프로젝트에는 `.cache`, `claude`, `downloads`, `results`, `logs`만 저장합니다.
- `git clone`, `npm ci`, 빌드는 준비 명령이 지정한 Cowork 세션 런타임에서만 수행합니다.
- 선택 프로젝트 안에서 저장소를 clone하거나 빌드하지 않습니다.
- 선택 프로젝트의 삭제 권한을 요청하지 않습니다. Hare 상태 파일은 삭제 없이 덮어씁니다.
- `/tmp`, `/dev/shm`, `/home/claude`, `/root/.local/share` 또는 이름이 비슷한 다른 폴더를 `dataDir`로 사용하지 않습니다.
- 프로젝트가 선택되지 않았으면 `FOLDER_REQUIRED`로 중단합니다.

필수 도메인은 `github.com`, `registry.npmjs.org`, `login.microsoftonline.com`, `graph.microsoft.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com`입니다. GitHub API나 Release asset이 아닌 `git ls-remote`와 `git clone` 경로를 사용합니다.

## 로그인 하드게이트

- 새 실행 환경에서 startup/auth status 또는 첫 업무 조회 전, 같은 명령 접두사와 `--data-dir`로 `network check --environment <codex|cowork|claude-code|unknown>`을 실행합니다. 실제 호스트를 지정하며 이 옵션 자체가 권한을 부여하지는 않습니다. 인증 캐시 접근·파일 변경 없이 최대 3초 연결 검사만 수행합니다.
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
- 삭제된 항목은 기본 제외합니다. 사용자가 삭제된 메일만 요청하면 `--folder deleted`, 삭제된 메일 포함을 명시하면 `--folder all-with-deleted`를 recent/flagged/search/count에 사용합니다. 빈 결과를 이유로 자동 확장하지 않습니다. `deleted`는 삭제된 항목 폴더 직속 메일만 대상으로 하며 하위 폴더 재귀 조회·영구 삭제 복구 영역 조회·복원은 지원하지 않습니다. 날짜는 삭제일이 아닌 수신일 기준이고, 커서를 이어갈 때도 같은 `--folder`와 `--mailbox`를 유지합니다.
- 플래그된 메일은 `outlook flagged --folder all`을 사용하며 모든 메일 결과의 `flagStatus`를 확인합니다.
- 공유 사서함이 명시된 요청은 `--mailbox <name-or-address>`를 사용합니다. 이름 후보가 모호하거나 접근이 거부되면 중단하며 본인 사서함으로 대체하지 않습니다. 공유 사서함 첨부파일 명령에도 같은 `--mailbox`를 유지합니다.
- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일이며 실제 범위를 답변에 포함합니다.
- 정확한 메일 건수는 `outlook count`를 사용합니다.
- 최신 Teams 채팅은 실제 마지막 메시지 생성 시각으로 판단합니다.
- SharePoint 사이트 존재 여부는 `sharepoint sites`로 확인합니다.
- SharePoint 사이트 파일을 외부 경로에 복사할 때는 `sharepoint export-files`의 전체 계획을 먼저 보여주고 명시적 승인을 받은 뒤 실행합니다. 동일한 승인 작업은 파일별 재승인 없이 재개하며 기존 파일을 덮어쓰지 않습니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드·삭제·공유, 권한 변경은 서비스하지 않습니다.
