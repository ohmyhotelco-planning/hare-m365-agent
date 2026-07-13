# Hare M365 Agent 운영 모델

## 목적

LLM이 사용자의 자연어 요청을 받아 Microsoft 365 업무 데이터를 조회할 수 있도록, 로컬 또는 Cowork 실행환경에서 Microsoft Graph delegated permission 기반 CLI를 실행합니다.

이 도구는 사람이 직접 쓰는 일반 앱이 아니라 LLM이 호출하는 실행 도구입니다. 기본 정책은 read-only입니다.

## 기본 실행 모델

```text
LLM -> Cowork sandbox shell(도메인 허용 목록 적용) -> 스냅샷 복원 또는 git clone -> HEAD 검증 -> 필요 시 npm ci + npm run build -> Hare CLI -> Microsoft Graph delegated access
                     ↕ 파일 스테이징/커밋 도구
연결된 HareM365Agent 프로젝트 폴더(영속 저장소: 스냅샷, 빌드 마커, 로그인 캐시, 규칙, 결과)
```

Cowork에는 실행 셸이 두 곳 있을 수 있습니다. 도메인 허용 목록은 세션 샌드박스 셸에 적용되고, 연결 폴더가 마운트된 디바이스 셸은 설정과 무관하게 모든 외부 도메인이 차단될 수 있습니다. 네트워크가 필요한 명령은 샌드박스 셸에서 실행하고, 연결 폴더는 영속 저장소로 사용합니다. 연결 폴더 마운트가 네트워크가 되는 셸에서 직접 보이는 환경이라면 그 마운트를 실행 위치로 함께 사용해도 됩니다.

사람은 인증, 승인, 권한 판단, 되돌리기 어려운 작업에만 개입합니다. 조회, 요약, 진단, 재시도, smoke test는 CLI와 LLM이 처리합니다.

## 권한 모델

- Azure Enterprise Application 권한이 Microsoft Graph 접근 상한을 정합니다.
- Delegated login 방식으로 실제 조회는 로그인한 사용자 권한으로 수행합니다.
- CLI 정책과 LLM 지침이 LLM이 호출할 수 있는 기능을 제한합니다.
- 현재 기준은 읽기 중심입니다. write, send, delete, upload, share 작업은 열지 않습니다.

## 사람이 담당하는 일

- Azure Application 및 Graph permission 승인
- Claude/Cowork 도메인 허용 목록에 `github.com`, `registry.npmjs.org`, `graph.microsoft.com`, `login.microsoftonline.com`, `ohmylab-my.sharepoint.com`, `ohmylab.sharepoint.com` 추가
- 기본 `hare.config.json` 설정 확인
- Microsoft 로그인 및 consent 완료
- 향후 write/send/delete/share 기능을 열기 전 명시 승인

## LLM이 담당하는 일

- 문서화된 CLI만 사용
- repo 접근 확인은 `git ls-remote` 또는 `git clone`으로 수행
- GitHub API 또는 GitHub Release asset 다운로드를 기본 경로로 사용하지 않음
- `npm ci --prefer-offline --no-audit --no-fund`, `npm run build`, `node dist/cli.js` 실행 (샌드박스 셸에서)
- `doctor`, `auth status` 진단
- 연결된 HareM365Agent 프로젝트 폴더와 HARE_ROOT 간 스냅샷·로그인 캐시·규칙·결과 동기화
- `auth login-start`와 `auth login-complete`로 45초 제한 안에서 인증 완료 후 상태 재확인
- 조회 limit을 작게 시작하고 필요한 만큼만 확장
- 메일/채팅/파일/토큰/캐시 원문 덤프 지양
- 인증, 승인, 권한 전략, 되돌리기 어려운 작업은 사람에게 요청

## 현재 검증된 기능

- 로컬 설정 진단
- 로그인 상태 확인
- Outlook Inbox 조회
- Teams 팀 조회
- Teams 채팅 조회
- Teams 채팅 메시지 조회
- OneDrive/SharePoint-visible 파일 검색
- 명시 요청 파일 다운로드

## 배포 기준

- 기본: GitHub 저장소를 `git clone`해 Cowork/Claude 실행환경에서 빌드 후 실행
- 보류: npmjs publish
- 비권장: GitHub Release asset 직접 다운로드
- 폐기: exe/pkg/SharePoint zip 중심 배포

## 운영 검증 기준

- `loggedIn`과 `tokenUsable`이 모두 `true`일 때만 조회 실행
- 캐시 파일 존재만으로 로그인 성공을 판단하지 않음
- Graph 일시 오류는 제한된 횟수와 `Retry-After` 기준으로 재시도
- 기간 미지정 검색은 `Asia/Seoul` 기준 최근 90일로 고정
- 전체 메일 집계에서 삭제된 메일 제외
- release ownership, versioning, update, rollback 기준은 운영 전 별도 확정
