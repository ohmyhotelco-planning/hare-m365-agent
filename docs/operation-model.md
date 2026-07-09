# Hare M365 Agent 운영 모델

## 목적

LLM이 사용자의 자연어 요청을 받아 Microsoft 365 업무 데이터를 조회할 수 있도록, 로컬 또는 Cowork 실행환경에서 Microsoft Graph delegated permission 기반 CLI를 실행합니다.

이 도구는 사람이 직접 쓰는 일반 앱이 아니라 LLM이 호출하는 실행 도구입니다. 기본 정책은 read-only입니다.

## 기본 실행 모델

```text
LLM -> local shell/Cowork sandbox -> git clone -> npm ci -> npm run build -> Hare CLI -> Microsoft Graph delegated access
```

사람은 인증, 승인, 권한 판단, 되돌리기 어려운 작업에만 개입합니다. 조회, 요약, 진단, 재시도, smoke test는 CLI와 LLM이 처리합니다.

## 권한 모델

- Azure Enterprise Application 권한이 Microsoft Graph 접근 상한을 정합니다.
- Delegated login 방식으로 실제 조회는 로그인한 사용자 권한으로 수행합니다.
- CLI 정책과 LLM 지침이 LLM이 호출할 수 있는 기능을 제한합니다.
- 현재 기준은 읽기 중심입니다. write, send, delete, upload, share 작업은 열지 않습니다.

## 사람이 담당하는 일

- Azure Application 및 Graph permission 승인
- Claude/Cowork 도메인 허용 목록에 `github.com`, `graph.microsoft.com`, `login.microsoftonline.com` 추가
- 기본 `hare.config.json` 설정 확인
- Microsoft 로그인 및 consent 완료
- 향후 write/send/delete/share 기능을 열기 전 명시 승인

## LLM이 담당하는 일

- 문서화된 CLI만 사용
- repo 접근 확인은 `git ls-remote` 또는 `git clone`으로 수행
- GitHub API 또는 GitHub Release asset 다운로드를 기본 경로로 사용하지 않음
- `npm ci`, `npm run build`, `node dist/cli.js` 실행
- `doctor`, `auth status` 진단
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

## 남은 과제

- Cowork 도메인 허용과 Microsoft Graph/Login 접근 가능성 반복 검증
- OS-backed credential store 검토
- silent token usability를 더 명확히 확인하는 `auth status` 강화
- `Chat.ReadWrite` scope 제거 가능성 검토
- SharePoint/Teams 파일 검색 범위 확장
- release ownership, versioning, update, rollback, audit 기준 정의
