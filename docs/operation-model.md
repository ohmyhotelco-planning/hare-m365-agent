# OMH M365 Agent 운영 모델

## 목적

LLM이 Microsoft 365 업무 질문에 답할 수 있도록 로컬의 정책 통제 실행 계층을 제공합니다. 인증은 Microsoft Graph delegated permission 기반이며, 실제 동작은 로그인한 사용자 권한 안에서 수행됩니다.

이 CLI는 사람이 직접 조작하는 데스크톱 앱이 아니라 LLM이 호출하는 안정적인 명령 표면입니다.

## 기본 실행 모델

```text
LLM -> 로컬 Linux shell/Cowork sandbox -> omh-m365 CLI -> Microsoft Graph delegated access
```

사람은 인증, 승인, 정책 판단, 되돌리기 어려운 작업에만 개입합니다. 조회, 요약, 진단, 재시도, smoke test는 CLI와 LLM이 처리합니다.

## 권한 모델

- Azure Enterprise Application 권한이 Microsoft Graph 접근 상한을 정합니다.
- Delegated login이므로 실제 조회는 로그인한 사용자 권한으로 수행됩니다.
- CLI 정책과 LLM 지침이 LLM이 호출할 수 있는 기능을 제한합니다.
- 현재 POC는 읽기 중심입니다. write, send, delete, upload, share 작업은 닫아둡니다.

## 사람이 담당하는 일

- Azure Application 및 Graph permission 승인
- Claude/Cowork 도메인 허용 목록에 `graph.microsoft.com`, `login.microsoftonline.com` 추가
- `.env` 배포 설정 확인
- Microsoft 로그인 및 consent 완료
- 향후 write/send/delete/share 기능을 열기 전 명시 승인
- 비밀 노출 의심 시 revoke/rotation 및 노출 범위 확인

## LLM이 담당하는 일

- 문서화된 CLI만 사용
- 임의 Graph script 작성 금지
- 사람에게 묻기 전에 `doctor`, `auth status` 등 진단 먼저 실행
- 작은 limit으로 조회 후 필요한 만큼만 확장
- 메일/채팅/파일/토큰/캐시 원문 덤프 지양
- 인증, 승인, 정책 누락, 되돌리기 어려운 작업이 필요한 경우에만 사람에게 요청

## 현재 검증된 기능

- 로컬 설정 진단
- 로그인 상태 확인
- Outlook Inbox 조회
- Teams 팀 조회
- Teams 채팅 조회
- Teams 채팅 메시지 조회
- OneDrive/SharePoint-visible 파일 검색
- 명시 요청된 파일 다운로드

## 파일/토큰 취급

- `.env`는 delegated public-client 설정 파일로 배포에 포함할 수 있습니다.
- `.env`와 `.cache`는 LLM이 읽거나 출력하지 않습니다.
- 배포 템플릿에는 `.cache`, `downloads`, `logs`를 포함하지 않습니다.
- POC의 token cache는 파일 기반입니다. 파일럿/운영 전에는 OS credential store 등 더 안전한 저장 방식을 검토합니다.

## 파일럿 전 남은 과제

- target LLM 환경의 로컬 sandbox 지속성 및 보안 특성 확정
- Cowork 도메인 허용 목록 정책과 Microsoft Graph/Login 접근 가능 여부 확인
- OS-backed credential store 검토
- silent token usability를 확인하는 더 강한 `auth status` 추가
- `Chat.ReadWrite` scope 제거 가능성 검토
- SharePoint/Teams 파일 검색 범위 확장
- release ownership, versioning, update, rollback, audit 기준 정의
