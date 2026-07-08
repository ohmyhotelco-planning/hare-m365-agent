# Hare M365 Agent

Hare M365 Agent는 LLM이 사용자의 로컬 실행 환경에서 Microsoft 365를 조회하도록 만든 Graph 기반 CLI입니다.

사용자가 자연어로 요청하면 LLM이 shell 명령으로 Hare를 호출하고, Hare는 사용자의 Microsoft 위임 로그인 권한 안에서 Outlook, Teams, OneDrive, SharePoint 데이터를 읽습니다.

```text
LLM -> local shell -> npm exec / hare-m365 -> Microsoft Graph delegated access
```

## 권장 배포 방식

현재 권장 방식은 **GitHub Release에 npm tarball을 올리고 `npm exec`로 실행**하는 방식입니다.

npm registry publish는 필수가 아닙니다. GitHub 저장소가 public이면 GitHub 계정이 없는 직원도 아래 방식으로 실행할 수 있습니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

같은 패키지 URL을 사용해 다른 명령도 실행합니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

GitHub Release가 아니라 npm registry에 공개 배포한 경우에는 `npx @ohmyhotel/hare-m365-agent ...` 형태도 사용할 수 있습니다.

## LLM에게 처음 줄 프롬프트

```text
아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해.
패키지 URL: https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
먼저 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 내 Microsoft 365 요청을 처리해.
.env, .cache, token, device code는 읽거나 출력하지 마.
```

## 사람이 해야 하는 일

1. Claude/Cowork처럼 도메인 허용 목록이 있는 환경에서는 아래 2개를 허용합니다.
   - `graph.microsoft.com`
   - `login.microsoftonline.com`
2. 처음 사용하거나 로그인이 만료되었으면 LLM이 안내하는 `auth login`을 실행하고 Microsoft device-code 로그인을 직접 완료합니다.
3. device code, token, `.cache` 내용은 채팅에 붙여넣지 않습니다.

## 현재 지원 범위

지원:

- Microsoft device-code 로그인
- 설정과 로그인 상태 확인
- Outlook Inbox 최근 메일 조회
- Teams 팀/채팅/채팅 메시지 조회
- OneDrive/SharePoint에서 사용자에게 보이는 파일 검색
- 사용자가 명시한 파일 다운로드

미지원:

- 메일 발송
- Teams 메시지 게시
- 일정 생성/수정
- 파일 업로드/수정/삭제/공유
- 권한 변경

## 개발 및 패키징

의존성 설치:

```bash
npm install
```

검증:

```bash
npm run typecheck
npm run build
```

GitHub Release용 산출물 생성:

```powershell
npm run package:github-release
```

출력:

```text
releases/github-release/v0.1.0/
```

이 폴더의 파일을 GitHub Release `v0.1.0`에 업로드합니다.

## 보안 원칙

- `.env`는 delegated public-client 설정 파일로 배포에 포함할 수 있습니다.
- `.cache/`, token cache, access token, refresh token, cookie, device code, credential은 공유하거나 출력하지 않습니다.
- LLM은 `.env`와 `.cache` 내용을 읽거나 출력하면 안 됩니다.
- 기본 정책은 read-only입니다.
- 토큰이나 credential이 노출되면 삭제만으로 충분하지 않습니다. revoke/rotation과 노출 범위 확인이 필요합니다.

## 참고 문서

- [GitHub Release npm 실행 가이드](docs/github-release-npm-guide.md)
- [npm CLI 가이드](docs/npm-cli-guide.md)
- [LLM 실행 규칙](AGENTS.md)
- [Claude 실행 규칙](CLAUDE.md)
- [운영 모델](docs/operation-model.md)
- [배포 채널 가이드](docs/distribution-channel-guide.md)
