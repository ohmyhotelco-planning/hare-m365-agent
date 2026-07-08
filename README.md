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
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

같은 패키지 URL을 사용해 다른 명령도 실행합니다.

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

GitHub Release가 아니라 npm registry에 공개 배포한 경우에는 `npx @ohmyhotel/hare-m365-agent ...` 형태도 사용할 수 있습니다.

## 연결형 runtime 폴더

Claude/Cowork처럼 프로젝트 또는 배포 폴더를 연결해 쓰는 환경에서는 같은 연결 폴더 안의 `runtime`을 인증 캐시와 다운로드 저장 위치로 사용합니다.

```text
<연결한 폴더>/
  runtime/
    .cache/
      msal-cache.json
    downloads/
    logs/
```

모든 Hare 명령은 같은 폴더에서 `HARE_M365_DATA_DIR=./runtime`을 붙여 실행합니다. CLI 프로세스는 `runtime/.cache/msal-cache.json`을 사용해 Graph 인증을 수행할 수 있지만, LLM은 이 파일을 열거나 출력하거나 요약하면 안 됩니다.

Windows PowerShell에서 같은 연결 폴더에 로그인 캐시를 만들려면:

```powershell
$env:HARE_M365_DATA_DIR = "$PWD\runtime"
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
```

## 맨 처음 1회용 프롬프트

설치 직후 처음 사용하는 채팅, 또는 도메인 허용 여부가 확실하지 않은 채팅에는 아래 프롬프트를 한 번 전달합니다.

```text
아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해.
패키지 URL: https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz

중요: Claude/Cowork처럼 도메인 허용 목록이 있는 환경이면 npm 실행 전에 먼저 사용자에게 아래 도메인이 허용되어 있는지 확인시켜.
github.com
release-assets.githubusercontent.com
registry.npmjs.org
graph.microsoft.com
login.microsoftonline.com

사용자가 "도메인 허용 완료"라고 확인하기 전에는 npm exec, 로그인, Outlook/Teams/Files 조회를 시작하지 마.

폴더 접근 권한이 없으면 파일 탐색기나 데스크톱 제어로 우회하지 말고, Claude/Cowork의 폴더 연결 기능으로 문서/Documents 안의 Hare M365 Agent 전용 폴더 연결/허용을 먼저 요청해. 사용자가 임의 경로를 판단하게 하지 마.

프로젝트 또는 배포 폴더가 연결되어 있으면 같은 연결 폴더 안의 ./runtime을 인증 캐시 위치로 사용하고 모든 Hare 명령에 HARE_M365_DATA_DIR=./runtime을 붙여 실행해. CLI는 runtime/.cache/msal-cache.json을 사용할 수 있지만, 너는 그 파일을 열거나 출력하거나 요약하지 마.

사용자가 도메인 허용을 완료했다고 말하면 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 내 Microsoft 365 요청을 처리해.
loggedIn이 false이면 로그인 hard gate로 멈추고 Outlook/Teams/Files 조회를 실행하지 마. Cowork/샌드박스에서는 auth login을 자동 실행하지 말고, 사용자가 직접 볼 수 있는 로컬 터미널에서 실행할 로그인 명령을 안내해. 연결 폴더를 쓰는 경우 사용자가 같은 폴더에서 같은 HARE_M365_DATA_DIR=./runtime 설정으로 로그인해야 해. 사용자가 "로그인 완료"라고 말하면 doctor/auth status를 다시 확인하고 원래 요청을 이어서 처리해.
.env, .cache, runtime/.cache, token, device code는 읽거나 출력하지 마.
```

## 이후 새 채팅세션에서 다시 사용할 때

새 채팅세션은 Hare M365 Agent와 패키지 URL을 기억하지 못합니다. 도메인 허용과 로그인을 이미 끝냈더라도 아래 짧은 프롬프트를 다시 전달합니다.

```text
도메인 허용은 이미 완료했다. Hare M365 Agent를 사용해줘.

패키지 URL:
https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz

먼저 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인해.
폴더 접근 권한이 없으면 Claude/Cowork의 폴더 연결 기능으로 문서/Documents 안의 Hare M365 Agent 전용 폴더 연결/허용을 먼저 요청해.
프로젝트 또는 배포 폴더가 연결되어 있으면 같은 연결 폴더 안의 ./runtime을 인증 캐시 위치로 사용하고 모든 Hare 명령에 HARE_M365_DATA_DIR=./runtime을 붙여 실행해.
loggedIn이 true면 바로 내 요청을 처리하고, false면 로그인 hard gate로 멈춰. Cowork/샌드박스에서는 auth login을 자동 실행하지 말고, 내가 직접 볼 수 있는 로컬 터미널에서 실행할 로그인 명령을 안내해. 연결 폴더를 쓰는 경우 같은 폴더에서 같은 HARE_M365_DATA_DIR=./runtime 설정으로 로그인해야 해. 내가 "로그인 완료"라고 말하면 doctor/auth status를 다시 확인하고 원래 요청을 이어서 처리해.
.env, .cache, runtime/.cache, token, device code는 읽거나 출력하지 마.

요청:
[여기에 Outlook/Teams/파일 관련 요청을 적기]
```

## 사람이 해야 하는 일

1. LLM에게 맨 처음 1회용 프롬프트를 주기 전에 Claude/Cowork 설정에서 아래 5개 도메인을 허용합니다.
   - `github.com`
   - `release-assets.githubusercontent.com`
   - `registry.npmjs.org`
   - `graph.microsoft.com`
   - `login.microsoftonline.com`
2. 폴더 허용 요청이 뜨면 `문서/Documents > Hare M365 Agent` 전용 폴더를 허용합니다.
3. 프로젝트/배포 폴더를 연결해 쓰는 경우 같은 폴더의 `runtime`을 인증 캐시 위치로 사용합니다.
4. 처음 사용하거나 로그인이 만료되었으면 같은 연결 폴더에서 같은 `HARE_M365_DATA_DIR=./runtime` 설정으로 `auth login`을 실행하고 Microsoft device-code 로그인을 직접 완료합니다.
5. device code, token, `.cache`, `runtime/.cache` 내용은 채팅에 붙여넣지 않습니다.

참고: GitHub Release 파일 URL은 `github.com`에서 시작하지만 실제 다운로드는 `release-assets.githubusercontent.com`으로 리다이렉트됩니다. 현재 v0.1.0 tarball은 실행 중 npm 의존성 설치를 위해 `registry.npmjs.org`도 필요할 수 있습니다.

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
- CLI가 `runtime/.cache/msal-cache.json`을 읽어 Graph 호출에 사용하는 것은 허용합니다.
- `.cache/`, `runtime/.cache/`, token cache, access token, refresh token, cookie, device code, credential은 공유하거나 출력하지 않습니다.
- LLM은 `.env`, `.cache`, `runtime/.cache` 내용을 읽거나 출력하면 안 됩니다.
- 기본 정책은 read-only입니다.
- 토큰이나 credential이 노출되면 삭제만으로 충분하지 않습니다. revoke/rotation과 노출 범위 확인이 필요합니다.

## 참고 문서

- [GitHub Release npm 실행 가이드](docs/github-release-npm-guide.md)
- [npm CLI 가이드](docs/npm-cli-guide.md)
- [LLM 실행 규칙](AGENTS.md)
- [Claude 실행 규칙](CLAUDE.md)
- [운영 모델](docs/operation-model.md)
- [배포 채널 가이드](docs/distribution-channel-guide.md)
