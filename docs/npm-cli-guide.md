# Hare M365 Agent npm CLI 가이드

## 목적

Hare M365 Agent는 LLM이 로컬 shell에서 Microsoft 365를 조회하기 위한 npm 기반 CLI입니다.

기본 실행 모델:

```text
LLM -> local shell -> npm exec / hare-m365 -> Microsoft Graph delegated access
```

## 1순위: GitHub Release tarball 실행

npm registry에 publish하지 않아도 GitHub Release에 올라간 `.tgz` 파일을 npm이 직접 실행할 수 있습니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

LLM은 같은 패키지 URL을 모든 명령에 반복해서 사용합니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
```

## 2순위: npm registry 공개 배포

나중에 public npm registry에 배포하면 더 짧게 실행할 수 있습니다.

```bash
npx @ohmyhotel/hare-m365-agent llm-guide
npx @ohmyhotel/hare-m365-agent doctor
npx @ohmyhotel/hare-m365-agent auth status
```

## 3순위: 전역 설치

개발자나 고정 사용자는 전역 설치도 가능합니다.

```bash
npm install -g "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz"
hare-m365 doctor
hare-m365 auth status
```

기존 호환 명령인 `omh-m365`도 유지됩니다.

## LLM 최초 프롬프트

```text
아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해.
패키지 URL: https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
먼저 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 내 Microsoft 365 요청을 처리해.
.env, .cache, token, device code는 읽거나 출력하지 마.
```

## 로그인

Microsoft device-code 로그인은 사용자가 직접 완료합니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth login
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

device code는 채팅에 붙여넣지 않습니다.

## 조회 명령

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 outlook inbox --limit 10
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams teams
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chats --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 teams chat-messages --chat-id "<chat-id>" --limit 20
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files search --query "keyword" --limit 10
```

파일 다운로드는 사용자가 특정 파일을 명시적으로 요청한 경우에만 수행합니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 files download --drive-id "<drive-id>" --item-id "<item-id>" --name "filename.ext"
```

## 로컬 저장 위치

npm CLI는 현재 폴더가 아니라 사용자별 OS 표준 위치에 캐시와 다운로드를 저장합니다.

Windows:

```text
%LOCALAPPDATA%\Ohmyhotel\HareM365Agent
```

Mac:

```text
~/Library/Application Support/Ohmyhotel/HareM365Agent
```

Linux:

```text
~/.local/share/ohmyhotel/hare-m365-agent
```

## 배포 파일 생성

GitHub Release용 산출물을 생성합니다.

```bash
npm run package:github-release
```

생성 파일:

```text
releases/github-release/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
releases/github-release/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz.sha256
releases/github-release/v0.1.0/Hare_M365_Start_Windows.cmd
releases/github-release/v0.1.0/Hare_M365_Start_Mac_Linux.sh
releases/github-release/v0.1.0/LLM_FIRST_PROMPT_KO.txt
releases/github-release/v0.1.0/README.md
```

이 파일들을 GitHub Release `v0.1.0`에 업로드합니다.

## 안전 규칙

- `.env`, `.cache`, token cache, access token, refresh token, cookie, device code를 읽거나 출력하지 않습니다.
- 기본 정책은 read-only입니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경은 현재 POC에서 수행하지 않습니다.
- Teams 최신 채팅 판단에는 `lastMessageCreatedDateTime`을 사용합니다.
