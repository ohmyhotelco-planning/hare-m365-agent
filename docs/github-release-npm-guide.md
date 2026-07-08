# GitHub Release npm 실행 가이드

## 결론

Hare M365 Agent는 npm registry 없이도 GitHub Release의 `.tgz` 파일을 npm으로 실행할 수 있습니다.

GitHub 저장소와 Release가 public이면 직원에게 GitHub 계정이 없어도 사용할 수 있습니다.

## Release에 올릴 파일

`npm run package:github-release` 실행 후 아래 폴더의 파일만 GitHub Release `v0.1.0`에 업로드합니다.

```text
releases/github-release/v0.1.0-upload-only/
```

GitHub 업로드 허용 확장자에 맞춘 파일 목록:

```text
ohmyhotel-hare-m365-agent-0.1.0.tgz
SHA256SUMS.txt
START_HERE.html
LLM_FIRST_PROMPT_KO.txt
README.md
github-release-npm-guide.md
Hare_M365_Start_Mac_Linux.sh
Hare_M365_Start_Windows.zip
```

`Hare_M365_Start_Windows.zip` 안에는 Windows 시작 스크립트 `.cmd`가 들어 있습니다. GitHub 업로드 화면에서 `.cmd`를 직접 받지 않는 경우가 있어 ZIP으로 감쌉니다.

사용자에게는 먼저 `START_HERE.html`을 열어 도메인 허용 단계를 완료하도록 안내합니다.

## LLM 실행 명령

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

이 방식은 전역 설치가 아니므로, 후속 명령도 같은 접두어를 붙입니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

## 사용자에게 줄 맨 처음 1회용 프롬프트

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

사용자가 도메인 허용을 완료했다고 말하면 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 Microsoft 365 요청을 처리해.
loggedIn이 false이면 사용자에게 어떤 명령을 실행할지 묻지 말고 auth login을 먼저 실행해. Microsoft device code는 사용자가 화면에서 직접 보고 브라우저에 입력하게 해. 코드를 채팅에 복사하거나 반복하지 마.
.env, .cache, token, device code는 읽거나 출력하지 마.
```

## 이후 새 채팅세션에서 다시 사용할 때

새 채팅세션은 Hare M365 Agent를 기억하지 못하므로 패키지 URL과 안전 지침을 다시 알려줘야 합니다. 이미 도메인 허용을 완료한 사용자는 아래 짧은 프롬프트를 사용합니다.

```text
도메인 허용은 이미 완료했다. Hare M365 Agent를 사용해줘.

패키지 URL:
https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz

먼저 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인해.
loggedIn이 true면 바로 내 요청을 처리하고, false면 사용자에게 선택지를 묻지 말고 auth login을 먼저 실행해. Microsoft 로그인은 내가 화면의 코드를 직접 브라우저에 입력해서 완료할게. 코드를 채팅에 복사하거나 반복하지 마.
.env, .cache, token, device code는 읽거나 출력하지 마.

요청:
[여기에 Outlook/Teams/파일 관련 요청을 적기]
```

## 사람이 해야 하는 일

- Node.js/npm이 설치되어 있어야 합니다.
- Claude/Cowork처럼 도메인 허용 목록이 있는 환경은 아래 도메인을 허용해야 합니다.
  - `github.com`
  - `release-assets.githubusercontent.com`
  - `registry.npmjs.org`
  - `graph.microsoft.com`
  - `login.microsoftonline.com`
- 처음 사용하거나 인증이 만료되면 Microsoft device-code 로그인을 직접 완료해야 합니다.

이 도메인 허용은 npm 실행보다 먼저 완료되어야 합니다. 허용 전에는 LLM이 패키지를 다운로드할 수 없으므로 `llm-guide`도 읽을 수 없습니다.

## Cowork 도메인 주의

GitHub Release asset은 `github.com`에서 `release-assets.githubusercontent.com`으로 리다이렉트됩니다. 따라서 `github.com`만 허용하면 다운로드가 실패할 수 있습니다.

현재 tarball은 실행 시 npm 의존성을 설치할 수 있으므로 `registry.npmjs.org`도 필요할 수 있습니다. 이 도메인을 제거하려면 다음 릴리즈에서 dependencies를 bundled로 패키징합니다.

## 주의

- public GitHub Release에 올리면 `.tgz` 안의 코드와 `.env` 설정 파일도 외부에서 받을 수 있습니다.
- `.env`는 delegated public-client 설정 파일이지만, 공개 전 내부 승인 기준은 확인해야 합니다.
- token, cache, device code, cookie, credential은 어떤 배포 파일에도 포함하지 않습니다.
