# GitHub Release npm 실행 가이드

## 결론

Hare M365 Agent는 npm registry 없이도 GitHub Release의 `.tgz` 파일을 npm으로 실행할 수 있습니다.

GitHub 저장소와 Release가 public이면 직원에게 GitHub 계정이 없어도 사용할 수 있습니다.

## Release에 올릴 파일

`npm run package:github-release` 실행 후 아래 폴더의 파일을 GitHub Release `v0.1.0`에 업로드합니다.

```text
releases/github-release/v0.1.0/
```

필수 파일:

```text
ohmyhotel-hare-m365-agent-0.1.0.tgz
ohmyhotel-hare-m365-agent-0.1.0.tgz.sha256
LLM_FIRST_PROMPT_KO.txt
README.md
```

선택 파일:

```text
Hare_M365_Start_Windows.cmd
Hare_M365_Start_Mac_Linux.sh
```

## LLM 실행 명령

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

이 방식은 전역 설치가 아니므로, 후속 명령도 같은 접두어를 붙입니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 doctor
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 auth status
```

## 사용자에게 줄 최초 프롬프트

```text
아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해.
패키지 URL: https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz
먼저 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 내 Microsoft 365 요청을 처리해.
.env, .cache, token, device code는 읽거나 출력하지 마.
```

## 사람이 해야 하는 일

- Node.js/npm이 설치되어 있어야 합니다.
- Claude/Cowork처럼 도메인 허용 목록이 있는 환경은 아래 도메인을 허용해야 합니다.
  - `github.com`
  - `objects.githubusercontent.com`
  - `registry.npmjs.org`
  - `graph.microsoft.com`
  - `login.microsoftonline.com`
- 처음 사용하거나 인증이 만료되면 Microsoft device-code 로그인을 직접 완료해야 합니다.

## 주의

- public GitHub Release에 올리면 `.tgz` 안의 코드와 `.env` 설정 파일도 외부에서 받을 수 있습니다.
- `.env`는 delegated public-client 설정 파일이지만, 공개 전 내부 승인 기준은 확인해야 합니다.
- token, cache, device code, cookie, credential은 어떤 배포 파일에도 포함하지 않습니다.
