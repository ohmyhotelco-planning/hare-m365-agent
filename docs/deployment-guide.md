# Hare M365 Agent 배포 가이드

## 현재 기준

현재 기준 배포 모델은 **public GitHub Release + npm exec**입니다.

```text
LLM -> local shell -> npm exec GitHub Release .tgz -> hare-m365 -> Microsoft Graph
```

## 배포 준비

개발 머신에서 실행합니다.

```powershell
npm install
npm run package:github-release
```

산출물:

```text
releases/github-release/v0.1.0/
```

이 폴더의 파일을 GitHub Release `v0.1.0`에 업로드합니다.

## GitHub Release 필수 파일

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
github-release-npm-guide.md
```

## 사용자 안내

사용자는 긴 명령어를 직접 외울 필요가 없습니다. `LLM_FIRST_PROMPT_KO.txt` 내용을 LLM에게 전달하면 됩니다.

핵심 실행 명령은 아래 형태입니다.

```bash
npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

## 사람이 해야 하는 일

- Node.js/npm이 설치되어 있어야 합니다.
- Claude/Cowork 도메인 허용 목록에는 `github.com`, `objects.githubusercontent.com`, `registry.npmjs.org`, `graph.microsoft.com`, `login.microsoftonline.com`을 허용합니다.
- 처음 사용하거나 로그인 만료 시 Microsoft device-code 로그인을 직접 완료합니다.
- device code, token, `.cache` 내용은 채팅에 붙여넣지 않습니다.

## 보조 배포

Windows exe, Mac pkg, SharePoint 배포는 보조 경로입니다. 현재 목적이 LLM에 붙여 쓰는 공통 실행 도구라면 우선 GitHub Release/npm exec 모델을 검증합니다.

## 보안

- `.env`는 delegated public-client 설정 파일로 배포에 포함할 수 있습니다.
- token, cache, device code, cookie, credential은 배포 파일에 포함하지 않습니다.
- 기본 정책은 read-only입니다.
