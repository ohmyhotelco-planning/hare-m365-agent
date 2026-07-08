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
releases/github-release/v0.1.0-upload-only/
```

GitHub Release에는 `v0.1.0-upload-only` 폴더의 파일만 업로드합니다.

## GitHub Release 필수 파일

```text
ohmyhotel-hare-m365-agent-0.1.0.tgz
SHA256SUMS.txt
START_HERE.html
LLM_FIRST_PROMPT_KO.txt
README.md
github-release-npm-guide.md
```

선택 파일:

```text
Hare_M365_Start_Windows.zip
Hare_M365_Start_Mac_Linux.sh
```

## 사용자 안내

사용자는 긴 명령어를 직접 외울 필요가 없습니다. 최초 사용 시에는 `LLM_FIRST_PROMPT_KO.txt` 내용을 LLM에게 한 번 전달하면 됩니다.

비개발자 사용자에게는 먼저 `START_HERE.html`을 열도록 안내합니다. 이 HTML은 도메인 허용, 맨 처음 1회용 프롬프트 복사, 새 채팅세션용 재사용 프롬프트 복사, 로그인 주의사항을 한 화면에서 보여줍니다.

Claude/Cowork처럼 프로젝트 또는 배포 폴더를 연결해 쓰는 환경에서는 같은 연결 폴더 안의 `runtime`을 인증 캐시 위치로 사용합니다. LLM이 실제 M365 조회를 수행하려면 LLM이 실행하는 CLI 프로세스가 이 캐시에 접근할 수 있어야 합니다.

단, Claude/Cowork처럼 도메인 허용 목록이 있는 환경에서는 사용자가 LLM에게 작업을 맡기기 전에 아래 도메인을 먼저 허용해야 합니다.

```text
github.com
release-assets.githubusercontent.com
registry.npmjs.org
graph.microsoft.com
login.microsoftonline.com
```

핵심 실행 명령은 아래 형태입니다.

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

## 사람이 해야 하는 일

- Node.js/npm이 설치되어 있어야 합니다.
- Claude/Cowork 도메인 허용 목록에는 `github.com`, `release-assets.githubusercontent.com`, `registry.npmjs.org`, `graph.microsoft.com`, `login.microsoftonline.com`을 LLM 실행 전에 먼저 허용합니다.
- 연결 폴더를 쓰는 경우 모든 Hare 명령에 `HARE_M365_DATA_DIR=./runtime`을 붙입니다.
- 새 채팅세션은 Hare M365 Agent를 기억하지 못하므로 `START_HERE.html`의 새 채팅용 프롬프트를 다시 붙여넣습니다.
- 처음 사용하거나 로그인 만료 시 같은 연결 폴더와 같은 `HARE_M365_DATA_DIR=./runtime` 설정으로 Microsoft device-code 로그인을 직접 완료합니다.
- device code, token, `.cache`, `runtime/.cache` 내용은 채팅에 붙여넣지 않습니다.

## Cowork 도메인 주의

GitHub Release 파일 URL은 `github.com`에서 시작하지만 실제 asset 다운로드는 `release-assets.githubusercontent.com`으로 리다이렉트됩니다.

현재 패키지는 npm 실행 중 의존성 설치를 위해 `registry.npmjs.org` 접근도 필요할 수 있습니다. Cowork에서 도메인 허용 목록을 최소화하려면 다음 릴리즈에서 npm 의존성을 tarball에 bundled로 포함하는 방식을 검토합니다.

## 보조 배포

Windows exe, Mac pkg, SharePoint 배포는 보조 경로입니다. 현재 목적이 LLM에 붙여 쓰는 공통 실행 도구라면 우선 GitHub Release/npm exec 모델을 검증합니다.

## 보안

- `.env`는 delegated public-client 설정 파일로 배포에 포함할 수 있습니다.
- token, cache, device code, cookie, credential은 배포 파일에 포함하지 않습니다.
- `runtime/.cache/msal-cache.json`은 로그인 후 생성되는 로컬 인증 캐시입니다. CLI가 접근하는 것은 허용하지만 LLM이 내용을 읽거나 출력하면 안 됩니다.
- 기본 정책은 read-only입니다.
