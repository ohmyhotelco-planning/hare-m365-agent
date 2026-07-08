# Hare M365 Agent GitHub Release

이 Release는 npm registry 없이 GitHub Release의 `.tgz` 파일을 npm으로 실행하기 위한 배포본입니다.

## 먼저 LLM에게 줄 문구

`LLM_FIRST_PROMPT_KO.txt` 내용을 LLM에게 전달하세요.

Claude/Cowork처럼 도메인 허용 목록이 있는 환경에서는 LLM에게 작업을 맡기기 전에 아래 도메인을 먼저 허용합니다.

```text
github.com
release-assets.githubusercontent.com
registry.npmjs.org
graph.microsoft.com
login.microsoftonline.com
```

이 설정이 끝나기 전에는 LLM이 npm 패키지를 다운로드할 수 없어서 `llm-guide`도 읽을 수 없습니다.

## 직접 실행 예

```bash
npm exec --yes --package "__PACKAGE_URL__" -- hare-m365 llm-guide
npm exec --yes --package "__PACKAGE_URL__" -- hare-m365 doctor
npm exec --yes --package "__PACKAGE_URL__" -- hare-m365 auth status
```

## Windows에서 스크립트로 시작

`Hare_M365_Start_Windows.zip`을 내려받아 압축을 풀고, 안의 `Hare_M365_Start_Windows.cmd`를 실행합니다.

처음 사용하거나 `loggedIn: false`가 나오면 스크립트 안내에 따라 Microsoft 로그인을 진행합니다.

## Mac/Linux에서 스크립트로 시작

터미널에서 실행합니다.

```bash
chmod +x Hare_M365_Start_Mac_Linux.sh
./Hare_M365_Start_Mac_Linux.sh
```

## 사람이 해야 하는 일

- Node.js/npm이 설치되어 있어야 합니다.
- Claude/Cowork 도메인 허용 목록에는 `github.com`, `release-assets.githubusercontent.com`, `registry.npmjs.org`, `graph.microsoft.com`, `login.microsoftonline.com`을 LLM 실행 전에 먼저 허용합니다.
- Microsoft device-code 로그인 코드는 본인이 브라우저에 직접 입력합니다.
- device code, token, `.cache` 내용은 채팅에 붙여넣지 않습니다.

## 포함 파일

```text
__PACKAGE_FILE__
SHA256SUMS.txt
LLM_FIRST_PROMPT_KO.txt
Hare_M365_Start_Windows.zip
Hare_M365_Start_Mac_Linux.sh
README.md
github-release-npm-guide.md
```
