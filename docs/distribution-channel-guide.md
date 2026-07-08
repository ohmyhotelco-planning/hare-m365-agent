# Hare M365 Agent 배포 채널 가이드

## 결론

현재 1순위 배포 방식은 **public GitHub Release + npm exec**입니다.

이 방식은 npm registry 유료 private package를 쓰지 않으면서도, Windows/Mac/Linux/LLM shell 환경에서 같은 npm 실행 모델을 사용할 수 있습니다.

## 1순위: Public GitHub Release + npm exec

GitHub Release에 npm tarball을 올리고 아래처럼 실행합니다.

```bash
HARE_M365_DATA_DIR=./runtime npm exec --yes --package "https://github.com/ohmyhotelco-planning/hare-m365-agent/releases/download/v0.1.0/ohmyhotel-hare-m365-agent-0.1.0.tgz" -- hare-m365 llm-guide
```

장점:

- 직원에게 GitHub 계정이 없어도 public release asset은 받을 수 있습니다.
- npm registry publish가 필요 없습니다.
- OS별 exe/pkg를 따로 만들지 않아도 됩니다.
- LLM이 shell에서 직접 실행하기 쉽습니다.

주의:

- GitHub 저장소 또는 Release가 private이면 GitHub 계정/권한 문제가 다시 생깁니다.
- Release가 public이면 `.tgz` 안의 코드와 `.env` 설정 파일도 공개됩니다.
- 직원 PC 또는 LLM 환경에 Node.js/npm이 필요합니다.
- Claude/Cowork처럼 폴더를 연결해 쓰는 환경에서는 같은 연결 폴더의 `runtime`을 인증 캐시 위치로 사용해야 합니다.
- Claude/Cowork처럼 도메인 허용 목록이 있는 환경에서는 `github.com`, `release-assets.githubusercontent.com`, `registry.npmjs.org`, `graph.microsoft.com`, `login.microsoftonline.com`을 허용해야 합니다.
- GitHub Release asset은 `github.com`에서 `release-assets.githubusercontent.com`으로 리다이렉트됩니다.
- 현재 v0.1.0은 npm 의존성 설치 때문에 `registry.npmjs.org`도 필요할 수 있습니다. 다음 릴리즈에서 dependencies를 bundled로 포함하면 이 의존성을 줄일 수 있습니다.

## 2순위: Public npm registry

코드 공개가 가능하고 npmjs.com 배포를 승인받으면 가장 짧게 실행할 수 있습니다.

```bash
npx @ohmyhotel/hare-m365-agent llm-guide
```

단, private npm package는 유료/인증 문제가 있어 비개발자 배포에는 적합하지 않습니다.

## 3순위: Intune / Company Portal / Jamf

비개발자에게 Node.js/npm 설치까지 포함해 관리형으로 배포해야 한다면 Intune, Company Portal, Jamf를 검토합니다.

장점:

- 버전 관리, 회수, 대상자 지정이 가능합니다.
- macOS 보안 정책, Windows SmartScreen, 실행 파일 차단 문제를 표준 배포 체계로 다룰 수 있습니다.

단점:

- OS별 패키징이 필요합니다.
- LLM이 직접 쓰는 npm 실행 모델과는 거리가 있습니다.

## 4순위: SharePoint 보조 배포

SharePoint는 npm 실험의 중심 배포 채널이 아니라 보조 채널입니다.

사용 용도:

- 사용자 안내 문서 공유
- GitHub Release 링크 공지
- 오프라인/차단 대비용 산출물 보관

주의:

- `.exe`, `.cmd`, `.ps1`, `.pkg`, `.tgz` 등은 조직 보안 정책에 따라 업로드 또는 공유가 차단될 수 있습니다.
- SharePoint를 중심으로 잡으면 결국 파일 다운로드형 배포가 되어 npm 실행 모델 검증과 어긋납니다.

## Teams 안내 예시

```text
Hare M365 Agent 사용 안내입니다.

LLM에게 아래 문구를 전달하세요.

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

처음 사용 시 Microsoft device-code 로그인이 필요할 수 있습니다.
로그인 화면의 코드는 본인이 직접 브라우저에 입력하고, 채팅에는 붙여넣지 마세요.
```

## Release 파일

GitHub Release `v0.1.0`에 업로드할 파일:

```text
ohmyhotel-hare-m365-agent-0.1.0.tgz
SHA256SUMS.txt
START_HERE.html
LLM_FIRST_PROMPT_KO.txt
README.md
Hare_M365_Start_Windows.zip
Hare_M365_Start_Mac_Linux.sh
github-release-npm-guide.md
```

업로드할 때는 `releases/github-release/v0.1.0-upload-only` 폴더의 파일만 사용합니다. 핵심은 `.tgz` URL과 LLM 최초 프롬프트입니다.
