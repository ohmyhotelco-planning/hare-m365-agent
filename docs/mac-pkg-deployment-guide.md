# Mac PKG 배포 및 실제 검증 가이드

## 목적

Mac 사용자는 `.zip` 안의 `.command` 파일을 직접 실행하지 않습니다.

운영 또는 파일럿 배포는 SharePoint 링크로 `.pkg`를 내려받고, macOS Installer로 설치하는 방식을 기준으로 합니다.

## 배포 구조

권장 SharePoint 폴더 구조:

```text
OMH M365 Agent 배포/
├─ Windows/
│  ├─ OMH-M365-Agent-Setup-0.1.0.exe
│  └─ OMH-M365-Agent-Setup-0.1.0.exe.sha256
├─ Mac/
│  ├─ OMH-M365-Agent-mac-0.1.0.pkg
│  └─ OMH-M365-Agent-mac-0.1.0.pkg.sha256
└─ 배포안내.md
```

Teams에는 실행 파일을 첨부하지 않고 SharePoint 배포 폴더 링크만 공유합니다.

## Mac PKG 빌드

이 작업은 macOS에서 실행합니다.

사전 조건:

- Xcode Command Line Tools
- `pkgbuild`, `productbuild`
- `python3`
- `node`는 버전 확인용으로 있으면 좋지만, 없으면 `package.json`에서 버전을 읽습니다.
- `releases/mac-llm` 폴더가 먼저 생성되어 있어야 합니다.

기본 빌드:

```bash
cd /path/to/omh-m365-agent
bash scripts/build-mac-pkg.sh
```

생성 파일:

```text
releases/OMH-M365-Agent-mac-0.1.0.pkg
releases/OMH-M365-Agent-mac-0.1.0.pkg.sha256
```

서명 예시:

```bash
SIGN_APP_IDENTITY="Developer ID Application: Company Name (TEAMID)" \
SIGN_INSTALLER_IDENTITY="Developer ID Installer: Company Name (TEAMID)" \
bash scripts/build-mac-pkg.sh
```

공증 예시:

```bash
SIGN_APP_IDENTITY="Developer ID Application: Company Name (TEAMID)" \
SIGN_INSTALLER_IDENTITY="Developer ID Installer: Company Name (TEAMID)" \
NOTARY_KEYCHAIN_PROFILE="notary-profile-name" \
bash scripts/build-mac-pkg.sh
```

실제 Apple ID 비밀번호, app-specific password, API key 값은 문서나 채팅에 쓰지 않습니다. `notarytool store-credentials`로 Keychain profile을 만든 뒤 profile 이름만 사용합니다.

## 설치 결과

PKG 설치 후 생성되는 항목:

```text
/Library/Application Support/OMH/M365Agent/
/Applications/OMH M365 Agent Start.app
/Applications/OMH M365 Agent Login.app
```

사용자별 데이터 위치:

```text
~/Library/Application Support/OMH/M365Agent/.cache
~/Library/Application Support/OMH/M365Agent/downloads
~/Library/Application Support/OMH/M365Agent/runtime
```

토큰 캐시는 설치 폴더가 아니라 사용자 홈 아래에 저장됩니다.

## 사용자 안내

사용자에게 전달할 최소 안내:

```text
1. SharePoint 링크에서 Mac용 OMH-M365-Agent-mac-0.1.0.pkg를 다운로드합니다.
2. pkg를 열어 설치합니다.
3. Applications에서 "OMH M365 Agent Start"를 엽니다.
4. 안내 화면에 따라 Claude/Cowork 도메인을 허용합니다.
5. Applications에서 "OMH M365 Agent Login"을 열고 Microsoft device-code 로그인을 완료합니다.
6. 안내 화면의 Cowork 프롬프트를 Claude/Cowork 또는 Claude Code에 붙여넣고 원하는 작업을 요청합니다.
```

device code, token, `.cache`, `runtime/.cache` 내용은 채팅에 붙여넣지 않습니다.

## 실제 검증

Mac에서 설치 후 아래 명령을 실행합니다.

```bash
cd /path/to/omh-m365-agent
bash scripts/verify-mac-install.sh
```

로그인이 필요하면:

```text
Applications에서 "OMH M365 Agent Login"을 실행하고 Microsoft 로그인을 완료합니다.
```

로그인 후 Graph smoke test:

```bash
bash scripts/verify-mac-install.sh --graph-smoke
```

확인할 항목:

- `doctor`에서 `configured: true`
- `auth status`에서 `loggedIn: true`
- `outlook inbox --limit 1` 성공
- `teams chats --limit 5` 성공
- `files search --query test --limit 1` 성공

## Mac LLM 검증 프롬프트

Mac에서 Codex 또는 Claude Code에 아래처럼 요청합니다.

```text
OMH M365 Agent Mac pkg 실제 설치 검증을 진행해.

규칙:
- .env, .cache, runtime/.cache, token, device code는 읽거나 출력하지 마.
- 먼저 /Library/Application Support/OMH/M365Agent 폴더와 /Applications의 Start/Login 앱 존재를 확인해.
- scripts/verify-mac-install.sh를 실행해 doctor/auth status를 확인해.
- loggedIn이 false면 사용자가 Applications의 "OMH M365 Agent Login" 앱으로 로그인해야 한다고 안내해.
- loggedIn이 true면 --graph-smoke 검증으로 Outlook 1건, Teams chats 5건, Files search 1건을 조회해.
- 결과는 성공/실패와 막힌 지점만 요약해.
```

## 남은 운영화 과제

- Apple Developer ID 서명/공증 적용
- Jamf 또는 Intune for macOS 배포 검토
- 최초 로그인 UX 개선
- OS credential store 기반 토큰 저장 검토
- Cowork hosted sandbox와 로컬 Claude Code 실행 환경 차이 검증
