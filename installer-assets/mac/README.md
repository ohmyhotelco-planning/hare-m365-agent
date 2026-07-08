# OMH M365 Agent Mac

Mac 사용자용 OMH M365 Agent 패키지입니다.

운영 또는 비개발자 배포는 `.pkg` 설치 방식을 권장합니다.
이 폴더형 패키지는 파일럿/검증용입니다.

## 시작

1. 압축을 풉니다.
2. `시작하기.html`을 더블클릭합니다.
3. 열린 안내 화면에 따라 진행합니다.

## PKG로 설치한 경우

1. Applications에서 `OMH M365 Agent Start`를 엽니다.
2. 로그인은 Applications의 `OMH M365 Agent Login`으로 진행합니다.
3. 열린 안내 화면의 Cowork 프롬프트를 사용합니다.

## Mac에서 실행이 막힐 때

처음 받은 `로그인 시작.command` 파일은 macOS가 확인 창을 띄울 수 있습니다.

- 더블클릭으로 열리지 않으면 `로그인 시작.command`를 Control-클릭 또는 오른쪽 클릭합니다.
- 메뉴에서 `열기`를 선택합니다.
- 확인 창이 뜨면 다시 `열기`를 선택합니다.

회사 보안 정책상 계속 차단되면 임의로 우회하지 말고 IT 배포 담당자에게 Jamf/Intune 배포본을 요청합니다.

## 포함된 것

- Mac용 실행 스크립트
- Apple Silicon / Intel Mac용 Node 런타임
- Microsoft Graph delegated public-client 설정 `.env`
- Cowork 실행용 `omh-m365.cjs`

## 주의

- device code, token, `.cache` 내용은 채팅에 붙여넣지 않습니다.
- 폴더형 패키지의 최초 로그인은 사용자가 직접 `로그인 시작.command`로 완료합니다.
- PKG 설치본의 최초 로그인은 `OMH M365 Agent Login` 앱으로 완료합니다.
- 실제 운영 배포는 Jamf, Intune for macOS, Company Portal 같은 관리형 배포 채널을 권장합니다.
