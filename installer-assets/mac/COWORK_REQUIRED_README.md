# Cowork 필수 안내 - Mac

이 폴더는 Mac 사용자용 OMH M365 Agent 실행 폴더입니다.

## Cowork가 해야 하는 일

1. 이 폴더가 Cowork shell에 보이지 않으면 Cowork의 폴더/프로젝트 연결 도구로 이 폴더를 연결합니다.
2. Cowork가 보여주는 mounted path로 이동합니다.
3. 아래 파일을 확인합니다.

```bash
pwd
ls -la
test -f run-cowork.sh
test -f omh-m365.cjs
```

4. 모든 Microsoft 365 조회 명령은 wrapper로 실행합니다.

```bash
bash run-cowork.sh doctor
bash run-cowork.sh auth status
```

## 하드 게이트

Outlook, Teams, Files 명령을 실행하기 전에 반드시 아래 조건이 모두 참이어야 합니다.

- `configured: true`
- `loggedIn: true`

둘 중 하나라도 실패하면 조회 명령을 실행하지 않습니다.

## 사람이 해야 하는 일

- Claude/Cowork 도메인 허용 목록에 `graph.microsoft.com`, `login.microsoftonline.com`을 추가합니다.
- PKG 설치본에서 로그인이 필요하면 Applications에서 `OMH M365 Agent Login`을 실행하고 Microsoft device-code 로그인을 직접 완료합니다.
- 폴더형 파일럿 패키지에서 로그인이 필요하면 Finder에서 `로그인 시작.command`를 더블클릭하고 Microsoft device-code 로그인을 직접 완료합니다.
- 폴더형 패키지의 `.command`가 회사 보안 정책상 차단되면 임의로 우회하지 말고 IT 배포 담당자에게 PKG/Jamf/Intune 배포본을 요청합니다.
- device code, token, `.cache` 내용은 채팅에 붙여넣지 않습니다.

## 로그인 안내

`loggedIn: false`이면 사용자에게 이렇게 안내합니다.

```text
PKG 설치본이면 Applications에서 "OMH M365 Agent Login"을 실행하세요.
폴더형 파일럿 패키지이면 Finder에서 Mac 배포 폴더를 열고 "로그인 시작.command"를 실행하세요.
터미널에 표시되는 device code를 Microsoft 로그인 화면에 직접 입력하세요.
코드는 채팅에 붙여넣지 마세요.
로그인이 끝나면 Cowork로 돌아와 계속 진행해달라고 말해주세요.
```

## 조회 명령 예시

```bash
bash run-cowork.sh outlook inbox --limit 1
bash run-cowork.sh teams chats --limit 20
bash run-cowork.sh teams chat-messages --chat-id "<chat-id>" --limit 20
bash run-cowork.sh files search --query "keyword" --limit 10
```

Teams 채팅방 최신성 판단:

- `teams chats` 결과는 `lastMessageCreatedDateTime` 기준으로 정렬됩니다.
- `lastUpdatedDateTime`은 마지막 메시지 시각이 아니라 메타데이터 변경 시각입니다.

## 금지 사항

- `.env`를 읽거나 출력하지 않습니다.
- `.cache/`, token cache, access token, refresh token, cookie, device code를 읽거나 출력하지 않습니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경을 하지 않습니다.
