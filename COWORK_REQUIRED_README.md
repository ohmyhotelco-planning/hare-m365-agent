# Cowork 필수 안내

이 파일은 Claude Cowork에서 OMH M365 Agent를 실행할 때 반드시 먼저 읽어야 하는 문서입니다.

## 목적

이 폴더를 Cowork의 로컬 Linux shell에서 Microsoft 365 읽기 전용 조회 도구로 사용합니다.

Windows 데스크톱 제어, 파일 탐색기 조작, 파일 더블클릭 실행은 사용하지 않습니다. Cowork의 폴더/프로젝트 연결은 허용되며 필요한 절차입니다.

## Cowork가 해야 하는 일

1. 폴더가 Linux shell에 아직 보이지 않으면 Cowork의 폴더/프로젝트 연결 도구로 이 로컬 폴더를 연결합니다.
2. Cowork가 보여주는 mounted Linux path로 이동합니다.
3. 필수 파일을 확인합니다.

```bash
pwd
ls -la
test -f run-cowork.sh
test -f omh-m365.cjs
```

4. 체크섬 파일이 있으면 번들 무결성을 확인합니다.

```bash
sha256sum -c omh-m365.cjs.sha256
node --check omh-m365.cjs
```

5. 모든 CLI 호출은 Cowork wrapper를 사용합니다.

```bash
bash run-cowork.sh doctor
bash run-cowork.sh auth status
```

`run-cowork.sh`는 `NODE_USE_ENV_PROXY=1`을 설정해 Microsoft Login과 Graph 호출이 Cowork host proxy를 타도록 합니다.

## 하드 게이트

Outlook, Teams, Files 명령을 실행하기 전에 반드시 아래 조건이 모두 참이어야 합니다.

- `configured: true`
- `loggedIn: true`

둘 중 하나라도 실패하면 Outlook, Teams, Files 명령을 실행하지 않습니다.

## 사람이 반드시 해야 하는 일

아래 작업은 사람만 할 수 있습니다.

1. Cowork가 폴더 연결 승인을 요청하면 승인합니다.
2. Claude/Cowork 설정의 도메인 허용 목록에 `graph.microsoft.com`과 `login.microsoftonline.com`을 추가합니다.
3. 배포본에는 delegated public-client 설정용 `.env`가 포함됩니다. `.env`가 없거나 미설정일 때만 `.env.example` 기준으로 값을 입력합니다.
4. 로그인이 필요하면 `LOGIN_FOR_COWORK.cmd`를 직접 실행해 Microsoft device-code 로그인을 완료합니다.

사용자에게 device code, token, `.cache` 내용을 채팅에 붙여넣으라고 요구하지 않습니다. `.env`는 LLM이 읽거나 출력할 필요가 없습니다.

## Cowork 도메인 허용 목록

Microsoft 365 조회와 로그인 상태 확인이 동작하려면 Claude/Cowork의 sandbox 네트워크 설정에서 아래 도메인이 허용되어야 합니다.

- `graph.microsoft.com`
- `login.microsoftonline.com`

일반 실행에는 `registry.npmjs.org`가 필요하지 않습니다. 이 릴리스는 번들 파일 `omh-m365.cjs`를 사용하며 Cowork에서 package install을 수행하지 않습니다.

`run-cowork.sh`는 `NODE_USE_ENV_PROXY=1`을 설정해 Cowork host proxy를 사용하게 합니다. 하지만 proxy 사용과 도메인 허용은 별개입니다. `network_error`, proxy 오류, `403`이 발생하면 먼저 wrapper 사용 여부를 확인하고, 그다음 사용자에게 위 두 도메인이 허용 목록에 있는지 확인해달라고 안내합니다.

## 로그인 규칙

Cowork의 one-shot shell은 device-code 로그인에 적합하지 않습니다. WSL은 기술 사용자용 fallback일 뿐, 비개발자 기본 경로가 아닙니다.

`loggedIn: false`이면 사용자에게 아래처럼 안내합니다.

```text
이 폴더를 Windows에서 열고 아래 파일을 실행하세요.

LOGIN_FOR_COWORK.cmd

Microsoft device-code 로그인은 사용자가 직접 완료하세요.
코드를 채팅에 붙여넣지 마세요.

helper에서 loggedIn true가 보이면 Cowork로 돌아와 계속 진행해달라고 말해주세요.
```

LLM은 device code를 읽거나, 반복하거나, 저장하거나, 요약하거나, 스크린샷으로 남기면 안 됩니다.

기술 사용자용 fallback:

```bash
cd /mnt/c/Users/OMH/Documents/GitHub/omh-m365-agent/releases/linux-llm
node omh-m365.cjs auth login
node omh-m365.cjs auth status
```

## 조회 명령 예시

`configured: true`와 `loggedIn: true`가 확인된 뒤에만 실행합니다.

```bash
bash run-cowork.sh outlook inbox --limit 1
bash run-cowork.sh teams chats --limit 20
bash run-cowork.sh teams chat-messages --chat-id "<chat-id>" --limit 20
bash run-cowork.sh files search --query "keyword" --limit 10
```

Teams 채팅방 최신성 판단:

- `teams chats` 결과는 Microsoft Graph의 `lastMessagePreview/createdDateTime` 기준으로 정렬됩니다.
- 최신 채팅방을 고를 때는 `lastMessageCreatedDateTime`을 사용합니다.
- `lastUpdatedDateTime`은 채팅방 메타데이터 변경 시각이므로 마지막 메시지 시각으로 사용하지 않습니다.

## 실패 시 처리

- 폴더가 마운트되지 않음: Cowork의 폴더/프로젝트 연결 도구를 사용합니다. 그래도 안 되면 Cowork shell에서 폴더 접근이 불가능하다고 보고합니다.
- 체크섬 또는 `node --check` 실패: 마운트된 배포 파일이 손상되었거나 불완전하다고 보고하고 중단합니다.
- `configured: false`: 배포본의 `.env` 누락 또는 미설정 상태입니다. 사용자가 `.env.example` 기준으로 `.env`를 로컬에서 작성해야 한다고 안내합니다.
- `loggedIn: false`: 위 로그인 규칙을 안내합니다.
- `network_error`, proxy, `403`: 먼저 `bash run-cowork.sh`를 사용했는지 확인합니다. 그래도 실패하면 Claude/Cowork 도메인 허용 목록에 `graph.microsoft.com`, `login.microsoftonline.com`이 있는지 사용자에게 확인 요청합니다.

## 금지 사항

- `.env`를 읽거나 출력하지 않습니다.
- `.cache/`, token cache, access token, refresh token, cookie, device code를 읽거나 출력하지 않습니다.
- 메일 발송, Teams 게시, 일정 생성, 파일 업로드/삭제/공유, 권한 변경을 하지 않습니다.
- 이 폴더, 다운로드 파일, 토큰 캐시, credential을 다른 환경으로 업로드하지 않습니다.
