# M365 그룹메일 직접 조회

로그인한 사용자에게 허용된 Microsoft 365 그룹의 본문과 첨부를 읽기 전용으로 조회한다.

## 명령

공통으로 기존 절대 `--data-dir`를 유지한다. 그룹은 공유 사서함이 아니므로 `outlook --mailbox` 대신 다음 명령을 사용한다.

```sh
hare-m365 --data-dir <existing-data-dir> groups resolve --group team@example.test
hare-m365 --data-dir <existing-data-dir> groups posts --group team@example.test --since 2026-10-01 --until 2026-10-15 --limit 20
hare-m365 --data-dir <existing-data-dir> groups posts --group team@example.test --since 2026-10-01 --until 2026-10-15 --cursor <nextCursor>
hare-m365 --data-dir <existing-data-dir> groups attachments list --group <group-id> --thread-id <thread-id> --post-id <post-id>
hare-m365 --data-dir <existing-data-dir> groups attachments download --group <group-id> --thread-id <thread-id> --post-id <post-id> --attachment-id <attachment-id>
```

`resolve`는 정확한 기본 SMTP·SMTP proxy 별칭·mailNickname 또는 디렉터리 ID를 확인한다. 표시 이름 유사 검색은 하지 않는다. 실제 `id`, `mail`, `groupTypes`, `mailEnabled`, `securityEnabled`, 판정 근거를 반환한다. `Unified`가 확인된 경우만 M365 그룹 원문을 읽는다. 0건·모호함·디렉터리 거부·유형 미확인을 구분하며 상세 Exchange 수신자 유형은 `unverified`다. 디렉터리에 없는 동적 배포 그룹까지 탐지하거나 사용자에게 연결된 모든 사서함을 열거하는 기능은 아니다.

## 원문과 완전성

- conversation → thread → post 순회로 관계를 보존한다. `bodyContent`와 `bodyContentType`은 post 원문이며 `bodyHtml`은 HTML 원문, `body`는 읽기용 텍스트다. raw MIME 메일 전체를 뜻하지 않는다.
- post의 `conversationId`가 `null`이면 요청한 부모 conversation 경로를 결과에 유지한다. 값이 있는 conversation ID나 thread ID가 요청 경로와 다르면 계속 오류로 중단한다.
- 실제 `createdDateTime` 기준으로 날짜를 비교한다. 기간 미지정은 최근 90일·Asia/Seoul이며 `range`를 반드시 보고한다. 대화 preview·lastDeliveredDateTime은 본문·post 생성시각이 아니다.
- 문서에 보장되지 않은 post 날짜 필터·최신순 정렬에 의존하지 않는다. 오래된 post를 만났다고 순회를 조기 종료하지 않는다. 정렬된 최근 N건 명령이나 정확한 전체 건수 명령이 아니다.
- 기본 호출 예산은 20초·100 요청·1,000 레코드, 반환은 기본 20건·최대 100건이다. `page.continuationAvailable`일 때 같은 그룹과 반환된 명시 날짜로 `nextCursor`를 사용한다. 커서는 본문을 포함하지 않지만 내부 식별자가 있으므로 공개하지 않는다. 그룹 전용 커서는 전체 진행 상태를 서명하며 24시간 동안 같은 데이터 폴더에서만 유효하다. 로컬 서명 키는 처음 이어받기가 생길 때 `.cache`에 생성되고 M365 인증과 별개다. 키·커서를 수정하거나 인증 파일로 대체하지 않는다.
- 결과 한도·시간/스캔 예산·부분 실패·원문/시각 누락·무진행을 명시한다. `complete=false`를 전체 조회로 표현하지 않는다. 누락 수는 호출별/누적으로 구분한다. 중복 제거 이력 10,000개 한도나 무진행이면 중단하고 무한 재시도하지 않는다.
- live API는 스냅샷이 아니다. 이어받는 페이지의 ID 구성이 달라지면 `CURSOR_STALE`로 중단한다. 미조회 페이지 변경까지 탐지하거나 과거 시점의 완전성을 보장하지 않는다.

## 권한과 실패 단계

2026-10-16 공식 v1.0 문서 확인 기준이다. 아래는 읽기 권한만 표기한다. Hare 인증은 그룹 본문·첨부 직접 조회를 위해 `Group.Read.All` 위임 scope를 요청한다. 해당 권한의 관리자 동의는 Entra에서 별도로 완료해야 하며 Hare가 자동으로 동의하거나 앱 권한을 변경하지 않는다. 기존 앱·테넌트·계정이 일치하면 인증 캐시를 보존하고 새 scope로 silent 취득을 먼저 시도한다. 추가 상호작용이 필요한 경우에만 로그인 절차를 안내한다. 현재 로그인으로 거부될 경우 구현 테스트와 실제 권한 검증 상태를 분리해 보고한다.

전체 Graph 조회가 같은 scope 검증을 공유하므로 새 권한 취득에 실패하면 기존 Outlook·Teams 조회도 차단될 수 있다. 적용 후 그룹 조회와 기존 조회를 함께 확인한다. 변경 전 scope로 시작한 미완료 device login은 이어받을 수 없으며 기존 로그인 상태를 임의 삭제하지 않는다.

| API | Delegated 회사 계정 | Application |
|---|---|---|
| `/groups`, `/groups/{id}` | GroupMember.Read.All / Group.ReadBasic.All / Group.Read.All 등 문서상 읽기 대안. 좁은 권한의 속성 노출은 실제 응답 확인 필요 | 문서상 읽기 대안 지원 |
| conversations 목록, threads 목록, posts 목록·단건 | Group-Conversation.Read.All 또는 Group.Read.All | 지원하지만 Hare는 delegated 전용 |
| post attachments 목록 | Group-Conversation.Read.All 또는 Group.Read.All | 미지원 |
| post attachment 단건 메타데이터·`/$value` | Group.Read.All | 미지원 |

목록 권한으로 단건 첨부 바이트까지 읽힌다고 추정하지 않는다. 디렉터리·conversation·thread·post·첨부 목록·메타데이터·바이트 실패를 별도 코드로 반환한다. 404는 해당 연결 경로를 찾지 못했거나 접근 불가이지 그룹의 절대 부재 증명이 아니다. 실패 뒤 `/me/messages`, 다른 커넥터나 UI로 대체하지 않는다. 현 문서의 디렉터리 최소 권한 표에는 쓰기 scope도 등장하므로 자동으로 채택하지 않는다.

## 첨부 안전 기준

group/thread/post/attachment ID와 출력명에 승인을 묶는다. fileAttachment는 원본 파일, itemAttachment는 MIME/vCard/iCal로 다운로드하며 referenceAttachment는 지원하지 않는다. 기존 기본 100MiB·승인 시 1GiB 상한(실제 적용 값은 policy 기준), 다운로드 폴더, 파일명 정리, 스트리밍 크기 검사, 기존 파일 덮어쓰기 금지를 재사용한다. 정책을 상향하지 않는다.

첨부 목록은 `list.partial`, `stopReason`, `continuationAvailable`, `nextCursor`를 확인한다. 이어받을 때 그룹·thread·post와 `--limit`도 그대로 유지한다. 뒤 페이지 조회 실패 시 확보한 메타데이터와 실패 단계를 함께 반환하며 첫 페이지 접근 자체가 거부되면 오류로 중단한다.

메타데이터 조회 성공, 파일 바이트 수신, 파일 형식 구조 검증은 서로 다르다. `downloadedBytes`는 실제 저장 크기이고 메타데이터 size와 같다고 가정하지 않는다. `structureValidation=NOT_PERFORMED`이며 별도 검증 없이 XLSX/ZIP 등의 내부 구조가 정상이라고 보고하지 않는다.

## 공식 근거

- [그룹 목록](https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0), [그룹 속성](https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0)
- [대화 목록](https://learn.microsoft.com/en-us/graph/api/group-list-conversations?view=graph-rest-1.0), [스레드 목록](https://learn.microsoft.com/en-us/graph/api/conversation-list-threads?view=graph-rest-1.0), [post 목록](https://learn.microsoft.com/en-us/graph/api/conversationthread-list-posts?view=graph-rest-1.0)
- [post 첨부 목록](https://learn.microsoft.com/en-us/graph/api/post-list-attachments?view=graph-rest-1.0), [첨부 메타데이터·원시 바이트](https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0)
