# M365 그룹메일 직접 조회

로그인한 사용자에게 허용된 Microsoft 365 그룹의 본문과 첨부를 읽기 전용으로 조회한다.

## 명령

공통으로 기존 절대 `--data-dir`를 유지한다. 그룹은 공유 사서함이 아니므로 `outlook --mailbox` 대신 다음 명령을 사용한다.

```sh
hare-m365 --data-dir <existing-data-dir> groups resolve --group team@example.test
hare-m365 --data-dir <existing-data-dir> groups posts --group team@example.test --since 2026-10-01 --until 2026-10-15 --limit 20 --out group-page-1.json
hare-m365 --data-dir <existing-data-dir> groups posts --group team@example.test --since 2026-10-01 --until 2026-10-15 --cursor-file <previous-outputPath> --out group-page-2.json
hare-m365 --data-dir <existing-data-dir> groups attachments list --group <group-id> --thread-id <thread-id> --post-id <post-id>
hare-m365 --data-dir <existing-data-dir> groups attachments download --group <group-id> --thread-id <thread-id> --post-id <post-id> --attachment-id <attachment-id>
```

`resolve`는 정확한 기본 SMTP·SMTP proxy 별칭·mailNickname 또는 디렉터리 ID를 확인한다. 표시 이름 유사 검색은 하지 않는다. 실제 `id`, `mail`, `groupTypes`, `mailEnabled`, `securityEnabled`, 판정 근거를 반환한다. `Unified`가 확인된 경우만 M365 그룹 원문을 읽는다. 0건·모호함·디렉터리 거부·유형 미확인을 구분하며 상세 Exchange 수신자 유형은 `unverified`다. 디렉터리에 없는 동적 배포 그룹까지 탐지하거나 사용자에게 연결된 모든 사서함을 열거하는 기능은 아니다.

`--cursor-file`은 이전 `groups posts --out`의 `outputPath`에 저장된 JSON에서 `page.nextCursor`를 읽는다. 커서 문자열만 담은 UTF-8 파일도 지원한다. 파일은 최대 32MiB이며 BOM·끝 줄바꿈을 허용한다. `--cursor`와 함께 쓸 수 없고, 서명·만료·데이터 폴더·그룹·기간 검증을 우회하지 않는다. Windows 명령줄 길이 제한에 걸리지 않도록 파일 경로만 전달하며 원문 커서를 대화·로그에 복사하지 않는다. `continuationAvailable=false`이면 이어받지 않는다.

## 원문과 완전성

- conversation → thread → post 순회로 관계를 보존한다. `bodyContent`와 `bodyContentType`은 post 원문이며 `bodyHtml`은 HTML 원문, `body`는 읽기용 텍스트다. raw MIME 메일 전체를 뜻하지 않는다.
- post의 `conversationId`가 `null`이면 요청한 부모 conversation 경로를 결과에 유지한다. 값이 있는 conversation ID가 다르면 중단한다. thread ID가 다르면 같은 그룹의 반환 스레드 post 목록에서 동일 post ID·정확한 반환 thread ID·충돌 없는 conversation ID를 재검증한다. 단건 조회 성공만으로 ID가 동등하다고 추정하지 않는다.
- 재검증은 최대 3페이지이며 기존 시간·요청·스캔 예산 안에서 수행한다. 성공한 경우 검증 목록의 원문을 사용하고 `threadId`에는 확인된 첨부 조회용 ID, `traversalThreadId`에는 원래 순회 ID, `threadVerification`에는 `POST_COLLECTION_REVALIDATED`를 반환한다. 원래 경로로 조회된 일반 결과는 `REQUEST_PATH`다. 원래 conversation 경로는 순회 출처이며 두 thread ID가 동등하다는 뜻은 아니다.
- 연결 불일치나 비재시도 재검증 실패는 확보한 결과와 `page.failure`를 반환하고 `complete=false`, `continuationAvailable=false`로 중단한다. 확인된 일시 오류만 아래 제한적 재개 기준을 적용한다. 실패 post는 결과에 포함하지 않으며 첨부 연결 검증을 완화하지 않는다. 예산 소진 시 미검증 post를 건너뛰지 않지만, 이어받기에서는 그 post의 재검증을 첫 페이지부터 다시 수행한다. 한 호출 예산 안에 검증하지 못하고 같은 위치에 머무르면 `NO_PROGRESS`로 중단한다.
- 같은 그룹의 동일 post는 순회 thread가 달라도 한 번만 반환한다. 이 중복 제거 기준 변경으로 이전 버전의 posts 커서는 거부되므로 새 조회를 시작해야 한다. 첨부 목록 커서는 변경하지 않는다.
- 실제 `createdDateTime` 기준으로 날짜를 비교한다. 기간 미지정은 최근 90일·Asia/Seoul이며 `range`를 반드시 보고한다. 대화 preview·lastDeliveredDateTime은 본문·post 생성시각이 아니다.
- 문서에 보장되지 않은 post 날짜 필터·최신순 정렬에 의존하지 않는다. 오래된 post를 만났다고 순회를 조기 종료하지 않는다. 정렬된 최근 N건 명령이나 정확한 전체 건수 명령이 아니다.
- 기본 호출 예산은 20초·100 요청·1,000 레코드, 반환은 기본 20건·최대 100건이다. `page.continuationAvailable`일 때 같은 그룹과 반환된 명시 날짜로 `nextCursor`를 사용한다. 커서는 본문을 포함하지 않지만 내부 식별자가 있으므로 공개하지 않는다. 그룹 전용 커서는 전체 진행 상태를 서명하며 24시간 동안 같은 데이터 폴더에서만 유효하다. 로컬 서명 키는 처음 이어받기가 생길 때 `.cache`에 생성되고 M365 인증과 별개다. 키·커서를 수정하거나 인증 파일로 대체하지 않는다.
- 결과 한도·시간/스캔 예산·부분 실패·원문/시각 누락·무진행을 명시한다. `complete=false`를 전체 조회로 표현하지 않는다. 누락 수는 호출별/누적으로 구분한다. 이력은 아래 로컬 체크포인트로 보존하며 `STATE_CAPACITY_LIMIT`이나 무진행이면 중단한다.
- 읽은 페이지의 순서 있는 ID 목록과 다음 페이지 링크를 서명된 커서에 고정한다. 이어받을 때 변경된 부모 페이지를 다시 읽고 이전 offset을 적용하지 않는다. 본문은 커서에 저장하지 않고, 미처리 post는 정확한 ID로 다시 읽어 기존 연결 검증을 적용한다. 대상이 사라졌거나 연결을 확인할 수 없으면 조용히 건너뛰지 않고 부분 실패로 표시한다.
- 큰 이력 또는 inline 커서 크기 상한에 도달하면 같은 dataDir의 로컬 체크포인트를 가리키는 짧은 v4 커서를 반환한다. v3 커서도 검증 후 재개할 수 있다. 로컬 저장 실패 시 새 결과/커서를 발행하지 않으며 이전 성공 결과를 보존한다.
- live API는 스냅샷이 아니다. 이후 페이지는 여전히 실시간 API에서 받아오므로 이미 읽은 페이지에 새 항목이 추가되거나 미조회 페이지 경계가 바뀐 경우까지 완전 수집을 보장하지 않는다. 처리 완료는 확보한 ID와 이어받은 페이지의 순회 완료를 뜻하며 특정 시점 전체 그룹의 정확한 건수 증명이 아니다. posts 커서 v3 이전의 커서는 새 조회로 시작해야 한다. 첨부 목록 커서 형식은 그대로 유지한다.

## 순회 상태와 제한적 재개

- 이력 합계 8,192개 이상 또는 inline 크기 초과 시 `.cache/group-post-checkpoints`에 1,024개 단위의 불변 해시 청크와 진행 상태를 저장한다. 본문·제목은 저장하지 않지만 내부 식별자가 있으므로 공개하지 않는다. 커서만 다른 dataDir로 복사해서 재개할 수 없다.
- 파일 해시, 커서 서명·만료·질의 조건과 부모/자식 연결을 모두 검증한다. 임시 파일을 완전히 쓰고 검증한 뒤 게시하며, 실패한 임시 파일도 용량에 포함한다. 자동 삭제·인증 캐시 정리·공유 저장소 잠금은 하지 않는다.
- 한 dataDir의 체크포인트 쓰기는 순차 실행한다. 항상 최신 성공 결과의 커서로 한 번씩 이어간다. 옛 커서 재사용은 그 시점 결과를 다시 반환할 수 있으며 전역 중복 실행 방지나 여러 작성자의 동시 용량 보장을 제공하지 않는다.
- seen(대화·스레드·post 식별자 합계)와 visited pages는 각각 최대 500,000개, 객체당 16MiB, 저장소 합계 256MiB로 제한한다. 이력은 호출 시 메모리에 복원하므로 무제한·고정 메모리 수집이 아니다. 한도는 메일 개수와 같지 않다. 저장 실패 시 이전 커서를 보존하고 운영 담당자에게 확인한다.
- `READ_FAILED`의 `failureDetails`는 단계·정제된 오류 코드·연속 실패 횟수·재개 가능 여부·대기 시간을 제공한다. HTTP 429/500/502/503/504, 명시적 요청 시간 예산 소진, 일부 일시 전송 오류만 재개 대상이다. `continuationAvailable=true`와 `retryAllowed=true`가 모두 있을 때 `retryAfterMs` 이후 반환된 새 커서로만 재개한다. 같은 위치에서 세 번째 연속 실패 시 `TRANSIENT_RETRY_LIMIT`으로 중단한다. 옛 커서 반복으로 이 제한을 우회하지 않는다.
- 인증·권한·식별자·알 수 없는 실패를 일시 오류로 추정하지 않는다. 기존 본문 보충 조회의 비일시적 실패는 `bodyFailure`/본문·시각 누락으로 명시하며 원문 확보 성공으로 계산하지 않는다. 알 수 없는 본문 보충 실패는 미처리 post를 소비하지 않고 중단한다. 제목 미확보 처리는 아래 기준을 유지한다.

## 개별 메일 제목 (선택 조회)

`groups posts --include-subject`는 반환할 post마다 검증된 최종 group/thread/post 경로에서 `singleValueExtendedProperties($filter=id eq 'String 0x0037')`를 읽는다. 기존 `Group.Read.All` 위임 권한을 사용하며 권한·인증 설정은 바꾸지 않는다. 기본 조회는 요청 수와 결과 형식이 그대로다.

- `postSubject.source=PidTagSubject`, `scope=POST`는 개별 메시지의 전체 제목 속성을 뜻한다. `value`는 접두사와 공백을 포함한 원문이다. thread/conversation `topic`과 본문은 제목 대체값으로 쓰지 않는다.
- 상태는 `PRESENT`(문자열 있음), `EMPTY`(확인된 빈 문자열), `MISSING`(확장 목록에 속성 없음), `UNVERIFIED`(확장 목록 미반환·불완전), `INVALID_RESPONSE`(잘못된 형식·중복 속성), `READ_FAILED`(API 실패), `NOT_ATTEMPTED_BUDGET`(예산 소진)을 구분한다. `PRESENT` 이외 상태로 제목 조건 충족을 추정하지 않는다.
- 추가 GET도 기존 시간·요청 예산에 포함한다. 제목 API 실패·예산 소진은 확보된 본문을 보존하고 상태를 반환한다. 단, 제목 응답의 post/thread/conversation 연결이 충돌하면 해당 post를 반환하지 않고 비재개 오류로 중단한다. 재검증된 thread는 제목 응답에서도 명시적 일치를 요구한다.
- `page.complete`는 기존 순회·본문 기준이며 제목 확보 성공을 뜻하지 않는다. 별도의 `subjectRead`는 이번 호출 반환 post의 제목 확보/미확보 수다. 제목 미확보 post도 반환·처리된 것이므로 일반 continuation에서 재시도하지 않는다.
- v3 커서에는 제목·옵션을 넣지 않는다. `--cursor-file` 재개 시에도 `--include-subject`를 명시해야 한다. 옵션을 바꾸면 이후 반환 post에만 적용되며 과거 결과를 보충하지 않는다. 제목이 필요한 분석은 옵션을 켜고 새 조회부터 시작한다.

공식 근거: [확장 속성 읽기](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0), [PidTagSubject 정의](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagsubject-canonical-property). 실제 그룹 post의 속성 반환 여부는 환경별로 확인하며 문서 지원만으로 제목 확보 성공을 단정하지 않는다.

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
