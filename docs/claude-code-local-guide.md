# Claude Code Desktop Local 연결 가이드

## 범위와 상태

- 설치 명령은 운영 배포 원본인 `master`를 가져온다. Claude Code 지원 코드와 가이드를 함께 병합하는 배포이며, 최종 사용자에게 배포 승인 여부를 다시 묻지 않는다. 환경 준비 프롬프트는 읽기 전용 점검을 바로 시작하고 개별 프로그램 설치·업그레이드에는 별도 승인을 받는다.
- 대상: Windows의 PowerShell 및 Mac의 POSIX 셸을 사용하는 Claude Desktop > Code > Local.
- 제공물: `release-templates/claude-code/`의 한국어, 영어, 일본어 독립형 HTML 세 장. 외부 스크립트, 외부 이미지, 복제한 스크린샷 없이 브라우저에서 열 수 있다.
- 실제 Windows/Mac Claude Desktop 수용 검증은 **미완료**다. 정적 가이드 테스트 통과를 실제 Desktop 설치, 인증, Mac 실행 성공으로 간주하지 않는다.
- 이 가이드는 Desktop Local 연결 절차에 한정한다. Cloud, SSH, WSL 및 Cowork 연결 절차는 대상이 아니다.

## 사용자 흐름

환경 준비와 Hare 연결을 순서대로 진행한다. 환경 준비는 Hare 폴더 선택을 선행 조건으로 삼지 않는다.

1. Claude Desktop의 Code 탭에서 Local을 선택한다. 제공된 호스트·실행 환경 정보로 확인하며 판별할 정보가 없을 때만 사용자에게 한 번 확인한다. OS나 경로만으로 Local이라고 추정하지 않는다. 앱 자체가 작업 폴더를 요구하는 경우에는 사용자가 로컬 폴더를 선택한다.
2. HTML 상단에서 Windows/Mac을 선택하고 **환경 준비 프롬프트**를 Code Local 채팅에 붙여넣는다. Claude가 Node.js `>=20.18.1`, npm, Git을 확인한다. 새 Node 설치에는 현재 지원되는 LTS를 사용하고 npm은 함께 설치한다. 이미 사용 가능한 도구는 재설치하지 않는다. 부족한 항목만 이유·공식 출처·설치 방법·변경 범위·관리자 권한 필요 여부를 설명하고 명시적 승인 후 설치한다. 설치창이나 관리자 승인이 필요하면 사용자에게 안내하고 기다리며, 정책 차단 시 IT 담당자에게 문의하도록 안내한다. 실제 설치는 이번 문서 작업에서 수행하지 않았다.
3. 설치 후 Claude를 완전히 종료했다가 Local로 다시 열고 환경 준비 프롬프트로 재확인한다. 기존 선택 폴더가 있다면 유지한다. 준비 완료 후 **기존 Hare 데이터 폴더 자체**를 선택하고 **Hare 연결 프롬프트**를 붙여넣는다. 처음 사용하는 경우에만 Git 저장소·OneDrive 밖의 전용 로컬 폴더를 사용자가 선택한다. 두 프롬프트 모두 터미널 명령이 아니다. 연결 단계에서 AI는 `<SELECTED_LOCAL_FOLDER>`를 선택한 정확한 절대 경로로 셸에 맞게 이스케이프하여 치환한다. 폴더가 없거나 치환이 불확실하면 실행하지 않는다.
4. Windows의 명령과 이후 `setup.nextCommand`는 PowerShell에서 실행한다. `&` 호출 연산자를 포함한 PowerShell을 Bash에 직접 넘기지 않는다. 네이티브 PowerShell 도구 또는 안전한 인수 전달로 PowerShell 실행 파일을 사용하며, Bash가 본문을 확장하지 않게 한다. Mac은 POSIX 셸을 사용한다.
5. 설치 명령은 startup까지 처리한다. 반환 후 추가 인증 명령이나 업무 조회 **전에** 선택 폴더의 생성된 `CLAUDE.md`와 `claude/hare-m365-agent-rules.md`를 읽는다. 원시 캐시는 읽지 않는다. 규칙이 없거나 호스트, 셸, 경로가 맞지 않으면 중단한다.
6. `setup.state`와 `setup.nextCommand`의 한 단계만 따른다. 대기 로그인은 사용자 완료 확인 후 완료 명령을 정확히 한 번 실행한다. `READY`이면 중단하고 업무 요청을 기다린다. `loggedIn=true`와 `tokenUsable=true`를 모두 확인하기 전에는 조회하지 않는다.
7. 다음 채팅에서도 **동일 폴더**를 연다. 기본 설정에서 자동 로드되는 `CLAUDE.md`의 규칙 연결을 확인하고 생성된 상태 확인 명령을 사용한다. 새 채팅만을 이유로 다시 설치하거나 로그인하지 않는다. 관리 설정 때문에 규칙 적용이 불확실하면 먼저 확인한다.

## 안전 경계

- 환경 준비와 Hare 연결을 분리한다. 환경 준비에서는 Hare clone/build, Microsoft 로그인, 업무 조회, 기존 Hare 데이터·인증 캐시 접근을 하지 않는다. 준비 완료는 로그인 성공이 아니다. 설치 전 명시적 사용자 승인과 호스트의 도구 승인은 각각 필요하며 프롬프트는 OS 보안 통제를 대신하지 않는다.
- 수동 설치 도움말은 접힌 상태로 제공한다. Windows는 Git 오류가 실제로 나온 경우에 Git for Windows를 안내하며, 모든 최신 Desktop Local 세션에 Git Bash가 선행 필수라고 단정하지 않는다. Hare 자체의 clone에는 Git이 필요하다. Mac은 필요한 경우 Apple Command Line Tools를 사용한다. Node.js 공식 LTS 설치 프로그램은 npm을 포함한다. 전체 Xcode, Homebrew, Claude CLI, WSL, Docker를 일괄 설치하도록 안내하지 않는다.
- 설치 권한 거부·정책 차단·도구 확인 불가를 미설치로 오인하지 않는다. 사용자 암호를 채팅으로 받지 않고 PATH/실행 정책을 임의 변경하지 않는다. Windows의 npm.ps1 실행 정책 오류는 설치된 npm.cmd 확인으로 구분한다. 재시작 후에도 문제가 있으면 재설치 반복 없이 중단한다.
- Node가 최소 버전 미만이면 `업그레이드 필요`이며 신규 설치와 같은 사전 설명·승인 절차를 적용한다. 요구 버전을 충족하는 정상 도구만 그대로 유지한다. 새 설치와 업그레이드에는 현재 지원되는 LTS를 안내한다.
- Node 20.x는 지원 종료를 안내하고 지원되는 LTS 업그레이드를 권고한다. 사용자가 거절해도 기술적 최소 버전을 충족하면 경고와 함께 진행하며 지원 중인 버전으로 표현하지 않는다. Git은 버전 및 `git branch -h`의 `--show-current` 지원을 확인한다. 도움말의 종료 코드만으로 실패라 하지 않으며 필요한 옵션이 없을 때 업그레이드를 제안한다.
- 새 세션에서 명령을 찾지 못해도 표준 설치 위치·OS 설치 기록·이미 확인된 버전 관리자 위치를 제한적으로 확인한다. 확인된 실행 파일의 절대 경로로 버전이 조회되면 `설치됨·PATH 미반영`으로 구분하고 재설치하지 않는다. 표준 위치에 없다는 것만으로 미설치라 하지 않는다. 상태가 불명확하면 사용자에게 최근 설치 여부나 별도 위치를 확인하고 중단한다. 전체 디스크 검색과 인증 캐시 접근은 하지 않는다.
- 사내 허용은 2026-10-01 사용자 확인에 따라 가이드에 명시한다: 이 Hare 준비에서 공식 Node.js·Git 설치 프로그램, Apple Command Line Tools, Windows의 기존 `winget` 이용을 허용하며 직원에게 정책을 다시 확인시키지 않는다. 개별 설치에는 패키지 ID·배포자·출처·버전·변경 범위를 설명하고 사용자 승인을 받는다. 누락·차단된 winget의 설치·복구나 정책 우회는 하지 않고 공식 설치 프로그램을 수동 안내한다. 해당 설치도 실제 차단되면 중단한다. UAC는 사용자가 직접 처리한다. 일괄 업데이트·패키지 소스 추가/변경은 범위 밖이다.
- winget 조회는 `OpenJS.NodeJS.LTS`와 `Git.Git`, `--exact --source winget`으로 대상을 명확히 한다. 승인 후 `--interactive`로 설치 옵션을 사용자가 확인하고 Node는 npm 포함 MSI를 선택한다. portable ZIP, `--silent`, 임의 `--override`는 사용하지 않는다. 이 작업에서 프로그램을 실제 설치한 것은 아니다.
- `설치됨·PATH 미반영`은 `ENVIRONMENT_INCOMPLETE`이며 Hare 연결로 진행하지 않는다. Claude 완전 종료·Local 재시작을 한 번 안내하고 일반 명령으로 다시 확인한다. 새 세션에서는 이미 같은 문제로 재시작했는지 확인해 반복하지 않는다. 계속 실패하면 도구·오류만 보고하고 IT 확인을 요청한다. 재설치, 별도 PATH 편집, 절대 경로를 통한 연결 우회는 하지 않는다.
- 기존 도구는 실제 실행/링크 경로, 요구 버전, 명령 성공으로 준비 상태를 판단한다. 표준 위치·OS 설치 기록·기존 버전 관리자 위치를 모두 인정하며, 기존 도구의 서명은 선택적 진단이다. 서명·설치 기록 부족이나 서명자 차이만으로 정상 도구를 차단하지 않는다. 이 PC에서 확인한 Windows 서명자(Node.js: OpenJS Foundation, Git: Johannes Schindelin)는 참고값이지 고정 허용 목록이 아니다. Node와 같은 설치의 npm은 그대로 쓰고, 전역 업데이트로 다른 npm이 선택돼도 위치 차이만 보고한다. npm 체인 분석과 패키지 재다운로드·압축 해제·해시 대조는 하지 않는다. 이는 사용자가 승인한 실행 편의와 검증 비용의 절충이며 기존 도구의 무결성을 보증하지 않는다.
- 새 설치 파일에는 공식 출처·유효 서명·공식 배포자 확인과 사용자 승인을 유지한다. Windows는 `Get-AuthenticodeSignature -LiteralPath`의 `Valid`, Mac `.pkg`는 `pkgutil --check-signature`로 확인하며 검증 실패나 실제 OS 차단을 우회하지 않는다. 기존 도구에 대한 선택적 진단과 새 설치 파일의 실행 조건을 구분한다.
- Windows Node 설치의 `Tools for Native Modules` 추가 도구는 선택하지 않는다. Git의 PATH 선택은 `Git from the command line and also from 3rd-party software`이며, 설치 프로그램이 수행할 PATH 변경을 승인 전 설명한다. 별도 임의 PATH 편집과 구분한다. 수동 안내에도 기존 설치 여부 확인과 재설치 방지 안내를 적용한다.
- 2026-10-01 사용자 승인에 따라 DRAFT 차단을 두 프롬프트에서 제거했다. 환경 준비는 점검 실행 요청이며, Hare 연결은 clone·의존성 설치·빌드·규칙 생성 요청이다. Microsoft 로그인 완료 확인, 실제 업무 요청 및 상태 변경 승인은 각각 유지한다. 코드를 `master`에 병합한 뒤 이 가이드를 제공하며, 아래 실기기 검증의 미완료 항목을 완료로 간주하지 않는다.
- 모든 후속 명령은 생성된 정확한 명령 접두사와 `--data-dir`, `--host claude-code`, `--command-shell powershell|posix`를 유지한다.
- 권한은 필요한 작업에 한정하여 표준 도구 승인 절차를 사용한다. 거부, 승인 도구 부재, 정책 차단 시 중단한다. 일괄 권한 우회나 보안 정책 변경을 안내하지 않는다.
- Local 네트워크 정책을 Cowork 허용 목록과 같다고 가정하지 않는다. 생성 규칙의 제한된 `network check --environment claude-code` 절차만 따른다. 인증 확인 불가를 로그아웃으로 해석하지 않으며 기존 인증 폴더와 캐시를 보존한다. 실패나 거부를 반복 재시도, 캐시 삭제, 재로그인으로 해결하지 않는다.
- M365 조회 및 Outlook 초안에는 Hare만 사용한다. 커넥터, Computer Use, GUI, 브라우저 자동화로 대체하지 않는다. 사용자가 Microsoft 로그인 페이지에서 직접 로그인하는 단계는 별개다.
- 초안은 수신자, 제목, 본문, 첨부 전체 미리보기와 명시적 승인 후에만 생성한다. 발송은 지원하지 않는다. 다운로드 및 내보내기의 조건부 승인, 전체 계획, 토큰, 제한, 재개 동작은 생성 규칙을 그대로 따르며 무조건 허용이나 무조건 재승인으로 바꾸지 않는다.
- Local은 실행 위치다. Hare 결과를 Claude가 읽으면 업무 내용은 모델에 전달된다. 비밀값과 원시 인증 캐시는 프롬프트, 출력, 테스트 증적에 포함하지 않는다.

## 생성과 정적 검증

생성기는 빌드된 `../dist/local-install.js`의 `buildLocalSetupCommand`를 직접 사용한다. 명령을 별도로 복제하거나 빌드 버전, SHA를 고정하지 않는다.

```js
buildLocalSetupCommand({
  dataDir: '<SELECTED_LOCAL_FOLDER>',
  repository: 'https://github.com/ohmyhotelco-planning/hare-m365-agent.git',
  branch: 'master',
  environment: 'claude-code',
  shell: 'powershell' // Mac: 'posix'
});
```

선행 조건: 프로젝트 의존성이 준비된 개발 환경에서 `npm run build`로 현재 소스의 `dist/local-install.js`를 빌드한다. 필요한 호스트/셸/환경 플래그가 빠진 이전 API는 생성기가 거부하며, 모든 언어 렌더링에 성공하기 전에는 HTML을 쓰지 않는다.

저장소 루트에서 다음 순서로 실행한다. 아래 명령은 생성된 설치 명령을 실행하지 않는다.

```text
node --check scripts/build-claude-code-guides.mjs
node --check test/claude-code-guides.test.mjs
npm run build
node scripts/build-claude-code-guides.mjs
node scripts/build-claude-code-guides.mjs --check
node --test test/claude-code-guides.test.mjs
```

- 일반 생성은 세 HTML만 쓴다. `--check`는 읽기 전용이며 바이트 차이, 누락, 이전 API에서 실패한다.
- 테스트는 3개 언어 x 2개 셸의 명령 동일성, 자리표시자 이스케이프, 호스트/셸 보존, 상태·권한·폴더·캐시·승인 규칙을 확인한다.
- 독립 실행 가능한 HTML 구조, 외부 자산 부재, 언어/접근성 속성, 선택과 복사 대상의 일치, 클립보드 성공/거부/미지원 및 수동 복사 대체 동작을 검사한다.
- 환경 준비 보완 후 가이드 테스트 43건 통과: 언어별 승인 경계, 구버전 업그레이드·설치 흔적 확인·winget 허용·설치 옵션 문구, 재시작 1회·미완료 판정·OS별 출처 기준, 단계 분리, 두 프롬프트의 독립 복사 및 선택 대체 경로를 확인했다. 설치 프로그램 실행 또는 LLM의 실제 지침 준수를 증명하는 테스트는 아니다.
- VM 기반 UI 테스트는 브라우저나 실제 Desktop 수용 검증을 대체하지 않는다. 2026-10-01 헤드리스 Edge에서 3개 언어 × 2개 화면 너비(1440/390px) × 2개 OS 선택의 12조합을 확인했다. 가로 넘침과 페이지 오류가 없고 선택된 프롬프트와 복사 API에 전달된 본문이 일치했다. 클립보드 API는 모의 처리했으므로 실제 OS 클립보드 권한 허용까지 검증한 것은 아니다.

## 실제 Desktop 수용 계획: 실행 대기

2026-10-01 배포 준비 작업본의 `npm run verify`는 타입 검사, 빌드, 테스트 413건 통과/1건 제외/0건 실패, 패키지 dry-run까지 통과했다. 제외된 1건은 Windows에서 허용되지 않는 큰따옴표 파일명에 대한 POSIX 사례다. 설치 테스트는 실제 PowerShell/Git Bash와 로컬 모의 저장소로 실행했으며, 실제 Mac이나 Claude Code Desktop의 로그인 성공을 증명하지 않는다. 독립 코드 검토에서 발견한 Git Bash의 Windows식 런타임 경로 판정 오류는 수정하고 회귀 테스트로 확인했다. 최종 독립 코드 검토에서 추가 배포 차단 결함은 발견되지 않았다.

아래 순서는 사용자/환경 소유자의 별도 승인 후 시행할 계획이다. 실제 사용자 환경의 설치, 인증, M365 호출은 실행하지 않았다. Windows와 실제 Mac을 각각 확인해야 한다.

| 순서 | 환경/사례 | 통과 기준 | 상태 |
| --- | --- | --- | --- |
| 1 | 3개 언어 HTML, 데스크톱/모바일 폭 | 글자 겹침·가로 넘침 없음, 선택된 OS와 복사 본문 일치, 클립보드 차단 시 수동 선택 가능 | 헤드리스 Edge 12조합 통과. 실제 OS 클립보드 권한은 미검증 |
| 2 | Windows/Mac 실제 Desktop Local | Code > Local에서 선택한 절대 폴더와 실제 작업 폴더 일치, worktree 아님 | 대기 |
| 3 | 각 OS의 도구 준비 | node/npm/git 확인, 부족하면 설치 전 설명·승인, 승인 거부 시 중단, 설치 후 재시작·재검증, 기존 도구 재설치 없음 | 실제 설치 대기 |
| 4 | 공백·한글/일본어·작은따옴표가 있는 선택 경로 | 정확한 셸 인용, 자리표시자 미실행, 데이터 폴더 변경 없음, Windows PowerShell의 Bash 직접 실행 없음 | 대기 |
| 5 | 좁은 권한 요청/거부 | 필요한 작업만 요청, 거부 시 실행·로그인·정책 우회 없이 중단 | 대기 |
| 6 | 최초 설치, 규칙 생성 | 별도 앱 런타임 사용, startup 중복 없음, 생성 규칙을 추가 인증 명령 전에 읽음 | 대기 |
| 7 | 신규 로그인/기존 대기 로그인 | 정확한 setup.state 준수, 사용자 완료 확인 후 완료 명령 1회, 코드 재생성·폴링 없음 | 대기 |
| 8 | 기존 유효 로그인/새 채팅 | 같은 폴더·캐시 보존, 규칙 적용 및 상태 확인, READY에서 재로그인 없이 중단 | 대기 |
| 9 | 네트워크 차단/인증 확인 불가 | unknown 유지, 제한된 검사만 수행, 캐시 삭제·재로그인·커넥터/GUI 우회 없음 | 대기 |
| 10 | 승인된 읽기/초안/다운로드/내보내기 사례 | 실제 읽기 승인 범위 준수, 초안 전체 미리보기, 조건부 승인·토큰·재개 규칙 유지, 메일 발송 없음 | 별도 외부 동작 승인 후 대기 |

수용 증적에는 OS, Desktop 버전, 셸, Node/Git 버전, 선택 경로 동일성 결과, 테스트 사례와 결과만 기록한다. 실제 사용자 경로, 계정, 토큰, 원시 캐시, 반환된 업무 내용을 불필요하게 복제하지 않는다. 실패 시 해당 단계와 비밀값 없는 오류만 남기고 중단한다.

## 근거

- [Claude Code Desktop](https://code.claude.com/docs/en/desktop): Code 탭, Local과 프로젝트 폴더, 권한, 세션 PATH 관련 안내.
- [Claude Code Setup](https://code.claude.com/docs/en/setup): 네이티브 설치 맥락. Hare의 Node 요구 사항은 이 저장소의 `package.json`이 별도 근거다.
- [Claude Code Memory](https://code.claude.com/docs/en/memory): 작업 폴더의 `CLAUDE.md` 로딩과 지침 적용.
- [Claude Code Permissions](https://code.claude.com/docs/en/permissions): 도구별 권한과 정책 경계.
- [Node.js 다운로드](https://nodejs.org/en/download), [npm 설치 안내](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm/): 지원되는 LTS와 npm 포함 설치.
- [Git for Windows](https://git-scm.com/install/windows), [Git for macOS](https://git-scm.com/install/mac): 공식 설치 경로 및 Apple Command Line Tools.
- [Node 추가 도구 설치 스크립트](https://github.com/nodejs/node/blob/v24.x/tools/msvs/install_tools/install_tools.bat), [Git 설치 옵션](https://github.com/git-for-windows/build-extra/blob/main/installer/install.iss): Windows의 선택적 도구와 PATH 옵션 확인 근거.
- [Microsoft WinGet](https://learn.microsoft.com/en-us/windows/package-manager/winget/): 기존 Windows 패키지 관리 도구. 이 준비 작업의 사내 이용 허용은 사용자에게 확인했으며 실제 OS 차단은 우회하지 않는다.
- [Get-AuthenticodeSignature](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-authenticodesignature): Windows 서명 상태의 읽기 전용 확인. 기존 도구에서는 선택적 진단이며, 새 설치 파일에서는 실행 전 확인에 사용한다. npm 스크립트의 서명이나 체인은 검사하지 않는다.
- [Apple Code Signing Tasks](https://developer.apple.com/library/archive/documentation/Security/Conceptual/CodeSigningGuide/Procedures/Procedures.html): Mac 서명 확인 참고. 기존 도구의 codesign 검증은 필수가 아니며 실제 Mac 설치 검증은 미완료다.
- Hare 구현 근거: `src/local-install.ts`, `src/setup-state.ts`, `src/session-rules.ts`, `package.json`. 설치 API 통합과 실제 Desktop 동작은 별도 검증 대상이다.
