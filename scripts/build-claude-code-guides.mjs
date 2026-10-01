import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildLocalSetupCommand } from "../dist/local-install.js";

const root = fileURLToPath(new URL("../", import.meta.url));
export const SELECTED_FOLDER = "<SELECTED_LOCAL_FOLDER>";
export const GUIDE_FILES = {
  ko: "Hare_M365_Claude_Code_연결가이드_KO.html",
  en: "Hare_M365_Claude_Code_Connection_Guide_EN.html",
  ja: "Hare_M365_Claude_Code_接続ガイド_JA.html"
};
export const SHELLS = { windows: "powershell", mac: "posix" };

const copy = {
  ko: {
    title: "Hare M365 Agent 연결 가이드",
    subtitle: "Claude Desktop · Code · Local | Windows / Mac",
    preview: "환경 준비 → Hare 폴더 선택 → 연결 → Microsoft 로그인 순서로 진행합니다. 이미 준비된 프로그램과 기존 Hare 로그인은 가능한 한 그대로 사용합니다.",
    request: "이 PC에서 Hare M365 Agent 연결을 진행해 줘. 아래 준비 명령의 clone·의존성 설치·빌드·규칙 생성을 요청하는 것이며, Microsoft 로그인은 사용자 확인을 받아 진행해. 이 요청은 메일 발송이나 업무 데이터 변경 승인이 아니야.",
    open: "HTML 파일을 내려받아 브라우저에서 여세요. Teams/OneDrive 미리보기에서는 복사가 제한될 수 있습니다.",
    prepare: "1. Claude Code Local 열기",
    prerequisites: "Claude Desktop의 Code 탭에서 진행합니다. Claude CLI를 따로 설치할 필요는 없습니다. Hare 실행에 필요한 Node.js와 Git은 다음 단계에서 Claude가 확인합니다. npm은 Node.js와 함께 설치됩니다.",
    folder: "Code 탭에서 Local을 선택하고 기존 Hare 데이터 폴더를 작업 폴더로 여세요. 처음 사용하는 경우에만 저장소 밖의 OneDrive가 아닌 전용 로컬 폴더를 선택하세요. 클라우드/샌드박스 경로, Git worktree, 임의로 추정한 새 폴더를 사용하지 마세요.",
    tools: "요구 버전을 충족하는 프로그램은 다시 설치하지 않습니다. 구버전은 Claude가 업그레이드 필요 여부를 설명하고 승인을 요청합니다. 설치 후 Claude를 완전히 종료하고 Local로 다시 여세요. 이미 선택한 폴더가 있다면 유지하세요. 새 채팅에서는 방금 설치했다는 사실도 함께 알려주세요.",
    setup: "3. Hare 연결",
    os: "사용 중인 운영체제",
    promptLabel: "Claude Code Local에 붙여넣을 프롬프트",
    paste: "아래 내용은 터미널 명령이 아닌 AI용 프롬프트입니다. 현재 선택한 절대 경로를 AI가 확인하고 자리표시자를 안전하게 치환한 뒤 실행해야 합니다.",
    copy: "Hare 연결 프롬프트 복사",
    copied: "프롬프트를 복사했습니다.",
    fallback: "자동 복사가 제한되었습니다. 선택된 프롬프트를 직접 복사하세요.",
    nojs: "JavaScript가 꺼져 있습니다. 해당 운영체제의 프롬프트를 직접 선택해 복사하세요.",
    next: "4. 로그인과 다음 대화",
    state: "준비 명령이 startup까지 처리합니다. 이후 생성된 CLAUDE.md와 claude/hare-m365-agent-rules.md를 읽고 setup.state의 다음 단계만 따릅니다. 로그인 대기 중에는 사용자의 완료 확인 후 명령을 한 번만 실행합니다. READY이면 멈추고 업무 요청을 기다립니다.",
    reuse: "다음 채팅에서도 반드시 같은 Hare 폴더를 Code Local 작업 폴더로 여세요. 기본 설정에서는 CLAUDE.md가 자동으로 로드되어 생성된 규칙으로 안내합니다. 적용 여부를 확인하고 규칙의 상태 확인 명령을 사용하세요. 새 채팅이라는 이유만으로 다시 로그인하지 마세요.",
    security: "보안 및 권한",
    privacy: "Local은 실행 위치를 뜻합니다. Hare가 반환하여 Claude가 읽는 업무 내용은 모델에 전달됩니다. 원시 인증 캐시나 비밀값을 읽거나 대화에 붙여넣지 마세요.",
    permission: "필요한 도구 권한만 표준 승인 절차로 요청합니다. 일괄 권한 우회는 사용하지 않으며 거부되면 중단합니다. Local의 네트워크 정책을 Cowork 허용 목록과 같다고 가정하지 않습니다. 네트워크로 인증 상태를 확인할 수 없으면 캐시를 보존하며 재로그인하지 않습니다.",
    operations: "조회와 Outlook 초안은 Hare만 사용하며 커넥터나 GUI로 우회하지 않습니다. 초안은 수신자·제목·본문·첨부 전체 미리보기를 보여주고 명시적 승인 후 생성합니다. 발송은 지원하지 않습니다. 다운로드와 내보내기는 생성된 규칙의 조건부 승인 절차를 그대로 따릅니다.",
    sources: "공식 참고 문서",
    before: [
      "현재 Claude Desktop의 Code > Local 세션에서 Hare를 준비해. 제공된 호스트·실행 환경 정보로 확인하고, 구분할 정보가 없을 때만 사용자에게 Local인지 한 번 물어봐. OS나 경로만으로 Local이라고 추정하지 마. Cloud/SSH/WSL이면 실행하지 말고 중단해.",
      "현재 선택한 작업 폴더의 정확한 절대 경로를 확인해. 기존 Hare 사용자라면 기존 데이터 폴더여야 해. 처음 사용하는 경우에만 사용자가 선택한 OneDrive 밖, Git 저장소 밖 전용 로컬 폴더를 사용해. 폴더가 없거나 불명확하면 FOLDER_REQUIRED를 알리고 멈춰. 경로를 추측하거나 새 폴더·Git worktree·클라우드/샌드박스 경로로 바꾸지 마.",
      "아래 명령의 <SELECTED_LOCAL_FOLDER>만 현재 선택한 정확한 절대 Local 작업 폴더로 치환해. 해당 셸의 인용 규칙에 맞게 경로 안의 따옴표 등 특수 문자를 이스케이프하되 경로 값 자체를 바꾸지 마. 다른 명령 내용은 변경하지 마. 치환하지 않은 자리표시자나 리터럴 꺾쇠 경로를 절대 실행하지 마. 안전하게 치환할 수 없으면 멈춰.",
      "같은 실행 세션에서 node --version, npm --version, git --version을 확인해. Hare는 Node.js >=20.18.1이 필요하며 Claude의 네이티브 설치와 무관한 요건이야. 실제 세션 PATH에 도구가 없거나 버전이 부족하면 멈추고 필요한 준비만 안내해. 임의로 설치하거나 PATH를 추측하지 마. Desktop 때문에 Claude CLI를 따로 설치하지 마.",
      "필요한 작업에 대해서만 호스트의 표준 도구 권한 승인을 요청해. 승인 거부·요청 도구 없음·정책 차단이면 중단해. 일괄 권한 우회나 Cowork 네트워크 정책 가정, 프록시/다른 환경 우회를 하지 마. 기존 인증 폴더와 캐시를 보존하고 원시 캐시·비밀값은 읽거나 출력하지 마."
    ],
    windows: "Windows: npm 버전 확인은 PowerShell에서 npm.cmd --version을 사용해. npm.ps1 정책 오류를 해결하려고 실행 정책을 변경하지 마. 아래 PowerShell 명령과 이후 PowerShell nextCommand는 PowerShell에서 실행해. & 호출 연산자가 있는 PowerShell을 Bash에 직접 전달하지 마. 네이티브 PowerShell 도구를 사용하거나 PowerShell 실행 파일에 스크립트를 안전한 인수로 전달해. Bash의 변수 확장·인용 해석으로 스크립트가 바뀌지 않게 하고, 안전한 전달이 불가능하면 중단해. --host claude-code와 --command-shell powershell을 유지해.",
    mac: "Mac: 아래 명령과 이후 POSIX nextCommand는 POSIX 셸에서 실행해. --host claude-code와 --command-shell posix를 유지해.",
    command: "준비 명령 (startup 포함, 전체 반복 금지):",
    after: [
      "준비 명령이 startup을 이미 처리하므로 startup을 중복 실행하지 마. 명령이 반환되면 추가 인증 명령이나 업무 조회 전에 선택 폴더의 CLAUDE.md와 claude/hare-m365-agent-rules.md를 읽어. 원시 인증 캐시를 읽지 마. 규칙이 없거나 호스트·셸·경로가 맞지 않으면 중단해. 이후에도 생성된 정확한 명령 접두사, --data-dir, --host claude-code, --command-shell을 유지해.",
      "startup 출력의 setup.state와 setup.nextCommand만 따라 한 단계씩 진행해. SETUP_REQUIRED/FOLDER_REQUIRED 또는 알 수 없는 상태는 이유를 알리고 중단해. LOGIN_START_REQUIRED일 때만 nextCommand를 변경 없이 포그라운드에서 한 번 실행하고 반환된 Microsoft 주소와 userCode를 보여준 뒤 사용자 본인의 회사 계정으로 로그인하도록 안내하고 멈춰. 특정 이메일 계정을 지정하지 마.",
      "LOGIN_COMPLETE_REQUIRED이면 기존 대기 로그인과 코드를 유지해. 사용자가 로그인을 완료했다고 확인하기 전에는 완료 명령을 실행하지 마. 확인 후 setup.nextCommand를 변경 없이 정확히 한 번 포그라운드 실행해. 폴링·백그라운드·중복 login-start를 하지 마. READY이면 준비 완료를 알리고 멈춰. 실제 조회는 별도 사용자 요청과 loggedIn=true 및 tokenUsable=true 확인 후에만 해.",
      "BLOCKED 또는 AUTH_CHECK_BLOCKED의 loggedIn=null/tokenUsable=null은 인증 확인 불가이지 로그아웃이나 만료가 아니야. 기존 인증 폴더와 캐시를 보존해. 생성된 규칙의 제한된 network check --environment claude-code 절차만 따르고 이미 실패·거부된 검사는 반복하지 마. 재로그인·캐시 초기화·다른 dataDir·커넥터·GUI 우회를 하지 마. 그 밖의 실패는 실패 단계와 비밀값 없는 오류 한 줄만 알리고 중단해.",
      "향후 M365 조회와 Outlook 초안에는 Hare CLI만 사용해. 커넥터·Computer Use·브라우저 자동화·Outlook/Teams/SharePoint GUI를 대체 경로로 사용하지 마. 기본은 읽기 전용이며 Outlook 초안은 수신자, 제목, 본문, 첨부 전체 미리보기와 명시적 승인 후에만 생성해. 발송하지 마. 다운로드·export의 조건부 AWAITING_USER_APPROVAL, 전체 계획, 승인 토큰, 제한·재개 규칙은 생성된 규칙을 그대로 따르고 무조건 승인 면제로 바꾸지 마.",
      "다음 채팅도 같은 폴더에서 시작해 CLAUDE.md의 규칙 연결이 적용됐는지 확인하고 생성된 상태 확인 명령을 사용해. 새 채팅만을 이유로 재설치하거나 재로그인하지 마. Local 실행이어도 반환된 업무 내용은 모델에 전달된다는 점을 사용자에게 알려줘."
    ]
  },
  en: {
    title: "Hare M365 Agent Connection Guide",
    subtitle: "Claude Desktop · Code · Local | Windows / Mac",
    preview: "Prepare prerequisites, select your Hare folder, connect, then sign in to Microsoft. Reuse working tools and existing Hare sign-in whenever possible.",
    request: "Connect Hare M365 Agent on this PC. I am requesting the clone, dependency installation, build, and rules generation in the setup command below. Coordinate Microsoft sign-in with the user. This is not approval to send mail or modify work data.",
    open: "Download this HTML file and open it in a browser. Teams/OneDrive previews may restrict copying.",
    prepare: "1. Open Claude Code Local",
    prerequisites: "Use the Code tab in Claude Desktop. You do not need a separate Claude CLI installation. In the next step, Claude checks Node.js and Git, which Hare needs to run. npm is included with Node.js.",
    folder: "In the Code tab, select Local and open your existing Hare data folder as the working folder. First-time users only: select a dedicated non-OneDrive local folder outside any repository. Do not use a cloud/sandbox path, a Git worktree, or a guessed replacement folder.",
    tools: "Do not reinstall tools that meet the requirements. Claude explains and asks for approval if an upgrade is needed. After installation, fully quit and reopen Claude in Local. Keep the same folder if one was already selected. In a new chat, also mention that you just installed the tools.",
    setup: "3. Connect Hare",
    os: "Your operating system",
    promptLabel: "Prompt for Claude Code Local",
    paste: "This is an AI prompt, not a command to paste into a terminal. The AI must verify your selected absolute folder and safely replace the placeholder before execution.",
    copy: "Copy Hare connection prompt",
    copied: "Prompt copied.",
    fallback: "Automatic copying is unavailable. Copy the selected prompt manually.",
    nojs: "JavaScript is disabled. Select and copy the prompt for your operating system manually.",
    next: "4. Sign-in and future chats",
    state: "The setup command already runs startup. Then read the generated CLAUDE.md and claude/hare-m365-agent-rules.md and follow only setup.state. For a pending login, run its completion command once after the user confirms completion. At READY, stop and wait for a work request.",
    reuse: "Open the SAME Hare folder as the Code Local working folder in every new chat. By default, CLAUDE.md loads automatically and points to the generated rules. Verify they apply and use the status command in those rules. A new chat alone is not a reason to log in again.",
    security: "Security and permissions",
    privacy: "Local describes where commands run. Work content returned by Hare and read by Claude is still sent to the model. Never read or paste raw authentication caches or secrets into chat.",
    permission: "Request only the needed standard tool permissions. Do not bypass permissions; stop if denied. Do not assume Local has Cowork's network allowlist policy. If network problems prevent authentication verification, preserve the cache and do not start a new login.",
    operations: "Use Hare exclusively for lookups and Outlook drafts, without connector or GUI fallbacks. Draft creation requires the complete recipients, subject, body, and attachments preview and explicit approval. Sending is unsupported. Preserve the generated rules' conditional download and export approval gates.",
    sources: "Official references",
    before: [
      "Prepare Hare in this Claude Desktop Code > Local session. Verify this from available host/execution context; only if that context cannot distinguish the mode, ask the user once whether it is Local. Do not infer Local from the OS or path alone. Stop without execution in Cloud/SSH/WSL.",
      "Verify the exact absolute working folder currently selected. Existing users must reuse their existing Hare data folder. First-time users only: use the user-selected dedicated non-OneDrive local folder outside Git repositories. If the selected folder is absent or uncertain, report FOLDER_REQUIRED and stop. Do not guess a path or replace it with a new folder, Git worktree, or cloud/sandbox location.",
      "Replace ONLY <SELECTED_LOCAL_FOLDER> in the command below with the current exact selected absolute Local working folder. Escape quotes and other special characters using the target shell's quoting rules without changing the path value. Do not change any other command content. Never execute the literal angle-bracket placeholder or an unreplaced path. Stop if safe substitution is not possible.",
      "Check node --version, npm --version, and git --version in the SAME execution session. Hare requires Node.js >=20.18.1 independently of Claude's native installation. If versions or tools on the actual session PATH are missing, stop and explain the prerequisite; do not install tools or guess PATH. Do not install a separate Claude CLI for Desktop.",
      "Request only the necessary standard host tool permission for the specific action. Stop if denied, the permission tool is unavailable, or policy blocks it. Do not bypass permissions, assume Cowork network policies, or escape via proxies or another environment. Preserve the existing authentication folder and cache; never read or print raw caches or secrets."
    ],
    windows: "Windows: check npm with npm.cmd --version in PowerShell; do not change execution policy to fix npm.ps1 errors. Execute the PowerShell command below and subsequent PowerShell nextCommand values in PowerShell. Never pass PowerShell containing the & call operator directly to Bash. Use a native PowerShell tool or invoke the PowerShell executable with the script passed as safe arguments. Prevent Bash expansion or quoting from changing the script; stop if safe invocation is unavailable. Preserve --host claude-code and --command-shell powershell.",
    mac: "Mac: execute the command below and subsequent POSIX nextCommand values in a POSIX shell. Preserve --host claude-code and --command-shell posix.",
    command: "Setup command (includes startup; do not repeat the entire setup):",
    after: [
      "The installer already handles startup: do not run startup a second time. After it returns, read the selected folder's generated CLAUDE.md and claude/hare-m365-agent-rules.md BEFORE any additional authentication command or work query. Never read raw authentication caches. Stop if rules are missing or the host, shell, or folder is wrong. Preserve the exact generated command prefix, --data-dir, --host claude-code, and --command-shell for all later commands.",
      "Follow only setup.state and setup.nextCommand from startup, one step at a time. Report and stop for SETUP_REQUIRED/FOLDER_REQUIRED or unknown states. Only for LOGIN_START_REQUIRED, run nextCommand unchanged once in the foreground, show the returned Microsoft URL and userCode, ask the user to sign in with their own company Microsoft account, and stop. Do not name or preselect an email account.",
      "For LOGIN_COMPLETE_REQUIRED, preserve the pending login and existing code. Wait until the user explicitly confirms sign-in completion, then run setup.nextCommand unchanged exactly once in the foreground. No polling, background execution, or duplicate login-start. At READY, report readiness and STOP. Run work queries only on a separate user request and after loggedIn=true AND tokenUsable=true are verified.",
      "BLOCKED or AUTH_CHECK_BLOCKED with loggedIn=null/tokenUsable=null means authentication is UNKNOWN, not signed out or expired. Preserve the authentication folder and cache. Follow only the generated rules' bounded network check --environment claude-code procedure; do not repeat a failed or denied check. Do not start a new login, reset caches, switch dataDir, or use a connector/GUI fallback. For other failures, report the failed step and one sanitized error line, then stop.",
      "Use Hare CLI exclusively for future M365 queries and Outlook drafts. No connectors, Computer Use, browser automation, or Outlook/Teams/SharePoint GUI fallbacks. Default to read-only; create an Outlook draft only after the complete recipients, subject, body, and attachments preview and explicit approval. Never send it. Preserve the generated rules' conditional download/export AWAITING_USER_APPROVAL gates, complete plans, approval tokens, limits, and resume rules; do not replace them with unconditional permission.",
      "Start future chats in the SAME folder, verify that CLAUDE.md's linked rules apply, and use the generated status command. Do not reinstall or log in just because a chat is new. Tell the user that Local execution still sends returned work content to the model."
    ]
  },
  ja: {
    title: "Hare M365 Agent 接続ガイド",
    subtitle: "Claude Desktop · Code · Local | Windows / Mac",
    preview: "環境準備 → Hare フォルダー選択 → 接続 → Microsoft ログインの順に進めます。利用可能なツールと既存の Hare ログインをできるだけ再利用します。",
    request: "この PC で Hare M365 Agent の接続を進めてください。以下の準備コマンドによる clone・依存関係のインストール・ビルド・ルール生成を依頼します。Microsoft ログインはユーザーと確認して進めてください。メール送信や業務データ変更の承認ではありません。",
    open: "この HTML ファイルをダウンロードし、ブラウザーで開いてください。Teams/OneDrive のプレビューではコピーが制限される場合があります。",
    prepare: "1. Claude Code Local を開く",
    prerequisites: "Claude Desktop の Code タブで進めます。Claude CLI の別途インストールは不要です。次の手順で、Hare に必要な Node.js と Git を Claude が確認します。npm は Node.js と一緒にインストールされます。",
    folder: "Code タブで Local を選び、既存の Hare データフォルダーを作業フォルダーとして開きます。初めて利用する場合のみ、リポジトリ外の OneDrive ではない専用ローカルフォルダーを選んでください。クラウド/サンドボックスのパス、Git worktree、推測した代替フォルダーは使用しません。",
    tools: "要件を満たすツールは再インストールしません。アップグレードが必要なら Claude が説明して承認を求めます。インストール後は Claude を完全に終了し、Local で開き直してください。選択済みフォルダーがあれば維持してください。新しいチャットでは、直前にインストールしたことも伝えてください。",
    setup: "3. Hare に接続",
    os: "使用する OS",
    promptLabel: "Claude Code Local に貼り付けるプロンプト",
    paste: "これは AI 用プロンプトであり、ターミナルに貼り付けるコマンドではありません。AI が選択中の絶対パスを確認し、プレースホルダーを安全に置換してから実行します。",
    copy: "Hare 接続プロンプトをコピー",
    copied: "プロンプトをコピーしました。",
    fallback: "自動コピーを利用できません。選択されたプロンプトを手動でコピーしてください。",
    nojs: "JavaScript が無効です。使用する OS のプロンプトを手動で選択してコピーしてください。",
    next: "4. ログインと次回のチャット",
    state: "準備コマンドは startup まで実行します。その後、生成された CLAUDE.md と claude/hare-m365-agent-rules.md を読み、setup.state の次の手順だけに従います。ログイン待機中は、ユーザーの完了確認後に完了コマンドを一度だけ実行します。READY では停止して業務依頼を待ちます。",
    reuse: "新しいチャットでも必ず同じ Hare フォルダーを Code Local の作業フォルダーとして開いてください。既定では CLAUDE.md が自動で読み込まれ、生成されたルールへ案内します。適用を確認し、ルール内の状態確認コマンドを使用してください。新しいチャットという理由だけで再ログインしないでください。",
    security: "セキュリティと権限",
    privacy: "Local はコマンドの実行場所を意味します。Hare が返し Claude が読む業務内容はモデルに送信されます。生の認証キャッシュや秘密情報を読んだり、チャットに貼り付けたりしないでください。",
    permission: "必要なツール権限だけを標準の承認手順で要求します。権限の一括回避は行わず、拒否されたら停止します。Local のネットワーク方針が Cowork の許可リストと同じとは仮定しません。ネットワーク障害で認証を確認できない場合はキャッシュを保持し、再ログインしません。",
    operations: "検索と Outlook 下書きには Hare だけを使い、コネクターや GUI で迂回しません。下書きは宛先・件名・本文・添付の完全なプレビューと明示的な承認後に作成します。送信には対応しません。ダウンロードとエクスポートは、生成されたルールの条件付き承認手順を維持します。",
    sources: "公式参考資料",
    before: [
      "現在の Claude Desktop の Code > Local セッションで Hare を準備してください。提供されたホスト・実行環境情報で確認し、判別できない場合だけ Local かユーザーに一度確認してください。OS やパスだけで Local と推測しません。Cloud/SSH/WSL では実行せず停止してください。",
      "現在選択されている作業フォルダーの正確な絶対パスを確認してください。既存ユーザーは既存の Hare データフォルダーを再利用します。初回のみ、ユーザーが選択した OneDrive 外かつ Git リポジトリ外の専用ローカルフォルダーを使ってください。選択がない、または不明な場合は FOLDER_REQUIRED を報告して停止してください。推測したパス、新しい代替フォルダー、Git worktree、クラウド/サンドボックスの場所を使わないでください。",
      "以下のコマンドの <SELECTED_LOCAL_FOLDER> だけを、現在選択中の正確な絶対 Local 作業フォルダーに置換してください。パス内の引用符などは対象シェルの規則でエスケープし、パス自体の値は変更しないでください。他のコマンド内容は変更しません。未置換のプレースホルダーや山括弧を含む仮のパスを絶対に実行しないでください。安全に置換できなければ停止してください。",
      "同じ実行セッションで node --version、npm --version、git --version を確認してください。Hare は Claude のネイティブインストールとは独立して Node.js >=20.18.1 を必要とします。実際のセッション PATH でツールが見つからない、またはバージョン不足なら停止し、必要な準備だけを案内してください。勝手なインストールや PATH の推測はしません。Desktop のために Claude CLI を別途インストールしないでください。",
      "対象の操作に必要なホスト標準のツール権限だけを要求してください。拒否、承認ツール不在、ポリシーによる遮断では停止してください。権限回避、Cowork ネットワーク方針の決めつけ、プロキシや別環境への迂回はしません。既存の認証フォルダーとキャッシュを保持し、生のキャッシュや秘密情報を読んだり出力したりしないでください。"
    ],
    windows: "Windows: npm は PowerShell で npm.cmd --version により確認し、npm.ps1 エラーのために実行ポリシーを変更しないでください。以下の PowerShell コマンドと以後の PowerShell nextCommand は PowerShell で実行してください。& 呼び出し演算子を含む PowerShell を Bash に直接渡してはいけません。ネイティブの PowerShell ツールを使うか、PowerShell 実行ファイルにスクリプトを安全な引数として渡してください。Bash の展開や引用符解釈による変更を防げない場合は停止してください。--host claude-code と --command-shell powershell を維持してください。",
    mac: "Mac: 以下のコマンドと以後の POSIX nextCommand は POSIX シェルで実行してください。--host claude-code と --command-shell posix を維持してください。",
    command: "準備コマンド (startup を含む。全体を繰り返さない):",
    after: [
      "準備コマンドは startup を実行済みなので、startup を重複実行しないでください。終了後、追加の認証コマンドや業務検索の前に、選択フォルダーに生成された CLAUDE.md と claude/hare-m365-agent-rules.md を読んでください。生の認証キャッシュは読みません。ルールがない、またはホスト・シェル・フォルダーが違う場合は停止してください。以後も生成された正確なコマンド接頭辞、--data-dir、--host claude-code、--command-shell を維持してください。",
      "startup の setup.state と setup.nextCommand だけに従い、一段階ずつ進めてください。SETUP_REQUIRED/FOLDER_REQUIRED や不明な状態は報告して停止します。LOGIN_START_REQUIRED の場合だけ nextCommand を変更せずフォアグラウンドで一度実行し、返された Microsoft URL と userCode を表示して、ユーザー自身の会社 Microsoft アカウントでログインするよう案内し停止してください。特定のメールアカウントを指定しないでください。",
      "LOGIN_COMPLETE_REQUIRED では既存のログイン待機状態とコードを保持してください。ユーザーがログイン完了を明示的に確認するまで完了コマンドを実行しません。確認後、setup.nextCommand を変更せずフォアグラウンドで正確に一度実行してください。ポーリング、バックグラウンド実行、login-start の重複は禁止です。READY では準備完了を伝えて停止してください。業務検索は別途のユーザー依頼と loggedIn=true かつ tokenUsable=true の確認後のみ実行してください。",
      "BLOCKED または AUTH_CHECK_BLOCKED の loggedIn=null/tokenUsable=null は認証未確認であり、ログアウトや期限切れではありません。認証フォルダーとキャッシュを保持してください。生成されたルールの限定された network check --environment claude-code 手順だけに従い、失敗・拒否済みの検査は繰り返さないでください。再ログイン、キャッシュ初期化、別 dataDir、コネクター/GUI への迂回はしません。他の失敗では失敗段階と秘密情報を含まないエラー一行だけを報告して停止してください。",
      "以後の M365 検索と Outlook 下書きは Hare CLI だけを使用してください。コネクター、Computer Use、ブラウザー自動化、Outlook/Teams/SharePoint GUI に迂回しません。既定は読み取り専用です。Outlook 下書きは宛先、件名、本文、添付の完全なプレビューと明示的承認の後だけ作成し、送信しないでください。ダウンロード/export の条件付き AWAITING_USER_APPROVAL、全体計画、承認トークン、制限、再開手順は生成されたルールを維持し、無条件許可に変更しないでください。",
      "次回も同じフォルダーでチャットを開始し、CLAUDE.md のルール参照が適用されていることを確認して生成された状態確認コマンドを使ってください。新しいチャットというだけで再インストールや再ログインをしないでください。Local 実行でも返された業務内容がモデルに送信されることをユーザーに伝えてください。"
    ]
  }
};

const environmentCopy = {
  ko: {
    heading: "2. Claude에게 환경 준비 맡기기",
    intro: "아래 환경 준비 프롬프트를 Claude Code Local 채팅에 붙여넣으세요. Claude가 필요한 프로그램을 확인하고, 없는 항목은 설치 방법을 설명한 뒤 승인을 요청합니다. 모두 준비됐다는 안내를 받으면 3번으로 넘어가세요.",
    label: "환경 준비 프롬프트",
    copy: "환경 준비 프롬프트 복사",
    policy: "사내 허용 정책: 이 Hare 환경 준비에 한해 nodejs.org의 공식 Node.js 설치 프로그램, git-scm.com에서 연결하는 공식 Git 설치 프로그램, Mac의 Apple Command Line Tools, Windows에 이미 있는 winget 사용이 허용돼 있어. 직원에게 회사 허용 여부를 다시 확인하도록 요구하지 마. 개별 설치·업그레이드 전 설명과 사용자 승인은 별개로 필요해. 이 허용은 관리자 권한 부여나 실제 OS·보안 정책 차단의 우회 허가가 아니야.",
    pathRecovery: "설치됨·PATH 미반영은 ENVIRONMENT_INCOMPLETE(환경 준비 미완료)이며 Hare 연결 단계로 진행하지 마. 아직 재시작하지 않았다면 Claude를 완전히 종료하고 Local로 한 번만 다시 열도록 안내해. 새 세션에서는 이번 문제로 이미 재시작했는지 먼저 물어보고, 했다면 또 재시작을 요청하지 마. 재시작 후 세션에서 일반 명령으로 버전을 재확인해(Windows npm은 npm.cmd --version). 계속 인식되지 않으면 ENVIRONMENT_INCOMPLETE와 해당 도구·오류를 보고하고 IT 확인을 안내해. 재설치·수동 PATH 편집·절대 경로로 Hare 연결 우회는 하지 마. 버전 요건을 만족하고 일반 명령이 모두 동작할 때만 환경 준비 완료라고 해.",
    origin: {
      windows: "기존 도구는 Get-Command로 실제 실행 경로와 링크 대상을 확인하고 버전 명령이 성공하는지 점검해. 표준 설치 위치, OS 설치 기록, nvm-windows 등 기존 버전 관리자 위치를 모두 인정해. 기존 도구의 서명 확인은 선택적 진단이며 서명 유무·조회 실패·서명자 차이만으로 정상 작동하는 도구를 차단하지 마. 이 PC에서 확인한 서명자는 Node.js의 OpenJS Foundation, Git for Windows의 Johannes Schindelin이지만 고정 허용 목록은 아니야. Node와 같은 설치 폴더의 npm.cmd는 그대로 사용해. 전역 업데이트 등으로 다른 위치의 npm이 선택되면 위치 차이만 보고하고 정상 실행을 막지 마. npm 스크립트 체인 분석, 패키지 재다운로드·압축 해제·해시 대조는 하지 마. 새 설치 파일은 별개로 공식 다운로드 출처, Get-AuthenticodeSignature -LiteralPath의 Valid, 공식 배포자 일치를 확인한 후 사용자 승인을 받아 실행해. 새 설치 파일의 서명·출처 확인이 실패하거나 OS가 실제 차단하면 실행하지 마.",
      mac: "기존 도구는 command -v로 실제 실행 경로와 링크 대상을 확인하고 버전 명령이 성공하는지 점검해. 표준 설치 위치, 설치 기록, 기존 버전 관리자 위치를 모두 인정해. 기존 도구의 codesign 검증은 필수가 아니며 서명·설치 기록 부족만으로 정상 작동하는 도구를 차단하지 마. Node와 같은 설치의 npm은 그대로 사용하고, 전역 업데이트 등으로 다른 위치의 npm이 선택되면 위치 차이만 보고해. npm 스크립트 체인 분석, 패키지 재다운로드·압축 해제·해시 대조는 하지 마. 새 .pkg는 별개로 공식 출처와 pkgutil --check-signature의 유효한 서명·공식 배포자를 확인한 뒤 사용자 승인을 받아 실행해. 확인 실패나 실제 OS 차단 시 실행하지 마. Gatekeeper를 해제하지 마."
    },
    manual: "Code가 열리지 않거나 직접 설치해야 하나요?",
    manualCheck: "이미 설치했거나 방금 설치한 프로그램이라면 다시 설치하지 마세요. 먼저 Claude를 완전히 종료하고 다시 여세요. 계속 오류가 나거나 설치 여부를 모르겠으면 사내 IT 담당자에게 확인하세요. 아래 설치 절차는 미설치 또는 승인된 업그레이드가 확인된 경우에만 진행합니다. 새로 내려받은 설치 파일은 실행 전 공식 출처·유효한 디지털 서명·공식 배포자를 확인하세요. 확인이 어렵다면 실행 전에 IT 담당자의 도움을 받으세요. 기존 도구에는 이 검사를 요구하지 않습니다.",
    windows: [
      "Git 또는 Git Bash가 필요하다는 오류가 나올 때: Git for Windows 공식 페이지에서 PC에 맞는 설치 파일을 내려받아 실행하세요. 시스템 종류는 Windows 설정 > 시스템 > 정보에서 확인할 수 있습니다. Git Bash는 Git for Windows에 포함됩니다. PATH 선택 화면에서는 Git from the command line and also from 3rd-party software (권장)를 선택하세요. 다른 실행 오류를 Git 누락으로 단정하지 마세요.",
      "Node.js가 필요하거나 최소 버전보다 오래된 경우: 공식 다운로드 페이지에서 Windows와 LTS를 선택하고 Windows Installer (.msi)를 내려받아 실행하세요. npm은 함께 설치됩니다. Tools for Native Modules의 추가 도구 자동 설치는 체크하지 마세요. 이 옵션은 Chocolatey·Python·Visual Studio Build Tools 등을 추가로 설치합니다."
    ],
    mac: [
      "Git이 필요한 경우: Spotlight에서 터미널을 찾아 열고 xcode-select --install 을 입력한 뒤 Enter를 누르세요. Apple 설치창에서 Command Line Tools 설치를 진행합니다. 전체 Xcode나 Homebrew를 따로 설치할 필요는 없습니다.",
      "Node.js가 필요한 경우: 공식 다운로드 페이지에서 macOS와 LTS를 선택하고 macOS Installer (.pkg)를 내려받아 실행하세요. npm도 함께 설치됩니다."
    ],
    blocked: "공식 Node.js·Git 설치 프로그램, Mac의 Apple Command Line Tools, 기존 winget 사용은 사내에서 허용되어 있습니다. 설치 전 Claude의 설명을 확인하고 승인하세요. 관리자 승인창은 직접 처리하며 암호는 Claude 채팅에 입력하지 마세요. 설치가 실제로 차단되면 우회하지 말고 IT 담당자에게 요청하세요. 설치 후 Claude를 완전히 종료하고 Local로 한 번 다시 여세요. 선택한 폴더가 있다면 유지하세요. 계속 프로그램을 인식하지 못하면 환경 준비 미완료이므로 Hare 연결을 진행하지 말고 IT 담당자에게 문의하세요.",
    prompt: [
      "이 PC의 Hare 실행 환경을 지금 점검해 줘. 이 프롬프트는 검토가 아닌 실제 환경 준비 요청이야. 읽기 전용 점검부터 시작하고, 설치·업그레이드는 대상과 방법을 설명한 뒤 별도 사용자 승인을 받아.",
      "Hare 연결 전에 이 PC의 필수 실행 환경만 준비해 줘. 제공된 호스트·실행 환경 정보로 Claude Desktop > Code > Local과 실제 OS를 확인해. 구분할 정보가 없을 때만 사용자에게 Local인지 한 번 물어보고, OS나 경로만으로 추정하지 마. Cloud/SSH/WSL이면 멈춰. 이 환경 준비 단계는 Hare 작업 폴더 선택 없이 진행할 수 있어. 폴더 확인·생성은 다음 Hare 연결 단계에서 하고, 앱 자체가 작업 폴더를 요구할 때만 사용자가 로컬 폴더를 선택하도록 안내해. Hare 설치, git clone, npm ci, 빌드, Microsoft 로그인, 업무 조회는 이 단계에서 하지 마. 기존 Hare 데이터 파일·인증 캐시를 읽거나 변경하지 마.",
      "같은 세션에서 node --version, npm --version, git --version을 읽기 전용으로 확인하고 준비됨/설치 필요/업그레이드 필요/설치됨·PATH 미반영/확인 불가를 구분해. Node.js >=20.18.1은 Hare의 최소 조건이야. 최소 버전 미만은 업그레이드 필요로 분류하고 설치와 같은 승인 절차를 따라. 새 설치·업그레이드는 공식 사이트의 현재 지원되는 LTS를 사용해. Node 20.x는 지원 종료 상태이므로 기능상 실행 가능과 별개로 지원되는 LTS 업그레이드를 권고하고 승인을 받아. 거절하면 지원 종료 경고를 남기되 기술적 최소 조건을 충족하는 실행 자체를 차단하지 마. 요구 버전을 충족하고 이미 사용할 수 있는 도구는 사용자 승인 없이 다시 설치·업데이트하지 마. Git은 git --version과 git branch -h 도움말의 --show-current 지원을 확인해. 도움말 출력의 종료 코드만으로 설치 실패라고 하지 마. 필요한 옵션이 없을 때만 업그레이드를 제안해. 권한 거부나 네트워크 오류를 미설치로 판정하지 마. Mac에서 git 호출이 설치창을 띄우면 자동 승인하지 말고 사용자에게 알려줘.",
      "명령을 찾지 못해도 즉시 설치를 제안하지 마. 새 세션에서도 표준 설치 위치와 OS의 설치 기록, 기존 버전 관리자 위치를 필요한 범위만 읽기 전용으로 확인해. 실행 파일의 존재만으로 준비 완료라 하지 말고 발견한 도구를 절대 경로로 버전 확인해. 절대 경로 실행은 되지만 일반 명령이 안 되면 설치됨·PATH 미반영으로 분류하고 재설치하지 마. 표준 위치에 없다는 이유만으로 미설치로 단정하지 마. 설치 여부나 이전 설치 시도가 불명확하면 사용자에게 최근 설치 여부·별도 설치 위치를 확인하고 멈춰. 디스크 전체를 탐색하거나 인증 캐시를 읽지 마.",
      "설치·업그레이드가 필요한 항목만 이름·이유·공식 출처·설치 방법·변경 범위·관리자 권한 필요 여부를 먼저 보여주고 명시적 승인을 기다려. 위 사내 허용 정책에 명시한 설치 방법을 사용해. 설치 방법이나 대상이 바뀌면 다시 확인받아. npm은 Node.js와 함께 설치하며 독립적으로 최신화하지 마. 기존 버전 관리 도구가 있으면 다른 방식의 설치를 겹치지 말고 현재 관리 방식을 먼저 확인해.",
      "승인 후에만 허용된 표준 도구로 해당 설치를 실행해. 사용자 설치창·관리자 승인이 필요하거나 도구로 처리할 수 없으면 필요한 클릭만 안내하고 기다려. 암호를 요청하거나 읽지 마. 권한 거부·정책 차단이면 멈추고 사내 IT 담당자 확인을 안내해. 보안 설정·실행 정책·PATH를 임의로 바꾸거나 권한을 일괄 우회하지 마. Claude CLI, WSL, Docker, 별도 패키지 관리자는 이 준비를 이유로 새로 설치하지 마.",
      "설치했다면 Claude를 완전히 종료하고 Local로 다시 열도록 안내해. 기존 선택 폴더가 있으면 그대로 유지해. 재시작 후 같은 프롬프트로 버전을 재확인하고 모두 사용할 수 있을 때만 환경 준비 완료, 다음 Hare 연결 단계로 진행하라고 말해. 재시작 후에도 도구가 안 보이면 재설치 반복 없이 오류만 보고하고 멈춰. 환경 준비 성공을 Hare 로그인 성공으로 표현하지 마."
    ],
    platformPrompt: {
      windows: "대상 OS: Windows 네이티브 Local. 공식 출처는 https://nodejs.org/en/download 및 https://git-scm.com/install/windows 이야. npm 확인에는 npm.cmd --version을 사용해. npm.ps1 정책 오류는 미설치가 아니야. 명령이 안 보이면 실제 ProgramFiles 아래 nodejs와 Git/cmd (일반적으로 C:\\Program Files\\nodejs, C:\\Program Files\\Git\\cmd), LocalAppData 아래 Programs/Git, OS 설치 기록과 이미 확인된 버전 관리자 위치만 읽기 전용으로 확인해. 기존 winget 사용은 이 준비 작업에 대해 사내 허용이 확인돼 있어. 패키지는 Node LTS=OpenJS.NodeJS.LTS, Git=Git.Git으로 고정하고 --exact --source winget으로 조회해. 승인 후 설치에는 --interactive를 지정해 사용자가 설치 옵션을 확인하게 하고 --silent나 임의 --override는 사용하지 마. Node는 npm 포함 MSI 설치를 선택하고 portable ZIP은 사용하지 마. 설치 전 패키지 ID·배포자·출처·설치 버전·변경 범위를 보여주고 명시적 승인을 받아. winget이 없거나 차단돼 있으면 새로 설치·복구하거나 우회하지 말고 허용된 공식 설치 프로그램의 수동 절차를 안내해. 그 설치도 실제로 차단되면 중단해. 전체 패키지 일괄 업데이트나 소스 추가·변경은 하지 마. UAC는 사용자가 직접 처리해. Node 설치의 Tools for Native Modules 추가 도구는 선택하지 마. Git PATH 옵션은 Git from the command line and also from 3rd-party software를 선택하고 해당 PATH 변경도 승인 전 설명에 포함해. 임의의 별도 PATH 편집은 하지 마. Git Bash에 PowerShell 구문을 직접 전달하지 마.",
      mac: "대상 OS: Mac Local, POSIX 셸. 공식 출처는 https://nodejs.org/en/download 및 https://git-scm.com/install/mac 이야. 명령이 없으면 /usr/local/bin/node, /usr/local/bin/npm, /usr/bin/git과 Apple 설치 기록, 이미 확인된 기존 버전 관리자 위치만 읽기 전용으로 확인해. Git용 Apple Command Line Tools는 xcode-select --install 로 설치창을 요청할 수 있지만 사용자 승인 전에 실행하지 마."
    }
  },
  en: {
    heading: "2. Let Claude prepare the environment",
    intro: "Paste the environment prompt below into your Claude Code Local chat. Claude checks the required tools and asks for approval after explaining any missing installation. Continue to step 3 only when all tools are ready.",
    label: "Environment preparation prompt",
    copy: "Copy environment prompt",
    policy: "Company authorization: for this Hare preparation, official Node.js installers from nodejs.org, official Git installers linked by git-scm.com, Apple's Command Line Tools on Mac, and existing Windows winget are permitted. Do not ask employees to reconfirm company permission. Explain each installation or upgrade and obtain the user's explicit approval separately. This authorization does not grant administrator rights or permit bypassing actual OS or security restrictions.",
    pathRecovery: "Installed but not on PATH means ENVIRONMENT_INCOMPLETE, not ready for Hare connection. If not yet restarted, ask the user to fully quit Claude and reopen Local once. In a new session, first ask whether they already restarted for this issue; if so, do not request another restart. Recheck versions through ordinary commands in the session after restart (npm.cmd --version on Windows). If still unavailable, report ENVIRONMENT_INCOMPLETE with the tool and error and refer to IT. Do not reinstall, manually edit PATH, or bypass the issue by connecting Hare with absolute paths. Report environment ready only when version requirements pass and all ordinary commands work.",
    origin: {
      windows: "For existing tools, use Get-Command to resolve executable paths and link targets, then check version-command success. Accept standard installation locations, OS installation records, and existing version-manager locations such as nvm-windows. Existing-tool signatures are optional diagnostics: do not block working tools solely for absent signatures, lookup failure, or a different signer. The signers observed on this PC were OpenJS Foundation for Node.js and Johannes Schindelin for Git for Windows; these are not a fixed allowlist. Use npm.cmd in Node's installation folder as-is. If a global update selects npm elsewhere, report the location difference without blocking successful execution. Do not analyze the npm script chain or redownload, extract, or hash-compare installation packages. New installers are separate: verify the official download source, Get-AuthenticodeSignature -LiteralPath status Valid, and matching official publisher, then obtain user approval before execution. Do not run a new installer if source/signature verification fails or the OS actually blocks it.",
      mac: "For existing tools, use command -v to resolve executable paths and link targets, then check version-command success. Accept standard installation locations, installation records, and existing version-manager locations. Existing-tool codesign verification is not required; do not block working tools solely for missing signatures or installation records. Use npm from the Node installation as-is; if a global update selects npm elsewhere, report the location difference only. Do not analyze the npm script chain or redownload, extract, or hash-compare installation packages. New .pkg installers are separate: verify the official source and valid signature with the matching official publisher via pkgutil --check-signature, then obtain user approval before execution. Stop on verification failure or an actual OS block. Never disable Gatekeeper."
    },
    manual: "Code will not open, or you need to install manually?",
    manualCheck: "Do not reinstall tools you already installed, including just now. First fully quit and reopen Claude. If the error persists or you are unsure whether a tool is installed, ask internal IT. Follow the installation steps below only for confirmed missing tools or an approved upgrade. Before running a newly downloaded installer, verify its official source, valid digital signature, and official publisher. Ask IT for help before execution if you cannot verify them. Existing tools do not require this check.",
    windows: [
      "If an error specifically requires Git or Git Bash: download and run the installer for your PC from the official Git for Windows page. Find your system type in Windows Settings > System > About. Git Bash is included. On the PATH screen, select Git from the command line and also from 3rd-party software (recommended). Do not assume unrelated startup errors mean Git is missing.",
      "If Node.js is missing or below the minimum version: choose Windows and LTS on the official download page, then download and run Windows Installer (.msi). npm is included. Leave the additional installation checkbox under Tools for Native Modules unchecked; it adds Chocolatey, Python, Visual Studio Build Tools, and other components."
    ],
    mac: [
      "If Git is needed: find Terminal with Spotlight, open it, enter xcode-select --install and press Enter. Follow Apple's Command Line Tools installation dialog. You do not need the full Xcode app or Homebrew.",
      "If Node.js is needed: choose macOS and LTS on the official download page, then download and run macOS Installer (.pkg). npm is included."
    ],
    blocked: "Official Node.js and Git installers, Apple's Command Line Tools on Mac, and existing winget are permitted by the company. Review Claude's explanation and approve before installation. Handle administrator dialogs yourself; never enter passwords in Claude chat. If installation is actually blocked, contact IT without bypassing restrictions. After installation, fully quit and reopen Claude in Local once, keeping any already selected folder. If tools remain unavailable, preparation is incomplete: do not proceed to Hare connection; contact IT.",
    prompt: [
      "Check this PC's Hare prerequisites now. This is an actual environment preparation request, not a request to review the prompt. Start with read-only checks; explain any installation or upgrade and obtain separate user approval.",
      "Prepare only the prerequisites on this PC before connecting Hare. Use available host/execution context to verify Claude Desktop > Code > Local and the actual OS. Only if that context cannot distinguish the mode, ask the user once whether it is Local; do not infer it from OS or path alone. Stop in Cloud/SSH/WSL. These prerequisite checks do not require a selected Hare folder. Defer folder selection/creation to Hare connection; only if the app itself requires a working folder, ask the user to select a local folder. Do not install Hare, run git clone, npm ci, builds, Microsoft sign-in, or work lookups in this step. Do not read or change existing Hare data files or authentication caches.",
      "Check node --version, npm --version, and git --version read-only in the same session. Distinguish ready / installation needed / upgrade needed / installed but not on PATH / unable to verify. Node.js >=20.18.1 is Hare's minimum. Below that minimum, classify as upgrade needed and follow the same approval process as installation. New installations and upgrades use the currently supported LTS from the official site. Node 20.x is end-of-life: distinguish functional compatibility from maintenance support and recommend a supported LTS upgrade with approval. If declined, report the EOL warning without blocking execution that meets the technical minimum. Do not reinstall or update working tools that meet the requirements without user approval. Check Git with git --version and --show-current in git branch -h help; a help command's exit code alone is not an installation failure. Propose an upgrade only if required options are absent. Do not interpret permission denial or network failure as missing software. If git opens an installation dialog on Mac, inform the user without approving it automatically.",
      "Do not immediately propose installation when a command is missing. Even in a new session, inspect standard locations, OS installation records, and existing version-manager locations read-only, limited to the relevant tools. File existence is not readiness: check versions using the absolute path of the discovered tool. If absolute-path execution works but the ordinary command does not, classify as installed but not on PATH and do not reinstall. Absence from standard locations alone does not prove absence. If installation status or prior attempts remain uncertain, ask the user about recent installation or custom locations and stop. Do not scan the whole disk or read authentication caches.",
      "For prerequisites needing installation or upgrade only, show the name, reason, official source, installation method, changes, and administrator requirements, then wait for explicit approval. Use the methods specified in the company authorization above. Obtain approval again if the method or target changes. npm comes with Node.js; do not upgrade it independently. If a version manager already exists, check that management method before adding a conflicting installation.",
      "Only after approval, execute that installation with permitted standard tools. If an installer dialog, administrator approval, or unavailable tool requires the user, explain the necessary clicks and wait. Never request or read passwords. Stop on permission denial or policy blocks and refer to internal IT. Do not independently change security settings, execution policy, or PATH, or bypass permissions. Do not newly install Claude CLI, WSL, Docker, or another package manager for this preparation.",
      "After installing, ask the user to fully quit Claude and reopen Local, keeping any already selected folder. On resuming this prompt, recheck versions and say environment ready, proceed to Hare connection only when all tools work. If tools remain unavailable after restart, report the error and stop without repeated reinstalls. Do not present environment readiness as successful Hare sign-in."
    ],
    platformPrompt: {
      windows: "Target OS: native Windows Local. Official sources: https://nodejs.org/en/download and https://git-scm.com/install/windows . Check npm with npm.cmd --version; an npm.ps1 policy error does not mean it is missing. If commands are unavailable, inspect nodejs and Git/cmd under the actual ProgramFiles (normally C:\\Program Files\\nodejs and C:\\Program Files\\Git\\cmd), Programs/Git under LocalAppData, OS installation records, and already identified version-manager locations read-only. Company permission for existing winget is confirmed for this preparation. Use Node LTS=OpenJS.NodeJS.LTS and Git=Git.Git with --exact --source winget for lookup. After approval use --interactive so the user can review installer options; do not use --silent or arbitrary --override. Select the Node MSI that includes npm, not a portable ZIP. Before installation, show the exact package ID, publisher, source, version, and changes and obtain explicit approval. If winget is missing or blocked, do not install or repair it or bypass restrictions; explain manual installation using the permitted official installers. Stop if that installation is actually blocked too. Do not upgrade all packages or add/change sources. The user handles UAC directly. Leave Node's Tools for Native Modules unchecked. Select Git from the command line and also from 3rd-party software for Git's PATH option and include this PATH change in the pre-approval explanation. Do not make separate arbitrary PATH edits. Never pass PowerShell syntax directly to Git Bash.",
      mac: "Target OS: Mac Local, POSIX shell. Official sources: https://nodejs.org/en/download and https://git-scm.com/install/mac . If commands are missing, inspect /usr/local/bin/node, /usr/local/bin/npm, /usr/bin/git, Apple installation records, and already identified version-manager locations read-only. xcode-select --install requests Apple's Command Line Tools installer for Git; do not run it before user approval."
    }
  },
  ja: {
    heading: "2. Claude に環境の準備を依頼",
    intro: "以下の環境準備プロンプトを Claude Code Local のチャットに貼り付けてください。Claude が必要なツールを確認し、不足があればインストール方法を説明して承認を求めます。すべて準備できたら手順 3 に進んでください。",
    label: "環境準備プロンプト",
    copy: "環境準備プロンプトをコピー",
    policy: "社内許可方針: この Hare 環境準備では、nodejs.org の公式 Node.js インストーラー、git-scm.com から案内される公式 Git インストーラー、Mac の Apple Command Line Tools、Windows の既存 winget の利用が許可されています。社員に会社の許可を再確認させないでください。個別のインストール・アップグレードの説明と明示的なユーザー承認は別途必要です。この許可は管理者権限の付与や実際の OS・セキュリティ制限の回避を意味しません。",
    pathRecovery: "インストール済み・PATH 未反映は ENVIRONMENT_INCOMPLETE (環境準備未完了) であり、Hare 接続に進んではいけません。未再起動なら Claude を完全に終了し、Local で一度だけ開き直すよう案内します。新しいセッションでは、この問題で既に再起動したかを先に確認し、実施済みなら再び要求しません。再起動後のセッションで通常のコマンドによりバージョンを再確認します (Windows は npm.cmd --version)。まだ認識されなければ ENVIRONMENT_INCOMPLETE と対象ツール・エラーを報告し、IT 確認を案内してください。再インストール・手動 PATH 編集・絶対パスによる Hare 接続への迂回はしません。バージョンの条件を満たし、通常のコマンドがすべて動く場合のみ環境準備完了とします。",
    origin: {
      windows: "既存ツールは Get-Command で実際の実行パスとリンク先を確認し、バージョンコマンドが成功するかを確認します。標準の保存先、OS のインストール記録、nvm-windows など既存のバージョン管理ツールの場所を認めます。既存ツールの署名確認は任意の診断です。署名の有無・確認失敗・署名者の違いだけで正常動作を止めないでください。この PC で確認した署名者は Node.js が OpenJS Foundation、Git for Windows が Johannes Schindelin ですが、固定の許可リストではありません。Node と同じフォルダーの npm.cmd はそのまま使います。グローバル更新などで別の npm が選ばれた場合は場所の違いだけを報告し、正常実行を止めません。npm スクリプトチェーンの解析やパッケージの再ダウンロード・展開・ハッシュ比較は行いません。新しいインストーラーは別途、公式取得元、Get-AuthenticodeSignature -LiteralPath の Valid、公式配布者の一致を確認し、ユーザー承認後に実行します。新しいインストーラーの出所・署名を確認できない場合や OS が実際に遮断する場合は実行しません。",
      mac: "既存ツールは command -v で実際の実行パスとリンク先を確認し、バージョンコマンドが成功するかを確認します。標準の保存先、インストール記録、既存のバージョン管理ツールの場所を認めます。既存ツールの codesign 検証は必須ではなく、署名や記録の不足だけで正常動作を止めません。Node と同じインストールの npm はそのまま使い、グローバル更新などで別の場所の npm が選ばれた場合は違いだけ報告します。npm スクリプトチェーンの解析やパッケージの再ダウンロード・展開・ハッシュ比較は行いません。新しい .pkg は別途、公式取得元と pkgutil --check-signature の有効な署名・公式配布者を確認し、ユーザー承認後に実行します。確認失敗や実際の OS 遮断では実行しません。Gatekeeper は解除しません。"
    },
    manual: "Code が開かない、または手動インストールが必要な場合",
    manualCheck: "既にインストールしたツールを再インストールしないでください。直前にインストールした場合も、まず Claude を完全に終了して開き直します。エラーが続く場合やインストール済みか不明な場合は社内 IT 担当者に確認してください。以下は未インストールまたは承認済みのアップグレードと確認できた場合だけ進めます。新しくダウンロードしたインストーラーは実行前に公式取得元・有効なデジタル署名・公式配布者を確認してください。確認が難しい場合は実行前に IT 担当者の支援を受けてください。既存ツールにこの確認は要求しません。",
    windows: [
      "Git または Git Bash が必要というエラーが出る場合: Git for Windows の公式ページから PC に合うインストーラーをダウンロードして実行します。システムの種類は Windows の設定 > システム > バージョン情報で確認できます。Git Bash は同梱されています。PATH の画面では Git from the command line and also from 3rd-party software (推奨) を選択してください。他の起動エラーを Git 不足と決めつけないでください。",
      "Node.js が未インストールまたは最低バージョン未満の場合: 公式ダウンロードページで Windows と LTS を選び、Windows Installer (.msi) をダウンロードして実行します。npm も含まれます。Tools for Native Modules の追加インストールはチェックしないでください。Chocolatey、Python、Visual Studio Build Tools などが追加されるためです。"
    ],
    mac: [
      "Git が必要な場合: Spotlight でターミナルを探して開き、xcode-select --install を入力して Enter を押します。Apple の画面で Command Line Tools のインストールを進めます。Xcode 全体や Homebrew は不要です。",
      "Node.js が必要な場合: 公式ダウンロードページで macOS と LTS を選び、macOS Installer (.pkg) をダウンロードして実行します。npm も含まれます。"
    ],
    blocked: "公式 Node.js・Git インストーラー、Mac の Apple Command Line Tools、既存 winget は社内で利用が許可されています。インストール前に Claude の説明を確認して承認してください。管理者承認画面は直接操作し、パスワードをチャットに入力しないでください。実際に遮断されたら回避せず IT 担当者に依頼してください。インストール後は Claude を完全に終了し、Local で一度開き直します。選択済みフォルダーは維持します。それでも認識されない場合は環境準備未完了のため Hare 接続へ進まず、IT 担当者に確認してください。",
    prompt: [
      "この PC の Hare 実行環境を今確認してください。プロンプトのレビューではなく実際の環境準備の依頼です。読み取り専用の確認から開始し、インストール・更新は対象と方法を説明して別途ユーザー承認を得てください。",
      "Hare 接続前に、この PC の必須ツールだけを準備してください。提供されたホスト・実行環境情報で Claude Desktop > Code > Local と実際の OS を確認します。判別できない場合だけ Local か一度確認し、OS やパスだけで推測しません。Cloud/SSH/WSL では停止してください。この環境準備は Hare フォルダー未選択でも進められます。フォルダー選択・作成は Hare 接続時に行い、アプリ自体が作業フォルダーを要求する場合のみローカルフォルダーの選択を案内してください。この段階では Hare のインストール、git clone、npm ci、ビルド、Microsoft ログイン、業務検索を行わないでください。既存の Hare データファイルや認証キャッシュを読んだり変更したりしないでください。",
      "同じセッションで node --version、npm --version、git --version を読み取り専用で確認し、準備済み/インストール必要/アップグレード必要/インストール済み・PATH 未反映/確認不可を区別してください。Node.js >=20.18.1 は Hare の最低条件です。最低バージョン未満はアップグレード必要と判定し、インストールと同じ承認手順を適用します。新規インストールとアップグレードには公式サイトで現在サポートされている LTS を使用します。Node 20.x はサポート終了のため実行可否と保守サポートを区別し、承認を得てサポート中の LTS への更新を勧めてください。拒否されたら警告を伝えますが、技術的最低条件を満たす実行は止めません。要件を満たして利用可能なツールはユーザー承認なしで再インストール・更新しないでください。Git は git --version と git branch -h の --show-current 対応を確認し、ヘルプの終了コードだけをインストール失敗としません。必要なオプションがなければ更新を提案します。権限拒否やネットワーク障害を未インストールと判断しません。Mac の git がインストール画面を開いたら、自動承認せずユーザーに伝えてください。",
      "コマンドが見つからなくても、すぐにインストールを提案しないでください。新しいセッションでも標準の保存先と OS のインストール記録、既存のバージョン管理ツールの場所を必要な範囲だけ読み取り専用で確認します。ファイルの存在だけで準備完了とせず、見つかったツールの絶対パスでバージョンを確認してください。絶対パスでは動くが通常のコマンドでは動かない場合はインストール済み・PATH 未反映と判定し、再インストールしません。標準の保存先にないだけで未インストールと断定しないでください。状態や過去の試行が不明なら、最近のインストールや別の保存先をユーザーに確認して停止します。ディスク全体の検索や認証キャッシュの読み取りは禁止です。",
      "インストール・アップグレードが必要な項目だけについて、名前・理由・公式配布元・方法・変更範囲・管理者権限の要否を示し、明示的承認を待ってください。上記の社内許可方針に明記した方法を使用します。方法や対象が変われば再承認を得てください。npm は Node.js と一緒にインストールし、単独で最新版に更新しません。既存のバージョン管理ツールがあれば、競合するインストールを追加せず現状を先に確認してください。",
      "承認後に限り、許可された標準ツールで対象のインストールを実行します。インストール画面・管理者承認・利用できないツールなどでユーザー操作が必要なら、必要なクリックを案内して待ってください。パスワードを要求したり読んだりしません。権限拒否やポリシーによる遮断では停止し、社内 IT 担当者への確認を案内してください。セキュリティ設定・実行ポリシー・PATH を勝手に変更したり、権限を回避したりしないでください。この準備のために Claude CLI、WSL、Docker、別のパッケージ管理ツールを新規インストールしません。",
      "インストールしたら Claude を完全に終了し、Local で開き直すよう案内してください。選択済みフォルダーは維持します。このプロンプトを再開したらバージョンを再確認し、すべて使える場合だけ環境準備完了と伝えて Hare 接続へ案内します。再起動後も見つからなければ再インストールを繰り返さずエラーを報告して停止します。環境準備完了を Hare ログイン成功として扱わないでください。"
    ],
    platformPrompt: {
      windows: "対象 OS: Windows ネイティブ Local。公式配布元は https://nodejs.org/en/download と https://git-scm.com/install/windows です。npm は npm.cmd --version で確認し、npm.ps1 ポリシーエラーを未インストールと扱わないでください。コマンドが見つからない場合、実際の ProgramFiles 内の nodejs と Git/cmd (通常 C:\\Program Files\\nodejs、C:\\Program Files\\Git\\cmd)、LocalAppData 内の Programs/Git、OS のインストール記録、確認済みのバージョン管理ツールの保存先だけを読み取り専用で確認します。この準備作業での既存 winget の社内許可は確認済みです。Node LTS=OpenJS.NodeJS.LTS、Git=Git.Git を --exact --source winget で照会してください。承認後のインストールは --interactive でオプションをユーザーが確認し、--silent や任意の --override は使いません。Node は npm を含む MSI を選び、portable ZIP は使いません。インストール前にパッケージ ID・配布者・取得元・バージョン・変更範囲を示し、明示的承認を得てください。winget がない、または遮断されている場合は新規インストール・修復・制限回避をせず、許可済み公式インストーラーの手動手順を案内します。そのインストールも実際に遮断されれば停止します。全パッケージの一括更新やソースの追加・変更は禁止です。UAC はユーザーが直接操作します。Node の Tools for Native Modules はチェックしません。Git の PATH は Git from the command line and also from 3rd-party software を選び、この PATH 変更も承認前に説明します。任意の別途 PATH 編集はしません。PowerShell 構文を Git Bash に直接渡さないでください。",
      mac: "対象 OS: Mac Local、POSIX シェル。公式配布元は https://nodejs.org/en/download と https://git-scm.com/install/mac です。コマンドがない場合、/usr/local/bin/node、/usr/local/bin/npm、/usr/bin/git、Apple のインストール記録、確認済みのバージョン管理ツールの保存先だけを読み取り専用で確認します。Git 用 Apple Command Line Tools のインストール画面は xcode-select --install で要求できますが、ユーザー承認前には実行しないでください。"
    }
  }
};

export function buildEnvironmentPrompt(language, platform) {
  const text = environmentCopy[language];
  if (!text || !Object.hasOwn(SHELLS, platform)) throw new Error("Unsupported guide language/platform.");
  return [text.prompt[0], text.policy, text.prompt[1], text.origin[platform], ...text.prompt.slice(2), text.platformPrompt[platform], text.pathRecovery].join("\n\n");
}

export function buildGuideCommand(shell) {
  if (!Object.values(SHELLS).includes(shell)) throw new Error(`Unsupported shell: ${shell}`);
  const command = buildLocalSetupCommand({
    dataDir: SELECTED_FOLDER,
    repository: "https://github.com/ohmyhotelco-planning/hare-m365-agent.git",
    branch: "master",
    environment: "claude-code",
    shell
  });
  for (const [flag, value] of [["host", "claude-code"], ["command-shell", shell], ["environment", "claude-code"]]) {
    if (!new RegExp(`--${flag}(?:\\s+['\"]?|['\"]\\s*,\\s*['\"])${value}(?:['\"]|\\s|$)`).test(command)) {
      throw new Error(`Installer API is not ready: missing --${flag} ${value}. Rebuild dist/local-install.js after integration.`);
    }
  }
  if (!command.includes(SELECTED_FOLDER)) throw new Error("Installer discarded the selected-folder placeholder.");
  return command;
}

export function buildPrompt(language, platform) {
  const text = copy[language];
  if (!text || !Object.hasOwn(SHELLS, platform)) throw new Error("Unsupported guide language/platform.");
  return [text.request, ...text.before, text[platform], text.command, "```" + SHELLS[platform] + "\n" + buildGuideCommand(SHELLS[platform]) + "\n```", ...text.after].join("\n\n");
}

function escapeHtml(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}

export function renderGuide(language) {
  const text = copy[language];
  if (!text) throw new Error(`Unsupported guide language: ${language}`);
  const environment = environmentCopy[language];
  const paragraph = (key) => `<p>${escapeHtml(text[key])}</p>`;
  return `<!doctype html>
<html lang="${language}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(text.title)} | Claude Code Local</title>
  <style>
    :root { --brand: #FF6000; --text: #1f2328; --muted: #59636e; --line: #d8dee4; --code: #f6f8fa; }
    * { box-sizing: border-box; letter-spacing: 0; }
    body { margin: 0; color: var(--text); background: #fff; font: 16px/1.6 "Segoe UI", Arial, sans-serif; overflow-wrap: anywhere; }
    header { border-top: 6px solid var(--brand); border-bottom: 1px solid var(--line); }
    header > div, main { max-width: 980px; margin: auto; padding: 24px 32px; }
    main { padding-bottom: 48px; }
    h1 { font-size: 28px; line-height: 1.3; margin: 0 0 8px; }
    h2 { font-size: 21px; line-height: 1.4; margin: 0 0 16px; border-left: 4px solid var(--brand); padding-left: 12px; }
    h3 { font-size: 17px; margin: 14px 0 8px; }
    p { margin: 10px 0; }
    section { padding: 26px 0; border-bottom: 1px solid var(--line); }
    .muted { color: var(--muted); }
    .notice { border-left: 4px solid var(--brand); padding: 8px 16px; margin: 0 0 12px; background: var(--code); }
    label { display: block; font-weight: 600; margin-bottom: 6px; }
    select, button { font: inherit; border-radius: 6px; min-height: 44px; padding: 8px 14px; max-width: 100%; }
    select { background: #fff; color: var(--text); border: 1px solid var(--muted); min-width: 180px; }
    button { background: var(--brand); color: #1f2328; border: 1px solid var(--brand); font-weight: 700; cursor: pointer; }
    :focus-visible { outline: 3px solid #0969da; outline-offset: 3px; }
    textarea { display: block; width: 100%; height: 360px; resize: vertical; padding: 14px; border: 1px solid var(--line); border-radius: 6px; color: var(--text); background: var(--code); font: 13px/1.6 Consolas, "Courier New", monospace; white-space: pre-wrap; overflow-wrap: anywhere; }
    [hidden] { display: none !important; }
    .controls { margin-top: 12px; display: flex; flex-wrap: wrap; gap: 12px; align-items: center; }
    .copy-status { min-height: 26px; flex: 1 1 220px; }
    details { margin: 18px 0; padding: 12px 0; border-top: 1px solid var(--line); }
    summary { cursor: pointer; font-weight: 600; }
    li { margin: 10px 0; }
    a { color: #0969da; text-underline-offset: 3px; }
    footer { padding-top: 24px; font-size: 14px; }
    footer ul { margin: 8px 0; padding-left: 20px; }
    @media (max-width: 600px) { header > div, main { padding-left: 18px; padding-right: 18px; } h1 { font-size: 25px; } textarea { height: 330px; } }
  </style>
</head>
<body>
  <header><div><h1>${escapeHtml(text.title)}</h1><p class="muted">${escapeHtml(text.subtitle)}</p></div></header>
  <main>
    <aside class="notice">${paragraph("preview")}${paragraph("open")}</aside>
    <label for="platform">${escapeHtml(text.os)}</label>
    <select id="platform"><option value="windows">Windows · PowerShell</option><option value="mac">Mac · POSIX</option></select>
    <section aria-labelledby="prepare"><h2 id="prepare">${escapeHtml(text.prepare)}</h2>${paragraph("prerequisites")}
      <details id="manual-setup"><summary>${escapeHtml(environment.manual)}</summary>
        <p>${escapeHtml(environment.manualCheck)}</p>
        ${Object.keys(SHELLS).map((platform) => `<div data-platform="${platform}"><h3>${platform === "windows" ? "Windows" : "Mac"}</h3><ol>${environment[platform].map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol><p><a href="https://git-scm.com/install/${platform}">${platform === "windows" ? "Git for Windows" : "Git / Apple Command Line Tools"}</a> · <a href="https://nodejs.org/en/download">Node.js LTS</a></p></div>`).join("\n")}
        <p>${escapeHtml(environment.blocked)}</p>
      </details>
    </section>
    <section aria-labelledby="environment"><h2 id="environment">${escapeHtml(environment.heading)}</h2><p>${escapeHtml(environment.intro)}</p>
      ${Object.keys(SHELLS).map((platform) => `<div data-platform="${platform}"><h3 id="environment-label-${platform}">${platform === "windows" ? "Windows" : "Mac"}</h3><textarea id="environment-${platform}" aria-label="${escapeHtml(environment.label)} (${platform})" aria-describedby="environment-label-${platform}" readonly spellcheck="false">${escapeHtml(buildEnvironmentPrompt(language, platform))}</textarea></div>`).join("\n")}
      <div class="controls"><button type="button" id="copy-environment" hidden>${escapeHtml(environment.copy)}</button><span class="copy-status" id="environment-status" role="status" aria-live="polite" data-success="${escapeHtml(text.copied)}" data-fallback="${escapeHtml(text.fallback)}"></span></div>
      ${paragraph("tools")}<noscript><p>${escapeHtml(text.nojs)}</p></noscript>
    </section>
    <section aria-labelledby="setup"><h2 id="setup">${escapeHtml(text.setup)}</h2>${paragraph("folder")}${paragraph("paste")}
      ${Object.keys(SHELLS).map((platform) => `<div data-platform="${platform}"><h3 id="label-${platform}">${platform === "windows" ? "Windows · PowerShell" : "Mac · POSIX"}</h3><textarea id="prompt-${platform}" aria-label="${escapeHtml(text.promptLabel)} (${platform})" aria-describedby="label-${platform}" readonly spellcheck="false">${escapeHtml(buildPrompt(language, platform))}</textarea></div>`).join("\n      ")}
      <div class="controls"><button type="button" id="copy" hidden>${escapeHtml(text.copy)}</button><span class="copy-status" id="copy-status" role="status" aria-live="polite" data-success="${escapeHtml(text.copied)}" data-fallback="${escapeHtml(text.fallback)}"></span></div>
      <noscript><p>${escapeHtml(text.nojs)}</p></noscript>
    </section>
    <section aria-labelledby="next"><h2 id="next">${escapeHtml(text.next)}</h2>${paragraph("state")}${paragraph("reuse")}</section>
    <section aria-labelledby="security"><h2 id="security">${escapeHtml(text.security)}</h2>${paragraph("privacy")}${paragraph("permission")}${paragraph("operations")}</section>
    <footer><strong>${escapeHtml(text.sources)}</strong><ul>
      <li><a href="https://code.claude.com/docs/en/desktop">Claude Code Desktop</a></li>
      <li><a href="https://code.claude.com/docs/en/setup">Setup</a></li>
      <li><a href="https://code.claude.com/docs/en/memory">CLAUDE.md / Memory</a></li>
      <li><a href="https://code.claude.com/docs/en/permissions">Permissions</a></li>
      <li><a href="https://docs.npmjs.com/downloading-and-installing-node-js-and-npm/">Node.js / npm</a></li>
    </ul></footer>
  </main>
  <script>
    const platform = document.getElementById("platform");
    const controls = [
      { button: document.getElementById("copy"), status: document.getElementById("copy-status"), prefix: "prompt-" },
      { button: document.getElementById("copy-environment"), status: document.getElementById("environment-status"), prefix: "environment-" }
    ];
    function updatePlatform() {
      document.querySelectorAll("[data-platform]").forEach((panel) => { panel.hidden = panel.dataset.platform !== platform.value; });
      controls.forEach(({ status }) => { status.textContent = ""; });
    }
    platform.addEventListener("change", updatePlatform);
    updatePlatform();
    controls.forEach(({ button, status, prefix }) => {
      button.hidden = false;
      button.addEventListener("click", async () => {
      const prompt = document.getElementById(prefix + platform.value);
      controls.forEach(({ button }) => { button.disabled = true; });
      platform.disabled = true;
      let copied = false;
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(prompt.value);
          copied = true;
        }
      } catch { /* Local-file clipboard restrictions use the selection fallback. */ }
      if (!copied) {
        prompt.focus();
        prompt.select();
        try { copied = document.execCommand("copy"); } catch { copied = false; }
      }
      status.textContent = copied ? status.dataset.success : status.dataset.fallback;
      controls.forEach(({ button }) => { button.disabled = false; });
      platform.disabled = false;
      });
    });
  </script>
</body>
</html>
`;
}

export function synchronizeGuides({ check = false } = {}) {
  // Render every language first so an unsupported installer cannot leave partial output.
  const guides = Object.entries(GUIDE_FILES).map(([language, file]) => ({
    file: path.join(root, "release-templates", "claude-code", file),
    bytes: Buffer.from(renderGuide(language), "utf8")
  }));
  const mismatches = [];
  for (const guide of guides) {
    if (check) {
      if (!fs.existsSync(guide.file) || !fs.readFileSync(guide.file).equals(guide.bytes)) mismatches.push(path.basename(guide.file));
    } else {
      fs.mkdirSync(path.dirname(guide.file), { recursive: true });
      fs.writeFileSync(guide.file, guide.bytes);
    }
  }
  if (mismatches.length) throw new Error(`Generated guides are missing or stale: ${mismatches.join(", ")}. Run node scripts/build-claude-code-guides.mjs.`);
  return guides.length;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const args = process.argv.slice(2);
    if (args.some((arg) => arg !== "--check")) throw new Error("Usage: node scripts/build-claude-code-guides.mjs [--check]");
    const check = args.includes("--check");
    console.log(`${check ? "Checked" : "Generated"} ${synchronizeGuides({ check })} Claude Code Local guides.`);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
