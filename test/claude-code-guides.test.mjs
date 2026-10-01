import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { buildLocalSetupCommand } from "../dist/local-install.js";
import { buildEnvironmentPrompt, buildGuideCommand, buildPrompt, GUIDE_FILES, renderGuide, SELECTED_FOLDER, SHELLS, synchronizeGuides } from "../scripts/build-claude-code-guides.mjs";

const root = fileURLToPath(new URL("../", import.meta.url));
const guidePath = (file) => path.join(root, "release-templates", "claude-code", file);
const decodeHtml = (value) => value.replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&quot;", '"').replaceAll("&#39;", "'").replaceAll("&amp;", "&");

const localizedInvariants = {
  ko: ["정확한 절대", "OneDrive 밖", "Git 저장소 밖", "치환하지 않은", "사용자가 로그인을 완료했다고 확인", "정확히 한 번", "캐시를 보존", "확인 불가이지 로그아웃이나 만료가 아니야", "전체 미리보기", "명시적 승인", "조건부", "모델에 전달", "원시 인증 캐시를 읽지 마", "거부", "같은 폴더"],
  en: ["exact absolute", "non-OneDrive", "outside Git repositories", "Never execute the literal angle-bracket placeholder", "user explicitly confirms sign-in completion", "exactly once", "Preserve the authentication folder and cache", "UNKNOWN, not signed out or expired", "complete recipients, subject, body, and attachments preview", "explicit approval", "conditional", "content to the model", "Never read raw authentication caches", "denied", "SAME folder"],
  ja: ["正確な絶対", "OneDrive 外", "Git リポジトリ外", "未置換のプレースホルダー", "ユーザーがログイン完了を明示的に確認", "正確に一度", "キャッシュを保持", "認証未確認であり、ログアウトや期限切れではありません", "完全なプレビュー", "明示的承認", "条件付き", "モデルに送信", "生の認証キャッシュは読みません", "拒否", "同じフォルダー"]
};

for (const [language, filename] of Object.entries(GUIDE_FILES)) {
  test(`${language}: standalone HTML is byte-for-byte generated and accessible`, () => {
    const html = renderGuide(language);
    assert.ok(fs.readFileSync(guidePath(filename)).equals(Buffer.from(html, "utf8")));
    assert.match(html, new RegExp(`<html lang="${language}">`));
    assert.match(html, /<meta charset="utf-8">/);
    assert.match(html, /--brand: #FF6000/);
    assert.doesNotMatch(html, /DRAFT|Draft \/|NOT RELEASED|draftGate/);
    assert.match(html, /<label for="platform">/);
    assert.match(html, /<select id="platform"><option value="windows">Windows · PowerShell<\/option><option value="mac">Mac · POSIX<\/option><\/select>/);
    assert.match(html, /role="status" aria-live="polite"/);
    assert.match(html, /<noscript>/);
    assert.equal([...html.matchAll(/<textarea\b/g)].length, 4);
    assert.ok(html.indexOf('id="platform"') < html.indexOf('id="prepare"'));
    assert.ok(html.indexOf('id="environment"') < html.indexOf('id="setup"'));
    assert.match(html, /<details id="manual-setup"><summary>/);
    const manual = html.match(/<details id="manual-setup">([\s\S]*?)<\/details>/)[1];
    const installerChecks = {
      ko: ["공식 출처·유효한 디지털 서명·공식 배포자", "실행 전에 IT 담당자의 도움", "기존 도구에는 이 검사를 요구하지 않습니다"],
      en: ["official source, valid digital signature, and official publisher", "Ask IT for help before execution", "Existing tools do not require this check"],
      ja: ["公式取得元・有効なデジタル署名・公式配布者", "実行前に IT 担当者の支援", "既存ツールにこの確認は要求しません"]
    };
    for (const phrase of installerChecks[language]) assert.ok(manual.includes(phrase), phrase);
    assert.doesNotMatch(html, /<details[^>]*\bopen\b/);
    for (const url of ["https://nodejs.org/en/download", "https://git-scm.com/install/windows", "https://git-scm.com/install/mac", "https://docs.npmjs.com/downloading-and-installing-node-js-and-npm/"]) assert.ok(html.includes(url));
    assert.doesNotMatch(html, /<img\b|<iframe\b|<script[^>]+src=|<link[^>]+href=|@import|url\(|fetch\(|XMLHttpRequest/i);
    assert.doesNotMatch(html, /dangerously-skip-permissions|bypassPermissions|Always allow|\/tmp\/|\/home\/claude|\/dev\/shm/);
    for (const page of ["desktop", "setup", "memory", "permissions"]) {
      assert.ok(html.includes(`https://code.claude.com/docs/en/${page}`));
    }
  });

  for (const [platform, shell] of Object.entries(SHELLS)) {
    test(`${language}/${platform}: upgrades and missing-command recovery avoid reinstall loops`, () => {
      const prompt = buildEnvironmentPrompt(language, platform);
      const safeguards = {
        ko: ["최소 버전 미만은 업그레이드 필요", "설치와 같은 승인 절차", "요구 버전을 충족", "표준 설치 위치", "OS의 설치 기록", "절대 경로", "설치됨·PATH 미반영", "미설치로 단정하지 마", "최근 설치 여부", "디스크 전체를 탐색"],
        en: ["Below that minimum, classify as upgrade needed", "same approval process as installation", "meet the requirements", "standard locations", "OS installation records", "absolute path", "installed but not on PATH", "alone does not prove absence", "recent installation", "Do not scan the whole disk"],
        ja: ["最低バージョン未満はアップグレード必要", "インストールと同じ承認手順", "要件を満たして", "標準の保存先", "OS のインストール記録", "絶対パス", "インストール済み・PATH 未反映", "未インストールと断定しない", "最近のインストール", "ディスク全体の検索"]
      };
      for (const phrase of safeguards[language]) assert.ok(prompt.includes(phrase), phrase);
      if (platform === "windows") {
        for (const phrase of ["ProgramFiles", "LocalAppData", "C:\\Program Files\\nodejs", "C:\\Program Files\\Git\\cmd", "winget", "UAC", "Tools for Native Modules", "Git from the command line and also from 3rd-party software"]) assert.ok(prompt.includes(phrase), phrase);
        const wingetRules = {
          ko: ["사내 허용이 확인돼 있어", "패키지 ID·배포자·출처·설치 버전", "명시적 승인을 받아", "새로 설치·복구하거나 우회하지 말고", "전체 패키지 일괄 업데이트나 소스 추가·변경은 하지 마", "UAC는 사용자가 직접", "추가 도구는 선택하지 마", "PATH 변경도 승인 전 설명"],
          en: ["Company permission for existing winget is confirmed", "package ID, publisher, source, version", "obtain explicit approval", "do not install or repair it or bypass restrictions", "Do not upgrade all packages or add/change sources", "user handles UAC directly", "Tools for Native Modules unchecked", "PATH change in the pre-approval explanation"],
          ja: ["既存 winget の社内許可は確認済み", "パッケージ ID・配布者・取得元・バージョン", "明示的承認を得て", "新規インストール・修復・制限回避をせず", "全パッケージの一括更新やソースの追加・変更は禁止", "UAC はユーザーが直接", "Tools for Native Modules はチェックしません", "PATH 変更も承認前に説明"]
        };
        for (const phrase of wingetRules[language]) assert.ok(prompt.includes(phrase), phrase);
        const manual = renderGuide(language).match(/<details id="manual-setup">([\s\S]*?)<\/details>/)[1];
        for (const phrase of ["Tools for Native Modules", "Chocolatey", "Python", "Visual Studio Build Tools", "Git from the command line and also from 3rd-party software"]) assert.ok(manual.includes(phrase), phrase);
      } else {
        for (const location of ["/usr/local/bin/node", "/usr/local/bin/npm", "/usr/bin/git"]) assert.ok(prompt.includes(location));
        assert.doesNotMatch(prompt, /C:\\Program Files|UAC|Tools for Native Modules/);
      }
    });
    test(`${language}/${platform}: approved company policy, bounded PATH recovery and origin criteria`, () => {
      const prompt = buildEnvironmentPrompt(language, platform);
      const criteria = {
        ko: ["직원에게 회사 허용 여부를 다시 확인하도록 요구하지 마", "사용자 승인은 별개로 필요", "한 번만", "이미 재시작했는지", "또 재시작을 요청하지 마", "ENVIRONMENT_INCOMPLETE", "Hare 연결 단계로 진행하지 마", "IT 확인을 안내", "서명"],
        en: ["Do not ask employees to reconfirm company permission", "user's explicit approval separately", "reopen Local once", "already restarted for this issue", "do not request another restart", "ENVIRONMENT_INCOMPLETE", "not ready for Hare connection", "refer to IT", "signature"],
        ja: ["社員に会社の許可を再確認させない", "ユーザー承認は別途必要", "一度だけ", "既に再起動したか", "実施済みなら再び要求しません", "ENVIRONMENT_INCOMPLETE", "Hare 接続に進んではいけません", "IT 確認を案内", "署名"]
      };
      for (const phrase of criteria[language]) assert.ok(prompt.includes(phrase), phrase);
      assert.doesNotMatch(prompt, /정책이 불명확하면|If policy is unclear|会社の方針で許可されている場合のみ/);
      const signatureCommand = platform === "windows" ? "Get-AuthenticodeSignature -LiteralPath" : "pkgutil --check-signature";
      assert.ok(prompt.includes(signatureCommand));
      const functionalChecks = {
        ko: ["버전 명령이 성공", "기존 버전 관리자 위치", "위치 차이만 보고", "패키지 재다운로드·압축 해제·해시 대조는 하지 마", "재시작 후 세션에서", "일반 명령이 모두 동작할 때만"],
        en: ["version-command success", "existing version-manager locations", "report the location difference", "Do not analyze the npm script chain or redownload, extract, or hash-compare", "in the session after restart", "all ordinary commands work"],
        ja: ["バージョンコマンドが成功", "既存のバージョン管理ツール", "違いだけ", "再ダウンロード・展開・ハッシュ比較は行いません", "再起動後のセッションで", "通常のコマンドがすべて動く場合のみ"]
      };
      for (const phrase of functionalChecks[language]) assert.ok(prompt.includes(phrase), phrase);
      assert.doesNotMatch(prompt, /NotSigned|npm-prefix\.js|node_modules\/npm\/bin\/npm-cli\.js|codesign --verify|재시작 뒤 같은 세션에서/);
      if (platform === "windows") {
        for (const phrase of ["Get-Command", "Valid", "npm.cmd", "nvm-windows", "OpenJS Foundation", "Johannes Schindelin"]) assert.ok(prompt.includes(phrase));
        const diagnostics = {
          ko: ["서명 확인은 선택적 진단", "정상 작동하는 도구를 차단하지 마", "고정 허용 목록은 아니야", "새 설치 파일의 서명·출처 확인이 실패하거나 OS가 실제 차단하면 실행하지 마"],
          en: ["signatures are optional diagnostics", "do not block working tools solely", "not a fixed allowlist", "Do not run a new installer if source/signature verification fails or the OS actually blocks it"],
          ja: ["署名確認は任意の診断", "正常動作を止めない", "固定の許可リストではありません", "新しいインストーラーの出所・署名を確認できない場合や OS が実際に遮断する場合は実行しません"]
        };
        for (const phrase of diagnostics[language]) assert.ok(prompt.includes(phrase), phrase);
      } else {
        for (const phrase of ["command -v", "pkgutil --check-signature", "Gatekeeper"]) assert.ok(prompt.includes(phrase));
        const optional = { ko: "codesign 검증은 필수가 아니며", en: "codesign verification is not required", ja: "codesign 検証は必須ではなく" };
        assert.ok(prompt.includes(optional[language]));
      }
    });
    test(`${language}/${platform}: environment preparation is separate and approval-gated`, () => {
      const prompt = buildEnvironmentPrompt(language, platform);
      assert.doesNotMatch(prompt, /DRAFT|NOT RELEASED|FOLDER_REQUIRED/);
      const runnable = {
        ko: ["실제 환경 준비 요청", "Hare 작업 폴더 선택 없이", "Local인지 한 번", "Node 20.x는 지원 종료", "실행 자체를 차단하지 마"],
        en: ["actual environment preparation request", "do not require a selected Hare folder", "ask the user once", "Node 20.x is end-of-life", "without blocking execution"],
        ja: ["実際の環境準備の依頼", "Hare フォルダー未選択でも", "Local か一度", "Node 20.x はサポート終了", "実行は止めません"]
      };
      for (const phrase of runnable[language]) assert.ok(prompt.includes(phrase), phrase);
      for (const phrase of ["git branch -h", "--show-current"]) assert.ok(prompt.includes(phrase));
      if (platform === "windows") {
        for (const phrase of ["OpenJS.NodeJS.LTS", "Git.Git", "--exact --source winget", "--interactive", "MSI"]) assert.ok(prompt.includes(phrase));
      }
      for (const invariant of ["Code > Local", "Cloud/SSH/WSL", "node --version", "npm --version", "git --version", "Node.js >=20.18.1", "LTS", "PATH", "https://nodejs.org/en/download", `https://git-scm.com/install/${platform}`]) assert.ok(prompt.includes(invariant), invariant);
      const approvals = {
        ko: ["명시적 승인을 기다려", "승인 후에만", "캐시를 읽거나 변경하지 마", "보안 설정·실행 정책·PATH를 임의로 바꾸거나", "관리자 승인이 필요", "암호를 요청하거나 읽지 마", "다시 설치·업데이트하지 마", "재설치 반복 없이", "사내 IT 담당자", "npm은 Node.js와 함께 설치", "Hare 로그인 성공으로 표현하지 마"],
        en: ["wait for explicit approval", "Only after approval", "Do not read or change existing Hare data files or authentication caches", "Do not independently change security settings, execution policy, or PATH", "administrator approval", "Never request or read passwords", "Do not reinstall or update working tools", "without repeated reinstalls", "internal IT", "npm comes with Node.js", "Do not present environment readiness as successful Hare sign-in"],
        ja: ["明示的承認を待って", "承認後に限り", "認証キャッシュを読んだり変更したりしない", "セキュリティ設定・実行ポリシー・PATH を勝手に変更", "管理者承認", "パスワードを要求したり読んだりしません", "再インストール・更新しない", "再インストールを繰り返さず", "社内 IT 担当者", "npm は Node.js と一緒にインストール", "Hare ログイン成功として扱わない"]
      };
      for (const phrase of approvals[language]) assert.ok(prompt.includes(phrase), phrase);
      assert.doesNotMatch(prompt, /```|--data-dir|auth login-start|winget install|curl\s|sudo\s|Set-ExecutionPolicy/);
      if (platform === "windows") assert.match(prompt, /npm\.cmd/);
      else assert.match(prompt, /xcode-select --install/);
      const html = renderGuide(language);
      const textarea = html.match(new RegExp(`<textarea id="environment-${platform}"[^>]*>([\\s\\S]*?)<\\/textarea>`));
      assert.equal(decodeHtml(textarea[1]), prompt);
    });
    test(`${language}/${platform}: installer source and safety invariants are preserved`, () => {
      const command = buildGuideCommand(shell);
      const expected = buildLocalSetupCommand({
        dataDir: "<SELECTED_LOCAL_FOLDER>",
        repository: "https://github.com/ohmyhotelco-planning/hare-m365-agent.git",
        branch: "master",
        environment: "claude-code",
        shell
      });
      assert.equal(command, expected);
      const prompt = buildPrompt(language, platform);
      assert.doesNotMatch(prompt, /DRAFT|NOT YET RELEASED|미배포|未リリース/);
      const hostChecks = {
        ko: ["제공된 호스트·실행 환경 정보", "Local인지 한 번", "OS나 경로만으로 Local이라고 추정하지 마"],
        en: ["available host/execution context", "ask the user once", "Do not infer Local from the OS or path alone"],
        ja: ["提供されたホスト・実行環境情報", "Local かユーザーに一度", "OS やパスだけで Local と推測しません"]
      };
      for (const phrase of hostChecks[language]) assert.ok(prompt.includes(phrase), phrase);
      assert.ok(prompt.includes("master"), "Installer uses the production release branch");
      assert.equal(prompt.split("```" + shell + "\n").length, 2);
      assert.ok(prompt.includes("```" + shell + "\n" + command + "\n```"));
      for (const invariant of [SELECTED_FOLDER, "--host claude-code", `--command-shell ${shell}`, "--data-dir", "node --version", "npm --version", "git --version", "Node.js >=20.18.1", "PATH", "setup.state", "setup.nextCommand", "LOGIN_START_REQUIRED", "LOGIN_COMPLETE_REQUIRED", "READY", "FOLDER_REQUIRED", "SETUP_REQUIRED", "AUTH_CHECK_BLOCKED", "loggedIn=true", "tokenUsable=true", "loggedIn=null", "tokenUsable=null", "network check --environment claude-code", "AWAITING_USER_APPROVAL", "CLAUDE.md", "claude/hare-m365-agent-rules.md", "Computer Use", ...localizedInvariants[language]]) {
        assert.ok(prompt.includes(invariant), `Missing ${language}/${platform} invariant: ${invariant}`);
      }
      const closingFence = prompt.indexOf("\n```", prompt.indexOf("```" + shell));
      const afterInstaller = prompt.slice(closingFence);
      assert.ok(afterInstaller.indexOf("claude/hare-m365-agent-rules.md") < afterInstaller.indexOf("LOGIN_START_REQUIRED"));
      if (platform === "windows") {
        assert.match(prompt, /npm\.cmd --version/);
        assert.match(prompt, /PowerShell/);
        assert.match(prompt, /Bash/);
        assert.ok(prompt.includes("&"));
        assert.doesNotMatch(command, /--command-shell\s+['"]?posix/);
      } else {
        assert.match(prompt, /POSIX/);
        assert.doesNotMatch(command, /--command-shell\s+['"]?powershell/);
      }
      assert.doesNotMatch(prompt, /--environment\s+['"]?cowork/);
      const html = renderGuide(language);
      const textarea = html.match(new RegExp(`<textarea id="prompt-${platform}"[^>]*>([\\s\\S]*?)<\\/textarea>`));
      assert.ok(textarea, "Selected prompt exists");
      assert.equal(decodeHtml(textarea[1]), prompt);
      assert.equal(textarea[1].includes(SELECTED_FOLDER), false, "Placeholder must be HTML-escaped");
    });
  }
}

test("--check verifies all generated bytes without changing files", () => {
  const before = Object.values(GUIDE_FILES).map((file) => ({ file, bytes: fs.readFileSync(guidePath(file)), mtime: fs.statSync(guidePath(file)).mtimeMs }));
  assert.equal(synchronizeGuides({ check: true }), 3);
  const result = spawnSync(process.execPath, ["scripts/build-claude-code-guides.mjs", "--check"], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /Checked 3 Claude Code Local guides/);
  for (const { file, bytes, mtime } of before) {
    assert.ok(fs.readFileSync(guidePath(file)).equals(bytes));
    assert.equal(fs.statSync(guidePath(file)).mtimeMs, mtime);
  }
});

test("invalid options fail without generating artifacts", () => {
  assert.throws(() => buildGuideCommand("bash"), /Unsupported shell/);
  assert.throws(() => buildPrompt("fr", "windows"), /Unsupported guide/);
  assert.throws(() => buildPrompt("en", "cloud"), /Unsupported guide/);
  assert.throws(() => buildEnvironmentPrompt("fr", "windows"), /Unsupported guide/);
  assert.throws(() => buildEnvironmentPrompt("en", "cloud"), /Unsupported guide/);
  const result = spawnSync(process.execPath, ["scripts/build-claude-code-guides.mjs", "--invalid"], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Usage:/);
});

function browserHarness(language, clipboardMode) {
  const html = renderGuide(language);
  const panels = ["manual", "environment", "connection"].flatMap(() => Object.keys(SHELLS).map((platform) => ({ dataset: { platform }, hidden: false })));
  const listeners = {};
  const elements = {
    platform: { value: "windows", disabled: false, addEventListener: (event, fn) => { listeners[event] = fn; } },
    copy: { hidden: true, disabled: false, addEventListener: (event, fn) => { listeners[event] = fn; } },
    "copy-status": { textContent: "", dataset: { success: "copied", fallback: "select manually" } },
    "copy-environment": { hidden: true, disabled: false, addEventListener: (event, fn) => { listeners["environment-" + event] = fn; } },
    "environment-status": { textContent: "", dataset: { success: "copied", fallback: "select manually" } }
  };
  const copied = [];
  for (const platform of Object.keys(SHELLS)) {
    for (const [prefix, build] of [["prompt", buildPrompt], ["environment", buildEnvironmentPrompt]]) {
    elements[`${prefix}-${platform}`] = {
      value: build(language, platform),
      focused: false,
      selected: false,
      focus() { this.focused = true; },
      select() { this.selected = true; }
    };
    }
  }
  const navigator = clipboardMode === "missing" ? {} : {
    clipboard: { async writeText(value) { if (clipboardMode !== "success") throw new Error("denied"); copied.push(value); } }
  };
  const document = {
    getElementById: (id) => elements[id],
    querySelectorAll: () => panels,
    execCommand: () => { if (clipboardMode === "throw") throw new Error("unsupported"); return clipboardMode === "fallback-success"; }
  };
  const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
  vm.runInNewContext(script, { document, navigator });
  return { elements, panels, listeners, copied };
}

for (const language of Object.keys(GUIDE_FILES)) {
  test(`${language}: selector changes visible and copied prompt together`, async () => {
    const { elements, panels, listeners, copied } = browserHarness(language, "success");
    assert.equal(elements.copy.hidden, false);
    assert.deepEqual(panels.map((panel) => panel.hidden), [false, true, false, true, false, true]);
    await listeners.click();
    assert.equal(copied[0], buildPrompt(language, "windows"));
    elements.platform.value = "mac";
    listeners.change();
    assert.deepEqual(panels.map((panel) => panel.hidden), [true, false, true, false, true, false]);
    assert.equal(elements["copy-status"].textContent, "");
    await listeners.click();
    assert.equal(copied[1], buildPrompt(language, "mac"));
    assert.equal(elements["copy-status"].textContent, "copied");
    assert.equal(elements.copy.disabled, false);
    assert.equal(elements.platform.disabled, false);
  });
  test(`${language}: environment copy remains independent of Hare connection`, async () => {
    const { elements, listeners, copied } = browserHarness(language, "success");
    assert.equal(elements["copy-environment"].hidden, false);
    await listeners["environment-click"]();
    assert.equal(copied[0], buildEnvironmentPrompt(language, "windows"));
    assert.equal(elements["copy-status"].textContent, "");
    elements.platform.value = "mac";
    listeners.change();
    assert.equal(elements["environment-status"].textContent, "");
    await listeners["environment-click"]();
    assert.equal(copied[1], buildEnvironmentPrompt(language, "mac"));
    assert.equal(elements["environment-status"].textContent, "copied");
    assert.equal(elements["copy-environment"].disabled, false);
    assert.equal(elements.copy.disabled, false);
  });
}

for (const mode of ["denied", "missing", "throw", "fallback-success"]) {
  test(`environment clipboard ${mode}: safe selection fallback`, async () => {
    const { elements, listeners } = browserHarness("en", mode);
    elements.platform.value = "mac";
    listeners.change();
    await listeners["environment-click"]();
    assert.equal(elements["environment-mac"].focused, true);
    assert.equal(elements["environment-mac"].selected, true);
    assert.equal(elements["environment-status"].textContent, mode === "fallback-success" ? "copied" : "select manually");
    assert.equal(elements["copy-environment"].disabled, false);
    assert.equal(elements.platform.disabled, false);
  });
  test(`clipboard ${mode}: safe local-file fallback`, async () => {
    const { elements, listeners } = browserHarness("en", mode);
    elements.platform.value = "mac";
    listeners.change();
    await listeners.click();
    assert.equal(elements["prompt-mac"].focused, true);
    assert.equal(elements["prompt-mac"].selected, true);
    assert.equal(elements["copy-status"].textContent, mode === "fallback-success" ? "copied" : "select manually");
    assert.equal(elements.copy.disabled, false);
    assert.equal(elements.platform.disabled, false);
  });
}
