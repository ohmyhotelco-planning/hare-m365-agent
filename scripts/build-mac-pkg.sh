#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
RELEASE_DIR="$ROOT/releases"
MAC_STAGE="$RELEASE_DIR/mac-llm"
ASSET_DIR="$ROOT/installer-assets/mac"
BUILD_DIR="$ROOT/build/mac-pkg"
ROOT_DIR="$BUILD_DIR/root"
SCRIPTS_DIR="$BUILD_DIR/scripts"
APP_SUPPORT="$ROOT_DIR/Library/Application Support/OMH/M365Agent"
APPS_DIR="$ROOT_DIR/Applications"
COMPONENT_PKG="$BUILD_DIR/OMH-M365-Agent-component.pkg"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "This script must be run on macOS because it uses pkgbuild/productbuild."
  exit 2
fi

for tool in pkgbuild productbuild python3 rsync shasum; do
  if ! command -v "$tool" >/dev/null 2>&1; then
    echo "Missing required macOS tool: $tool"
    exit 2
  fi
done

VERSION="$(
  cd "$ROOT"
  node -p "require('./package.json').version" 2>/dev/null ||
    sed -nE 's/.*"version"[[:space:]]*:[[:space:]]*"([^"]+)".*/\1/p' package.json | head -n 1
)"

PKG_PATH="$RELEASE_DIR/OMH-M365-Agent-mac-$VERSION.pkg"
START_APP="$APPS_DIR/OMH M365 Agent Start.app"
LOGIN_APP="$APPS_DIR/OMH M365 Agent Login.app"

require_file() {
  if [ ! -f "$1" ]; then
    echo "Required file is missing: $1"
    exit 2
  fi
}

require_file "$MAC_STAGE/omh-m365.cjs"
require_file "$MAC_STAGE/policy.json"
require_file "$MAC_STAGE/.env"
require_file "$MAC_STAGE/run-mac.sh"
require_file "$MAC_STAGE/run-cowork.sh"
require_file "$ASSET_DIR/start-guide-template.html"

rm -rf "$BUILD_DIR"
mkdir -p "$APP_SUPPORT" "$APPS_DIR" "$SCRIPTS_DIR" "$RELEASE_DIR"

rsync -a \
  --exclude ".cache" \
  --exclude "downloads" \
  --exclude "logs" \
  --exclude ".runtime" \
  --exclude "start-guide-template.html" \
  --exclude "로그인 시작.command" \
  "$MAC_STAGE/" "$APP_SUPPORT/"

python3 - "$ASSET_DIR/start-guide-template.html" "$APP_SUPPORT" <<'PY'
import html
import sys
from pathlib import Path

template_path = Path(sys.argv[1])
app_support = Path(sys.argv[2])
install_path = "/Library/Application Support/OMH/M365Agent"
login_app = "/Applications/OMH M365 Agent Login.app"
prompt_path = f"{install_path}/Cowork에 붙여넣을 프롬프트.txt"
prompt_text = f"""{install_path}

이 Mac에 설치된 OMH M365 Agent 폴더를 Cowork 또는 Claude Code에서 접근 가능한 shell 경로로 연결하고 COWORK_REQUIRED_README.md를 먼저 읽은 뒤 그대로 따라.
작업: <원하는 Microsoft 365 작업>
"""

html_text = template_path.read_text(encoding="utf-8")
html_text = html_text.replace("__ESCAPED_LOGIN_CMD__", html.escape(login_app))
html_text = html_text.replace("__ESCAPED_PROMPT__", html.escape(prompt_text))
html_text = html_text.replace("__ESCAPED_PROMPT_PATH__", html.escape(prompt_path))
html_text = html_text.replace("__ESCAPED_LINUX_PATH__", html.escape(install_path))
html_text = html_text.replace(
    "<p><strong>Finder에서 이 폴더를 열고</strong> 아래 파일을 더블클릭합니다.</p>",
    "<p><strong>Applications 폴더 또는 Launchpad에서</strong> 아래 앱을 실행합니다.</p>",
)
html_text = html_text.replace(
    "파일 이름은 <strong>로그인 시작.command</strong>입니다.",
    "앱 이름은 <strong>OMH M365 Agent Login</strong>입니다.",
)
html_text = html_text.replace(
    "더블클릭으로 열리지 않으면 파일을 Control-클릭 또는 오른쪽 클릭한 뒤 <strong>열기</strong>를 선택하세요. 확인 창이 뜨면 다시 <strong>열기</strong>를 누릅니다. 회사 보안 정책상 계속 차단되면 IT 배포 담당자에게 Jamf/Intune 배포본을 요청하세요.",
    "회사 보안 정책상 로그인 앱이 열리지 않으면 임의로 우회하지 말고 IT 배포 담당자에게 서명/공증 또는 Jamf/Intune 배포본을 요청하세요.",
)

(app_support / "Cowork에 붙여넣을 프롬프트.txt").write_text(prompt_text, encoding="utf-8")
(app_support / "COWORK_FIRST_PROMPT.txt").write_text(prompt_text, encoding="utf-8")
(app_support / "시작하기.html").write_text(html_text, encoding="utf-8")
(app_support / "start.html").write_text(html_text, encoding="utf-8")
PY

mkdir -p "$APP_SUPPORT/bin"
cat > "$APP_SUPPORT/bin/login-terminal.sh" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/Library/Application Support/OMH/M365Agent"

clear
echo "OMH M365 Agent - Microsoft 365 login"
echo
echo "This opens Microsoft device login in your browser."
echo "Enter the device code yourself in the browser."
echo "Do not paste the device code into Claude/Cowork/chat."
echo

open "https://microsoft.com/devicelogin" >/dev/null 2>&1 || true

"$APP_DIR/run-mac.sh" auth login

echo
echo "Checking login status..."
"$APP_DIR/run-mac.sh" auth status

echo
echo "If loggedIn is true, return to Claude/Cowork and continue."
echo
read -r -p "Press Enter to close this window. " _ || true
EOF

create_app() {
  local app_path="$1"
  local executable_name="$2"
  local bundle_id="$3"
  local script_body="$4"

  mkdir -p "$app_path/Contents/MacOS"
  cat > "$app_path/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleIdentifier</key>
  <string>$bundle_id</string>
  <key>CFBundleName</key>
  <string>$executable_name</string>
  <key>CFBundleDisplayName</key>
  <string>$executable_name</string>
  <key>CFBundleExecutable</key>
  <string>$executable_name</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>$VERSION</string>
  <key>CFBundleVersion</key>
  <string>$VERSION</string>
  <key>LSMinimumSystemVersion</key>
  <string>11.0</string>
</dict>
</plist>
EOF
  printf "%s\n" "$script_body" > "$app_path/Contents/MacOS/$executable_name"
  chmod 755 "$app_path/Contents/MacOS/$executable_name"
}

create_app \
  "$START_APP" \
  "OMH M365 Agent Start" \
  "com.ohmyhotel.m365agent.start" \
  '#!/usr/bin/env bash
set -euo pipefail
open "/Library/Application Support/OMH/M365Agent/start.html"'

create_app \
  "$LOGIN_APP" \
  "OMH M365 Agent Login" \
  "com.ohmyhotel.m365agent.login" \
  '#!/usr/bin/env bash
set -euo pipefail
LOGIN_SCRIPT="/Library/Application Support/OMH/M365Agent/bin/login-terminal.sh"
/usr/bin/osascript - "$LOGIN_SCRIPT" <<'"'"'OSA'"'"'
on run argv
  set loginScript to item 1 of argv
  tell application "Terminal"
    activate
    do script quoted form of loginScript
  end tell
end run
OSA'

chmod 755 "$APP_SUPPORT/run-mac.sh" "$APP_SUPPORT/run-cowork.sh" "$APP_SUPPORT/bin/login-terminal.sh"

cat > "$SCRIPTS_DIR/postinstall" <<'EOF'
#!/bin/sh
set -eu

APP_DIR="/Library/Application Support/OMH/M365Agent"

chmod -R go+rX "$APP_DIR" || true
chmod 755 "$APP_DIR/run-mac.sh" "$APP_DIR/run-cowork.sh" "$APP_DIR/bin/login-terminal.sh" || true

exit 0
EOF
chmod 755 "$SCRIPTS_DIR/postinstall"

if [ -n "${SIGN_APP_IDENTITY:-}" ]; then
  codesign --force --deep --options runtime --timestamp --sign "$SIGN_APP_IDENTITY" "$START_APP"
  codesign --force --deep --options runtime --timestamp --sign "$SIGN_APP_IDENTITY" "$LOGIN_APP"
fi

rm -f "$COMPONENT_PKG" "$PKG_PATH" "$PKG_PATH.sha256"
pkgbuild \
  --root "$ROOT_DIR" \
  --scripts "$SCRIPTS_DIR" \
  --identifier "com.ohmyhotel.m365agent" \
  --version "$VERSION" \
  --install-location "/" \
  "$COMPONENT_PKG"

if [ -n "${SIGN_INSTALLER_IDENTITY:-}" ]; then
  productbuild --sign "$SIGN_INSTALLER_IDENTITY" --package "$COMPONENT_PKG" "$PKG_PATH"
else
  productbuild --package "$COMPONENT_PKG" "$PKG_PATH"
fi

if [ -n "${NOTARY_KEYCHAIN_PROFILE:-}" ]; then
  xcrun notarytool submit "$PKG_PATH" --keychain-profile "$NOTARY_KEYCHAIN_PROFILE" --wait
  xcrun stapler staple "$PKG_PATH"
fi

shasum -a 256 "$PKG_PATH" > "$PKG_PATH.sha256"

echo "Created Mac pkg:"
echo "$PKG_PATH"
echo "SHA256:"
cat "$PKG_PATH.sha256"
