#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

finish() {
  echo
  echo "창을 닫으려면 Enter를 누르세요."
  read -r _ || true
}
trap finish EXIT

clear
echo "OMH M365 Agent - Mac 로그인"
echo
echo "1. Microsoft device login 페이지가 열립니다."
echo "2. 잠시 후 이 터미널 창에 device code가 표시됩니다."
echo "3. 브라우저에 그 코드를 직접 입력하고 회사 Microsoft 계정으로 로그인하세요."
echo
echo "중요: device code를 Claude/Cowork/채팅에 붙여넣지 마세요."
echo

open "https://microsoft.com/devicelogin" >/dev/null 2>&1 || true

./run-mac.sh auth login

echo
echo "로그인 상태를 확인합니다."
./run-mac.sh auth status

echo
echo "loggedIn: true가 보이면 로그인 완료입니다."
