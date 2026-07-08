#!/usr/bin/env bash
set -euo pipefail

HARE_M365_PACKAGE_URL="__PACKAGE_URL__"
export HARE_M365_PACKAGE_URL

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
START_GUIDE="${SCRIPT_DIR}/START_HERE.html"

echo "Hare M365 Agent 시작 도구"
echo
echo "이 파일은 GitHub Release에 올라간 Hare M365 Agent npm 패키지를 실행합니다."
echo "패키지 URL:"
echo "${HARE_M365_PACKAGE_URL}"
echo

if [ -f "${START_GUIDE}" ]; then
  echo "시작 가이드를 브라우저로 엽니다."
  if command -v open >/dev/null 2>&1; then
    open "${START_GUIDE}" || true
  elif command -v xdg-open >/dev/null 2>&1; then
    xdg-open "${START_GUIDE}" || true
  fi
  echo "가이드의 1번 도메인 허용 단계를 먼저 완료하세요."
  printf "완료 후 Enter를 누르세요: "
  read -r _
fi

if ! command -v npm >/dev/null 2>&1; then
  echo "npm을 찾을 수 없습니다."
  echo "Node.js LTS를 설치한 뒤 다시 실행하세요: https://nodejs.org/"
  exit 1
fi

echo "[1/3] LLM 사용 가이드를 표시합니다."
npm exec --yes --package "${HARE_M365_PACKAGE_URL}" -- hare-m365 llm-guide
echo

echo "[2/3] 설정 상태를 확인합니다."
npm exec --yes --package "${HARE_M365_PACKAGE_URL}" -- hare-m365 doctor
echo

echo "[3/3] 로그인 상태를 확인합니다."
npm exec --yes --package "${HARE_M365_PACKAGE_URL}" -- hare-m365 auth status
echo

printf "처음 사용이거나 loggedIn이 false이면 y를 눌러 Microsoft 로그인을 시작하세요 [y/N]: "
read -r answer
case "${answer}" in
  y|Y)
    echo
    echo "Microsoft device-code 로그인을 시작합니다."
    echo "표시되는 코드는 본인이 브라우저에 직접 입력하세요. 채팅에 붙여넣지 마세요."
    npm exec --yes --package "${HARE_M365_PACKAGE_URL}" -- hare-m365 auth login
    echo
    echo "로그인 후 상태를 다시 확인합니다."
    npm exec --yes --package "${HARE_M365_PACKAGE_URL}" -- hare-m365 auth status
    ;;
esac

echo
echo "LLM에게 줄 짧은 프롬프트:"
echo "아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해. 패키지 URL: ${HARE_M365_PACKAGE_URL} 먼저 도메인 허용 여부를 확인하고, 허용 완료 후 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인한 뒤 내 Microsoft 365 요청을 처리해. .env, .cache, token, device code는 읽거나 출력하지 마."
