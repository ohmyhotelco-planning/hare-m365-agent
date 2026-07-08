#!/usr/bin/env bash
set -euo pipefail

HARE_M365_PACKAGE_URL="__PACKAGE_URL__"
export HARE_M365_PACKAGE_URL

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
START_GUIDE="${SCRIPT_DIR}/START_HERE.html"
HARE_M365_DATA_DIR="${SCRIPT_DIR}/runtime"
export HARE_M365_DATA_DIR

echo "Hare M365 Agent 시작 도구"
echo
echo "이 파일은 GitHub Release에 올라간 Hare M365 Agent npm 패키지를 실행합니다."
echo "패키지 URL:"
echo "${HARE_M365_PACKAGE_URL}"
echo "runtime 위치:"
echo "${HARE_M365_DATA_DIR}"
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
echo "아래 GitHub Release 패키지를 npm exec로 실행해서 Hare M365 Agent를 사용해. 패키지 URL: ${HARE_M365_PACKAGE_URL} 먼저 도메인 허용 여부를 확인하고, 허용 완료 후 llm-guide를 읽고 doctor/auth status로 설정과 로그인을 확인해. 폴더 접근 권한 또는 지속 인증 캐시 위치가 필요하다고 확인될 때만 가능한 도구로 문서/Documents 안의 Hare M365 Agent 전용 폴더를 먼저 준비한 뒤 Claude/Cowork의 폴더 연결 기능으로 그 폴더 연결/허용을 요청해. 사용자가 폴더를 직접 만들거나 경로를 판단하게 하지 마. 프로젝트 또는 배포 폴더가 연결되어 있으면 같은 연결 폴더 안의 ./runtime을 인증 캐시 위치로 사용하고 모든 Hare 명령에 HARE_M365_DATA_DIR=./runtime을 붙여 실행해. CLI는 runtime/.cache/msal-cache.json을 사용할 수 있지만 너는 그 파일을 열거나 출력하거나 요약하지 마. loggedIn이 false이면 로그인 hard gate로 멈추고 Outlook/Teams/Files 조회를 실행하지 마. Cowork/샌드박스에서는 auth login을 자동 실행하지 말고, 내가 직접 볼 수 있는 로컬 터미널에서 실행할 로그인 명령을 안내해. 연결 폴더를 쓰는 경우 같은 폴더에서 같은 HARE_M365_DATA_DIR=./runtime 설정으로 로그인해야 해. 내가 \"로그인 완료\"라고 말하면 doctor/auth status를 다시 확인하고 원래 요청을 이어서 처리해. .env, .cache, runtime/.cache, token, device code는 읽거나 출력하지 마."
