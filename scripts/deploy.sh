#!/usr/bin/env bash
set -e

DEPLOYMENT_ID="AKfycbxRcwYEilQ8wxRyMHxmPXUmLJjZ9qL2oclioJ7iQIYXdENmgjgWemu9D3T2m-FXHVvg"
DESC="${1:-현장관리_자동배포}"

echo "🔍 1. 문법 및 QA 무결성 검증..."
cp scripts/SiteSchedule_sheet.js 현장서식.js
cp scripts/SiteSchedule_sheet.js scripts/SiteSchedule_sheet.gs
node -c 현장서식.js
node scripts/run_qa_tests.js

echo "📤 2. clasp push (코드 반영)..."
clasp push

echo "🏷️ 3. clasp version (새 버전 생성)..."
VERSION_OUT=$(clasp version "$DESC")
echo "$VERSION_OUT"
VERSION_NUM=$(echo "$VERSION_OUT" | grep -oE '[0-9]+' | tail -n 1)

echo "🚀 4. clasp deploy (웹앱 배포 재배치: Version $VERSION_NUM)..."
clasp deploy -i "$DEPLOYMENT_ID" -V "$VERSION_NUM" -d "$DESC"

echo "✅ 배포 완료! 모바일 웹앱 URL에 Version $VERSION_NUM 이 즉시 적용되었습니다."
