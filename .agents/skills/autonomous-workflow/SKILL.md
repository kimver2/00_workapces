---
name: autonomous-workflow
description: Guide and runbook for fully autonomous execution, development workflows, testing, and error recovery in this workspace.
---

# Autonomous Workflow Runbook

이 스킬은 작업 공간에서 자율적인 개발, 빌드, 테스트, 디버깅 사이클을 수행할 때 적용하는 가이드입니다.

## 핵심 실행 단계

1. **컨텍스트 파악**:
   - 관련 파일 및 디렉터리 구조를 먼저 조회합니다.
   - 의존성 및 프로젝트 설정을 확인합니다.

2. **구현 및 변경**:
   - 사용자의 목표에 맞게 필요한 모든 파일 생성/수정을 일괄 진행합니다.
   - 코드 스타일 및 기존 컨벤션을 유지합니다.

3. **검증 및 피드백 루프**:
   - 빌드/테스트/실행 명령어를 직접 실행하여 결과를 검증합니다.
   - 실패하거나 에러가 발생한 경우, 자율적으로 에러 로그를 분석하고 수정 패치를 적용한 뒤 재실행합니다.

4. **결과 보고**:
   - 최종 작업 내역과 변경된 파일 목록, 검증 결과를 간결하게 요약하여 전달합니다.
