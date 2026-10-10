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

3. **검증 및 피드백 루프 (QA & 코드리뷰 필수 파이프라인 연계)**:
   - **정적 문법 검증**: `node -c`, linter, 빌드 도구로 코드 구문 무결성을 즉시 검증합니다.
   - **QA 엣지 케이스 검증**: 경계값(`null`, 빈 문자열, 비정상 포맷, 날짜/시간 경계값), Rate Limit, 동시성 레이스 컨디션 가상 테스트를 실행합니다.
   - **다각도 셀프 코드리뷰**: 아키텍처, 성능(O(N^2) 방지, API 중복 호출 No-Op 방지), 보안 관점에서 결함을 분석합니다.
   - **자율 수정 (Self-Healing)**: 실패하거나 잠재적 결함이 발견된 경우, 질문 없이 자체 수정 패치를 적용한 뒤 재검증합니다. (스킬 `code-review-qa` 참조)

4. **결과 보고**:
   - 최종 작업 내역과 변경된 파일 목록, QA 테스트 시나리오 및 코드리뷰 검증 결과를 투명하게 요약하여 전달합니다.
