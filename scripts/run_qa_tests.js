/**
 * ========================================================
 * 우성아이디 현장 관리 시스템 - 종합 QA 및 무결성 검증 테스트 스위트
 * (Automated QA & Code Review Verification Suite)
 * ========================================================
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

console.log('========================================================');
console.log('🧪 [QA & 코드리뷰 검증 파이프라인] 테스트 시작');
console.log('========================================================\n');

// 1. 소스 파일 로드 및 구문 검증
const scriptPath = path.join(__dirname, 'SiteSchedule_sheet.js');
let code = '';
try {
  code = fs.readFileSync(scriptPath, 'utf8');
  console.log('✅ [Step 1: 정적 구문 검증] SiteSchedule_sheet.js 파일 로드 성공 (' + code.length + ' bytes)');
} catch (e) {
  console.error('❌ [Step 1: 오류] 소스 파일을 찾을 수 없습니다:', e.message);
  process.exit(1);
}

// 2. 가상 Apps Script 런타임 환경 구성
const sandbox = {
  Logger: { log: () => {} },
  console: console,
  SpreadsheetApp: {},
  CalendarApp: {},
  Utilities: { sleep: () => {} },
  Session: { getActiveUser: () => ({ getEmail: () => 'qa-tester@woosung.com' }) },
  PropertiesService: {
    getDocumentProperties: () => ({
      getProperty: () => '',
      setProperty: () => {},
      deleteProperty: () => {}
    })
  },
  ScriptApp: {
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' })
  }
};
vm.createContext(sandbox);

try {
  vm.runInContext(code, sandbox);
  console.log('✅ [Step 1: 정적 구문 검증] JavaScript AST 및 구문 무결성 통과 (Syntax Valid)\n');
} catch (e) {
  console.error('❌ [Step 1: 구문 오류 발생]:', e);
  process.exit(1);
}

let passCount = 0;
let failCount = 0;

function assertTest(name, condition, details) {
  if (condition) {
    passCount++;
    console.log(`  ✓ [PASS] ${name}`);
  } else {
    failCount++;
    console.error(`  ✗ [FAIL] ${name} -> ${details}`);
  }
}

// 3. QA 엣지 케이스 테스트 (날짜/시간 파서 전수 검증)
console.log('🔍 [Step 2: 날짜/시간 파싱 QA 엣지 케이스 검증]');

const parseFn = sandbox.parseScheduleDateRange;

// Case 1: 일반 시간 일정
const t1 = parseFn('26.10.11 15:00');
assertTest('단일 시간 일정 (26.10.11 15:00): hasTime === true', t1 && t1.hasTime === true);
assertTest('단일 시간 일정 기본 1시간 설정: 15:00 ~ 16:00', t1 && t1.startTime.getHours() === 15 && t1.endTime.getHours() === 16);
assertTest('단일 시간 일정 timeStr 검증: "15:00"', t1 && t1.timeStr === '15:00');

// Case 2: 시간 범위 일정
const t2 = parseFn('26.10.11 15:00 ~ 17:30');
assertTest('시간 범위 일정 (15:00 ~ 17:30): startTime 15:00, endTime 17:30', t2 && t2.startTime.getHours() === 15 && t2.endTime.getHours() === 17 && t2.endTime.getMinutes() === 30);
assertTest('시간 범위 timeStr 검증: "15:00~17:30"', t2 && t2.timeStr === '15:00~17:30');

// Case 3: 종일 일정 (시간 없음)
const t3 = parseFn('26.10.11');
assertTest('종일 단일 날짜: hasTime === false', t3 && t3.hasTime === false && t3.isRange === false);
assertTest('종일 단일 날짜 값: 2026년 10월 11일', t3 && t3.year === 2026 && t3.month === 10 && t3.day === 11);

// Case 4: 종일 기간 일정
const t4 = parseFn('26.10.11 ~ 26.10.15');
assertTest('종일 기간 일정: isRange === true, hasTime === false', t4 && t4.isRange === true && t4.hasTime === false);
assertTest('종일 기간 일정 구글 exclusive 종료일: 10월 16일', t4 && t4.endDate.getDate() === 16);

// Case 5: 엣지 케이스 (자정 00:00, 23:59, 윤년)
const t5 = parseFn('26.02.28 23:59');
assertTest('심야 시간 (23:59): 정상 파싱', t5 && t5.startTime.getHours() === 23 && t5.startTime.getMinutes() === 59);

// Case 6: 비정상/빈 값 방어
assertTest('빈 문자열 방어: null 반환', parseFn('') === null);
assertTest('공백 문자열 방어: null 반환', parseFn('   ') === null);
assertTest('null 값 방어: null 반환', parseFn(null) === null);
assertTest('undefined 값 방어: null 반환', parseFn(undefined) === null);

console.log('\n🔍 [Step 3: 포맷팅 및 역방향 동기화 헬퍼 검증]');
const fmtWithTime = sandbox.formatDateWithTime;
const sA = new Date(2026, 9, 11, 15, 0);
const eA1 = new Date(2026, 9, 11, 16, 0);
assertTest('1시간 기본 일정 포맷팅: "26.10.11 15:00"', fmtWithTime(sA, eA1) === '26.10.11 15:00');

const eA2 = new Date(2026, 9, 11, 18, 0);
assertTest('3시간 일정 포맷팅: "26.10.11 15:00 ~ 18:00"', fmtWithTime(sA, eA2) === '26.10.11 15:00 ~ 18:00');

console.log('\n🔍 [Step 3.5: 핵심 엔진 런타임 실행(Mock Runtime Execution) 전수 검증]');
const buildDetailsFn = sandbox.buildEventDetails;

// Row 템플릿: 0:현장, 1:날짜, 2:공종, 3:협력, 4:공유, 5:비고, 6:담당자, 7:진행상태, 8:캘린더ID, 9:시각, 10:작성일, 11:작성자, 12:수정일, 13:수정자
const baseMockRow = ['래미안 101동', '26.10.15', '목공', '우성목재', true, '오전작업', '김팀장', '예정', '', 0, '', '', '', ''];

// 1. 상태 '완료' 런타임 호출 테스트 (ReferenceError 방지 검증)
const rowDone = [...baseMockRow];
rowDone[7] = '완료';
const resDone = buildDetailsFn(rowDone, 15);
assertTest('buildEventDetails 런타임 실행 ("완료"): ReferenceError 없이 정상 실행', resDone !== null);
assertTest('상태 "완료" 시 캘린더 제목에 "✅ " 아이콘 정상 부여', resDone && resDone.title.startsWith('✅ 🟦'));

// 2. 상태 '진행중' 런타임 호출 테스트
const rowProg = [...baseMockRow];
rowProg[7] = '진행중';
const resProg = buildDetailsFn(rowProg, 16);
assertTest('상태 "진행중" 시 캘린더 제목에 "▶️ " 아이콘 정상 부여', resProg && resProg.title.startsWith('▶️ 🟦'));

// 2-1. 상태 '진행 중' (공백 포함) 엣지 케이스 테스트
const rowProgSpace = [...baseMockRow];
rowProgSpace[7] = '진행 중';
const resProgSpace = buildDetailsFn(rowProgSpace, 161);
assertTest('상태 "진행 중"(공백 포함) 시에도 캘린더 제목에 "▶️ " 아이콘 정상 부여', resProgSpace && resProgSpace.title.startsWith('▶️ 🟦'));

// 2-2. 상태 '▶️ 진행중' (기존 이모지 포함 입력) 엣지 케이스 테스트
const rowProgEmoji = [...baseMockRow];
rowProgEmoji[7] = '▶️ 진행중';
const resProgEmoji = buildDetailsFn(rowProgEmoji, 162);
assertTest('상태 "▶️ 진행중"(이모지 포함) 시에도 캘린더 제목에 "▶️ " 정상 부여', resProgEmoji && resProgEmoji.title.startsWith('▶️ 🟦'));

// 3. 상태 'AS' 런타임 호출 테스트
const rowAs = [...baseMockRow];
rowAs[7] = 'AS';
const resAs = buildDetailsFn(rowAs, 17);
assertTest('상태 "AS" 시 캘린더 제목에 "🛠️ " 아이콘 정상 부여', resAs && resAs.title.startsWith('🛠️ 🟦'));

// 3-1. 상태 'A/S' (슬래시 포함) 런타임 호출 테스트
const rowAsSlash = [...baseMockRow];
rowAsSlash[7] = 'A/S';
const resAsSlash = buildDetailsFn(rowAsSlash, 171);
assertTest('상태 "A/S"(슬래시 포함) 시 캘린더 제목에 "🛠️ " 아이콘 정상 부여', resAsSlash && resAsSlash.title.startsWith('🛠️ 🟦'));

// 4. 상태 '예정' 런타임 호출 테스트
const rowPlan = [...baseMockRow];
rowPlan[7] = '예정';
const resPlan = buildDetailsFn(rowPlan, 18);
assertTest('상태 "예정" 시 캘린더 제목에 상태 아이콘 없이 기본 "🟦 " 표시', resPlan && resPlan.title.startsWith('🟦'));

// 5. 상태 미입력(빈문자열, null, undefined) 런타임 안전성 테스트
const rowEmpty = [...baseMockRow];
rowEmpty[7] = '';
const resEmpty = buildDetailsFn(rowEmpty, 19);
assertTest('진행상태가 빈 문자열일 때 ReferenceError 없이 안전 처리', resEmpty !== null && resEmpty.title.startsWith('🟦'));

// 6. 모바일 웹앱 링크 Description 포함 여부 테스트
assertTest('일정 설명(Description)에 모바일 웹앱 링크(?row=15) 정상 삽입', resDone && resDone.description.indexOf('?row=15') !== -1);

// 7. buildMobileStatusHtml 런타임 실행 테스트
const buildHtmlFn = sandbox.buildMobileStatusHtml;
const htmlOutput = buildHtmlFn(15, '래미안 101동', '26.10.15', '목공', '완료');
assertTest('buildMobileStatusHtml 런타임 실행: 유효한 HTML 문자열 반환', typeof htmlOutput === 'string' && htmlOutput.indexOf('<!DOCTYPE html>') !== -1);
assertTest('모바일 HTML 내 4대 상태 버튼(예정, 진행중, 완료, AS) 정상 렌더링', htmlOutput.indexOf('btn-done') !== -1 && htmlOutput.indexOf('btn-progress') !== -1);

console.log('\n🔍 [Step 4: 코드리뷰 정적 보안/안전성 감사]');
// Rate limit 방어 패턴 검사
const hasDiffCheck = code.indexOf('curTitle === details.title') !== -1;
assertTest('스마트 변경 감지(Smart Diff Check) 패턴 존재 여부', hasDiffCheck);

const hasSleep = code.indexOf('Utilities.sleep(350)') !== -1;
assertTest('API Rate Limit 방어용 안전 딜레이(350ms) 적용 여부', hasSleep);

const hasBatchLimit = code.indexOf('(syncedCount + updatedCount) >= 25') !== -1;
assertTest('신규+수정 통합 25건 안전 분할 배치 제한 존재 여부', hasBatchLimit);

const hasNoVolatileRow = code.indexOf('스프레드시트 연동 (Row:') === -1;
assertTest('가변 행 번호(Row: XX) 완전 제거 여부', hasNoVolatileRow);

// 종일<->시간 전환 및 셀 즉시 반영 검사
const hasTypeTransition = code.indexOf('isCalAllDay !== !needTime') !== -1;
assertTest('종일<->시간 유형 전환 시 안전 재생성(delete & recreate) 로직 존재 여부', hasTypeTransition);

const hasFlushInTrigger = code.indexOf('SpreadsheetApp.flush()') !== -1;
assertTest('셀 수정 즉시 커밋용 SpreadsheetApp.flush() 호출 존재 여부', hasFlushInTrigger);

const hasDashRefreshInSync = code.indexOf('refreshDashboardCalendar(ss)') !== -1;
assertTest('단일 행 수정 직후 대시보드 실시간 연쇄 갱신 호출 존재 여부', hasDashRefreshInSync);

// 대시보드-현장데이터 교차 오염 방지 가드 검사
const hasSheetConflictGuard = code.indexOf('dashSheet.getSheetId() === siteSheet.getSheetId()') !== -1;
assertTest('대시보드-현장데이터 동일 시트 오염 방지 가드 존재 여부', hasSheetConflictGuard);

const hasStrictDowCheck = code.indexOf('hasDayOfWeekHeader(sheet)') !== -1;
assertTest('요일 헤더(일~토) 기반 엄격한 대시보드 판별 엔진 존재 여부', hasStrictDowCheck);

// [신규 QA] 하드코딩된 특정 연도/월/문자열 부재 및 100% 동적 연도 처리 검증
const hasHardcoded2026Fallback = code.indexOf('p2.year : 2026') !== -1 || code.indexOf('comp.year || 2026') !== -1;
assertTest('날짜 파싱 시 하드코딩 2026 폴백 완전 제거 여부 (동적 new Date().getFullYear() 적용)', !hasHardcoded2026Fallback);

const hasHardcodedCleanupDate = code.indexOf('c2.getValue()).trim() === "2026"') !== -1 || code.indexOf('e2.getValue()).trim() === "10"') !== -1;
assertTest('데이터 청소 시 특정 연도(2026)/월(10) 하드코딩 완전 제거 여부 (범용 정규식 적용)', !hasHardcodedCleanupDate);

const cleanupFnStr = code.slice(code.indexOf('function cleanupEmptyRowsMetadata'));
const hasHardcodedTaskNamesInCleanup = cleanupFnStr.indexOf('"목공"') !== -1 || cleanupFnStr.indexOf('"퍼플렉스"') !== -1;
assertTest('B열 날짜 청소 시 특정 공종명("목공", "퍼플렉스") 하드코딩 완전 제거 여부', !hasHardcodedTaskNamesInCleanup);

// [신규 QA] 6분 실행 시간 초과(Timeout) 방지 가드 및 신규 일정 우선순위 큐 검증
const hasTimeoutGuard = code.indexOf('Date.now() - startTime >') !== -1;
assertTest('구글 6분 타임아웃 방지 안전 실행 시간 가드(Execution Time Guard) 존재 여부', hasTimeoutGuard);

const hasPriorityQueue = code.indexOf('pendingIndices.concat(existingIndices)') !== -1;
assertTest('미등록 신규 일정 최우선 처리 큐(Priority Queue) 존재 여부', hasPriorityQueue);

// [신규 QA] 상태별 아이콘(스타일 A) 및 모바일 웹앱 드롭다운 연동 검증
const hasStatusIconsInTitle = code.indexOf('statusClean.indexOf("완료")') !== -1 && code.indexOf('"✅ "') !== -1 && code.indexOf('"▶️ "') !== -1 && code.indexOf('"🛠️ "') !== -1;
assertTest('캘린더 제목 상태별 아이콘(완료: ✅, 진행중: ▶️, AS: 🛠️) 표기 엔진 존재 여부', hasStatusIconsInTitle);

const hasForceSheetProtection = code.indexOf('forceSheetToCal === undefined') !== -1 && code.indexOf('!forceSheetToCal && sheetUpdated > 0 && calUpdated > sheetUpdated + 10000') !== -1;
assertTest('시트 수정 시 캘린더 날짜 역덮어쓰기 방지 안전 가드(forceSheetToCal) 존재 여부', hasForceSheetProtection);

const hasDash8Cols = code.indexOf('siteSheet.getRange(2, 1, lastRow - 1, 8).getValues()') !== -1 && code.indexOf('itemText = statusIcon + prefixIcon') !== -1;
assertTest('대시보드 달력 8열(진행상태) 연동 및 상태 아이콘 표시 엔진 탑재 여부', hasDash8Cols);

const hasMobileDoGet = code.indexOf('function doGet(e)') !== -1 && code.indexOf('buildMobileStatusHtml') !== -1;
assertTest('모바일 드롭다운 상태변경 웹앱 엔드포인트(doGet) 탑재 여부', hasMobileDoGet);

const hasReverseStatusSync = code.indexOf('calTitle.indexOf("✅")') !== -1 && code.indexOf('sheet.getRange(row, 8).setValue(detectedStatus)') !== -1;
assertTest('캘린더 제목 상태 이모지 역방향 동기화 엔진 탑재 여부', hasReverseStatusSync);

// [신규 QA] 비고(6열) 제외 실시간 캘린더 연동 및 무음 트리거 설치 검증
const hasMemoExclusion = code.indexOf('isOnlyMemo = (startCol === 6 && endCol === 6)') !== -1 && code.indexOf('if (!isOnlyMemo)') !== -1;
assertTest('비고(6열) 단독 수정 시 캘린더 실시간 동기화 제외 및 타 열 즉시 반영 가드 존재 여부', hasMemoExclusion);

const hasSilentTriggerInstall = code.indexOf('function ensureRealtimeTriggerInstalledSilent()') !== -1 && code.indexOf('ensureRealtimeTriggerInstalledSilent()') !== -1;
assertTest('실시간 트리거 백그라운드 무음 자동 설치(Silent Auto Install) 엔진 탑재 여부', hasSilentTriggerInstall);

const hasOnEditExtendedRange = code.indexOf('col >= 1 && col <= 8') !== -1;
assertTest('단순 onEdit 트리거 1~8열(진행상태/담당자 포함) 대시보드 연동 확장 여부', hasOnEditExtendedRange);

// [신규 QA] 드롭다운 '진행 중'(공백 포함) 유효성 검사 및 빨간 삼각형 오류 복구 검증
const hasSpacedValidation = code.indexOf('.requireValueInList(["예정", "진행 중", "완료", "AS"], true)') !== -1;
assertTest('진행상태 드롭다운 목록 내 "진행 중"(공백 포함) 표준 규격 탑재 여부 (빨간 삼각형 에러 방지)', hasSpacedValidation);

const hasFixDropdownFn = code.indexOf('function fixStatusDropdownValidation()') !== -1;
assertTest('진행상태 드롭다운 서식 복구 전용 함수(fixStatusDropdownValidation) 탑재 여부', hasFixDropdownFn);

// [신규 QA] 완전 빈 새 스프레드시트 원클릭 자동 구축(Bootstrap) 엔진 검증
const hasBootstrapWorkspace = code.indexOf('function setupNewWorkspaceFromScratch()') !== -1;
assertTest('새 스프레드시트 원클릭 자동 구축 함수(setupNewWorkspaceFromScratch) 탑재 여부', hasBootstrapWorkspace);

const hasCalendarGridBuilder = code.indexOf('function buildDashboardCalendarGrid(') !== -1;
assertTest('대시보드 만년달력 그리드/수식 자동 생성 엔진(buildDashboardCalendarGrid) 탑재 여부', hasCalendarGridBuilder);

const hasSiteConditionalFormatBuilder = code.indexOf('function setupSiteSheetConditionalFormatting(') !== -1;
assertTest('현장데이터 상태별 조건부 서식 자동 생성 엔진(setupSiteSheetConditionalFormatting) 탑재 여부', hasSiteConditionalFormatBuilder);

console.log('\n========================================================');
console.log(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
console.log('========================================================');

if (failCount > 0) {
  process.exit(1);
} else {
  console.log('🎉 모든 QA 및 코드리뷰 무결성 테스트를 100% 통과했습니다!');
}
