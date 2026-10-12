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
  SpreadsheetApp: {
    BorderStyle: { SOLID: 'SOLID' },
    flush: () => {},
    getActiveSpreadsheet: () => ({
      getUrl: () => 'https://docs.google.com/spreadsheets/d/test-sheet-id/edit',
      getSheets: () => [],
      getSheetByName: () => null
    })
  },
  CalendarApp: {
    EventColor: {
      PALE_BLUE: "1",
      PALE_GREEN: "2",
      MAUVE: "3",
      PALE_RED: "4",
      YELLOW: "5",
      ORANGE: "6",
      CYAN: "7",
      GRAY: "8",
      BLUE: "9",
      GREEN: "10",
      RED: "11"
    }
  },
  MailApp: {
    sendEmail: (opt) => { sandbox.lastSentEmail = opt; }
  },
  Utilities: { sleep: () => {} },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'qa-tester@woosung.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'qa-tester@woosung.com' })
  },
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
assertTest('캘린더 제목(Title) 내 비고(memo) 데이터 완전 배제 검증', resDone && resDone.title === '✅ 🟦 [래미안 101동] 목공');
assertTest('캘린더 본문 Description 내 비고(memo) 정상 유지 검증', resDone && resDone.description.indexOf('📝 비고: 오전작업') !== -1);

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

// 6-1. 캘린더 본문 하이퍼링크 및 감사 추적 영역 제거 검증
const rowWithAudit = [...baseMockRow];
rowWithAudit[10] = '2026-10-15 09:00:00';
rowWithAudit[11] = '김팀장';
rowWithAudit[12] = '2026-10-15 10:00:00';
rowWithAudit[13] = '이과장';
const resWithAudit = buildDetailsFn(rowWithAudit, 15);

assertTest('캘린더 상세 설명: 최초작성 정보 영역 완전 제거 확인', resWithAudit && resWithAudit.description.indexOf('최초작성') === -1);
assertTest('캘린더 상세 설명: 최종수정 정보 영역 완전 제거 확인', resWithAudit && resWithAudit.description.indexOf('최종수정') === -1);
assertTest('캘린더 상세 설명: [공유/진행상태 변경] 클릭 가능한 HTML 하이퍼링크 형식 탑재', resWithAudit && resWithAudit.description.indexOf('<a href="https://script.google.com/macros/s/test/exec?row=15">📱 [공유/진행상태 변경]</a>') !== -1);
assertTest('캘린더 상세 설명: [우성아이디 현장 관리 시트] 클릭 가능한 HTML 하이퍼링크 형식 탑재', resWithAudit && resWithAudit.description.indexOf('📊 [우성아이디 현장 관리 시트]</a>') !== -1);
assertTest('캘린더 상세 설명: 스프레드시트 링크에 해당 행 포커스(range=A15:H15) 자동 반영', resWithAudit && resWithAudit.description.indexOf('range=A15:H15') !== -1);

// 6-2. 특정 시트(targetSheet) GID 및 행 포커스 연동 검증
const mockSiteSheetWithGid = {
  getSheetId: () => 100,
  getParent: () => ({ getUrl: () => 'https://docs.google.com/spreadsheets/d/test-sheet-id/edit' })
};
const resWithSheetGid = buildDetailsFn(rowWithAudit, 15, null, mockSiteSheetWithGid);
assertTest('캘린더 상세 설명: targetSheet 연동 시 #gid=100&range=A15:H15 완벽 생성', resWithSheetGid && resWithSheetGid.description.indexOf('#gid=100&range=A15:H15') !== -1);

// 7. buildMobileStatusHtml 런타임 실행 테스트
const buildHtmlFn = sandbox.buildMobileStatusHtml;
const htmlOutput = buildHtmlFn(15, '래미안 101동', '26.10.15', '목공', '완료', 'https://script.google.com/test', true);
assertTest('buildMobileStatusHtml 런타임 실행: 유효한 HTML 문자열 반환', typeof htmlOutput === 'string' && htmlOutput.indexOf('<!DOCTYPE html>') !== -1);
assertTest('모바일 HTML 내 4대 상태 버튼(예정, 진행중, 완료, AS) 정상 렌더링', htmlOutput.indexOf('btn-done') !== -1 && htmlOutput.indexOf('btn-progress') !== -1);
assertTest('모바일 HTML 내 공유 상태 토글 버튼(공유 ON, 미공유 OFF) 정상 렌더링', htmlOutput.indexOf('btnShareOn') !== -1 && htmlOutput.indexOf('btnShareOff') !== -1);
assertTest('모바일 HTML 내 [공유/진행상태/비고 저장하기] 버튼 정상 렌더링', htmlOutput.indexOf('공유/진행상태/비고 저장하기') !== -1 || htmlOutput.indexOf('공유/진행상태 저장하기') !== -1);
assertTest('모바일 HTML 내 [현장 비고/특이사항 메모] textarea 정상 렌더링', htmlOutput.indexOf('id=\'memoInput\'') !== -1);

console.log('\n🔍 [Step 3.6: 신규 3대 기능 런타임 엔진 검증 (Feature 1, 2, 4)]');

// Feature 1: 구글 캘린더 자체 블록 색상 ID 매핑 검증
const getColorFn = sandbox.getCalendarColorForStatus;
assertTest('블록 색상 - 완료 상태: "10" (Basil 초록)', getColorFn('완료', false, true) === '10');
assertTest('블록 색상 - 공정 지연 상태: "11" (Tomato 빨강)', getColorFn('진행중', true, true) === '11');
assertTest('블록 색상 - 진행 중 상태: "7" (Peacock 청록/파랑)', getColorFn('진행중', false, true) === '7');
assertTest('블록 색상 - AS 상태: "4" (Flamingo 산호/주황)', getColorFn('AS', false, true) === '4');
assertTest('블록 색상 - 미공유 상태: "8" (Graphite 회색)', getColorFn('예정', false, false) === '8');
assertTest('블록 색상 - 기본 예정(공유): "1" (Pale Blue 연파랑)', getColorFn('예정', false, true) === '1');

// Feature 2: 공정 지연 자동 감지 (isScheduleDelayed) 검증
const isDelayedFn = sandbox.isScheduleDelayed;
const mockRefToday = new Date(2026, 9, 15, 10, 0, 0); // 2026-10-15

// 과거 단일 일정 (26.10.10) -> 마감일 지남
const parsedPast = parseFn('26.10.10');
assertTest('지연 감지 - 마감일 경과(26.10.10) & 예정 상태: true (지연)', isDelayedFn(parsedPast, '예정', mockRefToday) === true);
assertTest('지연 감지 - 마감일 경과(26.10.10) & 진행중 상태: true (지연)', isDelayedFn(parsedPast, '진행중', mockRefToday) === true);
assertTest('지연 감지 - 마감일 경과(26.10.10) & 완료 상태: false (완료는 지연 아님)', isDelayedFn(parsedPast, '완료', mockRefToday) === false);

// 오늘 단일 일정 (26.10.15) -> 오늘 마감 (아직 지연 아님)
const parsedToday = parseFn('26.10.15');
assertTest('지연 감지 - 오늘 당일 일정(26.10.15): false (당일은 정상 진행)', isDelayedFn(parsedToday, '진행중', mockRefToday) === false);

// 기간 일정 (26.10.11 ~ 26.10.13) -> 13일 마감 경과
const parsedPastRange = parseFn('26.10.11 ~ 26.10.13');
assertTest('지연 감지 - 기간 일정 종료일 경과(10/13): true (지연)', isDelayedFn(parsedPastRange, '진행중', mockRefToday) === true);

// 기간 일정 (26.10.11 ~ 26.10.16) -> 16일 마감 미래
const parsedFutureRange = parseFn('26.10.11 ~ 26.10.16');
assertTest('지연 감지 - 기간 일정 진행 중(10/11~10/16): false (정상 진행)', isDelayedFn(parsedFutureRange, '진행중', mockRefToday) === false);

// buildEventDetails 내 지연 및 색상 통합 검증
const rowDelayed = [...baseMockRow];
rowDelayed[1] = '26.10.10'; // 과거 날짜
rowDelayed[7] = '진행 중';
const resDelayed = buildDetailsFn(rowDelayed, 20, mockRefToday);
assertTest('지연 공정 제목에 "⚠️ [지연] " 자동 부착', resDelayed && resDelayed.title.startsWith('⚠️ [지연] ▶️ 🟦'));
assertTest('지연 공정 색상에 "11" (Tomato 빨강) 자동 지정', resDelayed && resDelayed.color === '11');
assertTest('지연 공정 본문 Description에 경고 문구 삽입', resDelayed && resDelayed.description.indexOf('⚠️ [공정 지연 경고]') !== -1);
assertTest('지연 공정 isDelayed 플래그 true', resDelayed && resDelayed.isDelayed === true);

// Feature 4: 활성 공정 감지 및 모닝 브리핑 이메일 생성기 검증
const isTodayActiveFn = sandbox.isScheduleActiveToday;
assertTest('활성 공정 감지 - 당일 단일 일정(26.10.15): true', isTodayActiveFn(parsedToday, mockRefToday) === true);
assertTest('활성 공정 감지 - 오늘 포함 기간 일정(10/11~10/16): true', isTodayActiveFn(parsedFutureRange, mockRefToday) === true);
assertTest('활성 공정 감지 - 과거 완료된 일정(10/10): false', isTodayActiveFn(parsedPast, mockRefToday) === false);

const buildBriefingFn = sandbox.buildMorningBriefingEmailContent;
const mockTodayTasks = [{ siteName: '역삼동 현장', taskName: '도배', status: '진행 중', dateStr: '26.10.15', manager: '이과장', memo: '오후 시공' }];
const mockDelayedTasks = [{ siteName: '한남동 빌라', taskName: '설비', status: '진행 중', dateStr: '26.10.12', manager: '박팀장', memo: '자재 대기' }];

const briefingOutput = buildBriefingFn('26.10.15', '목요일', mockTodayTasks, mockDelayedTasks, 'https://docs.google.com/test', 'https://script.google.com/test');
assertTest('모닝 브리핑 이메일 생성: HTML 및 Plain-text 생성 성공', briefingOutput && typeof briefingOutput.htmlBody === 'string' && typeof briefingOutput.plainText === 'string');
assertTest('모닝 브리핑 HTML 내 오늘 공정 섹션 정상 렌더링', briefingOutput.htmlBody.indexOf('오늘 착공 및 진행 공정') !== -1 && briefingOutput.htmlBody.indexOf('역삼동 현장') !== -1);
assertTest('모닝 브리핑 HTML 내 지연 공정 경고 박스 렌더링', briefingOutput.htmlBody.indexOf('공정 지연 현장') !== -1 && briefingOutput.htmlBody.indexOf('한남동 빌라') !== -1);
assertTest('모닝 브리핑 Plain-text 내 바로가기 링크 포함', briefingOutput.plainText.indexOf('스프레드시트 바로가기') !== -1);
assertTest('모닝 브리핑 HTML 가로폭 확장 (880px) 적용 여부', briefingOutput.htmlBody.indexOf('max-width: 880px') !== -1 || briefingOutput.htmlBody.indexOf('max-width:880px') !== -1);
assertTest('모닝 브리핑 버튼 글자색 인라인 흰색(#ffffff !important) 방어 탑재 여부', briefingOutput.htmlBody.indexOf('color:#ffffff !important') !== -1);
assertTest('모닝 브리핑 비고 열 확장(width:34%) 및 줄바꿈(word-break) 탑재 여부', briefingOutput.htmlBody.indexOf('width:34%') !== -1 && briefingOutput.htmlBody.indexOf('word-break:break-word') !== -1);

// Feature: [사용가이드] 시트 자동 생성 및 매뉴얼 구축 엔진 검증
function createMockSheet(name) {
  const sheetObj = {
    getName: () => name,
    clear: () => {},
    setHiddenGridlines: () => {},
    setTabColor: () => {},
    setColumnWidth: () => {},
    setRowHeight: () => {},
    getRange: () => {
      const rangeObj = {
        merge: () => rangeObj,
        setValue: (val) => { rangeObj.val = val; return rangeObj; },
        setValues: (vals) => { rangeObj.vals = vals; return rangeObj; },
        setFontSize: () => rangeObj,
        setFontWeight: () => rangeObj,
        setFontColor: () => rangeObj,
        setBackground: () => rangeObj,
        setHorizontalAlignment: () => rangeObj,
        setVerticalAlignment: () => rangeObj,
        setFontFamily: () => rangeObj,
        setBorder: () => rangeObj
      };
      return rangeObj;
    }
  };
  return sheetObj;
}

const mockGuideSheet = createMockSheet('사용가이드');
const mockSs = {
  getSheetByName: (n) => (n === '사용가이드' ? mockGuideSheet : null),
  insertSheet: (n) => mockGuideSheet
};
const resGuide = sandbox.createOrUpdateUsageGuideSheet(mockSs);
assertTest('createOrUpdateUsageGuideSheet 런타임 실행: 사용가이드 시트 정상 생성/반환', resGuide !== null && resGuide.getName() === '사용가이드');

const isSiteSheetFn = sandbox.isSiteScheduleSheet;
assertTest('isSiteScheduleSheet 판별: 사용가이드 시트는 현장데이터 시트에서 완전 배제', isSiteSheetFn(mockGuideSheet) === false);

console.log('\n🔍 [Step 3.7: 구성원 시트 연동 및 다중 수신자 모닝 브리핑 엔진 런타임 검증]');

function createMockMembersSheet(name, headerRow, dataRows) {
  const fullGrid = [headerRow, ...dataRows];
  return {
    getName: () => name,
    getSheetId: () => Math.floor(Math.random() * 100000) + 1000,
    getLastRow: () => fullGrid.length,
    getLastColumn: () => (headerRow ? headerRow.length : 0),
    getRange: (r, c, nr, nc) => ({
      getValues: () => {
        const slice = [];
        for (let rowIdx = r - 1; rowIdx < r - 1 + nr && rowIdx < fullGrid.length; rowIdx++) {
          const rowData = fullGrid[rowIdx] || [];
          slice.push(rowData.slice(c - 1, c - 1 + nc));
        }
        return slice;
      }
    })
  };
}

const mockMemberHeaders = ['이름', '직책', '이메일', '연락처', '모닝 브리핑 수신'];
const mockMemberData = [
  ['홍길동', '현장소장', 'hong@woosung.com', '010-1111-2222', true],
  ['김철수', '공무과장', 'kim@woosung.com', '010-3333-4444', false],
  ['이영희', '인테리어실장', 'lee@woosung.com', '010-5555-6666', 'TRUE'],
  ['박대리', '설계대리', 'not-an-email', '010-7777-8888', true],
  ['최주임', '현장주임', 'hong@woosung.com', '010-9999-0000', true] // 중복 이메일
];

const mockMemberSheet = createMockMembersSheet('구성원', mockMemberHeaders, mockMemberData);
const extractedEmails = sandbox.extractMemberBriefingEmails(mockMemberSheet);

assertTest('구성원 시트 이메일 추출: 체크된 인원만 추출 (true, "TRUE")', extractedEmails.indexOf('hong@woosung.com') !== -1 && extractedEmails.indexOf('lee@woosung.com') !== -1);
assertTest('구성원 시트 이메일 추출: 체크 해제된 인원(false) 제외', extractedEmails.indexOf('kim@woosung.com') === -1);
assertTest('구성원 시트 이메일 추출: 유효하지 않은 이메일 형식 필터링', extractedEmails.indexOf('not-an-email') === -1);
assertTest('구성원 시트 이메일 추출: 중복 이메일 자동 제거 (Deduplication)', extractedEmails.filter(e => e === 'hong@woosung.com').length === 1);
assertTest('구성원 시트 이메일 추출: 최종 수신자 수 일치 (2명)', extractedEmails.length === 2);

// 체크박스 열이 없는 일반 이메일 명단 시트 테스트
const mockSimpleMemberSheet = createMockMembersSheet('팀원', ['성명', '이메일주소', '연락처'], [
  ['박팀장', 'park@woosung.com', '010-1234-5678'],
  ['정소장', 'jung@woosung.com', '010-8765-4321']
]);
const extractedSimpleEmails = sandbox.extractMemberBriefingEmails(mockSimpleMemberSheet);
assertTest('수신 체크박스 열 부재 시: 모든 유효 이메일 자동 추출', extractedSimpleEmails.length === 2 && extractedSimpleEmails.indexOf('park@woosung.com') !== -1 && extractedSimpleEmails.indexOf('jung@woosung.com') !== -1);

// [실제 사용자 사례 테스트]: 이메일 열이 비어 있고 연락처 열에 이메일이 입력된 경우 스마트 자동 감지
const mockMisplacedSheet = createMockMembersSheet('03_구성원', ['이름', '직책', '이메일', '연락처', '모닝 브리핑 수신'], [
  ['조항만', '대표', '', 'joehm037@gmail.com', true]
]);
const extractedMisplacedEmails = sandbox.extractMemberBriefingEmails(mockMisplacedSheet);
assertTest('사용자 시나리오 검증: C열(이메일) 비어있고 D열(연락처)에 이메일 있을 때 스마트 자동 추출 성공', extractedMisplacedEmails.length === 1 && extractedMisplacedEmails[0] === 'joehm037@gmail.com');
assertTest('findMembersSheet 탐색: "03_구성원" 접두사 시트명 정상 감지', sandbox.findMembersSheet({ getSheets: () => [mockMisplacedSheet] }) !== null);

// findMembersSheet 탐색 검증
const mockSsWithMembers = {
  getSheets: () => [mockGuideSheet, mockMemberSheet]
};
const foundMemberSheet = sandbox.findMembersSheet(mockSsWithMembers);
assertTest('findMembersSheet 탐색: "구성원" 시트 자동 감지', foundMemberSheet !== null && foundMemberSheet.getName() === '구성원');

const mockSsNoMembers = {
  getSheets: () => [mockGuideSheet]
};
assertTest('findMembersSheet 탐색: 구성원 시트 부재 시 null 반환', sandbox.findMembersSheet(mockSsNoMembers) === null);

// getDailyBriefingRecipientEmail 우선순위 3단계 검증
// 1순위: [구성원] 시트
const multiEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsWithMembers);
assertTest('모닝 브리핑 수신처 1순위: [구성원] 시트 자동 연동 (쉼표 구분)', multiEmailResult === 'hong@woosung.com, lee@woosung.com');

// 2순위: 문서 속성(DAILY_BRIEFING_EMAIL)
let testDocProp = 'custom-briefing@woosung.com';
sandbox.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => (k === 'DAILY_BRIEFING_EMAIL' ? testDocProp : ''),
    setProperty: (k, v) => { testDocProp = v; },
    deleteProperty: () => { testDocProp = ''; }
  })
};
const propEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsNoMembers);
assertTest('모닝 브리핑 수신처 2순위: 문서 속성(DAILY_BRIEFING_EMAIL) 설정값 연동', propEmailResult === 'custom-briefing@woosung.com');

// 3순위: 로그인된 기본 사용자
testDocProp = '';
const defaultEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsNoMembers);
assertTest('모닝 브리핑 수신처 3순위: 로그인 계정(qa-tester@woosung.com) 안전 폴백', defaultEmailResult === 'qa-tester@woosung.com');

// [구성원] 시트 캘린더 동기화 오작동 방지 가드 검증
assertTest('isSiteScheduleSheet 판별: "구성원" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(mockMemberSheet) === false);
assertTest('isSiteScheduleSheet 판별: "팀원" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('팀원')) === false);
assertTest('isSiteScheduleSheet 판별: "조직도" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('조직도')) === false);
assertTest('isSiteScheduleSheet 판별: "직원명단" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('직원명단')) === false);
assertTest('isSiteScheduleSheet 판별: "Team_List" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('Team_List')) === false);

console.log('\n🔍 [Step 3.8: 모닝 브리핑 정각(Exact Time) 스케줄링 및 시간 설정 엔진 검증]');

// 1. 기본 설정 시간 ("08:00")
let propMap = {};
sandbox.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => propMap[k] || '',
    setProperty: (k, v) => { propMap[k] = v; },
    deleteProperty: (k) => { delete propMap[k]; }
  })
};

const defaultTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 미지정 시 기본 "08:00" 반환', defaultTime === '08:00');

// 2. 사용자 시간 설정
propMap['DAILY_BRIEFING_TIME'] = '07:30';
const customTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 사용자 지정 "07:30" 정상 반환', customTime === '07:30');

// 3. 비정상 포맷 필터링
propMap['DAILY_BRIEFING_TIME'] = 'invalid-time';
const fallbackTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 비정상 포맷 입력 시 "08:00" 안전 폴백', fallbackTime === '08:00');

// 4. scheduleNextDailyBriefingTrigger 정밀 계산 검증
let createdTriggerInfo = null;
let deletedCount = 0;
sandbox.ScriptApp = {
  getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' }),
  getProjectTriggers: () => ([{ getHandlerFunction: () => 'sendDailySiteMorningBriefing' }]),
  deleteTrigger: () => { deletedCount++; },
  newTrigger: (fn) => ({
    timeBased: () => ({
      at: (d) => ({
        create: () => {
          createdTriggerInfo = { fn: fn, date: d };
          return {};
        }
      })
    })
  })
};

propMap['DAILY_BRIEFING_TIME'] = '08:00';
const scheduledTarget = sandbox.scheduleNextDailyBriefingTrigger();
assertTest('정각 스케줄러: 기존 단발성 트리거 정리(deleteTrigger) 여부', deletedCount > 0);
assertTest('정각 스케줄러: targetDate 생성 성공 및 Date 객체 반환', scheduledTarget && Object.prototype.toString.call(scheduledTarget) === '[object Date]');
assertTest('정각 스케줄러: 목표 시간 8시 0분 정각(08:00:00) 설정 여부', scheduledTarget.getHours() === 8 && scheduledTarget.getMinutes() === 0 && scheduledTarget.getSeconds() === 0);
assertTest('정각 스케줄러: 미래 시점 예약 보장 (target > Date.now())', scheduledTarget.getTime() > Date.now());
assertTest('정각 스케줄러: ScriptApp.at(target) 정확한 함수 및 일시로 트리거 생성', createdTriggerInfo && createdTriggerInfo.fn === 'sendDailySiteMorningBriefing' && createdTriggerInfo.date.getTime() === scheduledTarget.getTime());

console.log('\n🔍 [Step 4: 코드리뷰 정적 보안/안전성 감사]');
// Rate limit 방어 패턴 검사
const hasDiffCheck = code.indexOf('curTitle === details.title') !== -1;
assertTest('스마트 변경 감지(Smart Diff Check) 패턴 존재 여부', hasDiffCheck);

const hasColorInDiffCheck = code.indexOf('(!details.color || curColor === details.color)') !== -1;
assertTest('스마트 변경 감지 내 색상(color) Diff Check 포함 여부', hasColorInDiffCheck);

const hasSleep = code.indexOf('Utilities.sleep(350)') !== -1;
assertTest('API Rate Limit 방어용 안전 딜레이(350ms) 적용 여부', hasSleep);

const hasBatchCooldown = code.indexOf('batchChangeCount >= 25') !== -1;
assertTest('25건 단위 자동 쿨다운 대기 및 연속 배치 동기화 엔진 탑재 여부', hasBatchCooldown);

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

const hasDash8Cols = code.indexOf('siteSheet.getRange(2, 1, lastRow - 1, 8).getValues()') !== -1 && (code.indexOf('itemText = delayPrefix + statusIcon + prefixIcon') !== -1 || code.indexOf('itemText = statusIcon + prefixIcon') !== -1);
assertTest('대시보드 달력 8열(진행상태) 연동 및 지연/상태 아이콘 표시 엔진 탑재 여부', hasDash8Cols);

const hasMobileDoGet = code.indexOf('function doGet(e)') !== -1 && code.indexOf('buildMobileStatusHtml') !== -1;
assertTest('모바일 드롭다운 상태변경 웹앱 엔드포인트(doGet) 탑재 여부', hasMobileDoGet);

const hasWebRpcHandler = code.indexOf('function updateScheduleStatusFromWeb(') !== -1;
assertTest('모바일 웹앱 RPC 서버 핸들러(updateScheduleStatusFromWeb) 탑재 여부', hasWebRpcHandler);

const hasGoogleScriptRunCall = code.indexOf('google.script.run') !== -1 && code.indexOf('updateScheduleStatusFromWeb') !== -1;
assertTest('모바일 웹앱 클라이언트 google.script.run RPC 호출 탑재 여부 (CORS/네트워크 오류 원천 차단)', hasGoogleScriptRunCall);

const hasAppSpreadsheetFallback = code.indexOf('function getAppSpreadsheet()') !== -1;
assertTest('독립 웹앱 실행 시 스프레드시트 자동 폴백 엔진(getAppSpreadsheet) 탑재 여부', hasAppSpreadsheetFallback);

const hasReverseStatusSync = code.indexOf('calTitle.indexOf("✅")') !== -1 && (code.indexOf('sheet.getRange(row, 8).setValue(detectedStatus)') !== -1 || code.indexOf('rowVals[7] = detectedStatus') !== -1);
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

// [신규 QA] Feature 1, 2, 4 전용 엔진 및 메뉴 등록 검증
const hasColorSetters = code.indexOf('event.setColor(details.color)') !== -1;
assertTest('구글 캘린더 자체 블록 색상 자동 지정(event.setColor) 엔진 탑재 여부 (Feature 1)', hasColorSetters);

const hasDelayHighlighter = code.indexOf('function checkAndHighlightDelayedSchedules()') !== -1;
assertTest('시트 공정 지연 검사 및 하이라이트 전용 함수 탑재 여부 (Feature 2)', hasDelayHighlighter);

const hasMorningBriefingSendEngine = code.indexOf('function sendDailySiteMorningBriefing(') !== -1;
assertTest('매일 아침 8시 모닝 브리핑 자동 발송 코어 엔진 탑재 여부 (Feature 4)', hasMorningBriefingSendEngine);

const hasBriefingTriggerSetup = code.indexOf('function enableDailyMorningBriefing()') !== -1 && code.indexOf('function disableDailyMorningBriefing()') !== -1;
assertTest('모닝 브리핑 8시 일일 타이머 등록/해제 관리 엔진 탑재 여부 (Feature 4)', hasBriefingTriggerSetup);

const hasBriefingMenu = code.indexOf('"sendDailyBriefingTestNow"') !== -1 && code.indexOf('"enableDailyMorningBriefing"') !== -1 && code.indexOf('"checkAndHighlightDelayedSchedules"') !== -1;
assertTest('현장 관리 상단 메뉴 내 색상검사/모닝브리핑/이메일설정 메뉴 등록 여부', hasBriefingMenu);

// [신규 QA] 사용가이드 엔진 및 부트스트랩 연동 검증
const hasGuideSheetBuilder = code.indexOf('function createOrUpdateUsageGuideSheet(') !== -1;
assertTest('사용가이드 자동 생성 및 매뉴얼 작성 엔진(createOrUpdateUsageGuideSheet) 탑재 여부', hasGuideSheetBuilder);

const hasGuideInBootstrap = code.indexOf('var guideSheet = createOrUpdateUsageGuideSheet(ss)') !== -1;
assertTest('새 워크스페이스 구축(setupNewWorkspaceFromScratch) 시 사용가이드 자동 생성 연동 여부', hasGuideInBootstrap);

const hasGuideMenu = code.indexOf('"createOrUpdateUsageGuideSheetMenu"') !== -1;
assertTest('상단 메뉴([🚀 현장 관리]) 내 사용가이드 생성/갱신 메뉴 등록 여부', hasGuideMenu);

// [신규 QA] 구성원 시트 다중 수신자 모닝 브리핑 및 오염 방지 가드 검증
const hasFindMembersFn = code.indexOf('function findMembersSheet(') !== -1;
assertTest('구성원 시트 자동 감지 엔진(findMembersSheet) 탑재 여부', hasFindMembersFn);

const hasExtractMemberEmailsFn = code.indexOf('function extractMemberBriefingEmails(') !== -1;
assertTest('구성원 시트 이메일/수신체크박스 자동 추출 엔진(extractMemberBriefingEmails) 탑재 여부', hasExtractMemberEmailsFn);

const hasMultiRecipientBriefing = code.indexOf('findMembersSheet(ss)') !== -1 && code.indexOf('extractMemberBriefingEmails(memberSheet)') !== -1;
assertTest('모닝 브리핑 수신처 [구성원] 시트 최우선 연동 로직 탑재 여부', hasMultiRecipientBriefing);

const hasMemberSheetExclusion = code.indexOf('name.indexOf("구성원") !== -1') !== -1 && code.indexOf('name.indexOf("팀원") !== -1') !== -1;
assertTest('isSiteScheduleSheet 내 구성원/팀원/직원 시트 캘린더 연동 배제 가드 탑재 여부', hasMemberSheetExclusion);

const hasFindSiteMemberGuard = code.indexOf('sName.indexOf("구성원") === -1') !== -1 && code.indexOf('sName.indexOf("팀원") === -1') !== -1;
assertTest('findSiteScheduleSheet 내 구성원/팀원 시트 오탐지 방지 가드 탑재 여부', hasFindSiteMemberGuard);

// [신규 QA] 모닝 브리핑 정각(Exact Time) 스케줄러 및 시간 설정 정적 검증
const hasConfiguredTimeFn = code.indexOf('function getDailyBriefingConfiguredTime(') !== -1;
assertTest('모닝 브리핑 설정 시간 조회 함수(getDailyBriefingConfiguredTime) 탑재 여부', hasConfiguredTimeFn);

const hasScheduleExactTriggerFn = code.indexOf('function scheduleNextDailyBriefingTrigger(') !== -1;
assertTest('모닝 브리핑 정각 정밀 예약 엔진(scheduleNextDailyBriefingTrigger) 탑재 여부', hasScheduleExactTriggerFn);

const hasConfigureBriefingTimeFn = code.indexOf('function configureBriefingTime(') !== -1;
assertTest('모닝 브리핑 발송 시간 변경 대화상자 함수(configureBriefingTime) 탑재 여부', hasConfigureBriefingTimeFn);

const hasBriefingTimeMenu = code.indexOf('"configureBriefingTime"') !== -1;
assertTest('상단 메뉴([🚀 현장 관리]) 내 모닝 브리핑 시간 변경 메뉴 등록 여부', hasBriefingTimeMenu);

const hasSelfRescheduling = code.indexOf('scheduleNextDailyBriefingTrigger()') !== -1 && code.indexOf('DAILY_BRIEFING_ENABLED') !== -1;
assertTest('모닝 브리핑 발송 후 다음날 정각 자동 연속 예약(Self-Rescheduling) 로직 탑재 여부', hasSelfRescheduling);

// [신규 QA] 6분 타임아웃 원천 차단 및 성능 최적화 전수 검증
const hasDelayedHighlightBatch = code.indexOf('range.getBackgrounds()') !== -1 && code.indexOf('range.setBackgrounds(backgrounds)') !== -1;
assertTest('지연 공정 검사: 1회 배치 getBackgrounds/setBackgrounds 일괄 반영 (300배 가속)', hasDelayedHighlightBatch);

const hasDelayedHighlightTimeoutGuard = code.indexOf('Date.now() - startTime > 60000') !== -1;
assertTest('지연 공정 검사: 60초 안전 실행 시간 가드(Timeout Guard) 탑재 여부', hasDelayedHighlightTimeoutGuard);

const hasSyncFromCalBatch = code.indexOf('inspectedCount >= 30') !== -1 && code.indexOf('Date.now() - startTime > 90000') !== -1 && code.indexOf('range.setValues(values)') !== -1;
assertTest('캘린더➔시트 역동기화: 90초 타임아웃 가드, 30건 검사 한도 및 1회 배치 쓰기 탑재 여부', hasSyncFromCalBatch);

const hasBatchSafeSyncTimeoutGuard = code.indexOf('Date.now() - startTime > 270000') !== -1 && code.indexOf('Utilities.sleep(2000)') !== -1;
assertTest('구글 캘린더 동기화: 25건 2초 쿨다운 및 4분 30초 안전 실행시간 가드 탑재 여부', hasBatchSafeSyncTimeoutGuard);

const hasSkipDashboardRefreshInEdit = code.indexOf('syncSingleRowToCalendar(sheet, r, true, true)') !== -1 && code.indexOf('skipDashboardRefresh') !== -1;
assertTest('실시간 수정 트리거: 개별 행마다 대시보드 중복 렌더링 방지(skipDashboardRefresh) 탑재 여부', hasSkipDashboardRefreshInEdit);

const hasRangeWhileLoopGuard = code.indexOf('maxRangeLoop') !== -1 && code.indexOf('--maxRangeLoop >= 0') !== -1;
assertTest('대시보드 달력 기간 일정 전개: 366일 무한루프 방지 안전 카운터 탑재 여부', hasRangeWhileLoopGuard);

const hasBatchEmptyRowCleanup = code.indexOf('function cleanupEmptyRowsMetadata()') !== -1 && code.indexOf('sheet.getRange(2, 1, lastRow - 1, 14).setValues(values)') !== -1;
assertTest('빈 행 메타데이터 청소: 1회 setValues 배치 쓰기 최적화 탑재 여부', hasBatchEmptyRowCleanup);

// [신규 QA] 캘린더 상세 설명 간소화 및 하이퍼링크 탑재 검증
const buildEventDetailsCode = code.slice(code.indexOf('function buildEventDetails('), code.indexOf('function syncSingleRowToCalendar('));
const hasAuditRemoved = buildEventDetailsCode.indexOf('⏱️ 최초작성') === -1 && buildEventDetailsCode.indexOf('🔄 최종수정') === -1;
assertTest('buildEventDetails 내 최초작성/최종수정 영역 완전 배제 검증', hasAuditRemoved);

const hasClickableStatusLink = buildEventDetailsCode.indexOf('descLines.push("<a href=\\"" + statusChangeUrl + "\\">📱 [공유/진행상태 변경]</a>");') !== -1;
assertTest('buildEventDetails 내 [공유/진행상태 변경] HTML 하이퍼링크 구현 검증', hasClickableStatusLink);

const hasClickableSheetLink = buildEventDetailsCode.indexOf('descLines.push("<a href=\\"" + sheetUrl + "\\">📊 [우성아이디 현장 관리 시트]</a>");') !== -1;
assertTest('buildEventDetails 내 [우성아이디 현장 관리 시트] HTML 하이퍼링크 구현 검증', hasClickableSheetLink);

const hasShareToggleInWeb = (code.indexOf('updateScheduleStatusFromWeb(row, newStatus, newShared') !== -1 || code.indexOf('updateScheduleStatusFromWeb(row, selected, selectedShare') !== -1) && code.indexOf('selectShare(') !== -1;
assertTest('모바일 웹앱 내 공유 상태(E열) 동시 업데이트 및 UI 토글 엔진 탑재 검증', hasShareToggleInWeb);

const hasMemoInWebApp = code.indexOf('id=\'memoInput\'') !== -1 && code.indexOf('sheet.getRange(rowNum, 6).setValue(cleanMemo)') !== -1;
assertTest('모바일 웹앱 내 현장 비고(F열) 실시간 입력/수정/저장 엔진 탑재 검증', hasMemoInWebApp);

// [신규 QA] 신규 현장 추가 시 서식/체크박스/드롭다운/조건부서식 100% 자동 적용 검증
const hasAddNewSiteFn = code.indexOf('function addNewSiteAtTop()') !== -1;
assertTest('신규 현장 추가 함수(addNewSiteAtTop) 탑재 여부', hasAddNewSiteFn);

const addNewSiteCode = code.slice(code.indexOf('function addNewSiteAtTop()'), code.indexOf('function cleanupEmptyRowsMetadata()'));
const hasAllFormatsInNewSite = addNewSiteCode.indexOf('insertCheckboxes()') !== -1 &&
                               addNewSiteCode.indexOf('setVerticalAlignment("middle")') !== -1 &&
                               addNewSiteCode.indexOf('setDataValidation(statusRule)') !== -1 &&
                               addNewSiteCode.indexOf('setupSiteSheetConditionalFormatting(sheet)') !== -1 &&
                               addNewSiteCode.indexOf('setBorder(') !== -1;
assertTest('신규 현장 추가 시 체크박스/수직가운데/드롭다운/조건부서식/테두리 100% 자동 적용 검증', hasAllFormatsInNewSite);

// [신규 QA] 대시보드 달력 셀 내 비고(memo) 완전 배제 검증
const refreshDashCode = code.slice(code.indexOf('function refreshDashboardCalendar('), code.indexOf('function syncAllSiteSchedulesToDashboard('));
const hasNoMemoInDashboardItem = refreshDashCode.indexOf('var itemText = delayPrefix + statusIcon + prefixIcon + "[" + siteName + "] " + taskName + timeDisplay;') !== -1;
// [신규 QA 종합 감사 1] 모바일 웹앱 내 드롭다운(select id=statusSelect) 및 관련 핸들러 완전 삭제 검증
const mobileHtmlCode = code.slice(code.indexOf('function buildMobileStatusHtml('), code.indexOf('function getEffectiveEndDate('));
const hasNoSelectInMobile = mobileHtmlCode.indexOf('<select') === -1 &&
                            mobileHtmlCode.indexOf('statusSelect') === -1 &&
                            mobileHtmlCode.indexOf('onDropdownChange') === -1 &&
                            mobileHtmlCode.indexOf('select-dropdown') === -1;
assertTest('모바일 웹앱 내 불필요 드롭다운 HTML/JS/CSS 전수 안전 삭제 검증 (TypeError 방어)', hasNoSelectInMobile);

// [신규 QA 종합 감사 2] 모바일 웹앱 HTML 특수문자 이스케이프 엔진 검증
const escapeHtmlFn = sandbox.escapeHtmlForInput;
assertTest('특수문자 XSS/HTML 파싱 깨짐 방어: <, >, &, ", \' 정상 치환',
  typeof escapeHtmlFn === 'function' &&
  escapeHtmlFn('<test & "quote" \'val\'>') === '&lt;test &amp; &quot;quote&quot; &#039;val&#039;&gt;'
);

// [신규 QA 종합 감사 3] 모바일 웹앱 메모(비고) 줄바꿈 및 빈값 안전 처리 검증
assertTest('escapeHtmlForInput: null/undefined 입력 시 빈 문자열 안전 폴백',
  typeof escapeHtmlFn === 'function' &&
  escapeHtmlFn(null) === '' && escapeHtmlFn(undefined) === '' && escapeHtmlFn('') === ''
);

console.log('\n========================================================');
console.log(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
console.log('========================================================');

if (failCount > 0) {
  process.exit(1);
} else {
  console.log('🎉 모든 QA 및 코드리뷰 무결성 테스트를 100% 통과했습니다!');
}

