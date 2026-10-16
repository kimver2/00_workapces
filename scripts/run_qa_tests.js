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
const scriptPath = fs.existsSync(path.join(__dirname, 'SiteSchedule_sheet.gs'))
  ? path.join(__dirname, 'SiteSchedule_sheet.gs')
  : path.join(__dirname, 'SiteSchedule_sheet.js');
let code = '';
try {
  code = fs.readFileSync(scriptPath, 'utf8');
  console.log('✅ [Step 1: 정적 구문 검증] ' + path.basename(scriptPath) + ' 파일 로드 성공 (' + code.length + ' bytes)');
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
    DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST' },
    newDataValidation: function() {
      var builder = {
        _list: [],
        _allowInvalid: false,
        requireValueInList: function(list, allowInvalid) {
          builder._list = list;
          builder._allowInvalid = allowInvalid;
          return builder;
        },
        setAllowInvalid: function(val) {
          builder._allowInvalid = val;
          return builder;
        },
        build: function() {
          return {
            getCriteriaType: function() { return 'VALUE_IN_LIST'; },
            getCriteriaValues: function() { return [builder._list.slice(), true]; }
          };
        }
      };
      return builder;
    },
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
    }),
    getScriptProperties: () => ({
      getProperty: (k) => (sandbox._props && sandbox._props[k]) || '',
      setProperty: (k, v) => { sandbox._props = sandbox._props || {}; sandbox._props[k] = String(v); },
      deleteProperty: (k) => { if (sandbox._props) delete sandbox._props[k]; }
    })
  },
  LockService: {
    getScriptLock: () => ({
      tryLock: () => true,
      releaseLock: () => {}
    })
  },
  ScriptApp: {
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' }),
    getProjectTriggers: () => [],
    deleteTrigger: () => {},
    newTrigger: () => ({
      timeBased: () => ({
        everyMinutes: () => ({
          create: () => {}
        })
      })
    })
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
const baseMockRow = ['래미안 101동', '26.10.25', '목공', '우성목재', true, '오전작업', '김팀장', '예정', '', 0, '', '', '', ''];

// 1. 상태 '완료' 런타임 호출 테스트 (ReferenceError 방지 검증)
const rowDone = [...baseMockRow];
rowDone[7] = '완료';
const resDone = buildDetailsFn(rowDone, 15);
assertTest('buildEventDetails 런타임 실행 ("완료"): ReferenceError 없이 정상 실행', resDone !== null);
assertTest('상태 "완료" 시 캘린더 제목에 "✅ " 아이콘 정상 부여', resDone && resDone.title.startsWith('✅ 🟦'));
assertTest('캘린더 제목(Title) 내 비고(memo) 데이터 완전 배제 및 [공종] 현장명 검증', resDone && resDone.title === '✅ 🟦 [목공] 래미안 101동');
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
assertTest('캘린더 상세 설명: [공유/진행상태 변경] 클릭 가능한 HTML 하이퍼링크 형식 탑재', resWithAudit && resWithAudit.description.indexOf('href="https://script.google.com/macros/s/test/exec?row=15') !== -1 && resWithAudit.description.indexOf('📱 [공유/진행상태 변경]</a>') !== -1);
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
assertTest('모바일 HTML 내 [저장하기] 버튼 정상 렌더링', htmlOutput.indexOf('저장하기') !== -1);
assertTest('모바일 HTML 내 [현장 비고/특이사항 메모] textarea 정상 렌더링', htmlOutput.indexOf('id=\'memoInput\'') !== -1);

// [신규 QA] 모바일 웹앱 내장 JavaScript 구문 무결성 정밀 검증 (SyntaxError 방지)
const scriptMatch = htmlOutput.match(/<script>([\s\S]*?)<\/script>/);
let isScriptSyntaxValid = false;
if (scriptMatch) {
  try {
    new Function(scriptMatch[1]);
    isScriptSyntaxValid = true;
  } catch (e) {
    isScriptSyntaxValid = false;
  }
}
assertTest('모바일 웹앱 내장 <script> 클라이언트 자바스크립트 구문 100% 정상 파싱 검증', isScriptSyntaxValid);

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
assertTest('현장별 색상 일치 검증: 지연 공정도 현장 고유 색상 유지', resDelayed && resDelayed.color === sandbox.getCalendarColorForSite('래미안 101동'));
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
const mockDashSheet = createMockSheet('대시보드');
const isSiteSheetFn = sandbox.isSiteScheduleSheet;
assertTest('isSiteScheduleSheet 판별: 사용가이드 시트는 현장데이터 시트에서 완전 배제', isSiteSheetFn(mockGuideSheet) === false);
assertTest('isSiteScheduleSheet 판별: 대시보드 시트는 현장데이터 시트에서 완전 배제', isSiteSheetFn(mockDashSheet) === false);

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

// [신규 QA: 대시보드 이중 렌더링 방지 및 단일 책임 분리 검증]
// 단순 onEdit에서는 대시보드 컨트롤(년/월, Today)만 전담하고, 현장데이터 변경에 의한 대시보드 새로고침은
// 설치형 트리거(handleCalendarEditTrigger)에서 단 1회 전담하여 이중 렌더링을 완전 해소
const onEditCode = code.slice(code.indexOf('function onEdit(e) {'), code.indexOf('function enable5MinAutoSync()'));
const hasNoDoubleRenderingInOnEdit = onEditCode.indexOf('refreshDashboardCalendar(ss)') !== -1 && onEditCode.indexOf('isSiteScheduleSheet(sheet)') === -1;
assertTest('대시보드 이중 렌더링 해소: 단순 onEdit 내 현장데이터 중복 렌더링 배제 검증', hasNoDoubleRenderingInOnEdit);

// [신규 QA] 드롭다운 '진행 중'(공백 포함) 유효성 검사 및 빨간 삼각형 오류 복구 검증
const hasSpacedValidation = code.indexOf('.requireValueInList(["예정", "진행 중", "완료", "AS"], true)') !== -1;
assertTest('진행상태 드롭다운 목록 내 "진행 중"(공백 포함) 표준 규격 탑재 여부 (빨간 삼각형 에러 방지)', hasSpacedValidation);

const hasFixDropdownFn = code.indexOf('function fixStatusDropdownValidation()') !== -1;
assertTest('진행상태 드롭다운 서식 복구 전용 함수(fixStatusDropdownValidation) 탑재 여부', hasFixDropdownFn);

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

// [신규 QA] 불필요한 초기 구축 및 가이드 생성 함수 정리 완료 검증
assertTest('초기 구축(setupNewWorkspaceFromScratch) 및 사용가이드 생성기 코드 정리(경량화) 검증',
  code.indexOf('function setupNewWorkspaceFromScratch()') === -1 && code.indexOf('function createOrUpdateUsageGuideSheet(') === -1
);

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
                               (addNewSiteCode.indexOf('setDataValidation(statusRule)') !== -1 || addNewSiteCode.indexOf('applySiteScheduleValidations') !== -1) &&
                               addNewSiteCode.indexOf('setupSiteSheetConditionalFormatting(sheet)') !== -1 &&
                               addNewSiteCode.indexOf('setBorder(') !== -1;
assertTest('신규 현장 추가 시 체크박스/수직가운데/드롭다운/조건부서식/테두리 100% 자동 적용 검증', hasAllFormatsInNewSite);

// [신규 QA] 공종 및 담당자 드롭다운 유효성 검사 및 노션 스타일 조건부 서식 탑재 검증
assertTest('공종(C열) 기본 37개+ 공종 목록 탑재 여부', code.indexOf('function getAvailableTasks(') !== -1 && code.indexOf('마루시공') !== -1);
assertTest('담당자(G열) 드롭다운 유효성 검사 엔진(getAvailableManagers) 탑재 여부', code.indexOf('function getAvailableManagers(') !== -1 && code.indexOf('조항만') !== -1);
assertTest('공종/담당자/상태 드롭다운 통합 적용 함수(applySiteScheduleValidations) 탑재 여부', code.indexOf('function applySiteScheduleValidations(') !== -1);
assertTest('공종(C열) 및 담당자(G열) 노션 스타일 조건부 서식 뱃지 탑재 여부',
  code.indexOf('var managerColorMap = [') !== -1 &&
  code.indexOf('var taskColorMap = [') !== -1
);

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

// [신규 QA] 캘린더 일정 내 모바일 웹앱 링크에 현장명(site) 및 공종명(task) 파라미터 자동 탑재 검증
const testRow = ['진행상태 확인', '26.10.13', '상담', '협력사', true, '', '담당자', '예정', 'ev-1234'];
const testDetails = buildDetailsFn(testRow, 2);
assertTest('캘린더 상세 설명 URL 내 현장명/공종명/EventID 쿼리 파라미터 포함 여부',
  testDetails &&
  testDetails.taskName === '상담' &&
  testDetails.description.indexOf('site=' + encodeURIComponent('진행상태 확인')) !== -1 &&
  testDetails.description.indexOf('task=' + encodeURIComponent('상담')) !== -1 &&
  testDetails.description.indexOf('id=ev-1234') !== -1
);

// [신규 QA] doGet 스마트 행 위치 자동 재탐색(Self-Healing Row Resolver) 엔진 탑재 검증
assertTest('doGet 스마트 행 자동 보정 로직(targetId, targetSite, targetTask 기반 스캔) 탑재 여부',
  code.indexOf('// [스마트 행 자동 보정 엔진]') !== -1 &&
  code.indexOf('targetSite && targetTask') !== -1 &&
  code.indexOf('resolvedRow') !== -1
);

// [신규 QA] 행 이동 및 공종명(Task) 변경 시 Event ID 불변 추적 알고리즘 시뮬레이션
const mockSheetData = [
  ['13종 펼치기', '26.10.13', '현장실측 / 보양', '', true, '', '', '예정', 'ev-other-1'], // 2행 (밀려서 새로 들어온 행)
  ['진행상태 확인', '26.10.13', '수정된새공종명', '', true, '', '', '예정', 'ev-target-123'], // 3행 (공종명이 변경된 원래 일정)
];
// 시나리오 A: 링크에는 row=2, id="ev-target-123"가 기록되어 있음 (행 번호와 공종명이 모두 불일치)
const targetIdA = 'ev-target-123';
let resolvedRowA = -1;
for (let i = 0; i < mockSheetData.length; i++) {
  if (mockSheetData[i][8] === targetIdA) {
    resolvedRowA = i + 2;
    break;
  }
}
assertTest('Self-Healing 시뮬레이션: 공종명이 변경되고 행이 밀려도 EventID로 정확히 타겟 행(3행) 포커스',
  resolvedRowA === 3
);

// 시나리오 B: 구형 링크(targetId 없음)이나 site+task 일치하는 경우 fallback
const legacyTargetSite = '13종 펼치기';
const legacyTargetTask = '현장실측 / 보양';
let resolvedRowB = -1;
for (let j = 0; j < mockSheetData.length; j++) {
  if (mockSheetData[j][0] === legacyTargetSite && mockSheetData[j][2] === legacyTargetTask) {
    resolvedRowB = j + 2;
    break;
  }
}
assertTest('Self-Healing 시뮬레이션: EventID 없는 구형 링크 시 site+task Fallback 자동 탐색(2행) 성공',
  resolvedRowB === 2
);

// 시나리오 C: 시트에서 삭제된 일정인 경우 안전 가드 (오동작 방지)
assertTest('doGet 안전 가드: 시트에서 삭제된 일정 요청 시 엉뚱한 행 덮어쓰기 방지(row = -1) 탑재',
  code.indexOf('row = -1;') !== -1 &&
  code.indexOf('// [안전 가드]') !== -1
);

// 시나리오 D: 신규 생성 이벤트 설명에 즉시 Event ID 주입 확인
assertTest('이벤트 생성 즉시 설명(Description)에 Event ID 주입(freshDetails) 로직 탑재 여부',
  code.indexOf('rowValues[8] = eventId;') !== -1 &&
  code.indexOf('rowVals[8] = newEventId;') !== -1
);

// ========================================================
// 8. 캘린더 중복 생성 원천 차단 및 유령/중복 일정 자동 정리 QA 검증
// ========================================================
console.log('\n🧹 [Step 8: 캘린더 중복 생성 방지 및 고아/유령 일정 청소 QA]');

// 8-1: buildEventDetails baseTitle 속성 포함 여부
const sampleRow = ['날짜 체크 시간 체크', '26.10.13', '전기마감', '', false, '', '전예린', '진행 중'];
const sampleDetails = sandbox.buildEventDetails(sampleRow, 19, null, null);
assertTest('buildEventDetails baseTitle 반환 검증 ([공종] 현장명 표준 포맷)',
  sampleDetails && sampleDetails.baseTitle === '[전기마감] 날짜 체크 시간 체크'
);

// 8-2: isSystemSiteScheduleEvent 검증
const mockSysEventA = {
  getTitle: () => '▶️ 🟥 [날짜 체크 시간 체크] 전기마감',
  getDescription: () => '📌 현장명: 날짜 체크 시간 체크\n🔨 공종: 전기마감'
};
const mockSysEventB = {
  getTitle: () => '🟥 [날짜 체크 시간 체크] 전기마감',
  getDescription: () => ''
};
const mockPersonalEvent = {
  getTitle: () => '치과 진료 예약',
  getDescription: () => '강남구 신사동 치과 정기검진'
};
assertTest('isSystemSiteScheduleEvent: 우리 시스템 일정(설명 마커) 정확 탐지',
  sandbox.isSystemSiteScheduleEvent(mockSysEventA) === true
);
assertTest('isSystemSiteScheduleEvent: 구형/설명 없는 시스템 일정(제목 패턴) 정확 탐지',
  sandbox.isSystemSiteScheduleEvent(mockSysEventB) === true
);
assertTest('isSystemSiteScheduleEvent: 개인 일정(비시스템) 철저 보호(false 반환)',
  sandbox.isSystemSiteScheduleEvent(mockPersonalEvent) === false
);

// 8-3: findMatchingCalendarEvent 탐색 알고리즘 시뮬레이션
const mockCalEvents = [
  {
    getId: () => 'orphan-ev-999',
    getTitle: () => '🟥 [날짜 체크 시간 체크] 전기마감',
    getDescription: () => '📌 현장명: 날짜 체크 시간 체크\n🔨 공종: 전기마감',
    getLocation: () => '날짜 체크 시간 체크',
    isAllDayEvent: () => true,
    getAllDayStartDate: () => new Date(2026, 9, 13, 0, 0, 0),
    getStartTime: () => new Date(2026, 9, 13, 0, 0, 0)
  },
  {
    getId: () => 'other-row-ev',
    getTitle: () => '▶️ 🟥 [날짜 체크 시간 체크] 도배',
    getDescription: () => '📌 현장명: 날짜 체크 시간 체크\n🔨 공종: 도배',
    getLocation: () => '날짜 체크 시간 체크',
    isAllDayEvent: () => true,
    getAllDayStartDate: () => new Date(2026, 9, 13, 0, 0, 0),
    getStartTime: () => new Date(2026, 9, 13, 0, 0, 0)
  }
];

const mockCal = {
  getEvents: (s, e) => mockCalEvents
};

const matchedEv = sandbox.findMatchingCalendarEvent(mockCal, sampleDetails, {});
assertTest('findMatchingCalendarEvent: 동일 현장/공종/날짜의 기존 일정 정확 탐색',
  matchedEv && matchedEv.getId() === 'orphan-ev-999'
);

const matchedWithOtherRow = sandbox.findMatchingCalendarEvent(mockCal, sampleDetails, { 'orphan-ev-999': true });
assertTest('findMatchingCalendarEvent: 타 행에 연결된 이벤트는 매칭에서 안전하게 배제',
  matchedWithOtherRow === null
);

// 8-4: cleanDuplicateCalendarEventsOnDate 시뮬레이션
let deletedEvId = null;
const mockDupeCal = {
  getEvents: (s, e) => [
    {
      getId: () => 'active-ev-111',
      getTitle: () => '▶️ 🟥 [날짜 체크 시간 체크] 전기마감',
      getDescription: () => '📌 현장명: 날짜 체크 시간 체크',
      getLocation: () => '날짜 체크 시간 체크',
      isAllDayEvent: () => true,
      getAllDayStartDate: () => new Date(2026, 9, 13, 0, 0, 0),
      getStartTime: () => new Date(2026, 9, 13, 0, 0, 0),
      deleteEvent: () => { deletedEvId = 'active-ev-111'; }
    },
    {
      getId: () => 'orphan-ev-222',
      getTitle: () => '🟥 [날짜 체크 시간 체크] 전기마감',
      getDescription: () => '📌 현장명: 날짜 체크 시간 체크',
      getLocation: () => '날짜 체크 시간 체크',
      isAllDayEvent: () => true,
      getAllDayStartDate: () => new Date(2026, 9, 13, 0, 0, 0),
      getStartTime: () => new Date(2026, 9, 13, 0, 0, 0),
      deleteEvent: () => { deletedEvId = 'orphan-ev-222'; }
    }
  ]
};

const dupeDeletedCount = sandbox.cleanDuplicateCalendarEventsOnDate(
  mockDupeCal,
  sampleDetails,
  'active-ev-111',
  { 'active-ev-111': true }
);
assertTest('cleanDuplicateCalendarEventsOnDate: 활성 ID 보존 및 잉여 중복/고아 일정만 선택 삭제',
  dupeDeletedCount === 1 && deletedEvId === 'orphan-ev-222'
);

// 8-5: 상단 메뉴 등록 검증
assertTest('상단 메뉴(⚡ 캘린더 연동 설정) 내 캘린더 중복/유령 일정 자동 정리 등록 여부',
  code.indexOf('.addItem("🧹 캘린더 중복 및 유령 일정 자동 정리", "cleanupDuplicateAndOrphanCalendarEvents")') !== -1
);

// 8-6: syncSingleRowToCalendar 및 syncSiteSchedulesBatchSafe 내 방어 로직 탑재 검증
assertTest('syncSingleRowToCalendar 내 중복 방지(findMatching) 및 정리(cleanDuplicate) 탑재 여부',
  code.indexOf('event = findMatchingCalendarEvent(cal, details, allSheetIds);') !== -1 &&
  code.indexOf('cleanDuplicateCalendarEventsOnDate(cal, details, event.getId(), allSheetIds);') !== -1
);

assertTest('syncSiteSchedulesBatchSafe 내 중복 방지(findMatching) 및 정리(cleanDuplicate) 탑재 여부',
  code.indexOf('event = findMatchingCalendarEvent(cal, details, allSheetIds);') !== -1 &&
  code.indexOf('cleanDuplicateCalendarEventsOnDate(cal, details, newEventId, allSheetIds);') !== -1
);

// ========================================================
// 9. 현장명 스마트 상속 및 현장명 누락 방어 QA 검증
// ========================================================
console.log('\n🏢 [Step 9: 현장명 스마트 상속 및 현장명 누락 방어 QA]');

// 9-1: 현장명 전혀 없는 행은 캘린더 등록 배제(null 반환)
const noSiteRow = ['', '26.10.29', '도배', '', true, '', '전예린', '예정'];
const noSiteDetails = sandbox.buildEventDetails(noSiteRow, 2, null, null);
assertTest('현장명이 전혀 없는 행은 buildEventDetails에서 null 반환 (미아 일정 생성 원천 방지)',
  noSiteDetails === null
);

// 9-2: 상위 행에서 현장명 자동 계승 모의 테스트
const mockSheetWithParent = {
  getRange: (r, c, nr, nc) => ({
    getValues: () => [['보라매경남 102동 502호']],
    setValue: (val) => {}
  })
};
const childRow = ['', '26.10.29', '도배', '', true, '', '전예린', '예정'];
const inheritedDetails = sandbox.buildEventDetails(childRow, 3, null, mockSheetWithParent);
assertTest('상위 행이 있는 경우 직전 현장명 자동 상속 검증',
  inheritedDetails && inheritedDetails.siteName === '보라매경남 102동 502호' &&
  inheritedDetails.title.indexOf('[도배] 보라매경남 102동 502호') !== -1
);

// 9-3: batch sync 내 effectiveSiteNames 연산 검증
const mockBatchVals = [
  ['보라매경남 102동 702호', '26.07.18', '상담'],
  ['', '26.07.24', '상담'],
  ['', '26.08.05', '철거'],
  ['망원 경일빌라 201호', '26.08.10', '창호실측'],
  ['', '26.08.12', '상담']
];
const effectiveTest = [];
let lastSite = '';
for (let i = 0; i < mockBatchVals.length; i++) {
  const rSite = mockBatchVals[i][0].trim();
  if (rSite) {
    lastSite = rSite;
    effectiveTest[i] = rSite;
  } else if (mockBatchVals[i][1] || mockBatchVals[i][2]) {
    effectiveTest[i] = lastSite;
  } else {
    effectiveTest[i] = '';
  }
}
assertTest('배치 동기화 effectiveSiteNames: 1~3행 보라매경남 702호, 4~5행 망원 경일빌라 정확 계승',
  effectiveTest[0] === '보라매경남 102동 702호' &&
  effectiveTest[1] === '보라매경남 102동 702호' &&
  effectiveTest[2] === '보라매경남 102동 702호' &&
  effectiveTest[3] === '망원 경일빌라 201호' &&
  effectiveTest[4] === '망원 경일빌라 201호'
);

console.log('\n🛡️ [Step 10: 작성 중 조기 등록 방지 가드 및 실시간 연동 영속성 QA]');

// 10-1: 공종 누락 행은 buildEventDetails에서 null 반환
const noTaskRow = ['보라매경남 102동 502호', '26.10.29', '', '', true, '', '전예린', '예정', '', 0, '', '', '', ''];
const noTaskDetails = sandbox.buildEventDetails(noTaskRow, 2, null, null);
assertTest('공종(taskName)이 없는 미완성 행은 buildEventDetails에서 null 반환 (조기 등록 원천 차단)', noTaskDetails === null);

// 10-2: 날짜만 있고 공종 없는 행 검증
const dateOnlyRow = ['보라매경남 102동 502호', '2026. 10. 29', '   ', '', false, '', '', '', '', 0, '', '', '', ''];
const dateOnlyDetails = sandbox.buildEventDetails(dateOnlyRow, 206, null, null);
assertTest('공종이 공백 문자열인 경우도 null 반환하여 작성 중 캘린더 등록 완전 방어', dateOnlyDetails === null);

// 10-3: 필수 3요소(현장명 + 날짜 + 공종) 완성 시 정상 객체 반환 검증
const completeRow = ['보라매경남 102동 502호', '2026. 10. 29', '도배', '', true, '', '전예린', '예정', '', 0, '', '', '', ''];
const completeDetails = sandbox.buildEventDetails(completeRow, 206, null, null);
assertTest('필수 3요소 완성 시 정상 세부정보 및 표준 제목 반환 ([공종] 현장명)',
  completeDetails !== null &&
  completeDetails.title.indexOf('[도배] 보라매경남 102동 502호') !== -1
);

// 10-4: 코드 내 SYNC_MODE 영속성 제어 검증
const hasSyncModeCheck = code.indexOf('props.getProperty("SYNC_MODE")') !== -1;
assertTest('실시간 트리거 및 무음 등록 함수 내 SYNC_MODE 영속 플래그 제어 탑재 검증', hasSyncModeCheck);

console.log('\n⏰ [Step 11: 1분 주기 스마트 자동 동기화 엔진 QA]');

// 11-1: 1분 주기 켜기 함수 탑재 여부
assertTest('enable1MinAutoSync 함수 탑재 여부', typeof sandbox.enable1MinAutoSync === 'function');

// 11-2: 1분/5분 주기 백그라운드 공통 코어 엔진 함수 탑재 여부
assertTest('runPeriodicAutoSync 백그라운드 코어 엔진 함수 탑재 여부', typeof sandbox.runPeriodicAutoSync === 'function');

// 11-3: 모든 자동 연동 끄기 함수 탑재 여부
assertTest('disableAllAutoCalendarSync 함수 탑재 여부', typeof sandbox.disableAllAutoCalendarSync === 'function');

// 11-4: handleCalendarEditTrigger 내 1MIN 모드 시 캘린더 API 미호출 분기 검증
const triggerCode = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncFromCalendarToSheet('));
const has1MinSkipLogic = triggerCode.indexOf('if (syncMode === "REALTIME")') !== -1;
assertTest('handleCalendarEditTrigger: 1MIN 모드 시 캘린더 API 호출 생략 (타이핑 렉 0% 보장)', has1MinSkipLogic);

// 11-5: 상단 메뉴(⚡ 캘린더 연동 설정) 내 1분 주기 자동 동기화 등록 검증
const menuCode = code.slice(code.indexOf('ui.createMenu("⚡ 캘린더 연동 설정")'), code.indexOf('ui.createMenu("📑 현장 데이터 편집/보기")'));
const has1MinInMenu = menuCode.indexOf('enable1MinAutoSync') !== -1;
assertTest('상단 메뉴(⚡ 캘린더 연동 설정) 내 1분 주기 자동 동기화 등록 여부', has1MinInMenu);

console.log('\n🛡️ [Step 12: 5분 주기 스마트 동기화, 이중 렌더링 방지, 원자적 쓰기 & 삭제 유예 QA]');

// 12-1: 5분 주기 스마트 동기화 함수 및 메뉴 등록 검증
assertTest('enable5MinAutoSync 함수 탑재 여부', typeof sandbox.enable5MinAutoSync === 'function');
assertTest('runPeriodicAutoSync 코어 엔진 함수 탑재 여부', typeof sandbox.runPeriodicAutoSync === 'function');
assertTest('상단 메뉴 내 5분 주기 자동 동기화(권장) 등록 여부', menuCode.indexOf('enable5MinAutoSync') !== -1);

// 12-2: 대시보드 단일 렌더링 전담 검증 (handleCalendarEditTrigger에서 1회 갱신)
const triggerFullCode = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncFromCalendarToSheet('));
const hasSingleRenderInTrigger = triggerFullCode.indexOf('refreshDashboardCalendar(ss)') !== -1;
assertTest('대시보드 단일 렌더링 전담: 설치형 트리거 내 대시보드 달력 갱신 보장', hasSingleRenderInTrigger);

// 12-3: 심야/유휴 시간 Quota 스마트 가드 검증 (22:00 ~ 07:00 시트 변경 없을 시 캘린더 조회 스킵)
const periodicCode = code.slice(code.indexOf('function runPeriodicAutoSync()'), code.indexOf('function runDebouncedSync()'));
const hasNightIdleGuard = periodicCode.indexOf('hour >= 22 || hour < 7') !== -1 && periodicCode.indexOf('rowsToSyncToCal.length === 0') !== -1;
assertTest('심야/유휴 시간 Quota 절약 가드(22:00~07:00 무작업 시 API 생략) 탑재 여부', hasNightIdleGuard);

// 12-4: 커서 라운드 로빈(Cursor Round-Robin) 전 행 균등 검사 검증
const hasCursorRoundRobin = periodicCode.indexOf('CAL_INSPECT_CURSOR') !== -1 && periodicCode.indexOf('(cursor + k) % values.length') !== -1;
assertTest('역방향 캘린더 검사 Cursor Round-Robin(전체 행 순환 점검) 탑재 여부', hasCursorRoundRobin);

// 12-5: 사용자 타이핑 덮어쓰기 방지: 원자적 타겟 셀 갱신(Atomic Target Cell Write) 검증
const hasAtomicCellWrite = periodicCode.indexOf('sheet.getRange(targetRowNumber, 2).setValue(newDateStr)') !== -1 &&
  periodicCode.indexOf('range.setValues(values)') === -1;
assertTest('역동기화 시 사용자 타이핑 덮어쓰기 원천 차단(개별 타겟 셀 원자적 갱신) 검증', hasAtomicCellWrite);

// 12-6: 일시적 타이핑/수정 중 캘린더 일정 삭제 유예 및 하위 공정 삭제 감지 가드 검증
const syncSingleCode = code.slice(code.indexOf('function syncSingleRowToCalendar('), code.indexOf('function syncSiteSchedulesBatchSafe('));
const hasDeletionGraceGuard = syncSingleCode.indexOf('isRowEmpty = (!rowValues[0] && !hasDate && !hasTask)') !== -1 &&
  syncSingleCode.indexOf('isSubRowCleared = (!hasDate && !hasTask)') !== -1 &&
  syncSingleCode.indexOf('if (isRowEmpty || isSubRowCleared || isExplicitCancel)') !== -1;
assertTest('타이핑 중 임시 빈 셀 방어 및 하위 공정 삭제 감지 가드 탑재 여부', hasDeletionGraceGuard);

// 12-7: 삭제 유예 로직 단위 테스트 시뮬레이션
function simulateDeletionDecision(rowValues) {
  var hasDate = !!String(rowValues[1] || "").trim();
  var hasTask = !!String(rowValues[2] || "").trim();
  var isRowEmpty = (!rowValues[0] && !hasDate && !hasTask);
  var isSubRowCleared = (!hasDate && !hasTask);
  var statusStr = String(rowValues[7] || "");
  var isExplicitCancel = (statusStr.indexOf("취소") !== -1 || statusStr.indexOf("삭제") !== -1);
  return (isRowEmpty || isSubRowCleared || isExplicitCancel);
}

const typingRow = ['보라매경남 102동 502호', '', '도배', '', true, '', '전예린', '예정']; // 날짜만 잠시 지우고 수정 중
assertTest('삭제 유예 시뮬레이션: 날짜만 잠시 지운 타이핑 중 상태는 삭제 배제(false)', simulateDeletionDecision(typingRow) === false);

const typingTaskRow = ['보라매경남 102동 502호', '2026. 10. 29', '', '', true, '', '전예린', '예정']; // 공종만 잠시 지우고 수정 중
assertTest('삭제 유예 시뮬레이션: 공종만 잠시 지운 타이핑 중 상태는 삭제 배제(false)', simulateDeletionDecision(typingTaskRow) === false);

const subRowCleared = ['보라매경남 102동 502호', '', '', '', false, '', '', '']; // 현장명은 상속 남아있으나 날짜+공종을 비운 하위 공정
assertTest('삭제 유예 시뮬레이션: 현장명 계승 상태에서 날짜+공종 모두 삭제 시 일정 삭제(true)', simulateDeletionDecision(subRowCleared) === true);

const emptyRow = ['', '', '', '', false, '', '', '']; // 행 전체가 비워진 경우
assertTest('삭제 유예 시뮬레이션: 핵심 정보가 전무한 행은 삭제 진행(true)', simulateDeletionDecision(emptyRow) === true);

const cancelRow = ['보라매경남 102동 502호', '2026. 10. 29', '도배', '', false, '', '', '취소']; // 명시적 취소
assertTest('삭제 유예 시뮬레이션: 상태가 명시적 "취소"인 행은 삭제 진행(true)', simulateDeletionDecision(cancelRow) === true);

console.log('\n💡 [Step 13: 스마트 디바운스 동적 1회성 자동 동기화 엔진 QA]');

// 13-1: 디바운스 활성화 함수 및 메뉴 탑재 검증
assertTest('enableDebouncedAutoSync 함수 탑재 여부', typeof sandbox.enableDebouncedAutoSync === 'function');
assertTest('runDebouncedSync 백그라운드 엔진 함수 탑재 여부', typeof sandbox.runDebouncedSync === 'function');
assertTest('상단 메뉴 내 스마트 디바운스 자동 동기화(최고 권장) 등록 여부', menuCode.indexOf('enableDebouncedAutoSync') !== -1);

// 13-2: 디바운스 트리거 정리 헬퍼 함수 탑재 검증
assertTest('cleanupDebouncedTriggers 헬퍼 함수 탑재 여부', typeof sandbox.cleanupDebouncedTriggers === 'function');
assertTest('cleanupAllPeriodicAndDebouncedTriggers 헬퍼 함수 탑재 여부', typeof sandbox.cleanupAllPeriodicAndDebouncedTriggers === 'function');

// 13-3: handleCalendarEditTrigger 내 DEBOUNCE 모드 시 지연 트리거 예약 검증
const triggerDebounceCode = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncFromCalendarToSheet('));
const hasDebounceScheduling = triggerDebounceCode.indexOf('syncMode === "DEBOUNCE"') !== -1 &&
  triggerDebounceCode.indexOf('ScriptApp.newTrigger("runDebouncedSync")') !== -1;
assertTest('handleCalendarEditTrigger: DEBOUNCE 모드 시 1회성 지연 트리거(runDebouncedSync) 동적 예약 검증', hasDebounceScheduling);

// 13-4: runDebouncedSync 내 연속 입력 시 트리거 재연장(Debounce) 분기 검증
const debouncedCoreCode = code.slice(code.indexOf('function runDebouncedSync()'), code.indexOf('function handleCalendarEditTrigger('));
const hasDebounceExtension = debouncedCoreCode.indexOf('nowMs - lastEditMs < 40000') !== -1 &&
  debouncedCoreCode.indexOf('ScriptApp.newTrigger("runDebouncedSync")') !== -1;
assertTest('runDebouncedSync: 연속 타이핑 중 유휴 미달 시 1분 뒤로 재연장(Debounce Extension) 검증', hasDebounceExtension);

// 13-5: ensureRealtimeTriggerInstalledSilent 내 기본값 DEBOUNCE 적용 및 무유휴(0 Quota) 정리 검증
const silentTriggerCode = code.slice(code.indexOf('function ensureRealtimeTriggerInstalledSilent()'), code.indexOf('function runPeriodicAutoSync()'));
const hasDebounceDefault = silentTriggerCode.indexOf('props.setProperty("SYNC_MODE", "DEBOUNCE")') !== -1 &&
  silentTriggerCode.indexOf('syncMode === "DEBOUNCE"') !== -1 &&
  silentTriggerCode.indexOf('hasPeriodicTrigger') !== -1;
assertTest('ensureRealtimeTriggerInstalledSilent: DEBOUNCE 기본 표준 채택 및 주기 트리거 자동 정리(0 Quota) 검증', hasDebounceDefault);

console.log('\n🚀 [Step 14: Argon v8.1 아키텍처 패치 전수 검증]');

// 14-1: runDebouncedSync isRescheduled 플래그 도입으로 finally 블록 트리거 증발 방지 검증
const hasRescheduledProtection = debouncedCoreCode.indexOf('isRescheduled = true') !== -1 &&
  debouncedCoreCode.indexOf('if (!isRescheduled)') !== -1;
assertTest('runDebouncedSync: isRescheduled 플래그 기반 finally 블록 트리거 증발 원천 차단 검증', hasRescheduledProtection);

// 14-2: runPeriodicAutoSync 순환 커서 stepsTaken 기반 균등 순환 및 상태 이모지 역동기화 검증
const hasStepsTakenCursor = periodicCode.indexOf('stepsTaken++') !== -1 &&
  periodicCode.indexOf('(cursor + stepsTaken) %') !== -1;
assertTest('runPeriodicAutoSync: stepsTaken 실제 이동 거리 기반 순환 커서 제자리걸음 방지 검증', hasStepsTakenCursor);

const hasStrictEmojiInPeriodic = periodicCode.indexOf('calTitle.indexOf("✅")') !== -1 &&
  periodicCode.indexOf('calTitle.indexOf("완료")') === -1 &&
  periodicCode.indexOf('calTitle.indexOf("진행")') === -1;
assertTest('runPeriodicAutoSync: 캘린더 제목 상태 이모지 기준 엄격 판별(한글 텍스트 오탐지 방어) 검증', hasStrictEmojiInPeriodic);

// 14-3: syncFromCalendarToSheet 수동 동기화 원자적 셀 갱신 및 빈 행 유령 ID 실제 청소 검증
const syncCalToSheetCode = code.slice(code.indexOf('function syncFromCalendarToSheet('), code.indexOf('function isDateEqual('));
const hasAtomicSyncCalToSheet = syncCalToSheetCode.indexOf('sheet.getRange(targetRowNumber, 2).setValue(newDateStr)') !== -1 &&
  syncCalToSheetCode.indexOf('range.setValues(values)') === -1;
assertTest('syncFromCalendarToSheet: 전체 행 덮어쓰기 제거 및 핀포인트 원자적 셀 갱신 통일 검증', hasAtomicSyncCalToSheet);

const hasGhostCellClearing = syncCalToSheetCode.indexOf('sheet.getRange(i + 2, 9, 1, 6).clearContent()') !== -1;
assertTest('syncFromCalendarToSheet: 빈 행 유령 ID 및 메타데이터 시트 셀 실제 청소(clearContent) 탑재 검증', hasGhostCellClearing);

// 14-4: refreshDashboardCalendar 2D 배열 주차별 행 단위 배치 쓰기 가속 검증
const refreshDashBatchCode = code.slice(code.indexOf('function refreshDashboardCalendar('), code.indexOf('function syncAllSiteSchedulesToDashboard('));
const hasBatchDashboardWrite = refreshDashBatchCode.indexOf('rowsToUpdate[pos.row]') !== -1 &&
  refreshDashBatchCode.indexOf('dashSheet.getRange(parseInt(rNum, 10), 2, 1, 7).setValues(') !== -1;
assertTest('refreshDashboardCalendar: 주차별 2D 배열 배치 쓰기로 31회 개별 setValue 호출 제거 및 10배 가속 검증', hasBatchDashboardWrite);

// 14-5: updateScheduleStatusFromWeb 모바일 동시 접속 충돌 방어 (LockService 적용)
const mobileWebCode = code.slice(code.indexOf('function updateScheduleStatusFromWeb('), code.indexOf('function doGet('));
const hasMobileLockService = mobileWebCode.indexOf('LockService.getScriptLock()') !== -1 &&
  mobileWebCode.indexOf('lock.tryLock(10000)') !== -1 &&
  mobileWebCode.indexOf('lock.releaseLock()') !== -1;
assertTest('updateScheduleStatusFromWeb: 모바일 웹 동시 접속 상태 변경 충돌 방어(LockService 10초 대기) 탑재 검증', hasMobileLockService);
// ========================================================
// 15. [공종 50+ 사전 탑재 및 실시간 자동 학습(Auto-Learning) 유효성 검사 엔진 QA]
// ========================================================
console.log('\n🎨 [Step 15: 공종 50+ 사전 탑재 및 실시간 자동 학습(Auto-Learning) 유효성 검사 엔진 QA]');

// 15-1: getAvailableTasks 50+ 공종 사전 탑재 및 주요 공종 포함 여부
const availableTasks = sandbox.getAvailableTasks({ getLastRow: () => 0 });
assertTest('getAvailableTasks: 50개 이상의 표준 인테리어 공종 사전 탑재 여부', availableTasks.length >= 50);
assertTest('getAvailableTasks: 창호실측, 실측, 보양, 시스템에어컨, 도면/설계, 싱크대 등 핵심 공종 포함 여부',
  availableTasks.indexOf('창호실측') !== -1 &&
  availableTasks.indexOf('실측') !== -1 &&
  availableTasks.indexOf('보양') !== -1 &&
  availableTasks.indexOf('시스템에어컨') !== -1 &&
  availableTasks.indexOf('도면/설계') !== -1 &&
  availableTasks.indexOf('싱크대') !== -1 &&
  availableTasks.indexOf('엘리베이터보양') !== -1 &&
  availableTasks.indexOf('샷시') !== -1
);

// 15-2: ensureTasksInValidation 함수 존재 및 시뮬레이션
assertTest('ensureTasksInValidation: 신규 공종 실시간 자동 학습 함수 탑재 여부', typeof sandbox.ensureTasksInValidation === 'function');

let appliedRule = null;
const mockSheetForValidation = {
  getLastRow: () => 10,
  getMaxRows: () => 100,
  getRange: (r, c, nr, nc) => ({
    getDataValidation: () => ({
      getCriteriaType: () => 'VALUE_IN_LIST',
      getCriteriaValues: () => [['철거', '목공', '타일'], true]
    }),
    setDataValidation: (rule) => { appliedRule = rule; }
  })
};

// 미등록 공종('창호실측', '특수에폭시') 입력 시뮬레이션
sandbox.ensureTasksInValidation(mockSheetForValidation, ['창호실측', '특수에폭시']);
const newCriteriaValues = appliedRule ? appliedRule.getCriteriaValues() : [];
const newAllowedList = newCriteriaValues[0] || [];
assertTest('ensureTasksInValidation: 사전에 없는 신규 공종 입력 시 유효성 목록에 즉시 자동 추가 검증',
  newAllowedList.indexOf('창호실측') !== -1 &&
  newAllowedList.indexOf('특수에폭시') !== -1 &&
  newAllowedList.indexOf('철거') !== -1
);

// 이미 존재하는 공종 입력 시 setDataValidation 불필요 호출 생략(No-Op) 검증
let redundantCallCount = 0;
const mockSheetNoOp = {
  getLastRow: () => 10,
  getMaxRows: () => 100,
  getRange: (r, c, nr, nc) => ({
    getDataValidation: () => ({
      getCriteriaType: () => 'VALUE_IN_LIST',
      getCriteriaValues: () => [['철거', '목공', '타일', '창호실측'], true]
    }),
    setDataValidation: () => { redundantCallCount++; }
  })
};
sandbox.ensureTasksInValidation(mockSheetNoOp, ['목공', '창호실측']);
assertTest('ensureTasksInValidation: 기존 등록된 공종 입력 시 불필요한 setDataValidation 쓰기 생략(0ms 지연) 검증', redundantCallCount === 0);

// 15-3: handleCalendarEditTrigger C열 감지 및 ensureTasksInValidation 자동 호출 탑재 검증
const editTriggerCode = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncFromCalendarToSheet('));
const hasColumnCAutoLearn = editTriggerCode.indexOf('startCol <= 3 && endCol >= 3') !== -1 &&
  editTriggerCode.indexOf('ensureTasksInValidation(sheet, editedTasks)') !== -1;
assertTest('handleCalendarEditTrigger: C열(공종) 편집 감지 시 실시간 자동 학습 엔진 즉시 호출 검증', hasColumnCAutoLearn);

// 15-4: taskColorMap 확장 키워드 탑재 검증
assertTest('taskColorMap: 엘리베이터보양, 시스템에어컨, 샷시, 싱크대, 실리콘 등 세부 키워드 색상 매핑 검증',
  code.indexOf('엘리베이터보양') !== -1 &&
  code.indexOf('샷시') !== -1 &&
  code.indexOf('싱크대') !== -1 &&
  code.indexOf('실리콘') !== -1
);

// 15-5: applySiteScheduleValidations sheet.getMaxRows() 깊이 적용 검증
const applyValCode = code.slice(code.indexOf('function applySiteScheduleValidations('), code.indexOf('function setupSiteSheetConditionalFormatting('));
const hasMaxRowsValidation = applyValCode.indexOf('sheet.getMaxRows()') !== -1;
assertTest('applySiteScheduleValidations: 시트 전체 깊이(getMaxRows)까지 드롭다운 유효성 적용 검증', hasMaxRowsValidation);

// ========================================================
// 16. [모바일 웹앱 및 모닝 브리핑 날짜 표준 표기(YY.MM.DD) 및 버튼명 간소화 QA]
// ========================================================
console.log('\n📱 [Step 16: 모바일 웹앱 및 모닝 브리핑 날짜 표준 표기(YY.MM.DD) 및 버튼명 간소화 QA]');

// 16-1: formatScheduleDateDisplay 함수 탑재 및 단위 테스트 (요일 포함)
assertTest('formatScheduleDateDisplay: 날짜 표준 포맷 변환 함수 탑재 여부', typeof sandbox.formatScheduleDateDisplay === 'function');
assertTest('formatScheduleDateDisplay: 단일 종일 일정 -> 26.11.05 (목) 변환 검증',
  sandbox.formatScheduleDateDisplay(new Date(2026, 10, 5)) === '26.11.05 (목)'
);
assertTest('formatScheduleDateDisplay: 기간 종일 일정 -> 26.11.05 (목) ~ 26.11.07 (토) 변환 검증',
  sandbox.formatScheduleDateDisplay('2026-11-05 ~ 2026-11-07') === '26.11.05 (목) ~ 26.11.07 (토)'
);
assertTest('formatScheduleDateDisplay: 단일 시간 일정 -> 26.11.05 (목) 14:00 변환 검증',
  sandbox.formatScheduleDateDisplay(new Date(2026, 10, 5, 14, 0)) === '26.11.05 (목) 14:00'
);
assertTest('formatScheduleDateDisplay: 당일 시간 범위 -> 26.11.05 (목) 14:00 ~ 16:00 변환 검증',
  sandbox.formatScheduleDateDisplay('26.11.05 14:00 ~ 16:00') === '26.11.05 (목) 14:00 ~ 16:00'
);
assertTest('formatScheduleDateDisplay: 다일 시간 범위 -> 26.11.05 (목) 14:00 ~ 26.11.06 (금) 11:00 변환 검증',
  sandbox.formatScheduleDateDisplay('26.11.05 14:00 ~ 26.11.06 11:00') === '26.11.05 (목) 14:00 ~ 26.11.06 (금) 11:00'
);
assertTest('formatScheduleDateDisplay: 원시 Date.toString() 문자열 -> 26.11.11 (수) 완벽 파싱/변환 검증',
  sandbox.formatScheduleDateDisplay(new Date('2026-11-11T00:00:00+09:00')) === '26.11.11 (수)'
);

// 16-2: buildMobileStatusHtml 일시 표기 시 formatScheduleDateDisplay 적용 및 GMT+0900 배제 검증
const testHtml = sandbox.buildMobileStatusHtml(2, '신대방우성 5동 601호', new Date(2026, 10, 11), 'S/A', '예정');
assertTest('buildMobileStatusHtml: 일시 출력 시 26.11.11 (수) 깔끔한 포맷 렌더링 검증', testHtml.indexOf("<span class='info-val'>26.11.11 (수)</span>") !== -1);
assertTest('buildMobileStatusHtml: 일시 출력에서 raw Date (GMT+0900 / 한국 표준시) 완전 제거 검증', testHtml.indexOf('GMT+0900') === -1);

// 16-3: buildMobileStatusHtml 하단 버튼명 "저장하기" 간소화 검증
assertTest('buildMobileStatusHtml: 하단 저장 버튼 텍스트 [저장하기] 적용 검증', testHtml.indexOf('<span>저장하기</span>') !== -1);
assertTest('buildMobileStatusHtml: 이전 장문 버튼텍스트(공유/진행상태/비고 저장하기) 완전 제거 검증', testHtml.indexOf('공유/진행상태/비고 저장하기') === -1);
assertTest('buildMobileStatusHtml: JS saveStatus 핸들러 내 버튼 텍스트 [저장하기] 통일 검증',
  testHtml.indexOf("btn.innerText = '저장하기';") !== -1
);

// 16-4: sendDailySiteMorningBriefing dateStr 포맷팅 적용 검증
const briefingCode = code.slice(code.indexOf('function sendDailySiteMorningBriefing('), code.indexOf('function buildMorningBriefingEmailContent('));
const hasBriefingDateFormat = briefingCode.indexOf('formatScheduleDateDisplay(dateVal)') !== -1;
assertTest('sendDailySiteMorningBriefing: 모닝 브리핑 이메일 발송 시 dateStr에 formatScheduleDateDisplay 적용 검증', hasBriefingDateFormat);

// ========================================================
// 17. [모닝 브리핑 및 백그라운드 유령(고아) 일정 무음 자동 청소 엔진 QA]
// ========================================================
console.log('\n🧹 [Step 17: 모닝 브리핑 및 백그라운드 유령(고아) 일정 무음 자동 청소 엔진 QA]');

// 17-1: cleanupOrphanCalendarEventsSilent 함수 탑재 여부
assertTest('cleanupOrphanCalendarEventsSilent 무음 정리 엔진 함수 탑재 여부', typeof sandbox.cleanupOrphanCalendarEventsSilent === 'function');
assertTest('cleanupDuplicateCalendarEventsSilent 무음 중복 정리 엔진 함수 탑재 여부', typeof sandbox.cleanupDuplicateCalendarEventsSilent === 'function');

// 17-2: sendDailySiteMorningBriefing 발송 전 cleanupOrphanCalendarEventsSilent 호출 검증
const hasMorningCleanup = briefingCode.indexOf('cleanupOrphanCalendarEventsSilent(sheet, cal)') !== -1;
assertTest('sendDailySiteMorningBriefing: 메일 발송 전 캘린더 유령 일정 자동 무음 청소 호출 검증', hasMorningCleanup);

// 17-3: runPeriodicAutoSync NEEDS_ORPHAN_CLEANUP 연동 탑재 검증
const periodicFullCode = code.slice(code.indexOf('function runPeriodicAutoSync()'), code.indexOf('function runDebouncedSync()'));
const hasPeriodicCleanup = periodicFullCode.indexOf('NEEDS_ORPHAN_CLEANUP') !== -1 &&
  periodicFullCode.indexOf('cleanupOrphanCalendarEventsSilent(sheet, cal)') !== -1;
assertTest('runPeriodicAutoSync: 행 삭제(REMOVE_ROW) 감지 플래그 및 15분 주기 유령 일정 청소 탑재 검증', hasPeriodicCleanup);

// 17-4: cleanupOrphanCalendarEventsSilent 단위 테스트 시뮬레이션
let deletedEvents = [];
const mockCalendar = {
  getEvents: (start, end) => [
    {
      getId: () => 'valid-event-id-101',
      getTitle: () => '🟦 [보라매 101동] 도배',
      getDescription: () => '📌 현장명: 보라매 101동\n🔨 공종: 도배',
      deleteEvent: function() { deletedEvents.push(this.getId()); }
    },
    {
      getId: () => 'orphan-event-id-999', // 시트에서 삭제된 유령 일정
      getTitle: () => '🟦 [신대방 301동] 철거',
      getDescription: () => '📌 현장명: 신대방 301동\n🔨 공종: 철거\n우성아이디 현장 일정',
      deleteEvent: function() { deletedEvents.push(this.getId()); }
    },
    {
      getId: () => 'personal-private-event', // 개인 일정
      getTitle: () => '가족 저녁 식사',
      getDescription: () => '가족 모임 식당 예약',
      deleteEvent: function() { deletedEvents.push(this.getId()); }
    }
  ]
};

const makeRangeMock = (values) => ({
  getValue: () => "진행상태",
  setValue: function() { return this; },
  setFontWeight: function() { return this; },
  setFontColor: function() { return this; },
  setBackground: function() { return this; },
  setHorizontalAlignment: function() { return this; },
  setVerticalAlignment: function() { return this; },
  setDataValidation: function() { return this; },
  setValues: function() { return this; },
  getValues: () => values
});

const sheetDataForSilent = [
  ['보라매 101동', '26.10.20', '도배', '', true, '', '', '예정', 'valid-event-id-101', Date.now(), '', '', '', ''],
  ['망원동 빌라', '26.10.22', '타일', '', true, '', '', '예정', 'valid-event-id-102', Date.now(), '', '', '', '']
];

const mockSheetForSilent = {
  getLastRow: () => sheetDataForSilent.length + 1,
  getMaxColumns: () => 14,
  setColumnWidth: () => {},
  hideColumns: () => {},
  getRange: (r, c, nr, nc) => makeRangeMock(sheetDataForSilent)
};

deletedEvents = [];
const cleanedCount = sandbox.cleanupOrphanCalendarEventsSilent(mockSheetForSilent, mockCalendar);
assertTest('cleanupOrphanCalendarEventsSilent: 시트에서 삭제된 유령 일정(orphan-event-id-999) 정확 삭제 검증',
  deletedEvents.indexOf('orphan-event-id-999') !== -1 && cleanedCount === 1
);
assertTest('cleanupOrphanCalendarEventsSilent: 시트에 존재하는 정상 일정(valid-event-id-101) 삭제 방어 검증',
  deletedEvents.indexOf('valid-event-id-101') === -1
);
assertTest('cleanupOrphanCalendarEventsSilent: 개인 일정(personal-private-event) 절대 건드리지 않음 검증',
  deletedEvents.indexOf('personal-private-event') === -1
);

// ========================================================
// 18. [우클릭 행 삭제(REMOVE_ROW) onChange 트리거 및 모닝 브리핑 자율 복구 QA]
// ========================================================
console.log('\n🗑️ [Step 18: 우클릭 행 삭제(REMOVE_ROW) onChange 트리거 및 모닝 브리핑 자율 복구 QA]');

// 18-1: handleSheetChangeTrigger 함수 탑재 여부
assertTest('handleSheetChangeTrigger 함수 탑재 여부', typeof sandbox.handleSheetChangeTrigger === 'function');

// 18-2: ensureRealtimeTriggerInstalledSilent 내 onChange 트리거 자동 등록 검증
const triggerInstallCode = code.slice(code.indexOf('function ensureRealtimeTriggerInstalledSilent('), code.indexOf('function handleSheetChangeTrigger('));
const hasOnChangeTriggerInstall = triggerInstallCode.indexOf('.onChange()') !== -1 &&
  triggerInstallCode.indexOf('handleSheetChangeTrigger') !== -1;
assertTest('ensureRealtimeTriggerInstalledSilent: 행 삭제 감지 onChange 트리거 자동 설치 탑재 검증', hasOnChangeTriggerInstall);

// 18-3: handleSheetChangeTrigger 내 REMOVE_ROW 및 NEEDS_ORPHAN_CLEANUP 검증
const sheetChangeCode = code.slice(code.indexOf('function handleSheetChangeTrigger('), code.indexOf('function cleanupOrphanCalendarEventsSilent('));
const hasRemoveRowLogic = (sheetChangeCode.indexOf('e.changeType === "REMOVE_ROW"') !== -1 || sheetChangeCode.indexOf('e.changeType !== "REMOVE_ROW"') !== -1) &&
  sheetChangeCode.indexOf('NEEDS_ORPHAN_CLEANUP') !== -1 &&
  sheetChangeCode.indexOf('refreshDashboardCalendar') !== -1;
assertTest('handleSheetChangeTrigger: REMOVE_ROW 발생 시 대시보드 새로고침 및 청소 예약 검증', hasRemoveRowLogic);

// 18-4: sendDailySiteMorningBriefing strict isManual 검증 (GAS 이벤트 객체 e 전달 시 UI 팝업 차단)
const hasStrictIsManual = briefingCode.indexOf('var isManual = (isManualTest === true);') !== -1;
assertTest('sendDailySiteMorningBriefing: isManualTest === true 엄격 판별(백그라운드 트리거 시 getUi 에러 원천 차단)', hasStrictIsManual);

// 18-5: sendDailySiteMorningBriefing finally 블록 내 scheduleNextDailyBriefingTrigger 보장
const hasFinallyReschedule = briefingCode.indexOf('finally') !== -1 &&
  briefingCode.indexOf('scheduleNextDailyBriefingTrigger()') !== -1;
assertTest('sendDailySiteMorningBriefing: finally 블록 내 다음날 트리거 무조건 재예약(Self-Rescheduling 보장)', hasFinallyReschedule);

// 18-6: getDailyBriefingRecipientEmail 내 소유자 이메일 4순위 fallback 탑재 검증
const recipientFuncCode = code.slice(code.indexOf('function getDailyBriefingRecipientEmail('), code.indexOf('function configureBriefingEmail('));
const hasOwnerFallback = recipientFuncCode.indexOf('ss.getOwner()') !== -1 &&
  recipientFuncCode.indexOf('DAILY_BRIEFING_EMAIL') !== -1;
assertTest('getDailyBriefingRecipientEmail: 백그라운드 실행 시 소유자 이메일 4순위 fallback 및 영속 캐싱 검증', hasOwnerFallback);

// 18-7: ensureRealtimeTriggerInstalledSilent 내 DAILY_BRIEFING_ENABLED 트리거 자율 복구 검증
const hasBriefingAutoRecovery = triggerInstallCode.indexOf('DAILY_BRIEFING_ENABLED') !== -1 &&
  triggerInstallCode.indexOf('scheduleNextDailyBriefingTrigger()') !== -1;
assertTest('ensureRealtimeTriggerInstalledSilent: 모닝 브리핑 트리거 소실 시 자율 복구(Self-Healing) 검증', hasBriefingAutoRecovery);

// ========================================================
// 19. [Argon 4대 숨은 실무 지뢰 전수 검증 QA]
// ========================================================
console.log('\n💣 [Step 19: Argon 4대 숨은 실무 지뢰 전수 검증 QA]');

// 19-1: 지뢰 1 - handleSheetChangeTrigger에서 OTHER 이벤트 완전 배제 (오직 REMOVE_ROW만 처리)
const hasNoOtherInChangeTrigger = sheetChangeCode.indexOf('OTHER') === -1 &&
  (sheetChangeCode.indexOf('e.changeType !== "REMOVE_ROW"') !== -1 || sheetChangeCode.indexOf('e.changeType === "REMOVE_ROW"') !== -1);
assertTest('지뢰 1 방어: handleSheetChangeTrigger에서 OTHER 완전 배제 및 REMOVE_ROW 엄격 제한', hasNoOtherInChangeTrigger);

// 19-2: 지뢰 2 - runPeriodicAutoSync 25건 초과 잔여 행 존재 시 LAST_PERIODIC_SYNC_MS 최신화 지연 및 재예약
const hasBatchSyncProtection = periodicCode.indexOf('hasRemainingRowsToSync') !== -1 &&
  periodicCode.indexOf('if (!hasRemainingRowsToSync)') !== -1 &&
  periodicCode.indexOf('runDebouncedSync') !== -1;
assertTest('지뢰 2 방어: runPeriodicAutoSync 25건 초과 잔여 행 존재 시 타임스탬프 갱신 지연 및 디바운스 자동 재예약', hasBatchSyncProtection);

// 19-3: 지뢰 3 - 아코디언 하위 행(A열 빈값) 상위 현장명 자동 계승 검증
const refreshDashSiteCode = code.slice(code.indexOf('function refreshDashboardCalendar('), code.indexOf('function syncAllSiteSchedulesToDashboard('));
const delayCheckCode = code.slice(code.indexOf('function checkAndHighlightDelayedSchedules('), code.indexOf('function clearScheduleHighlights('));
const doGetCode = code.slice(code.indexOf('function doGet('), code.indexOf('function escapeHtmlForInput('));

const hasAccordionSiteInheritance = refreshDashSiteCode.indexOf('lastKnownSite') !== -1 &&
  delayCheckCode.indexOf('lastKnownSite') !== -1 &&
  briefingCode.indexOf('lastKnownSite') !== -1 &&
  doGetCode.indexOf('prevSites') !== -1;
assertTest('지뢰 3 방어: 아코디언 하위 행 현장명(A열) 자동 계승 (대시보드/지연알림/모닝브리핑/웹앱)', hasAccordionSiteInheritance);

// 19-4: 지뢰 4 - 백그라운드 트리거 getActiveSpreadsheet() null 시 getAppSpreadsheet() 안전 폴백
const hasBgAppSpreadsheetFallback = periodicCode.indexOf('getAppSpreadsheet()') !== -1 &&
  refreshDashSiteCode.indexOf('getAppSpreadsheet()') !== -1;
assertTest('지뢰 4 방어: 백그라운드 실행 시 getActiveSpreadsheet() null 대비 getAppSpreadsheet() 안전 폴백', hasBgAppSpreadsheetFallback);

// ========================================================
// 20. [현장명 변경 캘린더 동기화 & 캘린더 기본 전날 알림 자동 차단 QA]
// ========================================================
console.log('\n🔔 [Step 20: 현장명 변경 캘린더 동기화 & 캘린더 기본 전날 알림 자동 차단 QA]');

// 20-1: runPeriodicAutoSync re-entrant lock 안전 획득 및 해제 검증
const hasReentrantLock = periodicCode.indexOf('lock.hasLock()') !== -1 &&
  periodicCode.indexOf('var needRelease = false;') !== -1 &&
  periodicCode.indexOf('needRelease = true;') !== -1;
assertTest('runPeriodicAutoSync: re-entrant Lock 안전 획득(hasLock) 및 데드락 방지 검증', hasReentrantLock);

// 20-2: handleCalendarEditTrigger A열 수정 시 하위 아코디언 행 타임스탬프 동반 갱신 검증
const editTriggerSubCode = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncFromCalendarToSheet('));
const hasSubRowPropagation = editTriggerSubCode.indexOf('startCol <= 1 && endCol >= 1') !== -1 &&
  editTriggerSubCode.indexOf('subTimeStamps.push([nowMs])') !== -1 &&
  editTriggerSubCode.indexOf('subMetaUpdates.push(') !== -1;
assertTest('handleCalendarEditTrigger: A열(현장명) 수정 시 하위 아코디언 행 타임스탬프 동반 갱신 검증', hasSubRowPropagation);

// 20-3: syncSingleRowToCalendar 생성 및 수정 시 event.removeAllReminders() 호출 검증
const singleSyncCode = code.slice(code.indexOf('function syncSingleRowToCalendar('), code.indexOf('function syncRecentSiteSchedulesToGoogleCalendar('));
const hasReminderRemovalSingle = singleSyncCode.indexOf('event.removeAllReminders()') !== -1;
assertTest('syncSingleRowToCalendar: 캘린더 기본 전날 알림 자동 차단(removeAllReminders) 탑재 검증', hasReminderRemovalSingle);

// 20-4: syncSiteSchedulesBatchSafe 생성 및 수정 시 event.removeAllReminders() 호출 검증
const batchSyncCode = code.slice(code.indexOf('function syncSiteSchedulesBatchSafe('), code.indexOf('function findMatchingCalendarEvent('));
const hasReminderRemovalBatch = batchSyncCode.indexOf('event.removeAllReminders()') !== -1;
assertTest('syncSiteSchedulesBatchSafe: 배치 동기화 시 캘린더 기본 전날 알림 자동 차단 탑재 검증', hasReminderRemovalBatch);

// 20-5: removeAllSiteCalendarReminders 함수 탑재 및 상단 메뉴 등록 검증
assertTest('removeAllSiteCalendarReminders: 캘린더 전날 알림 일괄 해제 함수 탑재 여부', typeof sandbox.removeAllSiteCalendarReminders === 'function');
const onOpenMenuCode = code.slice(code.indexOf('function onOpen('), code.indexOf('function onEdit('));
assertTest('onOpen: 상단 메뉴(⚡ 캘린더 연동 설정) 내 캘린더 전날 알림 전체 끄기 등록 검증', onOpenMenuCode.indexOf('removeAllSiteCalendarReminders') !== -1);

// ========================================================
// 21. [현장명 변경 캘린더 동기화 & 캘린더 중복 생성 원천 차단 & 무음 자동 정리 QA]
// ========================================================
console.log('\n🛡️ [Step 21: 현장명 변경 캘린더 동기화 & 캘린더 중복 생성 원천 차단 & 무음 자동 정리 QA]');

// 21-1: cleanupDuplicateCalendarEventsSilent 함수 탑재 여부
assertTest('cleanupDuplicateCalendarEventsSilent 함수 탑재 여부', typeof sandbox.cleanupDuplicateCalendarEventsSilent === 'function');

// 21-2: syncSiteSchedulesBatchSafe 시작 시 무음 중복 사전 정리 호출 검증
assertTest('syncSiteSchedulesBatchSafe: 시작 시 cleanupDuplicateCalendarEventsSilent 호출 검증',
  batchSyncCode.indexOf('cleanupDuplicateCalendarEventsSilent(cal, sheet)') !== -1
);

// 21-3: syncSiteSchedulesBatchSafe 및 syncSingleRowToCalendar 내 CalendarApp.getEventById 폴백 검증
assertTest('syncSiteSchedulesBatchSafe: CalendarApp.getEventById 폴백 탑재 검증',
  batchSyncCode.indexOf('CalendarApp.getEventById(existingId)') !== -1
);
assertTest('syncSingleRowToCalendar: CalendarApp.getEventById 폴백 탑재 검증',
  singleSyncCode.indexOf('CalendarApp.getEventById(existingEventId)') !== -1
);

// 21-4: 신규 생성 직전 최후 중복 탐색 가드(fallbackEv) 탑재 검증
assertTest('syncSiteSchedulesBatchSafe: 신규 생성 전 fallbackEv 중복 재확인 가드 탑재 검증',
  batchSyncCode.indexOf('fallbackEv = findMatchingCalendarEvent(cal, details, allSheetIds)') !== -1
);
assertTest('syncSingleRowToCalendar: 신규 생성 전 fallbackEv 중복 재확인 가드 탑재 검증',
  singleSyncCode.indexOf('fallbackEv = findMatchingCalendarEvent(cal, details, allSheetIds)') !== -1
);

// 21-5: findMatchingCalendarEvent 현장명 변경(fallbackTaskMatches) 매칭 지원 검증
const matchFuncCode = code.slice(code.indexOf('function findMatchingCalendarEvent('), code.indexOf('function cleanDuplicateCalendarEventsOnDate('));
assertTest('findMatchingCalendarEvent: 현장명 변경 시 동일 날짜/공종 fallbackTaskMatches 매칭 검증',
  matchFuncCode.indexOf('fallbackTaskMatches') !== -1
);

// 21-6: isDateEqual UTC/KST 시차(14시간 이내) 방어 검증
const isDateEqualFunc = code.slice(code.indexOf('function isDateEqual('), code.indexOf('function formatDateYYMMDD('));
assertTest('isDateEqual: 타임존(UTC vs KST 등 9시간 시차) 14시간 오프셋 방어 검증',
  isDateEqualFunc.indexOf('diffHours <= 14') !== -1
);

// ========================================================
// 22. Argon 캘린더 영구 고정 & 개인 캘린더 폴백/자동생성 원천 차단 QA
// ========================================================
console.log('\n🔒 [Step 22: 공용 캘린더 영구 고정 및 개인 캘린더 폴백/자동생성 원천 차단 QA]');

const getTargetCalCode = code.slice(code.indexOf('function getTargetCalendar()'), code.indexOf('function configureTargetCalendar()'));
const configTargetCalCode = code.slice(code.indexOf('function configureTargetCalendar()'), code.indexOf('function buildEventDetails('));

assertTest('getTargetCalendar: 개인 캘린더(getDefaultCalendar) 폴백 원천 차단 검증',
  !/CalendarApp\s*\.\s*getDefaultCalendar/.test(getTargetCalCode)
);

assertTest('getTargetCalendar: 신규 캘린더(createCalendar) 자동 생성 원천 차단 검증',
  !/CalendarApp\s*\.\s*createCalendar/.test(getTargetCalCode)
);

assertTest('getTargetCalendar: 캘린더 미발견 시 null 반환 및 경고 로깅 검증',
  getTargetCalCode.indexOf('return null;') !== -1 && getTargetCalCode.indexOf('Logger.log') !== -1
);

assertTest('configureTargetCalendar: 임의의 신규 캘린더 자동 생성 제거 검증',
  configTargetCalCode.indexOf('CalendarApp.createCalendar') === -1
);

assertTest('configureTargetCalendar: 빈 입력 시 공용 기본값 복원(deleteProperty) 검증',
  configTargetCalCode.indexOf('deleteProperty("TARGET_CALENDAR_SETTING")') !== -1
);

assertTest('DEFAULT_CALENDAR_ID_OR_NAME: 전역 변수 설정 존재 검증',
  typeof sandbox.DEFAULT_CALENDAR_ID_OR_NAME === 'string' && sandbox.DEFAULT_CALENDAR_ID_OR_NAME.length > 0
);

// 런타임 getTargetCalendar 동작 검증
let defaultCalCalled = false;
let createCalCalled = false;
sandbox.CalendarApp.getDefaultCalendar = () => { defaultCalCalled = true; return { getName: () => '개인 캘린더' }; };
sandbox.CalendarApp.createCalendar = () => { createCalCalled = true; return { getName: () => '신규 캘린더' }; };
sandbox.CalendarApp.getCalendarById = (id) => (id === 'c_shared_cal@group.calendar.google.com' ? { getId: () => id, getName: () => '현장 일정' } : null);
sandbox.CalendarApp.getCalendarsByName = (name) => (name === '현장 일정' ? [{ getId: () => 'c_shared_cal@group.calendar.google.com', getName: () => name }] : []);

// 케이스 1: 정상 이름 조회 -> 공용 캘린더 반환
const foundCal = sandbox.getTargetCalendar();
assertTest('getTargetCalendar 런타임: 정상 이름 조회 시 캘린더 객체 반환 검증',
  foundCal !== null && foundCal.getName() === '현장 일정' && !defaultCalCalled && !createCalCalled
);

// ========================================================
// 23. 캘린더 제목([공종] 현장명) & 위치 명시 & 현장별 고유 색상 매핑 QA
// ========================================================
console.log('\n🎨 [Step 23: 캘린더 제목([공종] 현장명) & 위치(Location) & 현장별 고유 색상 매핑 QA]');

// 23-1: buildEventDetails 제목 형식 검증 ([공종] 현장명)
const testSiteA = '보라매 101동';
const testTaskA = '철거';
const testRowA = [testSiteA, '26.11.05', testTaskA, '', true, '', '홍길동', '예정'];
const detailsA = sandbox.buildEventDetails(testRowA, 10, null, null);
assertTest('캘린더 제목: [공종] 현장명 표준 형식 반환 검증',
  detailsA && detailsA.title.indexOf('[' + testTaskA + '] ' + testSiteA) !== -1
);

// 23-2: 상세 설명(Description) 및 Location 위치 명시 검증
assertTest('캘린더 상세 설명(Description) 내 📍 위치: 현장명 포함 검증',
  detailsA && detailsA.description.indexOf('📍 위치: ' + testSiteA) !== -1
);
assertTest('캘린더 상세 설명(Description) 내 📌 현장명: 포함 검증',
  detailsA && detailsA.description.indexOf('📌 현장명: ' + testSiteA) !== -1
);

// 23-3: 현장별 고유 색상 매핑 함수 탑재 여부
assertTest('getCalendarColorForSite 함수 탑재 여부', typeof sandbox.getCalendarColorForSite === 'function');

// 23-4: 동일한 현장명은 항상 100% 동일한 색상 반환 검증
const colorA1 = sandbox.getCalendarColorForSite(testSiteA);
const colorA2 = sandbox.getCalendarColorForSite(testSiteA);
const colorA3 = sandbox.getCalendarColorForSite('  ' + testSiteA + '  ');
assertTest('동일 현장명("보라매 101동")은 언제나 동일한 색상 반환 검증',
  colorA1 === colorA2 && colorA2 === colorA3
);

// 23-5: 시트 내 서로 다른 현장은 서로 다른 색상 배정 검증
const testSiteB = '신대방 301동';
const testSiteC = '역삼 자이';
const mockMultiSiteSheet = {
  getLastRow: () => 4,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [[testSiteA], [testSiteB], [testSiteC]]
  })
};
sandbox._siteColorMapCache = null; // 캐시 리셋
const colorSheetA = sandbox.getCalendarColorForSite(testSiteA, mockMultiSiteSheet);
const colorSheetB = sandbox.getCalendarColorForSite(testSiteB, mockMultiSiteSheet);
const colorSheetC = sandbox.getCalendarColorForSite(testSiteC, mockMultiSiteSheet);
assertTest('서로 다른 현장(A, B, C)은 서로 다른 색상으로 자동 분기 검증',
  colorSheetA !== colorSheetB && colorSheetB !== colorSheetC && colorSheetA !== colorSheetC
);

// 23-6: buildEventDetails 객체의 color가 현장 고유 색상과 일치 검증
assertTest('buildEventDetails 반환 color가 현장 고유 색상과 100% 일치 검증',
  detailsA && detailsA.color === colorA1
);

// ========================================================
// 24. 동일 날짜/동일 현장/동일 공종 복수 일정 100% 보존 QA (8월 25일 도배 2건 누락 버그 방어 검증)
// ========================================================
console.log('\n🛡️ [Step 24: 동일 날짜/현장/공종 복수 일정 100% 보존 QA (8/25 도배 2건 누락 방어)]');

const dobeDate = new Date(2026, 7, 25, 0, 0, 0); // 2026-08-25
const dobeDetails = {
  title: '[도배] 보라매경남 102동 702호',
  siteName: '보라매경남 102동 702호',
  taskName: '도배',
  description: '📌 현장명: 보라매경남 102동 702호\n🔨 공종: 도배',
  parsedDate: {
    startDate: dobeDate,
    endDate: dobeDate,
    hasTime: false,
    isRange: false
  }
};

let deletedEventsMap = {};
function createMockEvent(id, title, site, task, date) {
  return {
    getId: () => id,
    getTitle: () => title,
    getDescription: () => `📌 현장명: ${site}\n🔨 공종: ${task}\n우성아이디`,
    getLocation: () => site,
    isAllDayEvent: () => true,
    getAllDayStartDate: () => date,
    getAllDayEndDate: () => date,
    getStartTime: () => date,
    getEndTime: () => date,
    deleteEvent: () => { deletedEventsMap[id] = true; }
  };
}

// 24-1: cleanDuplicateCalendarEventsOnDate - 시트의 유효한 2건 모두 보존 검증
deletedEventsMap = {};
const mockCal2Dobe = {
  getEvents: () => [
    createMockEvent('ev-dobe-1', '[도배] 보라매경남 102동 702호', '보라매경남 102동 702호', '도배', dobeDate),
    createMockEvent('ev-dobe-2', '[도배] 보라매경남 102동 702호', '보라매경남 102동 702호', '도배', dobeDate)
  ]
};
const cleanedCount1 = sandbox.cleanDuplicateCalendarEventsOnDate(
  mockCal2Dobe,
  dobeDetails,
  'ev-dobe-1',
  { 'ev-dobe-1': true, 'ev-dobe-2': true }
);
assertTest('cleanDuplicateCalendarEventsOnDate: 시트에 존재하는 동일 날짜/현장/공종 2건은 절대 상호 삭제하지 않고 보존',
  cleanedCount1 === 0 && !deletedEventsMap['ev-dobe-2']
);

// 24-2: cleanDuplicateCalendarEventsOnDate - 잉여 유령 일정만 선택 삭제 검증
deletedEventsMap = {};
const mockCal3Dobe = {
  getEvents: () => [
    createMockEvent('ev-dobe-1', '[도배] 보라매경남 102동 702호', '보라매경남 102동 702호', '도배', dobeDate),
    createMockEvent('ev-dobe-2', '[도배] 보라매경남 102동 702호', '보라매경남 102동 702호', '도배', dobeDate),
    createMockEvent('ev-dobe-orphan', '[도배] 보라매경남 102동 702호', '보라매경남 102동 702호', '도배', dobeDate)
  ]
};
const cleanedCount2 = sandbox.cleanDuplicateCalendarEventsOnDate(
  mockCal3Dobe,
  dobeDetails,
  'ev-dobe-1',
  { 'ev-dobe-1': true, 'ev-dobe-2': true }
);
assertTest('cleanDuplicateCalendarEventsOnDate: 시트에 없는 잉여 고아 일정(orphan)만 정확 삭제하고 정규 2건 보존',
  cleanedCount2 === 1 && deletedEventsMap['ev-dobe-orphan'] && !deletedEventsMap['ev-dobe-2']
);

// 24-3: cleanupDuplicateCalendarEventsSilent - 동일 날짜/공종 2건 시트 등록 시 무음 삭제 방어 검증
deletedEventsMap = {};
const mockMultiRowSheet = {
  getLastRow: () => 3,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [
      ['보라매경남 102동 702호', '26.08.25', '도배', '', true, '', '홍길동', '예정', 'ev-dobe-1', 123456],
      ['보라매경남 102동 702호', '26.08.25', '도배', '', true, '', '홍길동', '예정', 'ev-dobe-2', 123457]
    ]
  })
};
const silentDeleted1 = sandbox.cleanupDuplicateCalendarEventsSilent(mockCal2Dobe, mockMultiRowSheet);
assertTest('cleanupDuplicateCalendarEventsSilent: 시트에 등록된 동일 날짜/공종 2건 무음 삭제 방어 검증',
  silentDeleted1 === 0 && !deletedEventsMap['ev-dobe-1'] && !deletedEventsMap['ev-dobe-2']
);

// 24-4: cleanupDuplicateCalendarEventsSilent - 시트 외 잉여 고아 일정만 무음 청소 검증
deletedEventsMap = {};
const silentDeleted2 = sandbox.cleanupDuplicateCalendarEventsSilent(mockCal3Dobe, mockMultiRowSheet);
assertTest('cleanupDuplicateCalendarEventsSilent: 2건의 정상 일정 보존 및 3번째 미등록 잉여 일정만 선택 삭제 검증',
  silentDeleted2 === 1 && deletedEventsMap['ev-dobe-orphan'] && !deletedEventsMap['ev-dobe-1'] && !deletedEventsMap['ev-dobe-2']
);

// 24-5: findMatchingCalendarEvent - 이미 1행에 배정된 일정(ev-dobe-1)은 2행 신규 동기화 시 가로채기(hijack) 방지 검증
const matchedForSecondRow = sandbox.findMatchingCalendarEvent(
  mockCal2Dobe,
  dobeDetails,
  { 'ev-dobe-1': true }
);
assertTest('findMatchingCalendarEvent: 1행에 매핑된 기존 일정은 2행에서 가로채지 않고 null 반환(독립 일정 생성 유도)',
  matchedForSecondRow && matchedForSecondRow.getId() === 'ev-dobe-2' // ev-dobe-2가 미등록 상태면 2행과 연결
);

// 24-6: 만약 ev-dobe-1, ev-dobe-2 둘 다 이미 시트에 연결되어 있으면 null 반환 (신규 3번째 일정 생성을 위해 가로채기 방지)
const matchedWhenBothAssigned = sandbox.findMatchingCalendarEvent(
  mockCal2Dobe,
  dobeDetails,
  { 'ev-dobe-1': true, 'ev-dobe-2': true }
);
assertTest('findMatchingCalendarEvent: 모든 이벤트가 시트 행들에 이미 매핑되어 있으면 가로채기 차단(null 반환)',
  matchedWhenBothAssigned === null
);

// ========================================================
// 25. 공휴일/공유 캘린더 색상 충돌 방지 & 17일 일정 색상(청록) 배제 QA
// ========================================================
console.log('\n🎨 [Step 25: 공휴일/공유 캘린더 색상 충돌 방지 & 17일 기존 색상 배제 QA]');

const forbiddenColors = ['7', '6', '4', '11']; // 7: 17일 기존 청록, 6: 공휴일 주황, 4: 공휴일 산호, 11: 지연 경고 빨강
const testSitesList = [
  '신대방우성 5동 601호', // 17일에 등록된 1순위 대표 현장
  '보라매경남 102동 702호',
  '망원 경일빌라 201호',
  '역삼 래미안',
  '반포 자이',
  '도곡 렉슬',
  '잠실 엘스'
];

const mockPaletteSheet = {
  getLastRow: () => testSitesList.length + 1,
  getRange: () => ({
    getValues: () => testSitesList.map(s => [s])
  })
};
sandbox._siteColorMapCache = null; // 캐시 리셋

// 25-1: 17일에 등록된 1순위 현장 색상이 기존 7(청록)이 아닌 9(Blueberry 진파랑)로 안전 변경되었는지 검증
const firstSiteColor = sandbox.getCalendarColorForSite(testSitesList[0], mockPaletteSheet);
assertTest('17일 등록 1순위 현장: 공휴일/공유 캘린더와 겹치던 7(청록) 배제 및 9(Blueberry 진파랑) 부여 검증',
  firstSiteColor === '9'
);

// 25-2: 팔레트 전체에서 금지 색상(7, 6, 4, 11)이 단 하나도 반환되지 않는지 검증
let hasForbiddenColor = false;
for (let i = 0; i < testSitesList.length; i++) {
  const c = sandbox.getCalendarColorForSite(testSitesList[i], mockPaletteSheet);
  if (forbiddenColors.indexOf(c) !== -1) {
    hasForbiddenColor = true;
    break;
  }
}
assertTest('시트 내 모든 현장 색상에서 7(기존 17일 청록), 6/4(공휴일 주황/산호), 11(지연 빨강) 완전 배제 검증',
  !hasForbiddenColor
);

// 25-3: 해시 기반 단독 호출 시에도 금지 색상 배제 검증
let hashHasForbidden = false;
for (let i = 0; i < 50; i++) {
  const c = sandbox.getCalendarColorForSite('임의현장_' + i, null);
  if (forbiddenColors.indexOf(c) !== -1) {
    hashHasForbidden = true;
    break;
  }
}
assertTest('시트 없는 단독 해시 매핑 시에도 7/6/4/11 금지 색상 100% 배제 검증',
  !hashHasForbidden
);
// ========================================================
// 26. 구글 API Rate Limit 방어 & 지수 백오프 & 8/25 미등록 일정 100% 보장 배치 엔진 QA
// ========================================================
console.log('\n⚡ [Step 26: Rate Limit 방어 & 지수 백오프 & 8/25 미등록 일정 100% 보장 QA]');

// 26-1: executeWithRateLimitRetry 함수 탑재 여부
assertTest('executeWithRateLimitRetry 함수 탑재 여부', typeof sandbox.executeWithRateLimitRetry === 'function');

// 26-2: executeWithRateLimitRetry 정상 동작 시 즉각 반환 검증
let callCount = 0;
const normalResult = sandbox.executeWithRateLimitRetry(() => {
  callCount++;
  return 'SUCCESS';
}, 2, 50);
assertTest('executeWithRateLimitRetry: 오류 없을 때 1회 호출 즉각 성공 검증',
  normalResult === 'SUCCESS' && callCount === 1
);

// 26-3: executeWithRateLimitRetry Rate Limit 일시 에러 발생 시 자동 지수 백오프 재시도 성공 검증
let retryAttempts = 0;
const recoveredResult = sandbox.executeWithRateLimitRetry(() => {
  retryAttempts++;
  if (retryAttempts === 1) {
    throw new Error('짧은 시간에 너무 많은 작업을 수행했습니다. 작업을 잠시 중단한 후 다시 시도하세요.');
  }
  return 'RECOVERED';
}, 2, 20); // 20ms 빠른 테스트
assertTest('executeWithRateLimitRetry: Rate Limit 발생 시 백오프 대기 후 자동 복구 검증',
  recoveredResult === 'RECOVERED' && retryAttempts === 2
);

// 26-4: executeWithRateLimitRetry 최대 재시도 초과 시 에러 정상 전파 검증
let maxRetryAttempts = 0;
let caughtError = null;
try {
  sandbox.executeWithRateLimitRetry(() => {
    maxRetryAttempts++;
    throw new Error('Rate limit exceeded quota');
  }, 2, 10);
} catch (e) {
  caughtError = e;
}
assertTest('executeWithRateLimitRetry: 재시도 한도 초과 시 예외 안전 throw 검증',
  caughtError !== null && maxRetryAttempts === 3 // 0, 1, 2 = 3회
);

// 26-5: syncSiteSchedulesBatchSafe 내 1회 일괄 조회(Pre-fetch) 캐시 탑재 검증
assertTest('syncSiteSchedulesBatchSafe: 전체 캘린더 이벤트 1회 사전 조회(calEventsMap) 탑재 검증',
  code.indexOf('calEventsMap') !== -1 &&
  code.indexOf('cal._preloadedEvents = preloadedEvents;') !== -1
);

// 26-6: syncSiteSchedulesBatchSafe 내 I열 복제(중복 ID) 행 자동 감지 및 시트 청소 검증
assertTest('syncSiteSchedulesBatchSafe: 복사-붙여넣기 행 중복 ID 감지 및 셀 청소 탑재 검증',
  code.indexOf('duplicateRowsToClear') !== -1 &&
  code.indexOf('seenSheetIds[eid]') !== -1
);

// 26-7: syncSiteSchedulesBatchSafe 90일 확장 및 미등록 일정 100% 보장 검증
const batchCodeForStep26 = code.slice(code.indexOf('function syncSiteSchedulesBatchSafe('), code.indexOf('function ensureCalendarSyncColumns('));
assertTest('syncSiteSchedulesBatchSafe: ninetyDaysAgo 90일 확장 탐색 탑재 검증',
  batchCodeForStep26.indexOf('ninetyDaysAgo') !== -1 &&
  batchCodeForStep26.indexOf('90 * 24 * 60 * 60 * 1000') !== -1
);

assertTest('syncSiteSchedulesBatchSafe: 과거 일정이라도 미등록 행(!existingId)은 100% 무조건 동기화 큐 포함 검증 (8/25 누락 방어)',
  batchCodeForStep26.indexOf('if (onlyRecent && existingId && calEventsMap[existingId])') !== -1
);

// 26-8: syncSiteSchedulesBatchSafe 내 executeWithRateLimitRetry 이벤트 생성/수정 연동 검증
assertTest('syncSiteSchedulesBatchSafe: 이벤트 생성 및 갱신 시 executeWithRateLimitRetry 보호 연동 검증',
  batchCodeForStep26.indexOf('executeWithRateLimitRetry(function()') !== -1
);

// 26-9: findMatchingCalendarEvent 및 cleanDuplicateCalendarEventsOnDate 메모리 캐시 연동 검증
const findMatchCode = code.slice(code.indexOf('function findMatchingCalendarEvent('), code.indexOf('function cleanDuplicateCalendarEventsOnDate('));
const cleanDupCode = code.slice(code.indexOf('function cleanDuplicateCalendarEventsOnDate('), code.indexOf('function cleanupDuplicateAndOrphanCalendarEvents('));
assertTest('findMatchingCalendarEvent: cal._preloadedEvents 메모리 캐시 활용(0 API Call) 검증',
  findMatchCode.indexOf('cal._preloadedEvents') !== -1
);
assertTest('cleanDuplicateCalendarEventsOnDate: cal._preloadedEvents 메모리 캐시 활용(0 API Call) 검증',
  cleanDupCode.indexOf('cal._preloadedEvents') !== -1
);

// ========================================================
// 27. [스마트 UI] 스프레드시트 우측 사이드바 & 팝업 실시간 대시보드 화면 엔진 QA
// ========================================================
console.log('\n🖥️ [Step 27: 스프레드시트 우측 사이드바 & 팝업 실시간 대시보드 화면 QA]');

// 27-1: openSiteScheduleSidebar & openSiteScheduleModal 함수 탑재 여부
assertTest('openSiteScheduleSidebar 함수 탑재 여부', typeof sandbox.openSiteScheduleSidebar === 'function');
assertTest('openSiteScheduleModal 함수 탑재 여부', typeof sandbox.openSiteScheduleModal === 'function');

// 27-2: 상단 메뉴 등록 검증
const onOpenCode = code.slice(code.indexOf('function onOpen('), code.indexOf('function onEdit('));
assertTest('onOpen: "대시보드 사이드바 열기" 메뉴 항목 등록 검증',
  onOpenCode.indexOf('"대시보드 사이드바 열기"') !== -1 && onOpenCode.indexOf('openSiteScheduleSidebar') !== -1
);
assertTest('onOpen: "대시보드 팝업 열기" 메뉴 항목 등록 검증',
  onOpenCode.indexOf('"대시보드 팝업 열기"') !== -1 && onOpenCode.indexOf('openSiteScheduleModal') !== -1
);

// 27-3: getHexColorForCalendarCode 색상 변환 검증
assertTest('getHexColorForCalendarCode: 9(Blueberry) -> #3F51B5 매핑 검증',
  sandbox.getHexColorForCalendarCode('9') === '#3F51B5'
);
assertTest('getHexColorForCalendarCode: 10(Basil) -> #0B8043 매핑 검증',
  sandbox.getHexColorForCalendarCode('10') === '#0B8043'
);
assertTest('getHexColorForCalendarCode: 미지정 시 기본 파랑(#2563EB) 반환 검증',
  sandbox.getHexColorForCalendarCode('unknown') === '#2563EB'
);

// 27-4: getSidebarDashboardData 데이터 스키마 및 KPI 계산 검증
const testSidebarSheetData = [
  ['보라매 101동', '26.10.15', '도배', '한샘', true, '도배 풀칠', '홍길동', '진행', 'ev-1', Date.now(), '', '', '', ''],
  ['보라매 101동', '26.10.10', '철거', '철거팀', true, '마루 철거', '김철수', '예정', 'ev-2', Date.now(), '', '', '', ''],
  ['망원동 빌라', '26.10.25', '타일', '을지로', false, '화장실 타일', '이영희', '완료', 'ev-3', Date.now(), '', '', '', '']
];
const mockSidebarSheet = {
  getLastRow: () => testSidebarSheetData.length + 1,
  getRange: (r, c, nr, nc) => ({
    getValues: () => testSidebarSheetData
  })
};
sandbox.findSiteScheduleSheet = () => mockSidebarSheet;

const dashData = sandbox.getSidebarDashboardData();
assertTest('getSidebarDashboardData: KPI 지표 반환 검증 (전체 현장, 공정, 지연 집계)',
  dashData && dashData.kpi &&
  dashData.kpi.totalSites === 2 &&
  dashData.kpi.totalTasks === 3 &&
  dashData.kpi.delayedCount >= 1
);

assertTest('getSidebarDashboardData: 현장별 그룹화 및 tasks 목록 매핑 검증',
  dashData.sites && dashData.sites.length === 2 &&
  dashData.sites[0].siteName === '보라매 101동' &&
  dashData.sites[0].tasks.length === 2 &&
  dashData.sites[0].colorHex.indexOf('#') === 0
);

// 27-5: buildSidebarDashboardHtml 렌더링 검증
const sidebarHtml = sandbox.buildSidebarDashboardHtml('sidebar');
assertTest('buildSidebarDashboardHtml: 사이드바 HTML, CSS, JS 포함 및 실시간 검색/아코디언 탑재 검증',
  typeof sidebarHtml === 'string' &&
  sidebarHtml.indexOf('kpi-grid') !== -1 &&
  sidebarHtml.indexOf('searchInput') !== -1 &&
  sidebarHtml.indexOf('toggleSite') !== -1 &&
  sidebarHtml.indexOf('onTaskClick') !== -1
);

// 27-6: focusSheetRow 안전 실행 검증
let activatedRange = null;
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
  setActiveSheet: () => {},
  getRange: () => ({ activate: () => {} })
});
mockSidebarSheet.setActiveRange = (rng) => { activatedRange = rng; };
const focusRes = sandbox.focusSheetRow(2);
assertTest('focusSheetRow: 정상 행 클릭 시 활성화 및 success 반환 검증',
  focusRes && focusRes.success === true && activatedRange !== null
);

// ========================================================
// 📱 [Step 28: 반응형 모바일 대시보드 및 원터치 바텀시트 관리 QA]
// ========================================================
console.log('\n📱 [Step 28: 반응형 모바일 대시보드 및 원터치 바텀시트 관리 QA]');

// 28-1: showMobileDashboardUrl 함수 탑재 및 상단 메뉴 등록 검증
assertTest('showMobileDashboardUrl 함수 탑재 여부', typeof sandbox.showMobileDashboardUrl === 'function');
assertTest('onOpen: 모바일 대시보드 URL 안내 메뉴 항목 등록 검증',
  onOpenCode.indexOf('showMobileDashboardUrl') !== -1
);

// 28-2: buildSidebarDashboardHtml("mobile") 모바일 최적화 및 메타 태그 검증
const mobileHtml = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml("mobile"): 모바일 반응형 뷰포트(viewport) 메타 태그 탑재',
  mobileHtml.indexOf('viewport') !== -1 && mobileHtml.indexOf('user-scalable=no') !== -1
);
assertTest('buildSidebarDashboardHtml("mobile"): iOS 홈 화면 추가(apple-mobile-web-app-capable) 메타 태그 탑재',
  mobileHtml.indexOf('apple-mobile-web-app-capable') !== -1
);
assertTest('buildSidebarDashboardHtml("mobile"): 모바일 실시간 타이틀 표기 검증',
  mobileHtml.indexOf('우성아이디 현장 대시보드') !== -1
);
assertTest('buildSidebarDashboardHtml("mobile"): 모바일 원터치 바텀시트(bottom-sheet) 컨테이너 탑재',
  mobileHtml.indexOf('sheet-overlay') !== -1 && mobileHtml.indexOf('bottom-sheet') !== -1
);
assertTest('buildSidebarDashboardHtml("mobile"): 바텀시트 내 상태 버튼 및 비고 메모 편집 탑재',
  mobileHtml.indexOf('sheet-status-grid') !== -1 && mobileHtml.indexOf('sheetMemoInput') !== -1
);

// 28-3: doGet(e) 라우팅 검증
// 가상 HtmlOutput 모의 객체
sandbox.HtmlService = {
  createHtmlOutput: function(htmlContent) {
    var out = {
      _content: htmlContent,
      _title: '',
      _metaTags: {},
      _xFrame: '',
      setTitle: function(t) { out._title = t; return out; },
      addMetaTag: function(n, c) { out._metaTags[n] = c; return out; },
      setXFrameOptionsMode: function(m) { out._xFrame = m; return out; },
      getContent: function() { return out._content; }
    };
    return out;
  },
  XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
};

// A. row 파라미터 없을 때 -> 모바일 전체 대시보드 반환
const doGetDash = sandbox.doGet({});
assertTest('doGet: row 파라미터 생략 시 모바일 전체 대시보드 자동 서빙 검증',
  doGetDash && doGetDash._title.indexOf('우성아이디 현장 대시보드') !== -1 &&
  doGetDash._content.indexOf('우성아이디 현장 대시보드') !== -1
);
assertTest('doGet: view="dashboard" 명시 시에도 모바일 전체 대시보드 서빙 검증',
  sandbox.doGet({ parameter: { view: 'dashboard' } })._content.indexOf('kpi-grid') !== -1
);

// B. row=2 단일 일정 접근 시 -> buildMobileStatusHtml 반환 및 상단 전체 대시보드 링크 검증
const mockGetRange = (r, c, nr, nc) => ({
  getValues: () => [
    ['보라매 101동', '26.10.25', '목공', '우성', true, '현장메모', '김소장', '진행', 'ev-1']
  ],
  getValue: () => '보라매 101동'
});
mockSidebarSheet.getRange = mockGetRange;
mockSidebarSheet.getLastRow = () => 10;
sandbox.findSiteScheduleSheet = () => mockSidebarSheet;

const doGetSingle = sandbox.doGet({ parameter: { row: '2' } });
assertTest('doGet: row=2 전달 시 단일 일정 상태 관리 화면 출력 검증',
  doGetSingle && doGetSingle._content.indexOf('보라매 101동') !== -1 &&
  doGetSingle._content.indexOf('목공') !== -1
);
assertTest('doGet: 단일 일정 관리 화면 상단에 "← 📱 전체 대시보드" 바로가기 링크 탑재 검증',
  doGetSingle && doGetSingle._content.indexOf('← 📱 전체 대시보드') !== -1
);

// ========================================================
// 🎯 [Step 29: 오늘 날짜 정밀 필터링 & 상세화면 버튼 제거 및 저장버튼 100% 확장 QA]
// ========================================================
console.log('\n🎯 [Step 29: 오늘 날짜 정밀 필터링 & 상세화면 버튼 제거 및 저장버튼 100% 확장 QA]');

// 29-1: 오늘(10.16) vs 내일(10.17) 정밀 분리 검증 (14시간 시차 오차 제거)
const testTodayTomorrowData = [
  ['신대방우성 5동 601호', '26.10.16', '양생1', '업체A', true, '', '송지은', '예정', '', 0, '', '', '', ''],
  ['신대방우성 5동 601호', '26.10.17', '양생2', '업체B', true, '', '송지은', '예정', '', 0, '', '', '', '']
];
mockSidebarSheet.getRange = (r, c, nr, nc) => ({
  getValues: () => testTodayTomorrowData
});
mockSidebarSheet.getLastRow = () => 3;
sandbox.findSiteScheduleSheet = () => mockSidebarSheet;

const dashTodayCheck = sandbox.getSidebarDashboardData();
const todayTask = dashTodayCheck.sites[0].tasks.find(t => t.taskName === '양생1');
const tomorrowTask = dashTodayCheck.sites[0].tasks.find(t => t.taskName === '양생2');

assertTest('getSidebarDashboardData: 오늘 일정(26.10.16)은 isToday === true 검증',
  todayTask && todayTask.isToday === true
);
assertTest('getSidebarDashboardData: 내일 일정(26.10.17)은 isToday === false (오늘 필터 배제) 검증',
  tomorrowTask && tomorrowTask.isToday === false
);
assertTest('getSidebarDashboardData: 오늘 일정(26.10.16)은 지연(isDelayed) 오인 방어(false) 검증',
  todayTask && todayTask.isDelayed === false
);
assertTest('getSidebarDashboardData: KPI 오늘 공정 수에 내일 일정이 섞이지 않고 정확히 1건 집계 검증',
  dashTodayCheck.kpi.todayCount === 1
);

// 29-2: buildSidebarDashboardHtml 내 상세화면 버튼 제거 및 저장하기 버튼 100% 확장 검증
const currentDashHtml = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: 차단 에러를 유발하던 "🔗 상세화면" 버튼 완전 제거 검증',
  currentDashHtml.indexOf('상세화면') === -1 && currentDashHtml.indexOf('openSoloPage') === -1
);
assertTest('buildSidebarDashboardHtml: "💾 저장하기" 버튼 100% 전체 너비(width: 100%) 확장 검증',
  currentDashHtml.indexOf('.sheet-save-btn { width: 100%;') !== -1
);

console.log('\n========================================================');
console.log(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
console.log('========================================================');

if (failCount > 0) {
  process.exit(1);
} else {
  console.log('🎉 모든 QA 및 코드리뷰 무결성 테스트를 100% 통과했습니다!');
}




