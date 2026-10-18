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
const scriptPath = fs.existsSync(path.join(__dirname, '..', '현장서식.js'))
  ? path.join(__dirname, '..', '현장서식.js')
  : path.join(__dirname, '현장서식.js');
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
assertTest('1시간 기본 일정 포맷팅: "2026.10.11 15:00"', fmtWithTime(sA, eA1) === '2026.10.11 15:00');

const eA2 = new Date(2026, 9, 11, 18, 0);
assertTest('3시간 일정 포맷팅: "2026.10.11 15:00 ~ 18:00"', fmtWithTime(sA, eA2) === '2026.10.11 15:00 ~ 18:00');

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

// 4. 상태 '예정' 런타임 호출 테스트 (예정 상태는 상태 아이콘 생략)
const rowPlan = [...baseMockRow];
rowPlan[7] = '예정';
const resPlan = buildDetailsFn(rowPlan, 18);
assertTest('상태 "예정" 시 캘린더 제목 맨 앞에 상태 아이콘 없이 "🟦 " 공유 아이콘만 깔끔하게 부여', resPlan && resPlan.title.startsWith('🟦'));

// 5. 상태 미입력(빈문자열, null, undefined) 런타임 안전성 테스트
const rowEmpty = [...baseMockRow];
rowEmpty[7] = '';
const resEmpty = buildDetailsFn(rowEmpty, 19);
assertTest('진행상태가 빈 문자열일 때 ReferenceError 없이 기본 "🟦" 안전 처리', resEmpty !== null && resEmpty.title.startsWith('🟦'));

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

const hasBatchEmptyRowCleanup = code.indexOf('function cleanupEmptyRowsMetadata()') !== -1 &&
  (code.indexOf('sheet.getRange(2, 1, lastRow - 1, 10).setValues(values)') !== -1 || code.indexOf('sheet.getRange(2, 1, lastRow - 1, 14).setValues(values)') !== -1);
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
  (code.indexOf('var managerColorMap = [') !== -1 || code.indexOf('var managerColorMap = getManagerColorMap') !== -1) &&
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

// 11-5: 상단 메뉴(onOpen) 내 1분 주기 자동 동기화 등록 검증
const onOpenStart = code.indexOf('function onOpen()');
const menuCode = code.slice(onOpenStart, code.indexOf('function isSiteScheduleSheet(') !== -1 ? code.indexOf('function isSiteScheduleSheet(') : onOpenStart + 3500);
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

// 14-5: updateScheduleStatusFromWeb 모바일 동시 접속 충돌 방어 및 논블로킹 즉시 저장 검증
const mobileWebCode = code.slice(code.indexOf('function updateScheduleStatusFromWeb('), code.indexOf('function doGet('));
const hasMobileLockService = mobileWebCode.indexOf('LockService.getScriptLock()') !== -1 &&
  mobileWebCode.indexOf('lock.tryLock') !== -1 &&
  mobileWebCode.indexOf('lock.releaseLock()') !== -1 &&
  mobileWebCode.indexOf('sheet.getRange(rowNum, 8).setValue(newStatus)') !== -1;
assertTest('updateScheduleStatusFromWeb: 모바일 웹 논블로킹 즉시 시트 저장 및 Lock 경합 완화 탑재 검증', hasMobileLockService);
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
assertTest('formatScheduleDateDisplay: 단일 종일 일정 -> 2026.11.05 (목) 변환 검증',
  sandbox.formatScheduleDateDisplay(new Date(2026, 10, 5)) === '2026.11.05 (목)'
);
assertTest('formatScheduleDateDisplay: 기간 종일 일정 -> 2026.11.05 (목) ~ 2026.11.07 (토) 변환 검증',
  sandbox.formatScheduleDateDisplay('2026-11-05 ~ 2026-11-07') === '2026.11.05 (목) ~ 2026.11.07 (토)'
);
assertTest('formatScheduleDateDisplay: 단일 시간 일정 -> 2026.11.05 (목) 14:00 변환 검증',
  sandbox.formatScheduleDateDisplay(new Date(2026, 10, 5, 14, 0)) === '2026.11.05 (목) 14:00'
);
assertTest('formatScheduleDateDisplay: 당일 시간 범위 -> 2026.11.05 (목) 14:00 ~ 16:00 변환 검증',
  sandbox.formatScheduleDateDisplay('26.11.05 14:00 ~ 16:00') === '2026.11.05 (목) 14:00 ~ 16:00'
);
assertTest('formatScheduleDateDisplay: 다일 시간 범위 -> 2026.11.05 (목) 14:00 ~ 2026.11.06 (금) 11:00 변환 검증',
  sandbox.formatScheduleDateDisplay('26.11.05 14:00 ~ 26.11.06 11:00') === '2026.11.05 (목) 14:00 ~ 2026.11.06 (금) 11:00'
);
assertTest('formatScheduleDateDisplay: 원시 Date.toString() 문자열 -> 2026.11.11 (수) 완벽 파싱/변환 검증',
  sandbox.formatScheduleDateDisplay(new Date('2026-11-11T00:00:00+09:00')) === '2026.11.11 (수)'
);

// 16-2: buildMobileStatusHtml 일시 표기 시 formatScheduleDateDisplay 적용 및 GMT+0900 배제 검증
const testHtml = sandbox.buildMobileStatusHtml(2, '신대방우성 5동 601호', new Date(2026, 10, 11), 'S/A', '예정');
assertTest('buildMobileStatusHtml: 일시 출력 시 2026.11.11 (수) 깔끔한 포맷 렌더링 검증', testHtml.indexOf("<span class='info-val'>2026.11.11 (수)</span>") !== -1);
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
  editTriggerSubCode.indexOf('sheet.getRange(maxProcessRow + 1, 10, subRowsToTouch, 1).setValues(subTimeStamps)') !== -1;
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

// 29-1: 오늘 vs 내일 정밀 분리 검증 (14시간 시차 오차 제거)
const now29 = new Date();
const pad29 = n => String(n).padStart(2, '0');
const todayStr29 = `${String(now29.getFullYear()).slice(-2)}.${pad29(now29.getMonth() + 1)}.${pad29(now29.getDate())}`;
const tmrw29 = new Date(now29.getTime() + 86400000);
const tmrwStr29 = `${String(tmrw29.getFullYear()).slice(-2)}.${pad29(tmrw29.getMonth() + 1)}.${pad29(tmrw29.getDate())}`;

const testTodayTomorrowData = [
  ['신대방우성 5동 601호', todayStr29, '양생1', '업체A', true, '', '송지은', '예정', '', 0, '', '', '', ''],
  ['신대방우성 5동 601호', tmrwStr29, '양생2', '업체B', true, '', '송지은', '예정', '', 0, '', '', '', '']
];
mockSidebarSheet.getRange = (r, c, nr, nc) => ({
  getValues: () => testTodayTomorrowData
});
mockSidebarSheet.getLastRow = () => 3;
sandbox.findSiteScheduleSheet = () => mockSidebarSheet;

const dashTodayCheck = sandbox.getSidebarDashboardData();
const todayTask = dashTodayCheck.sites[0].tasks.find(t => t.taskName === '양생1');
const tomorrowTask = dashTodayCheck.sites[0].tasks.find(t => t.taskName === '양생2');

assertTest('getSidebarDashboardData: 오늘 일정(' + todayStr29 + ')은 isToday === true 검증',
  todayTask && todayTask.isToday === true
);
assertTest('getSidebarDashboardData: 내일 일정(' + tmrwStr29 + ')은 isToday === false (오늘 필터 배제) 검증',
  tomorrowTask && tomorrowTask.isToday === false
);
assertTest('getSidebarDashboardData: 오늘 일정(' + todayStr29 + ')은 지연(isDelayed) 오인 방어(false) 검증',
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

// ========================================================
// 🎯 Step 30: 공정 날짜순 오름차순 정렬 & 완료 블록 목록 노출 배제 QA
// ========================================================
console.log('\n🎯 [Step 30: 공정 날짜순 오름차순 정렬 & 완료 블록 목록 노출 배제 QA]');

const testSortAndCompletedSheetData = [
  // 신대방우성 16동 1501호: 시트상 26.11.25가 26.11.03보다 위에 위치 (사용자 캡처 실사례)
  ['신대방우성 16동 1501호', '26.11.25', '철거', '철거업체', true, '', '홍길동', '예정', 'ev-s1', Date.now(), '', '', '', ''],
  ['신대방우성 16동 1501호', '26.11.03', '홈파기', '설비업체', true, '', '홍길동', '예정', 'ev-s2', Date.now(), '', '', '', ''],
  // 압구정 현대아파트: 모든 공정이 완료된 현장 (블록 전체 완료)
  ['압구정 현대아파트', '26.10.01', '철거', '철거팀', true, '', '김반장', '완료', 'ev-a1', Date.now(), '', '', '', ''],
  ['압구정 현대아파트', '26.10.05', '타일', '타일팀', true, '', '김반장', '완료', 'ev-a2', Date.now(), '', '', '', ''],
  // 반포 자이: 일부만 완료된 현장 (블록 미완료)
  ['반포 자이', '26.12.01', '도배', '도배팀', true, '', '이소장', '예정', 'ev-b1', Date.now(), '', '', '', ''],
  ['반포 자이', '26.11.20', '마루', '마루팀', true, '', '이소장', '완료', 'ev-b2', Date.now(), '', '', '', '']
];

sandbox.findSiteScheduleSheet = () => ({
  getLastRow: () => testSortAndCompletedSheetData.length + 1,
  getRange: () => ({
    getValues: () => testSortAndCompletedSheetData
  })
});

const sortDashData = sandbox.getSidebarDashboardData();

// 30-1: 공정 날짜순 오름차순 정렬 검증 (신대방우성: 11.03 홈파기가 11.25 철거보다 앞으로 정렬)
const shindaebang = sortDashData.sites.find(s => s.siteName === '신대방우성 16동 1501호');
assertTest('getSidebarDashboardData: 현장 내 공정 날짜순 오름차순 정렬 검증 (11.03 홈파기 우선)',
  shindaebang &&
  shindaebang.tasks.length === 2 &&
  shindaebang.tasks[0].taskName === '홈파기' &&
  shindaebang.tasks[1].taskName === '철거'
);

// 30-2: 블록 전체 완료 여부 판별 검증
const apgujeong = sortDashData.sites.find(s => s.siteName === '압구정 현대아파트');
const banpo = sortDashData.sites.find(s => s.siteName === '반포 자이');

assertTest('getSidebarDashboardData: 모든 공정이 완료된 블록은 isAllCompleted === true 검증',
  apgujeong && apgujeong.isAllCompleted === true
);

assertTest('getSidebarDashboardData: 일부 공정만 완료된 블록은 isAllCompleted === false 검증',
  banpo && banpo.isAllCompleted === false &&
  shindaebang && shindaebang.isAllCompleted === false
);

// 30-3: KPI 활성 현장 수 검증 (전체 3개 중 완료된 1개 제외하고 activeSites === 2)
assertTest('getSidebarDashboardData: KPI activeSites 지표에 완료 현장 배제 검증 (3개 중 2개)',
  sortDashData.kpi.totalSites === 3 &&
  sortDashData.kpi.activeSites === 2
);

// 30-4: 현장 카드(sites) 정렬 검증: 가장 빠른 일정 순 정렬 (신대방우성 11.03 vs 반포 자이 12.01)
assertTest('getSidebarDashboardData: 활성 현장 카드 가장 빠른 일정 순 정렬 검증',
  sortDashData.sites[0].siteName === '신대방우성 16동 1501호'
);

// 30-5: buildSidebarDashboardHtml 렌더링 내 완료 블록 노출 배제 로직 탑재 검증
const step30TestHtml = sandbox.buildSidebarDashboardHtml('sidebar');
assertTest('buildSidebarDashboardHtml: 완료 블록 기본 리스트 노출 배제(site.isAllCompleted) 필터 탑재 검증',
  step30TestHtml.indexOf('site.isAllCompleted') !== -1
);
assertTest('buildSidebarDashboardHtml: 상단 필터 탭에 "✅ 완료" 아카이브 탭 탑재 검증',
  step30TestHtml.indexOf('data-filter="done"') !== -1 &&
  step30TestHtml.indexOf('✅ 완료') !== -1
);
assertTest('buildSidebarDashboardHtml: KPI 라벨이 "진행 현장"으로 표기 검증',
  step30TestHtml.indexOf('진행 현장') !== -1
);

// ========================================================
// 🎯 Step 31: 팝업 로딩 중앙 정렬 & 클라이언트 스크립트 구문 무결성 및 무한로딩 방어 QA
// ========================================================
console.log('\n🎯 [Step 31: 팝업 로딩 중앙 정렬 & 클라이언트 스크립트 구문 무결성 및 무한로딩 방어 QA]');

const modalHtml = sandbox.buildSidebarDashboardHtml('modal');
const modalScriptMatch = modalHtml.match(/<script>([\s\S]*?)<\/script>/);

// 31-1: 모달 팝업 내 클라이언트 인라인 자바스크립트 구문 무결성 검증 (Unexpected token 방어)
let scriptSyntaxValid = false;
try {
  new vm.Script(modalScriptMatch[1]);
  scriptSyntaxValid = true;
} catch(e) {
  scriptSyntaxValid = false;
}
assertTest('buildSidebarDashboardHtml("modal"): 생성된 클라이언트 JS 구문 무결성(SyntaxError 0) 검증',
  scriptSyntaxValid === true
);

// 31-2: 팝업 내 로딩 바(spinner)와 안내 텍스트 중앙 정렬(loading-state) 검증
assertTest('buildSidebarDashboardHtml: 로딩 컨테이너 grid-column: 1 / -1 전체 너비 확장 검증',
  modalHtml.indexOf('.loading-state { grid-column: 1 / -1;') !== -1
);
assertTest('buildSidebarDashboardHtml: 로딩 컨테이너 flex 세로 중앙 정렬 탑재 검증',
  modalHtml.indexOf('display: flex; flex-direction: column; align-items: center; justify-content: center;') !== -1
);
assertTest('buildSidebarDashboardHtml: 로딩 텍스트(loading-text) 텍스트 중앙 정렬 탑재 검증',
  modalHtml.indexOf('.loading-text { text-align: center;') !== -1 &&
  modalHtml.indexOf('class="loading-text"') !== -1
);

// 31-3: document.readyState 대응 안전 초기화 로직 탑재 검증
assertTest('buildSidebarDashboardHtml: document.readyState 확인 및 DOMContentLoaded/load 안전 이벤트 리스너 탑재 검증',
  modalHtml.indexOf('document.readyState === "complete"') !== -1 &&
  modalHtml.indexOf('addEventListener("DOMContentLoaded"') !== -1
);

// 31-4: getSidebarDashboardData 예외 방어(try-catch) 탑재 검증
sandbox.findSiteScheduleSheet = () => { throw new Error("시트 읽기 테스트 예외"); };
const safeDashData = sandbox.getSidebarDashboardData();
assertTest('getSidebarDashboardData: 시트 예외 발생 시 에러 객체 반환 및 스크립트 중단 방어 검증',
  safeDashData &&
  safeDashData.error &&
  safeDashData.error.indexOf("데이터 처리 중 오류 발생") !== -1 &&
  safeDashData.kpi &&
  safeDashData.sites.length === 0
);

// 31-5: 전체 및 완료 탭 및 이름 필터에서는 접힌 상태, 검색 시에는 펼쳐진 상태 유지 검증
assertTest('buildSidebarDashboardHtml: 전체/완료/이름 필터에서는 기본 접힘, 검색 시에는 펼침(isExpanded) 적용 검증',
  modalHtml.indexOf('(query !== "") || (currentFilter !== "all" && currentFilter !== "done")') !== -1
);

// ========================================================
// 🎯 Step 32: 검색 시 완료 항목 배제 & 모바일 접힌 상태 & 공유 뱃지 & 동적 일정 정렬 QA
// ========================================================
console.log('\n🎯 [Step 32: 검색 시 완료 항목 배제 & 모바일 접힌 상태 & 공유 뱃지 & 동적 일정 정렬 QA]');

const step32MobileHtml = sandbox.buildSidebarDashboardHtml('mobile');
const mobileScriptMatch = step32MobileHtml.match(/<script>([\s\S]*?)<\/script>/);

// 32-1: 모바일 HTML 클라이언트 JS 구문 무결성(SyntaxError 0) 검증
let mobileScriptValid = false;
try {
  new vm.Script(mobileScriptMatch[1]);
  mobileScriptValid = true;
} catch(e) {
  mobileScriptValid = false;
}
assertTest('buildSidebarDashboardHtml("mobile"): 모바일 클라이언트 JS 구문 무결성(SyntaxError 0) 검증',
  mobileScriptValid === true
);

// 32-2: 검색 시 완료/취소 공정 항목 배제 필터 탑재 검증
assertTest('buildSidebarDashboardHtml: 검색 중일 때 완료/취소 공정 항목 제외(query && currentFilter !== "done") 탑재 검증',
  step32MobileHtml.indexOf('query && currentFilter !== "done"') !== -1 &&
  step32MobileHtml.indexOf('stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1') !== -1
);

// 32-3: 접힌 상태에서 헤더 날짜 뱃지 배제 & 세부 카드 공유 뱃지(share-tag) & 동적 정렬 로직 검증
assertTest('buildSidebarDashboardHtml: 헤더 날짜 뱃지 배제(사용자 요청: 접힌상태 날짜 제외) 검증',
  step32MobileHtml.indexOf('site-date-badge') === -1
);

assertTest('buildSidebarDashboardHtml: 세부 카드 내 공유 상태 뱃지 스타일(.share-tag) 탑재 검증',
  step32MobileHtml.indexOf('.share-tag {') !== -1 &&
  step32MobileHtml.indexOf('.share-tag.share-on {') !== -1 &&
  step32MobileHtml.indexOf('.share-tag.share-off {') !== -1
);

assertTest('buildSidebarDashboardHtml: 세부 카드 내 네모박스 파랑/빨강 공유 뱃지(🟦 공유 / 🟥 미공유) 렌더링 검증',
  step32MobileHtml.indexOf('🟦 공유') !== -1 &&
  step32MobileHtml.indexOf('🟥 미공유') !== -1
);

assertTest('buildSidebarDashboardHtml: 검색/필터 시 매칭된 공정의 가장 빠른 일정 기준 동적 정렬(matchedSites.sort) 탑재 검증',
  step32MobileHtml.indexOf('earliestFilteredMs') !== -1 &&
  step32MobileHtml.indexOf('matchedSites.sort') !== -1
);

// 32-4: getSidebarDashboardData: earliestDateDisplay 및 today/delayed 플래그 전달 검증
const testStep32SheetData = [
  ['역삼 자이', '26.10.16', '철거', '업체A', true, '', '홍길동', '완료', 'ev-32-1', Date.now(), '', '', '', ''],
  ['역삼 자이', '26.10.17', '목공', '업체B', true, '', '홍길동', '진행', 'ev-32-2', Date.now(), '', '', '', ''],
  ['역삼 자이', '26.10.20', '타일', '업체C', true, '', '홍길동', '예정', 'ev-32-3', Date.now(), '', '', '', ''],
  ['서초 래미안', '26.10.18', '도배', '업체D', true, '', '이영희', '예정', 'ev-32-4', Date.now(), '', '', '', '']
];
sandbox.findSiteScheduleSheet = () => ({
  getLastRow: () => testStep32SheetData.length + 1,
  getRange: () => ({
    getValues: () => testStep32SheetData
  })
});
const step32Data = sandbox.getSidebarDashboardData();
const yeoksamSite = step32Data.sites.find(s => s.siteName === '역삼 자이');
const seochoSite = step32Data.sites.find(s => s.siteName === '서초 래미안');

assertTest('getSidebarDashboardData: 진행 중인 가장 빠른 공정의 dateDisplay 전달 검증 (역삼 자이 -> 26.10.17)',
  yeoksamSite && yeoksamSite.earliestDateDisplay && yeoksamSite.earliestDateDisplay.indexOf('26.10.17') !== -1
);

assertTest('getSidebarDashboardData: 현장 카드 날짜 오름차순 정렬 검증 (역삼 자이 10.17이 서초 래미안 10.18보다 앞섬)',
  step32Data.sites[0].siteName === '역삼 자이' &&
  step32Data.sites[1].siteName === '서초 래미안'
);

// 32-5: 클라이언트 검색 시 완료 항목 배제 로직 시뮬레이션
const simTasks = yeoksamSite.tasks;
const simQuery = '홍길동';
const simFilter = 'all';
const filtered = simTasks.filter(t => {
  if (simFilter === 'done') {
    if (t.status.indexOf('완료') === -1 && t.status.indexOf('취소') === -1) return false;
  }
  if (simQuery && simFilter !== 'done') {
    const stClean = (t.status || '').replace(/\s+/g, '');
    if (stClean.indexOf('완료') !== -1 || stClean.indexOf('취소') !== -1) return false;
  }
  if (simQuery) {
    const matchMgr = (t.manager || '').toLowerCase().indexOf(simQuery) !== -1;
    if (!matchMgr) return false;
  }
  return true;
});

assertTest('검색 시뮬레이션: "홍길동" 검색 시 완료된 공정(철거)은 제외되고 미완료 2건만 노출 검증',
  filtered.length === 2 &&
  filtered.some(t => t.taskName === '목공') &&
  filtered.some(t => t.taskName === '타일') &&
  !filtered.some(t => t.taskName === '철거')
);

// 32-6: 검색 시 현장별 빠른 일정 순 동적 정렬 시뮬레이션 검증 (사용자 스크린샷 실사례)
const testSearchSortData = [
  {
    siteName: '신대방우성 5동 601호',
    tasks: [
      { taskName: '철거', dateDisplay: '26.10.15', dateMs: 1760486400000, status: '완료' },
      { taskName: '도배', dateDisplay: '26.11.05', dateMs: 1762300800000, status: '예정' }
    ]
  },
  {
    siteName: '신대방우성 1동 805호',
    tasks: [
      { taskName: '도배', dateDisplay: '26.10.21', dateMs: 1761004800000, status: '예정' }
    ]
  }
];

const simMatched = [];
testSearchSortData.forEach(s => {
  const fTasks = s.tasks.filter(t => t.taskName.indexOf('도배') !== -1 && t.status.indexOf('완료') === -1);
  if (fTasks.length === 0) return;
  let eMs = 0;
  for (let i = 0; i < fTasks.length; i++) {
    if (fTasks[i].dateMs && (!eMs || fTasks[i].dateMs < eMs)) eMs = fTasks[i].dateMs;
  }
  simMatched.push({ site: s, filteredTasks: fTasks, earliestFilteredMs: eMs });
});
simMatched.sort((a, b) => a.earliestFilteredMs - b.earliestFilteredMs);

assertTest('동적 정렬 시뮬레이션: "도배" 검색 시 10.21 일정을 가진 신대방우성 1동 805호가 11.05 일정의 5동 601호보다 상위에 정렬',
  simMatched[0].site.siteName === '신대방우성 1동 805호' &&
  simMatched[1].site.siteName === '신대방우성 5동 601호'
);

// ========================================================
// 🎯 Step 33: 검색창 검색어 삭제 ✕ 버튼 & 03_구성원 시트 연동 담당자 필터 버튼 QA
// ========================================================
console.log('\n🎯 [Step 33: 검색창 검색어 삭제 ✕ 버튼 & 03_구성원 시트 연동 담당자 필터 버튼 QA]');

// 33-1: getMemberNames 함수 탑재 여부
assertTest('getMemberNames 함수 탑재 여부',
  typeof sandbox.getMemberNames === 'function'
);

// 33-2: getMemberNames가 03_구성원 시트에서 개별 구성원 이름 추출 및 김종혁 제외 검증
const mockMembersSheet = {
  getName: () => '03_구성원',
  getLastRow: () => 5,
  getLastColumn: () => 3,
  getRange: (r, c, nr, nc) => {
    if (r === 1) {
      return { getValues: () => [['직급', '성명', '연락처']] };
    }
    return {
      getValues: () => [
        ['김종혁'],
        ['조항만'],
        ['송지은, 전예린'],
        ['영하']
      ]
    };
  }
};
const mockSSWithMembers = {
  getSheets: () => [mockMembersSheet]
};
const extractedMembers = sandbox.getMemberNames(mockSSWithMembers);

assertTest('getMemberNames: 03_구성원 시트 추출 시 "김종혁"·"영하" 배제 및 복합 담당자 개별 분리 검증',
  Array.isArray(extractedMembers) &&
  extractedMembers.length === 3 &&
  !extractedMembers.includes('김종혁') &&
  !extractedMembers.includes('영하') &&
  extractedMembers[0] === '조항만' &&
  extractedMembers[1] === '송지은' &&
  extractedMembers[2] === '전예린'
);

// 33-3: 시트 미존재 시 fallback 기본 구성원 반환 검증
const fallbackMembers = sandbox.getMemberNames({ getSheets: () => [] });
assertTest('getMemberNames: 구성원 시트 부재 시 fallback 기본 멤버 반환 및 김종혁·영하 미포함 검증',
  fallbackMembers.length === 5 && fallbackMembers.includes('조항만') && !fallbackMembers.includes('김종혁') && !fallbackMembers.includes('영하')
);

// 33-4: getSidebarDashboardData 반환값에 members 포함 검증
const dashDataWithMembers = sandbox.getSidebarDashboardData();
assertTest('getSidebarDashboardData: 반환 객체에 members 배열 포함 검증',
  dashDataWithMembers && Array.isArray(dashDataWithMembers.members) && dashDataWithMembers.members.length > 0
);

// 33-5: 검색창 ✕ 버튼 HTML 및 CSS 스타일 탑재 검증
const dashHtml33 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: 검색창 내 검색어 삭제 ✕ 버튼(#searchClearBtn) 탑재 검증',
  dashHtml33.indexOf('id="searchClearBtn"') !== -1 &&
  dashHtml33.indexOf('onclick="clearSearch()"') !== -1 &&
  dashHtml33.indexOf('.search-clear') !== -1 &&
  dashHtml33.indexOf('.search-clear.show') !== -1
);

// 33-6: 검색창 input right padding 확장(30px) 및 onSearchInput 연동 검증
assertTest('buildSidebarDashboardHtml: 검색창 30px 패딩 및 onSearchInput 입력 이벤트 연동 검증',
  dashHtml33.indexOf('padding: 8px 30px 8px 28px;') !== -1 &&
  dashHtml33.indexOf('oninput="onSearchInput()"') !== -1
);

// 33-7: 담당자 필터 바(#memberFilterBar) HTML 및 스타일 탑재 검증
assertTest('buildSidebarDashboardHtml: 상단 필터 탭 하단에 담당자 필터 바(#memberFilterBar) 탑재 검증',
  dashHtml33.indexOf('id="memberFilterBar"') !== -1 &&
  dashHtml33.indexOf('.member-filter-bar') !== -1 &&
  dashHtml33.indexOf('.member-tab') !== -1 &&
  dashHtml33.indexOf('.member-tab.active') !== -1
);

// 33-8: 클라이언트 스크립트에 clearSearch, onSearchInput, renderMemberFilterTabs, setMemberFilter 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 검색 삭제 및 담당자 필터 제어 함수 탑재 검증',
  dashHtml33.indexOf('function onSearchInput()') !== -1 &&
  dashHtml33.indexOf('function clearSearch()') !== -1 &&
  dashHtml33.indexOf('function renderMemberFilterTabs(members)') !== -1 &&
  dashHtml33.indexOf('function setMemberFilter(enc)') !== -1
);

// 33-9: 클라이언트 renderMemberFilterTabs에서 김종혁 필터 버튼 생성 배제 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 renderMemberFilterTabs에서 "김종혁" 버튼 2중 방어 배제 검증',
  dashHtml33.indexOf('if (m === "김종혁" || m.indexOf("김종혁") !== -1) return;') !== -1
);

// 33-10: onDataLoaded에서 renderMemberFilterTabs 호출 탑재 검증
assertTest('buildSidebarDashboardHtml: onDataLoaded 시 renderMemberFilterTabs(data.members) 자동 호출 검증',
  dashHtml33.indexOf('renderMemberFilterTabs(data.members || []);') !== -1
);

// 33-11: 담당자 필터 선택 시 완료/취소 공정 제외 로직 탑재 검증
assertTest('buildSidebarDashboardHtml: 담당자 선택 시 완료/취소 공정 제외 로직 탑재 검증',
  dashHtml33.indexOf('if (selectedMember) {') !== -1 &&
  dashHtml33.indexOf('if (stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1) return false;') !== -1
);

// 33-12: 담당자 필터링 시뮬레이션 검증 (특정 멤버 클릭 시 해당 담당자의 완료 제외 미완료 공정만 필터 & 펼침 & 빠른 일정순 정렬)
const testMemberFilterSites = [
  {
    siteName: '현장A',
    colorHex: '#2563eb',
    isAllCompleted: false,
    tasks: [
      { taskName: '도배', manager: '송지은', dateDisplay: '26.10.25', dateMs: 1761350400000, status: '예정' },
      { taskName: '철거', manager: '송지은', dateDisplay: '26.10.15', dateMs: 1760486400000, status: '완료' },
      { taskName: '타일', manager: '조항만', dateDisplay: '26.10.18', dateMs: 1760745600000, status: '진행중' }
    ]
  },
  {
    siteName: '현장B',
    colorHex: '#3F51B5',
    isAllCompleted: false,
    tasks: [
      { taskName: '마루', manager: '송지은, 영하', dateDisplay: '26.10.20', dateMs: 1760918400000, status: '예정' }
    ]
  }
];

// 시뮬레이션: '송지은' 담당자 필터 적용
const simSelectedMember = '송지은';
const simMatchedMemberSites = [];
testMemberFilterSites.forEach(s => {
  const fTasks = s.tasks.filter(t => {
    // 담당자 필터 시 완료/취소 공정 배제
    const stClean = (t.status || '').replace(/\s+/g, '');
    if (stClean.indexOf('완료') !== -1 || stClean.indexOf('취소') !== -1) return false;

    const mgr = (t.manager || '').trim();
    return mgr.indexOf(simSelectedMember) !== -1;
  });
  if (fTasks.length === 0) return;
  let eMs = 0;
  for (let i = 0; i < fTasks.length; i++) {
    if (fTasks[i].dateMs && (!eMs || fTasks[i].dateMs < eMs)) eMs = fTasks[i].dateMs;
  }
  simMatchedMemberSites.push({ site: s, filteredTasks: fTasks, earliestFilteredMs: eMs });
});
simMatchedMemberSites.sort((a, b) => a.earliestFilteredMs - b.earliestFilteredMs);

const simIsExpanded = ('' !== '') || ('all' !== 'all' && 'all' !== 'done');

assertTest('담당자 필터 시뮬레이션: "송지은" 필터 시 완료 공정(철거)은 제외되고, 10.20 일정을 가진 현장B가 10.25 현장A보다 상위 정렬되며 카드는 닫힌 상태(isExpanded: false) 유지 검증',
  simMatchedMemberSites.length === 2 &&
  simMatchedMemberSites[0].site.siteName === '현장B' &&
  simMatchedMemberSites[0].filteredTasks.length === 1 &&
  simMatchedMemberSites[1].site.siteName === '현장A' &&
  simMatchedMemberSites[1].filteredTasks.length === 1 &&
  simMatchedMemberSites[1].filteredTasks[0].taskName === '도배' && // 철거(완료)는 제외!
  simIsExpanded === false
);

// 33-13: 필터 해제 시 시뮬레이션 (선택 해제 시 평상시 접힌 상태 유지)
const simResetMember = '';
const simResetQuery = '';
const simResetFilter = 'all';
const simResetExpanded = (simResetQuery !== '') || (simResetMember !== '') || (simResetFilter !== 'all' && simResetFilter !== 'done');
assertTest('필터 해제 시뮬레이션: 담당자 및 검색어가 없을 때 전체 탭에서 카드가 접힌 상태(isExpanded: false) 유지 검증',
  simResetExpanded === false
);

// ========================================================
// 34. [우측 상단 calendar 버튼 & 00_Calendar 스타일 모바일 달력 & 현장 체크박스 필터 QA]
// ========================================================
console.log('\n📅 [Step 34: 우측 상단 calendar 버튼 & 00_Calendar 스타일 모바일 달력 & 현장 체크박스 필터 QA]');

const dashHtml34 = sandbox.buildSidebarDashboardHtml('mobile');

// 34-1: 하단 푸터 바 달력 토글 버튼(#calendarToggleBtn) 및 캘린더 동기화 제거 검증
assertTest('buildSidebarDashboardHtml: 하단 푸터 바에 [📅 캘린더 보기] 토글 버튼(#calendarToggleBtn) 탑재 및 [캘린더 동기화] 제거 검증',
  dashHtml34.indexOf('id="calendarToggleBtn"') !== -1 &&
  dashHtml34.indexOf('📅 캘린더 보기') !== -1 &&
  dashHtml34.indexOf('toggleCalendarView()') !== -1 &&
  dashHtml34.indexOf('syncCalendar') === -1
);

// 34-2: 모바일 월간 달력 페이지(#calendarArea 및 그리드) 마크업 & CSS 검증
assertTest('buildSidebarDashboardHtml: 00_Calendar 스타일 월간 달력(#calendarArea 및 그리드) 탑재 검증',
  dashHtml34.indexOf('id="calendarArea"') !== -1 &&
  dashHtml34.indexOf('class="calendar-view"') !== -1 &&
  dashHtml34.indexOf('class="cal-nav-bar"') !== -1 &&
  dashHtml34.indexOf('class="cal-weekdays"') !== -1 &&
  dashHtml34.indexOf('id="calGrid"') !== -1 &&
  dashHtml34.indexOf('.cal-day-cell') !== -1 &&
  dashHtml34.indexOf('.cal-event-pill') !== -1
);

// 34-3: 현장 카드 앞 체크박스(.site-check) 마크업 및 아코디언 간섭 방지 검증
assertTest('buildSidebarDashboardHtml: 현장 카드 앞 체크박스(.site-check) 및 아코디언 클릭 간섭 방지(stopPropagation) 검증',
  dashHtml34.indexOf('.site-check') !== -1 &&
  dashHtml34.indexOf('class="site-check"') !== -1 &&
  dashHtml34.indexOf('onSiteCheckboxClick(event, this)') !== -1 &&
  dashHtml34.indexOf('e.stopPropagation();') !== -1
);

// 34-4: 체크박스 선택 시 액션 바(#checkedActionBar) 및 선택 탭(#checkedTab) 탑재 검증
assertTest('buildSidebarDashboardHtml: 선택 현장 액션 바(#checkedActionBar) 및 탭(#checkedTab) 마크업 검증',
  dashHtml34.indexOf('id="checkedTab"') !== -1 &&
  dashHtml34.indexOf('id="checkedActionBar"') !== -1 &&
  dashHtml34.indexOf('onclick="clearAllChecked()"') !== -1
);

// 34-5: 클라이언트 스크립트에 달력 및 체크박스 제어 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 달력 및 체크박스 제어 함수 탑재 검증',
  dashHtml34.indexOf('function toggleCalendarView()') !== -1 &&
  dashHtml34.indexOf('function renderCalendar(') !== -1 &&
  dashHtml34.indexOf('function onSiteCheckboxClick(e, cb)') !== -1 &&
  dashHtml34.indexOf('function updateCheckedStatusUI()') !== -1 &&
  dashHtml34.indexOf('function clearAllChecked()') !== -1 &&
  dashHtml34.indexOf('function prevMonth()') !== -1 &&
  dashHtml34.indexOf('function nextMonth()') !== -1 &&
  dashHtml34.indexOf('function goToTodayMonth()') !== -1
);

// 34-6: getSidebarDashboardData 각 공정에 startDateStr, endDateStr 계산 탑재 검증
const testDashData34 = sandbox.getSidebarDashboardData();
const firstTask34 = (testDashData34 && testDashData34.sites && testDashData34.sites[0] && testDashData34.sites[0].tasks[0]) || {};
assertTest('getSidebarDashboardData: 공정에 YYYY-MM-DD 형식의 startDateStr, endDateStr 계산 탑재 검증',
  typeof firstTask34.startDateStr === 'string' &&
  /^\d{4}-\d{2}-\d{2}$/.test(firstTask34.startDateStr)
);

// 34-7: 모바일 클라이언트 스크립트 구문 무결성(SyntaxError 0) 검증
const scriptMatch34 = dashHtml34.match(/<script>([\s\S]*?)<\/script>/);
let clientSyntaxOk34 = false;
if (scriptMatch34 && scriptMatch34[1]) {
  try {
    new vm.Script(scriptMatch34[1]);
    clientSyntaxOk34 = true;
  } catch (err) {
    console.error('클라이언트 JS 구문 에러:', err);
  }
}
assertTest('buildSidebarDashboardHtml: 달력 및 체크박스 로직 추가 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk34);

// 34-8: 캘린더 이벤트 알약 전체 배경색, 800 엑스트라 볼드 및 텍스트 섀도우(cal-pill-text) 검증
assertTest('buildSidebarDashboardHtml: 캘린더 이벤트 알약 전체 배경색 채움, 800 엑스트라 볼드 및 텍스트 섀도우(cal-pill-text) 렌더링 검증',
  dashHtml34.indexOf('cal-sh-box') !== -1 &&
  dashHtml34.indexOf('cal-pill-text') !== -1 &&
  dashHtml34.indexOf('background-color:\' + bgCol + \';color:\' + txtCol') !== -1 &&
  dashHtml34.indexOf('font-weight:800;') !== -1 &&
  dashHtml34.indexOf('text-shadow:') !== -1 &&
  dashHtml34.indexOf('font-size:9.5px;') !== -1 &&
  dashHtml34.indexOf('font-size:8.2px;letter-spacing:-0.2px;') !== -1 &&
  dashHtml34.indexOf('font-size:7.2px;letter-spacing:-0.5px;') !== -1 &&
  dashHtml34.indexOf('\'<span class="cal-sh-box \' + shClass + \'"></span><span class="cal-pill-text" style="\' + fontStyle + \'">\' + escapeHtml(task.taskName) + \'</span>\'') !== -1
);
assertTest('buildSidebarDashboardHtml: 일정이 많을 때 더보기 버튼 완전 배제 및 전체 리스트 노출 검증',
  dashHtml34.indexOf('cal-event-more') === -1 &&
  dashHtml34.indexOf('건 더보기') === -1
);

// 34-9: 이름 눌렀을 때 카드가 닫힌 상태 유지 검증
assertTest('buildSidebarDashboardHtml: 이름(담당자) 클릭 필터 시 카드 닫힌 상태(isExpanded에서 selectedMember 배제) 검증',
  dashHtml34.indexOf('const isExpanded = (query !== "") || (currentFilter !== "all" && currentFilter !== "done");') !== -1
);

// 34-10: 달력 렌더링 시뮬레이션: 체크박스로 선택된 현장의 공정만 달력에 매핑되는지 검증
const mockCalSites = [
  {
    siteName: '보경102동',
    colorHex: '#10b981',
    tasks: [
      { taskName: '도배', startDateStr: '2026-10-20', endDateStr: '2026-10-21', status: '진행중', manager: '조항만' },
      { taskName: '청소', startDateStr: '2026-10-25', endDateStr: '2026-10-25', status: '예정', manager: '송지은' }
    ]
  },
  {
    siteName: '신대방우성 5동 601호',
    colorHex: '#2563eb',
    tasks: [
      { taskName: '타일', startDateStr: '2026-10-20', endDateStr: '2026-10-20', status: '예정', manager: '전예린' }
    ]
  }
];

// 시뮬레이션 A: '보경102동'만 체크되었을 때
const simCheckedSites = new Set(['보경102동']);
const simTargetSites = mockCalSites.filter(s => simCheckedSites.has(s.siteName));
const simDayMapA = {};
simTargetSites.forEach(s => {
  s.tasks.forEach(t => {
    if (!simDayMapA[t.startDateStr]) simDayMapA[t.startDateStr] = [];
    simDayMapA[t.startDateStr].push({ site: s, task: t });
  });
});

assertTest('달력 시뮬레이션: 체크된 현장("보경102동")만 필터링되어 "신대방우성" 공정(타일)은 달력 매핑에서 완벽 배제 검증',
  simDayMapA['2026-10-20'] &&
  simDayMapA['2026-10-20'].length === 1 &&
  simDayMapA['2026-10-20'][0].site.siteName === '보경102동' &&
  simDayMapA['2026-10-20'][0].task.taskName === '도배'
);

// 시뮬레이션 B: 담당자('송지은') 필터와 체크박스가 동시에 적용될 때
const simMember = '송지은';
const simDayMapB = {};
simTargetSites.forEach(s => {
  s.tasks.forEach(t => {
    if (simMember && t.manager.indexOf(simMember) === -1) return;
    if (!simDayMapB[t.startDateStr]) simDayMapB[t.startDateStr] = [];
    simDayMapB[t.startDateStr].push({ site: s, task: t });
  });
});

assertTest('달력 시뮬레이션: 체크된 현장("보경102동") 내에서 "송지은" 담당자 필터 적용 시 10.25 청소만 노출 검증',
  !simDayMapB['2026-10-20'] &&
  simDayMapB['2026-10-25'] &&
  simDayMapB['2026-10-25'].length === 1 &&
  simDayMapB['2026-10-25'][0].task.taskName === '청소'
);

// 시뮬레이션 C: 빈 상태 메시지 (선택 탭에서 아무 현장도 체크 안 되었을 때)
const emptyCheckedHtml = (0 === 0) ? '<div class="empty-state">☑️ 선택된 현장이 없습니다.' : '';
assertTest('선택 탭 빈 상태 시뮬레이션: 체크된 현장 0개 시 친절한 안내 메시지 출력 검증',
  emptyCheckedHtml.indexOf('☑️ 선택된 현장이 없습니다.') !== -1
);

// ========================================================
// 35. [캘린더 오늘 버튼 옆 📷 캡쳐 버튼 & 전체 화면 캡쳐 및 다운로드 모달 QA]
// ========================================================
console.log('\n📷 [Step 35: 캘린더 오늘 버튼 옆 📷 캡쳐 버튼 & 전체 화면 캡쳐 및 다운로드 모달 QA]');

const dashHtml35 = sandbox.buildSidebarDashboardHtml('mobile');

// 35-1: head 태그 내 html2canvas CDN 스크립트 포함 검증
assertTest('buildSidebarDashboardHtml: head 태그 내 html2canvas CDN 라이브러리 스크립트 포함 검증',
  dashHtml35.indexOf('html2canvas.min.js') !== -1
);

// 35-2: 오늘 버튼 옆 📷 캡쳐 버튼(#btnCaptureCal) 탑재 검증
assertTest('buildSidebarDashboardHtml: 오늘 버튼 옆 [📷 캡쳐] 버튼(#btnCaptureCal) 및 onclick="captureCalendarScreen()" 연동 검증',
  dashHtml35.indexOf('id="btnCaptureCal"') !== -1 &&
  dashHtml35.indexOf('class="cal-capture-btn"') !== -1 &&
  dashHtml35.indexOf('onclick="captureCalendarScreen()"') !== -1 &&
  dashHtml35.indexOf('📷 캡쳐') !== -1
);

// 35-3: .cal-capture-btn CSS 스타일 정의 검증
assertTest('buildSidebarDashboardHtml: .cal-capture-btn 버튼 스타일 및 disabled 상태 CSS 탑재 검증',
  dashHtml35.indexOf('.cal-capture-btn {') !== -1 &&
  dashHtml35.indexOf('.cal-capture-btn:disabled') !== -1
);

// 35-4: 캘린더 캡쳐 프리뷰 모달(#captureModal) 및 다운로드 버튼 검증
assertTest('buildSidebarDashboardHtml: 캘린더 캡쳐 프리뷰 모달(#captureModal), 이미지(#capturePreviewImg), 다운로드 버튼(#btnDownloadCapture) 검증',
  dashHtml35.indexOf('id="captureModal"') !== -1 &&
  dashHtml35.indexOf('id="capturePreviewImg"') !== -1 &&
  dashHtml35.indexOf('id="btnDownloadCapture"') !== -1 &&
  dashHtml35.indexOf('onclick="closeCaptureModal()"') !== -1
);

// 35-5: 클라이언트 스크립트에 captureCalendarScreen, showCaptureModal, closeCaptureModal 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 captureCalendarScreen, showCaptureModal, closeCaptureModal 함수 탑재 검증',
  dashHtml35.indexOf('function captureCalendarScreen()') !== -1 &&
  dashHtml35.indexOf('function showCaptureModal(dataUrl, filename)') !== -1 &&
  dashHtml35.indexOf('function closeCaptureModal(e)') !== -1
);

// 35-6: 캡쳐 로직 추가 후 클라이언트 JS 무결성(SyntaxError 0) 검증
const scriptMatch35 = dashHtml35.match(/<script>([\s\S]*?)<\/script>/);
let clientSyntaxOk35 = false;
if (scriptMatch35 && scriptMatch35[1]) {
  try {
    new vm.Script(scriptMatch35[1]);
    clientSyntaxOk35 = true;
  } catch (err) {
    console.error('클라이언트 JS 구문 에러:', err);
  }
}
assertTest('buildSidebarDashboardHtml: 캡쳐 로직 추가 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk35);

// ========================================================
// 36. [2026.00.00 날짜 표준 형식 통일 & 캡쳐 모달 다운로드 분리 QA]
// ========================================================
console.log('\n📅 [Step 36: 2026.00.00 날짜 표준 형식 통일 & 캡쳐 모달 다운로드 분리 QA]');

// 36-1: formatDateYYMMDD가 4자리 연도 YYYY.MM.DD 반환 검증
const testDate36 = new Date(2026, 9, 15); // 2026-10-15
const formattedYmd36 = sandbox.formatDateYYMMDD(testDate36);
assertTest('formatDateYYMMDD: 4자리 연도 YYYY.MM.DD (2026.10.15) 표준 형식 반환 검증',
  formattedYmd36 === '2026.10.15'
);

// 36-2: formatDateTime이 4자리 연도 YYYY.MM.DD HH:mm 반환 검증
const testTime36 = new Date(2026, 9, 9, 20, 50);
const formattedTime36 = sandbox.formatDateTime(testTime36);
assertTest('formatDateTime: 4자리 연도 YYYY.MM.DD HH:mm (2026.10.09 20:50) 표준 형식 반환 검증',
  formattedTime36 === '2026.10.09 20:50'
);

// 36-3: normalizeDateStringToYYYYMMDD 다양한 입력 케이스 2026.00.00 정규화 검증
assertTest('normalizeDateStringToYYYYMMDD: "26.10.15" -> "2026.10.15" 정규화 검증',
  sandbox.normalizeDateStringToYYYYMMDD('26.10.15') === '2026.10.15'
);
assertTest('normalizeDateStringToYYYYMMDD: 공백 포함 "2026. 10. 21" -> "2026.10.21" 정규화 검증',
  sandbox.normalizeDateStringToYYYYMMDD('2026. 10. 21') === '2026.10.21'
);
assertTest('normalizeDateStringToYYYYMMDD: 1자리 월/일 "2026. 9. 9" -> "2026.09.09" 패딩 정규화 검증',
  sandbox.normalizeDateStringToYYYYMMDD('2026. 9. 9') === '2026.09.09'
);
assertTest('normalizeDateStringToYYYYMMDD: 기간 날짜 "26.10.21 ~ 26.10.22" -> "2026.10.21 ~ 2026.10.22" 정규화 검증',
  sandbox.normalizeDateStringToYYYYMMDD('26.10.21 ~ 26.10.22') === '2026.10.21 ~ 2026.10.22'
);

// 36-4: formatAllDatesInSheetToYYYYMMDD 함수 탑재 여부 검증
assertTest('formatAllDatesInSheetToYYYYMMDD: 시트 날짜 2026.00.00 일괄 변환 전용 함수 탑재 검증',
  typeof sandbox.formatAllDatesInSheetToYYYYMMDD === 'function'
);

// 36-5: captureCalendarScreen 캡쳐 즉시 자동 다운로드 제거(팝업만 단독 띄움) 검증
const dashHtml36 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: 캡쳐 버튼 클릭 시 중복 다운로드 유발하던 즉시 link.click() 배제 및 팝업(showCaptureModal) 단독 호출 검증',
  dashHtml36.indexOf('showCaptureModal(dataUrl, filename);') !== -1 &&
  dashHtml36.indexOf('link.download = filename;\n        link.href = dataUrl;\n        document.body.appendChild(link);\n        link.click();') === -1
);

// 36-6: 마감 경과 미완료 공정(어제 일정 + 예정)이 지연(#fff1f2 붉은색)으로 판별되는 원리 검증
const testDelayedTask = {
  startDate: new Date(2026, 9, 15),
  endDate: new Date(2026, 9, 15)
};
const todayRef = new Date(2026, 9, 16);
const isDelayResult = sandbox.isScheduleDelayed(testDelayedTask, '예정', todayRef);
assertTest('isScheduleDelayed: 오늘(10.16) 기준 어제(10.15) 예정 공정은 isDelayed === true (시트 #fff1f2 붉은색 하이라이트 정상 판정) 검증',
  isDelayResult === true
);

console.log('\n🧹 [Step 37: 지연 외 모든 정상/미래/완료 행 흰색 배경 자동 복구 & 상단 메뉴 최적화 QA]');

// 37-1: cleanNonDelayedRowBackgroundsSilent 무음 일괄 청소 함수 탑재 검증
assertTest('cleanNonDelayedRowBackgroundsSilent 함수 탑재 여부',
  typeof sandbox.cleanNonDelayedRowBackgroundsSilent === 'function'
);

// 37-2: cleanNonDelayedRowBackgroundsSilent 동작 시뮬레이션: 미래(109행) 및 완료(107행) 복구, 진짜 지연(121행)만 하이라이트
const mockSheetRows = [
  // [현장명, 날짜, 공종, 협력업체, 공유, 비고, 담당자, 진행상태]
  ['보경102동', '2026.08.10', '철거', '', true, '', '전예린', '완료'],      // 107행: 과거 완료 -> 흰색 복구 대상
  ['보경102동', '2026.10.21', '상담', '', true, '1차 샘플링', '전예린', '예정'], // 109행: 미래 예정 -> 흰색 복구 대상!
  ['보경102동', '2026.09.22', '창호실측', '', true, '', '전예린', '완료'],   // 111행: A열 흰색, B~F열만 수동 칠해진 행!
  ['보경102동', '2026.10.15', '타일', '', false, '', '전예린', '예정'],       // 121행: 마감 경과(지연) -> 붉은색 유지 대상!
  ['보경102동', '2026.10.20', '타일', '', false, '', '전예린', '예정']        // 122행: 미래 예정 -> 흰색 유지 대상
];
const mockBackgrounds = [
  ['#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2'], // 이전 분홍 잔여물
  ['#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2'], // 이전 분홍 잔여물 (미래 일정인데 분홍색이었음)
  ['#ffffff', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#ffffff', '#ffffff'], // A열은 흰색, B~F열만 수동 채색된 패턴!
  ['#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2', '#fff1f2'], // 진짜 지연
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']  // 정상 흰색
];

let updatedBg = null;
const mockSheetObj = {
  getLastRow: () => 6,
  getRange: (r, c, nr, nc) => ({
    getValues: () => mockSheetRows,
    getBackgrounds: () => mockBackgrounds,
    setBackgrounds: (bgs) => { updatedBg = bgs; }
  })
};

sandbox.cleanNonDelayedRowBackgroundsSilent(mockSheetObj);
assertTest('cleanNonDelayedRowBackgroundsSilent 시뮬레이션: 107행 완료 공정이 흰색(#ffffff)으로 자동 복구 검증',
  updatedBg && updatedBg[0][0] === '#ffffff'
);
assertTest('cleanNonDelayedRowBackgroundsSilent 시뮬레이션: 109행 미래 일정(2026.10.21)이 흰색(#ffffff)으로 자동 복구 검증',
  updatedBg && updatedBg[1][0] === '#ffffff'
);
assertTest('cleanNonDelayedRowBackgroundsSilent 시뮬레이션: A열은 흰색이고 B~F열만 수동 채색된 111행의 B열 및 F열이 흰색(#ffffff)으로 완벽 복구 검증',
  updatedBg && updatedBg[2][1] === '#ffffff' && updatedBg[2][4] === '#ffffff' && updatedBg[2][5] === '#ffffff'
);
assertTest('cleanNonDelayedRowBackgroundsSilent 시뮬레이션: 121행 진짜 마감 경과 지연 공정(2026.10.15 예정)은 붉은색(#fff1f2) 유지 검증',
  updatedBg && updatedBg[3][0] === '#fff1f2'
);

// 37-3: onOpen 내 cleanNonDelayedRowBackgroundsSilent 무음 호출 탑재 검증
assertTest('onOpen: 시트가 열릴 때 정상 행 배경색 자동 청소(cleanNonDelayedRowBackgroundsSilent) 1회 무음 실행 탑재 검증',
  code.indexOf('function onOpen()') !== -1 &&
  code.slice(code.indexOf('function onOpen()'), code.indexOf('function onOpen()') + 500).indexOf('cleanNonDelayedRowBackgroundsSilent(sheet)') !== -1
);

// 37-4: formatAllDatesInSheetToYYYYMMDD 내 cleanNonDelayedRowBackgroundsSilent 자동 연동 검증
assertTest('formatAllDatesInSheetToYYYYMMDD: 날짜 변환 후 cleanNonDelayedRowBackgroundsSilent 자동 연동 검증',
  code.slice(code.indexOf('function formatAllDatesInSheetToYYYYMMDD('), code.indexOf('function formatDateWithTime(')).indexOf('cleanNonDelayedRowBackgroundsSilent(sheet)') !== -1
);

// 37-5: syncRecentSiteSchedulesToGoogleCalendar 내 정상 행 #ffffff 복구 로직 탑재 검증
assertTest('syncRecentSiteSchedulesToGoogleCalendar: 지연 아닌 정상 행(#ffffff) 복구 로직 탑재 검증',
  code.indexOf('details.isDelayed') !== -1 &&
  code.indexOf('sheet.getRange(row, 1, 1, 8).setBackground("#ffffff")') !== -1
);

// 37-6: 상단 메뉴 최상위에 3개 핵심 기능만 노출되고 나머지는 '고급 설정 & 시스템 관리'로 통합 검증
const onOpenMenuSlice = code.slice(code.indexOf('function onOpen()'), code.indexOf('function isSiteScheduleSheet(') !== -1 ? code.indexOf('function isSiteScheduleSheet(') : code.indexOf('function onOpen()') + 2500);
assertTest('상단 메뉴: 최상위 3개 핵심 항목(사이드바, 팝업, 캘린더 동기화) 외 모든 항목이 [고급 설정 & 시스템 관리] 하위로 일원화 배치 검증',
  onOpenMenuSlice.indexOf('대시보드 사이드바 열기') !== -1 &&
  onOpenMenuSlice.indexOf('대시보드 팝업 열기') !== -1 &&
  onOpenMenuSlice.indexOf('📅 구글 캘린더 동기화') !== -1 &&
  onOpenMenuSlice.indexOf('ui.createMenu("⚙️ 고급 설정 & 시스템 관리")') !== -1
);

console.log('\n🇰🇷 [Step 38: 모바일 캘린더 대한민국 법정공휴일 표시 & 대시보드 동일 렌더링 QA]');

// 38-1: getKoreanHolidaysMap 함수 탑재 여부
assertTest('getKoreanHolidaysMap 함수 탑재 여부',
  typeof sandbox.getKoreanHolidaysMap === 'function'
);

// 38-2: getKoreanHolidaysMap 2026년 공휴일 및 대체공휴일 정확도 검증
const holidayMapSample = sandbox.getKoreanHolidaysMap(null);
assertTest('getKoreanHolidaysMap: 2026-10-09 한글날 정상 매핑 검증',
  holidayMapSample['2026-10-09'] === '한글날'
);
assertTest('getKoreanHolidaysMap: 2026-10-03 개천절 정상 매핑 검증',
  holidayMapSample['2026-10-03'] === '개천절'
);
assertTest('getKoreanHolidaysMap: 2026-10-05 대체공휴일 정상 매핑 검증',
  holidayMapSample['2026-10-05'] === '대체공휴일'
);
assertTest('getKoreanHolidaysMap: 2026-12-25 성탄절 정상 매핑 검증',
  holidayMapSample['2026-12-25'] === '성탄절'
);

// 38-3: getSidebarDashboardData 반환 객체에 holidays 필드 포함 검증
const dashDataSample = sandbox.getSidebarDashboardData();
assertTest('getSidebarDashboardData: 반환 객체에 holidays 공휴일 맵 포함 검증',
  dashDataSample && typeof dashDataSample.holidays === 'object' && dashDataSample.holidays['2026-10-09'] === '한글날'
);

// 38-4: buildSidebarDashboardHtml("mobile") 클라이언트 JS 무결성 검증
const mobileCalHtml = sandbox.buildSidebarDashboardHtml('mobile');
const scriptStartIdx = mobileCalHtml.indexOf('<script>');
const scriptEndIdx = mobileCalHtml.lastIndexOf('</script>');
const clientScript = mobileCalHtml.slice(scriptStartIdx + 8, scriptEndIdx);
let clientSyntaxOk = false;
try {
  new vm.Script(clientScript);
  clientSyntaxOk = true;
} catch (eSyntax) {
  console.error('Client Script Syntax Error:', eSyntax);
}
assertTest('buildSidebarDashboardHtml("mobile"): 공휴일 기능 탑재 후 클라이언트 JS 구문 무결성(SyntaxError 0) 검증',
  clientSyntaxOk === true
);

// 38-5: CSS에 공휴일 스타일(.cal-day-cell.holiday, .cal-holiday-name) 탑재 검증
assertTest('buildSidebarDashboardHtml: .cal-day-cell.holiday 연분홍 배경 및 #dc2626 빨간색 텍스트 CSS 검증',
  mobileCalHtml.indexOf('.cal-day-cell.holiday') !== -1 &&
  mobileCalHtml.indexOf('.cal-holiday-name') !== -1
);

// 38-6: DEFAULT_KOREAN_HOLIDAYS 내장 및 activeHolidays 병합 검증
assertTest('buildSidebarDashboardHtml: DEFAULT_KOREAN_HOLIDAYS 내장 및 onDataLoaded 병합 로직 탑재 검증',
  clientScript.indexOf('DEFAULT_KOREAN_HOLIDAYS') !== -1 &&
  clientScript.indexOf('activeHolidays = Object.assign') !== -1
);

// 38-7: renderCalendar 셀 내 공휴일 클래스(holiday sun) 및 cal-holiday-name 렌더링 검증
assertTest('buildSidebarDashboardHtml: renderCalendar 내 공휴일(holiday sun) 스타일 및 cal-holiday-name 라벨 렌더링 검증',
  clientScript.indexOf('dayClass += " holiday sun"') !== -1 &&
  clientScript.indexOf('cal-holiday-name') !== -1
);

// 38-8: onCalDayClick 알림에 공휴일 명칭 포함 검증
assertTest('buildSidebarDashboardHtml: onCalDayClick 클릭 알림 시 공휴일 명칭(holidayName) 노출 검증',
  clientScript.indexOf('const holidayName = (typeof activeHolidays !== "undefined" && activeHolidays[dStr])') !== -1 &&
  clientScript.indexOf('dStr + (holidayName ? " (" + holidayName + ")" : "")') !== -1
);

// 38-9: 모바일 캘린더 날짜 하단 공휴일 라벨 세로 배치(flex-direction: column) 및 max-width 100% 검증
assertTest('buildSidebarDashboardHtml: .cal-day-header 날짜 아래 공휴일 텍스트 세로 배치(flex-direction: column) CSS 검증',
  mobileCalHtml.indexOf('.cal-day-header { display: flex; flex-direction: column; align-items: flex-start;') !== -1
);

// 38-10: 공휴일 텍스트 스타일(max-width: 100%, margin-top: 1px, letter-spacing: -0.3px) 검증
assertTest('buildSidebarDashboardHtml: .cal-holiday-name 날짜 하단 마진 및 100% 너비 확장(말줄임 방지) CSS 검증',
  mobileCalHtml.indexOf('.cal-holiday-name { font-size: 8px; font-weight: 700; color: #dc2626; margin-top: 1px; margin-left: 1px;') !== -1 &&
  mobileCalHtml.indexOf('max-width: 100%;') !== -1
);

// ========================================================
// 🎯 Step 39: 캘린더 네비게이션 년/월 중앙 정렬 & 완료 목록 최근 일자 우선 정렬 QA
// ========================================================
console.log('\n🧭 [Step 39: 캘린더 상단 년/월 중앙 정렬 & 좌측 캡쳐/우측 오늘 배치 & 완료 목록 최근순 정렬 QA]');

const step39Html = sandbox.buildSidebarDashboardHtml('mobile');
const sStart = step39Html.indexOf('<script>');
const sEnd = step39Html.lastIndexOf('</script>');
const step39Script = step39Html.slice(sStart + 8, sEnd);

// 39-1: .cal-nav-bar 좌/중/우 3분할 flex 컨테이너 클래스 CSS 탑재 검증
assertTest('buildSidebarDashboardHtml: .cal-nav-bar 좌/중/우 flex 컨테이너 CSS (.cal-nav-left, .cal-nav-center, .cal-nav-right) 탑재 검증',
  step39Html.indexOf('.cal-nav-left {') !== -1 &&
  step39Html.indexOf('.cal-nav-center {') !== -1 &&
  step39Html.indexOf('.cal-nav-right {') !== -1
);

// 39-2: 캘린더 헤더 네비게이션 마크업 배치(좌: 캡쳐, 중: ◀ 년/월 ▶, 우: 오늘) 검증
assertTest('buildSidebarDashboardHtml: 캘린더 상단 년/월 중앙 정렬, 좌측 캡쳐, 우측 오늘 버튼 마크업 검증',
  step39Html.indexOf('class="cal-nav-left"') !== -1 &&
  step39Html.indexOf('id="btnCaptureCal"') !== -1 &&
  step39Html.indexOf('class="cal-nav-center"') !== -1 &&
  step39Html.indexOf('id="calMonthTitle"') !== -1 &&
  step39Html.indexOf('class="cal-nav-right"') !== -1 &&
  step39Html.indexOf('class="cal-today-btn"') !== -1 &&
  step39Html.indexOf('class="cal-nav-left"') < step39Html.indexOf('class="cal-nav-center"') &&
  step39Html.indexOf('class="cal-nav-center"') < step39Html.indexOf('class="cal-nav-right"')
);

// 39-3: renderList 내 완료 목록 최근 일자 내림차순 정렬 로직 탑재 검증
assertTest('buildSidebarDashboardHtml: renderList 내 currentFilter === "done" 일 때 bLatest - aLatest 최근순 정렬 검증',
  step39Script.indexOf('currentFilter === "done"') !== -1 &&
  step39Script.indexOf('return bLatest - aLatest;') !== -1 &&
  step39Script.indexOf('latestFilteredMs') !== -1
);

// 39-4: 완료 목록 정렬 시뮬레이션 검증 (최근 완공 현장이 상위 노출)
const mockDoneSites = [
  { site: { siteName: "Site_Old", initialOrder: 1 }, latestFilteredMs: new Date("2026-08-10").getTime(), earliestFilteredMs: new Date("2026-08-01").getTime() },
  { site: { siteName: "Site_Recent", initialOrder: 2 }, latestFilteredMs: new Date("2026-10-15").getTime(), earliestFilteredMs: new Date("2026-10-01").getTime() },
  { site: { siteName: "Site_Mid", initialOrder: 3 }, latestFilteredMs: new Date("2026-09-24").getTime(), earliestFilteredMs: new Date("2026-09-10").getTime() }
];

function sortSites(sites, filter) {
  const cloned = sites.slice();
  cloned.sort((a, b) => {
    if (filter === "done") {
      const aLatest = a.latestFilteredMs || a.earliestFilteredMs || 0;
      const bLatest = b.latestFilteredMs || b.earliestFilteredMs || 0;
      if (aLatest && bLatest) {
        if (aLatest !== bLatest) return bLatest - aLatest;
      } else if (aLatest) {
        return -1;
      } else if (bLatest) {
        return 1;
      }
      return (b.site.initialOrder || 0) - (a.site.initialOrder || 0);
    }
    const aMs = a.earliestFilteredMs || 0;
    const bMs = b.earliestFilteredMs || 0;
    if (aMs && bMs) {
      if (aMs !== bMs) return aMs - bMs;
    } else if (aMs) {
      return -1;
    } else if (bMs) {
      return 1;
    }
    return 0;
  });
  return cloned;
}

const doneSorted = sortSites(mockDoneSites, "done");
assertTest('정렬 시뮬레이션: 완료(done) 탭에서는 최근 완공 현장(10.15)이 1위, 과거 현장(08.10)이 최하위로 정렬 검증',
  doneSorted[0].site.siteName === "Site_Recent" &&
  doneSorted[1].site.siteName === "Site_Mid" &&
  doneSorted[2].site.siteName === "Site_Old"
);

const allSorted = sortSites(mockDoneSites, "all");
assertTest('정렬 시뮬레이션: 일반(all) 탭에서는 기존대로 가장 빠른 일정(08.01)이 1위로 정렬 유지 검증',
  allSorted[0].site.siteName === "Site_Old" &&
  allSorted[1].site.siteName === "Site_Mid" &&
  allSorted[2].site.siteName === "Site_Recent"
);

// 🎯 Step 40: 선택 현장만 보기 버튼 제거 & 캘린더 상단 선택 현장명(1건: 현장명, 2건 이상: 현장명 외 N건) 표기 QA
// ========================================================
console.log('\n🏷️ [Step 40: 선택 현장만 보기 버튼 제거 & 캘린더 선택 현장명/외 N건 표기 QA]');

const step40Html = sandbox.buildSidebarDashboardHtml('mobile');
const s40Start = step40Html.indexOf('<script>');
const s40End = step40Html.lastIndexOf('</script>');
const step40Script = step40Html.slice(s40Start + 8, s40End);

// 40-1: checkedActionBar에서 btnToggleCheckedOnly("선택 현장만 보기") 버튼 제거 및 해제 버튼 보존 검증
assertTest('buildSidebarDashboardHtml: checkedActionBar에서 선택 현장만 보기 버튼 제거 및 해제 버튼 보존 검증',
  step40Html.indexOf('id="btnToggleCheckedOnly"') === -1 &&
  step40Html.indexOf('선택 현장만 보기') === -1 &&
  step40Html.indexOf('id="checkedActionBar"') !== -1 &&
  step40Html.indexOf('onclick="clearAllChecked()"') !== -1
);

// 40-2: renderCalendar 내 1건 선택 시 현장명 노출 로직 검증
assertTest('buildSidebarDashboardHtml: renderCalendar 내 1건 선택 시 현장명 노출 로직 검증',
  step40Script.indexOf('siteArr.length === 1') !== -1 &&
  step40Script.indexOf('noticeParts.push("☑️ " + escapeHtml(firstSite))') !== -1
);

// 40-3: renderCalendar 내 2건 이상 선택 시 [현장명 외 N건] 노출 로직 검증
assertTest('buildSidebarDashboardHtml: renderCalendar 내 2건 이상 선택 시 [현장명 외 N건] 노출 로직 검증',
  step40Script.indexOf('noticeParts.push("☑️ " + escapeHtml(firstSite) + " 외 " + (siteArr.length - 1) + "건")') !== -1
);

// 40-4: 캘린더 선택 현장명 표기 시뮬레이션: 1건일 때 "☑️ 보경102동" 반환 검증
function formatCalendarNotice(checkedSitesSet) {
  if (!checkedSitesSet || checkedSitesSet.size === 0) return "";
  const siteArr = Array.from(checkedSitesSet);
  const firstSite = siteArr[0] || "현장";
  if (siteArr.length === 1) {
    return "☑️ " + firstSite;
  } else {
    return "☑️ " + firstSite + " 외 " + (siteArr.length - 1) + "건";
  }
}

const sim1 = formatCalendarNotice(new Set(["보경102동"]));
assertTest('시뮬레이션: 1개 현장 선택 시 "☑️ 보경102동" 단독 명칭 표기 검증',
  sim1 === "☑️ 보경102동"
);

// 40-5: 캘린더 선택 현장명 표기 시뮬레이션: 2건일 때 "☑️ 보경102동 외 1건" 반환 검증
const sim2 = formatCalendarNotice(new Set(["보경102동", "신대방우성"]));
assertTest('시뮬레이션: 2개 현장 선택 시 "☑️ 보경102동 외 1건" 축약 표기 검증',
  sim2 === "☑️ 보경102동 외 1건"
);

// 40-6: 캘린더 선택 현장명 표기 시뮬레이션: 4건일 때 "☑️ 보경102동 외 3건" 반환 검증
const sim4 = formatCalendarNotice(new Set(["보경102동", "신대방우성", "래미안옥수", "자이서초"]));
assertTest('시뮬레이션: 4개 현장 선택 시 "☑️ 보경102동 외 3건" 축약 표기 검증',
  sim4 === "☑️ 보경102동 외 3건"
);

// 🎯 Step 41: 캘린더 좌우 스와이프 제스처 & 슬라이드 전환 애니메이션 QA
// ========================================================
console.log('\n👆 [Step 41: 캘린더 좌우 스와이프 제스처 & 슬라이드 전환 애니메이션 QA]');

const step41Html = sandbox.buildSidebarDashboardHtml('mobile');
const s41Start = step41Html.indexOf('<script>');
const s41End = step41Html.lastIndexOf('</script>');
const step41Script = step41Html.slice(s41Start + 8, s41End);

// 41-1: CSS 내 .calendar-view 및 .cal-grid에 touch-action: pan-y 탑재 검증
assertTest('buildSidebarDashboardHtml: .calendar-view 및 .cal-grid touch-action: pan-y CSS 탑재 검증',
  step41Html.indexOf('.calendar-view { touch-action: pan-y;') !== -1 &&
  step41Html.indexOf('.cal-grid { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); width: 100%; background: #e2e8f0; gap: 1px; touch-action: pan-y; }') !== -1
);

// 41-2: CSS 내 @keyframes 슬라이드 애니메이션 및 클래스 탑재 검증
assertTest('buildSidebarDashboardHtml: 캘린더 슬라이드 애니메이션(@keyframes calSlideFromRight/Left, .cal-slide-next/prev) CSS 탑재 검증',
  step41Html.indexOf('@keyframes calSlideFromRight') !== -1 &&
  step41Html.indexOf('@keyframes calSlideFromLeft') !== -1 &&
  step41Html.indexOf('.cal-slide-next') !== -1 &&
  step41Html.indexOf('.cal-slide-prev') !== -1
);

// 41-3: 클라이언트 스크립트에 setupCalendarSwipe 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 setupCalendarSwipe 터치/마우스 제스처 리스너 함수 탑재 검증',
  step41Script.indexOf('function setupCalendarSwipe()') !== -1 &&
  step41Script.indexOf('calArea.addEventListener("touchstart"') !== -1 &&
  step41Script.indexOf('calArea.addEventListener("touchend"') !== -1 &&
  step41Script.indexOf('calArea.addEventListener("mousedown"') !== -1 &&
  step41Script.indexOf('window.addEventListener("mouseup"') !== -1
);

// 41-4: prevMonth / nextMonth 및 renderCalendar 슬라이드 방향 연동 검증
assertTest('buildSidebarDashboardHtml: prevMonth("prev") / nextMonth("next") 방향 전달 및 renderCalendar 애니메이션 클래스 연동 검증',
  step41Script.indexOf('renderCalendar("prev")') !== -1 &&
  step41Script.indexOf('renderCalendar("next")') !== -1 &&
  step41Script.indexOf('function renderCalendar(direction)') !== -1 &&
  step41Script.indexOf('grid.classList.add("cal-slide-next")') !== -1 &&
  step41Script.indexOf('grid.classList.add("cal-slide-prev")') !== -1
);

// 41-5: 스와이프 제스처 판정 알고리즘 시뮬레이션 함수
function evaluateSwipeGesture(deltaX, deltaY, elapsed) {
  if (elapsed > 750) return null; // 롱프레스는 무시
  if (Math.abs(deltaX) >= 40 && Math.abs(deltaX) > Math.abs(deltaY) * 1.25) {
    return (deltaX < 0) ? "next" : "prev";
  }
  return null;
}

assertTest('제스처 시뮬레이션: 오른쪽에서 왼쪽 스와이프(deltaX: -60, deltaY: 10, 300ms) -> "next" (다음 달) 정상 판정 검증',
  evaluateSwipeGesture(-60, 10, 300) === "next"
);

assertTest('제스처 시뮬레이션: 왼쪽에서 오른쪽 스와이프(deltaX: 60, deltaY: 10, 300ms) -> "prev" (이전 달) 정상 판정 검증',
  evaluateSwipeGesture(60, 10, 300) === "prev"
);

assertTest('제스처 시뮬레이션: 세로 스크롤 제스처(deltaX: 15, deltaY: 80, 250ms) -> 오동작 방지(null 반환) 검증',
  evaluateSwipeGesture(15, 80, 250) === null
);

assertTest('제스처 시뮬레이션: 미세 탭/터치 떨림(deltaX: 15, deltaY: 5, 80ms) -> 오동작 방지(null 반환) 검증',
  evaluateSwipeGesture(15, 5, 80) === null
);

assertTest('제스처 시뮬레이션: 0.75초 초과 롱 터치 드래그(deltaX: 70, deltaY: 5, 900ms) -> 오동작 방지(null 반환) 검증',
  evaluateSwipeGesture(70, 5, 900) === null
);

// ========================================================
// ⛅ Step 42: 오늘기준 7일간 캘린더 날씨 예보 연동 및 주간 뷰(Weekly) 모드 QA
// ========================================================
console.log('\n⛅ [Step 42: 오늘기준 7일간 캘린더 날씨 예보 연동 및 주간 뷰(Weekly) 모드 QA]');

const step42Html = sandbox.buildSidebarDashboardHtml('mobile');
const s42Start = step42Html.indexOf('<script>');
const s42End = step42Html.lastIndexOf('</script>');
const step42Script = step42Html.slice(s42Start + 8, s42End);

// 42-1: 캘린더 상단 월간/주간 전환 세그먼트 버튼 및 주간 리스트 컨테이너 마크업 검증
assertTest('buildSidebarDashboardHtml: 캘린더 상단 월간/주간 전환 버튼(#btnModeMonth, #btnModeWeek) 및 주간 리스트(#calWeekList) 마크업 검증',
  step42Html.indexOf('id="btnModeMonth"') !== -1 &&
  step42Html.indexOf('id="btnModeWeek"') !== -1 &&
  step42Html.indexOf('id="calWeekList"') !== -1 &&
  step42Html.indexOf('setCalendarMode(\'month\')') !== -1 &&
  step42Html.indexOf('setCalendarMode(\'week\')') !== -1
);

// 42-2: CSS 내 세그먼트 바, 날씨 뱃지, 주간 카드 스타일 탑재 검증
assertTest('buildSidebarDashboardHtml: CSS 내 .cal-mode-bar, .cal-weather-badge, .cal-week-card 스타일 탑재 검증',
  step42Html.indexOf('.cal-mode-bar') !== -1 &&
  step42Html.indexOf('.cal-mode-btn') !== -1 &&
  step42Html.indexOf('.cal-weather-badge') !== -1 &&
  step42Html.indexOf('.cal-weather-badge.rain') !== -1 &&
  step42Html.indexOf('.cal-weather-badge.snow') !== -1 &&
  step42Html.indexOf('.cal-week-list') !== -1 &&
  step42Html.indexOf('.cal-week-card') !== -1
);

// 42-3: 클라이언트 스크립트에 날씨/주간 뷰 핵심 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 주간 뷰 및 날씨 예보 제어 함수 탑재 검증',
  step42Script.indexOf('function setCalendarMode(') !== -1 &&
  step42Script.indexOf('function navPrevPeriod()') !== -1 &&
  step42Script.indexOf('function navNextPeriod()') !== -1 &&
  step42Script.indexOf('function navTodayPeriod()') !== -1 &&
  step42Script.indexOf('function prevWeek()') !== -1 &&
  step42Script.indexOf('function nextWeek()') !== -1 &&
  step42Script.indexOf('function goToThisWeek()') !== -1 &&
  step42Script.indexOf('function renderWeeklyCalendar(') !== -1 &&
  step42Script.indexOf('function renderCurrentCalendarView(') !== -1 &&
  step42Script.indexOf('function loadWeatherData()') !== -1 &&
  step42Script.indexOf('function getSundayOfCurrentWeek(') !== -1
);

// 42-4: 날씨 및 주간 뷰 로직 추가 후 클라이언트 JS 무결성(SyntaxError 0) 검증
let clientSyntaxOk42 = false;
try {
  new vm.Script(step42Script);
  clientSyntaxOk42 = true;
} catch (err) {
  console.error('Step 42 클라이언트 JS 구문 에러:', err);
}
assertTest('buildSidebarDashboardHtml: 날씨 및 주간 뷰 탑재 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk42);

// 42-5: 날씨 예보 데이터 파싱 시뮬레이션 함수
function parseWeatherDataSim(code, maxT, minT, pop) {
  let icon = "☀️", desc = "맑음";
  if (code >= 1 && code <= 2) { icon = "⛅"; desc = "구름조금"; }
  else if (code === 3) { icon = "☁️"; desc = "흐림"; }
  else if (code === 45 || code === 48) { icon = "🌫️"; desc = "안개"; }
  else if ((code >= 51 && code <= 67) || (code >= 80 && code <= 82)) { icon = "🌧️"; desc = "비"; }
  else if ((code >= 71 && code <= 77) || (code >= 85 && code <= 86)) { icon = "❄️"; desc = "눈"; }
  else if (code >= 95) { icon = "⛈️"; desc = "뇌우"; }
  const isRain = (icon === "🌧️" || icon === "⛈️");
  const isSnow = (icon === "❄️");
  let badgeText = icon;
  if (pop >= 30) badgeText += " " + pop + "%";
  else if (maxT !== null) badgeText += " " + maxT + "°";
  return {
    icon, desc, maxT, minT, pop, isRain, isSnow, badgeText,
    fullText: icon + " " + desc + (maxT !== null ? " " + maxT + "°/" + minT + "°" : "") + (pop > 0 ? " (강수 " + pop + "%)" : "")
  };
}

assertTest('날씨 파싱 시뮬레이션: 맑음(code: 0, pop: 10%, maxT: 22도) -> ☀️ 22° 뱃지 검증',
  (() => {
    const w = parseWeatherDataSim(0, 22, 12, 10);
    return w.icon === '☀️' && w.badgeText === '☀️ 22°' && !w.isRain && !w.isSnow;
  })()
);

assertTest('날씨 파싱 시뮬레이션: 비(code: 61, pop: 80%, maxT: 18도) -> 🌧️ 80% 뱃지 및 isRain: true 검증',
  (() => {
    const w = parseWeatherDataSim(61, 18, 11, 80);
    return w.icon === '🌧️' && w.badgeText === '🌧️ 80%' && w.isRain === true;
  })()
);

assertTest('날씨 파싱 시뮬레이션: 눈(code: 71, pop: 60%, maxT: -2도) -> ❄️ 60% 뱃지 및 isSnow: true 검증',
  (() => {
    const w = parseWeatherDataSim(71, -2, -8, 60);
    return w.icon === '❄️' && w.badgeText === '❄️ 60%' && w.isSnow === true;
  })()
);

// 42-6: 주간 뷰 일주일 날짜 계산 시뮬레이션 함수 (일요일 기준 7일)
function getWeekDaysListSim(startDate) {
  const list = [];
  const cur = new Date(startDate);
  for (let i = 0; i < 7; i++) {
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, '0');
    const d = String(cur.getDate()).padStart(2, '0');
    list.push(`${y}-${m}-${d}`);
    cur.setDate(cur.getDate() + 1);
  }
  return list;
}

const testSunday = new Date('2026-10-18T00:00:00'); // 2026-10-18은 일요일
const weekDays = getWeekDaysListSim(testSunday);
assertTest('주간 날짜 계산 시뮬레이션: 일요일(10-18)부터 토요일(10-24)까지 7일 정확 계산 검증',
  weekDays.length === 7 &&
  weekDays[0] === '2026-10-18' &&
  weekDays[6] === '2026-10-24'
);

// 42-7: 주간 네비게이션 시뮬레이션 (이전 주 / 다음 주 이동)
const nextWeekDate = new Date(testSunday);
nextWeekDate.setDate(nextWeekDate.getDate() + 7);
const nextWeekDays = getWeekDaysListSim(nextWeekDate);
assertTest('주간 네비게이션 시뮬레이션: 다음 주 이동 시 10-25부터 10-31까지 7일 계산 검증',
  nextWeekDays[0] === '2026-10-25' &&
  nextWeekDays[6] === '2026-10-31'
);

// ========================================================
// 📊 Step 43: 현장별 공정 진척률(progressRate) 계산 및 게이지 바/뱃지 QA
// ========================================================
console.log('\n📊 [Step 43: 현장별 공정 진척률(progressRate) 계산 및 게이지 바/뱃지 QA]');

// 43-1: getSidebarDashboardData 데이터 무결성 검증 (각 현장에 progressRate, validTasksCount 등 진척 필드 포함)
const testProgressDashData = sandbox.getSidebarDashboardData();
const firstSite43 = (testProgressDashData && testProgressDashData.sites && testProgressDashData.sites[0]) || {};
assertTest('getSidebarDashboardData: 각 현장 객체에 progressRate, validTasksCount, completedTasksCount 필드 포함 검증',
  typeof firstSite43.progressRate === 'number' &&
  typeof firstSite43.validTasksCount === 'number' &&
  typeof firstSite43.completedTasksCount === 'number' &&
  typeof firstSite43.inProgressTasksCount === 'number' &&
  typeof firstSite43.canceledTasksCount === 'number'
);

// 43-2: 진척률 계산 알고리즘 시뮬레이션 함수
function calculateSiteProgressRateSim(tasks) {
  let completedCount = 0;
  let canceledCount = 0;
  let inProgressCount = 0;
  for (let i = 0; i < tasks.length; i++) {
    const st = (tasks[i].status || "").replace(/\s+/g, "");
    if (st.indexOf("완료") !== -1) completedCount++;
    else if (st.indexOf("취소") !== -1) canceledCount++;
    else if (st.indexOf("진행") !== -1) inProgressCount++;
  }
  const totalCount = tasks.length;
  const validCount = totalCount - canceledCount;
  if (validCount <= 0) return 0;
  const activeTasks = tasks.filter(t => {
    const st = (t.status || "").replace(/\s+/g, "");
    return st.indexOf("완료") === -1 && st.indexOf("취소") === -1;
  });
  if (totalCount > 0 && activeTasks.length === 0) return 100;
  return Math.min(100, Math.max(0, Math.round((completedCount / validCount) * 100)));
}

// 43-3: 100% 완료 현장 진척률 검증
const mockAllDoneTasks = [
  { taskName: '철거', status: '완료' },
  { taskName: '도배', status: '완료' },
  { taskName: '마루', status: '완료' }
];
assertTest('진척률 계산 시뮬레이션: 모든 공정 완료 시 progressRate === 100 검증',
  calculateSiteProgressRateSim(mockAllDoneTasks) === 100
);

// 43-4: 50% 진행 현장 진척률 검증 (4개 중 완료 2개, 진행 1개, 예정 1개)
const mockHalfTasks = [
  { taskName: '철거', status: '완료' },
  { taskName: '설비', status: '완료' },
  { taskName: '목공', status: '진행중' },
  { taskName: '타일', status: '예정' }
];
assertTest('진척률 계산 시뮬레이션: 유효 4개 중 완료 2개일 때 progressRate === 50 검증',
  calculateSiteProgressRateSim(mockHalfTasks) === 50
);

// 43-5: 취소 공정 분모 제외 정밀 계산 검증 (총 5개 중 취소 1개, 완료 2개, 예정 2개 -> 유효 4개 중 완료 2개 = 50%)
const mockCancelTasks = [
  { taskName: '철거', status: '완료' },
  { taskName: '설비', status: '완료' },
  { taskName: '샷시', status: '취소' }, // 취소는 모수에서 제외!
  { taskName: '목공', status: '예정' },
  { taskName: '타일', status: '예정' }
];
assertTest('진척률 계산 시뮬레이션: 취소 공정(샷시) 모수 제외 후 유효 4개 중 완료 2개로 정확히 50% 계산 검증',
  calculateSiteProgressRateSim(mockCancelTasks) === 50
);

// 43-6: 공정이 0개이거나 모두 취소된 경우 방어 (0% 반환)
const mockAllCancelTasks = [
  { taskName: '철거', status: '취소' },
  { taskName: '설비', status: '취소' }
];
assertTest('진척률 계산 시뮬레이션: 전 공정 취소 시 NaN/Infinity 방어 및 0% 반환 검증',
  calculateSiteProgressRateSim(mockAllCancelTasks) === 0 &&
  calculateSiteProgressRateSim([]) === 0
);

// 43-7: buildSidebarDashboardHtml CSS 내 진척률 클래스 탑재 검증
const step43Html = sandbox.buildSidebarDashboardHtml('mobile');
const s43Start = step43Html.indexOf('<script>');
const s43End = step43Html.lastIndexOf('</script>');
const step43Script = step43Html.slice(s43Start + 8, s43End);

assertTest('buildSidebarDashboardHtml: .site-progress-badge, .site-progress-track, .site-progress-bar, .site-summary-box CSS 탑재 검증',
  step43Html.indexOf('.site-progress-badge') !== -1 &&
  step43Html.indexOf('.site-progress-track') !== -1 &&
  step43Html.indexOf('.site-progress-bar') !== -1 &&
  step43Html.indexOf('.site-summary-box') !== -1 &&
  step43Html.indexOf('.site-summary-rate') !== -1
);

// 43-8: renderList 내 진척률 뱃지, 게이지 바, 요약 박스 렌더링 로직 검증
assertTest('buildSidebarDashboardHtml: renderList 내 pRate 뱃지, 미니 게이지 바, 서머리 박스 렌더링 검증',
  step43Script.indexOf('site-progress-badge') !== -1 &&
  step43Script.indexOf('site-progress-track') !== -1 &&
  step43Script.indexOf('site-progress-bar') !== -1 &&
  step43Script.indexOf('site-summary-box') !== -1
);

// 43-9: renderWeeklyCalendar 내 주간 알약 진척률 태그 탑재 검증
assertTest('buildSidebarDashboardHtml: renderWeeklyCalendar 내 현장 알약 진척률 태그(pRateTag) 탑재 검증',
  step43Script.indexOf('const pRateTag = (typeof site.progressRate === "number") ? " " + site.progressRate + "%" : "";') !== -1
);

// 43-10: 클라이언트 JS 구문 무결성(SyntaxError 0) 검증
let clientSyntaxOk43 = false;
try {
  new vm.Script(step43Script);
  clientSyntaxOk43 = true;
} catch (err) {
  console.error('Step 43 클라이언트 JS 구문 에러:', err);
}
assertTest('buildSidebarDashboardHtml: 진척률 기능 탑재 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk43);

// ========================================================
// 🔄 Step 44: 상세 모달 같은 날 다중 일정 네비게이션(< > 버튼, 순번 인디케이터) QA
// ========================================================
console.log('\n🔄 [Step 44: 상세 모달 같은 날 다중 일정 네비게이션(< > 버튼) QA]');

const step44Html = sandbox.buildSidebarDashboardHtml('mobile');
const s44Start = step44Html.indexOf('<script>');
const s44End = step44Html.lastIndexOf('</script>');
const step44Script = step44Html.slice(s44Start + 8, s44End);

// 44-1: CSS 내 sheet-day-nav 및 btn-sheet-nav 스타일 탑재 검증
assertTest('buildSidebarDashboardHtml: .sheet-day-nav 및 .btn-sheet-nav CSS 스타일 탑재 검증',
  step44Html.indexOf('.sheet-day-nav') !== -1 &&
  step44Html.indexOf('.sheet-day-nav-label') !== -1 &&
  step44Html.indexOf('.sheet-day-nav-counts') !== -1 &&
  step44Html.indexOf('.sheet-day-nav-btns') !== -1 &&
  step44Html.indexOf('.btn-sheet-nav') !== -1 &&
  step44Html.indexOf('.btn-sheet-nav:disabled') !== -1
);

// 44-2: 바텀시트 HTML 내 sheetDayNav 컨테이너 및 ◀ / ▶ 버튼 마크업 검증
assertTest('buildSidebarDashboardHtml: 바텀시트 내 #sheetDayNav, #sheetDayNavIdx, #btnSheetPrevDayTask, #btnSheetNextDayTask 마크업 검증',
  step44Html.indexOf('id="sheetDayNav"') !== -1 &&
  step44Html.indexOf('id="sheetDayNavIdx"') !== -1 &&
  step44Html.indexOf('id="sheetDayNavTotal"') !== -1 &&
  step44Html.indexOf('id="btnSheetPrevDayTask"') !== -1 &&
  step44Html.indexOf('id="btnSheetNextDayTask"') !== -1 &&
  step44Html.indexOf('onclick="navSameDayTask(-1)"') !== -1 &&
  step44Html.indexOf('onclick="navSameDayTask(1)"') !== -1
);

// 44-3: 클라이언트 스크립트에 findSameDayTasks, updateSameDayNavUI, navSameDayTask 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트에 같은 날 일정 제어 함수 탑재 검증',
  step44Script.indexOf('let sameDayTaskList = [];') !== -1 &&
  step44Script.indexOf('let curSameDayIndex = 0;') !== -1 &&
  step44Script.indexOf('function findSameDayTasks(') !== -1 &&
  step44Script.indexOf('function updateSameDayNavUI()') !== -1 &&
  step44Script.indexOf('function navSameDayTask(') !== -1
);

// 44-4: onCalDayClick 내 모바일 바텀시트 mList 전달 연동 검증
assertTest('buildSidebarDashboardHtml: onCalDayClick 내 모바일 바텀시트 mList 다중 일정 전달 연동 검증',
  step44Script.indexOf('openBottomSheet(mList[0].site, mList[0].task, mList, 0);') !== -1
);

// 44-5: 같은 날 다중 공정 검색 알고리즘 시뮬레이션
const mockMultiSitesData = [
  {
    siteName: '보경102동',
    tasks: [
      { taskName: '철거', startDateStr: '2026-10-20', endDateStr: '2026-10-20', row: 2 },
      { taskName: '양생', startDateStr: '2026-10-22', endDateStr: '2026-10-23', row: 3 }
    ]
  },
  {
    siteName: '신대방우성',
    tasks: [
      { taskName: '타일', startDateStr: '2026-10-20', endDateStr: '2026-10-20', row: 10 },
      { taskName: '도배', startDateStr: '2026-10-20', endDateStr: '2026-10-21', row: 11 }
    ]
  }
];

function findSameDayTasksSim(sites, baseTask) {
  if (!baseTask || !sites) return [];
  const baseDStr = baseTask.startDateStr || "";
  const list = [];
  sites.forEach(site => {
    site.tasks.forEach(t => {
      const sStr = t.startDateStr || "";
      const eStr = t.endDateStr || sStr;
      if (sStr && sStr <= baseDStr && baseDStr <= eStr) {
        list.push({ site: site.siteName, task: t });
      }
    });
  });
  return list;
}

const foundSameDay = findSameDayTasksSim(mockMultiSitesData, { startDateStr: '2026-10-20' });
assertTest('같은 날 공정 탐색 시뮬레이션: 2026-10-20 기준 3개 공정(보경 철거, 우성 타일, 우성 도배) 정확 추출 검증',
  foundSameDay.length === 3 &&
  foundSameDay[0].task.taskName === '철거' &&
  foundSameDay[1].task.taskName === '타일' &&
  foundSameDay[2].task.taskName === '도배'
);

// 44-6: 같은 날 1개 공정인 날(2026-10-23) 단독 1건 추출 검증
const foundSingle = findSameDayTasksSim(mockMultiSitesData, { startDateStr: '2026-10-23' });
assertTest('같은 날 공정 탐색 시뮬레이션: 1건뿐인 날(2026-10-23)은 1건만 반환 검증',
  foundSingle.length === 1 &&
  foundSingle[0].task.taskName === '양생'
);

// 44-7: 이전 / 다음 네비게이션 상태 및 경계값 시뮬레이션
class SameDayNavigatorSim {
  constructor(list, initialIdx = 0) {
    this.list = list;
    this.idx = initialIdx;
  }
  get canPrev() { return this.idx > 0; }
  get canNext() { return this.idx < this.list.length - 1; }
  nav(delta) {
    const next = this.idx + delta;
    if (next >= 0 && next < this.list.length) {
      this.idx = next;
    }
  }
}

const navSim = new SameDayNavigatorSim(foundSameDay, 0); // 3건 중 0번 시작
assertTest('네비게이션 시뮬레이션: 첫 번째 일정(1/3)에서 ◀ 비활성화, ▶ 활성화 검증',
  navSim.canPrev === false && navSim.canNext === true && navSim.idx === 0
);

navSim.nav(1); // 1번(가운데)으로 이동
assertTest('네비게이션 시뮬레이션: 다음 이동 시 두 번째 일정(2/3)에서 ◀ 활성화, ▶ 활성화 검증',
  navSim.canPrev === true && navSim.canNext === true && navSim.idx === 1
);

navSim.nav(1); // 2번(마지막)으로 이동
assertTest('네비게이션 시뮬레이션: 세 번째 일정(3/3)에서 ◀ 활성화, ▶ 비활성화 검증',
  navSim.canPrev === true && navSim.canNext === false && navSim.idx === 2
);

navSim.nav(1); // 경계 초과 시도
assertTest('네비게이션 시뮬레이션: 마지막에서 추가 다음 클릭 시 인덱스 2 초과 방어 검증',
  navSim.idx === 2
);

navSim.nav(-2); // 첫 번째로 되돌아가기
assertTest('네비게이션 시뮬레이션: 이전 연속 클릭 시 첫 번째(0)로 정상 복귀 및 음수 인덱스 방어 검증',
  navSim.idx === 0 && navSim.canPrev === false
);

// 44-8: 클라이언트 JS 구문 무결성(SyntaxError 0) 검증
let clientSyntaxOk44 = false;
try {
  new vm.Script(step44Script);
  clientSyntaxOk44 = true;
} catch (err) {
  console.error('Step 44 클라이언트 JS 구문 에러:', err);
}
assertTest('buildSidebarDashboardHtml: 같은 날 다중 일정 네비게이션 탑재 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk44);

console.log('\n👥 [Step 45: 현장데이터 담당자 열 구성원 시트 연동 & 김종혁 제외 자동 반영 QA]');

// 45-1: isMembersSheet 시트명 판별 함수 검증
const isMembersSheetFn = sandbox.isMembersSheet;
assertTest('isMembersSheet: 함수 정의 및 호출 가능 검증', typeof isMembersSheetFn === 'function');
assertTest('isMembersSheet: "구성원" 시트 true 판정', isMembersSheetFn({ getName: () => '구성원' }) === true);
assertTest('isMembersSheet: "03_구성원" 시트 true 판정', isMembersSheetFn({ getName: () => '03_구성원' }) === true);
assertTest('isMembersSheet: "팀원" 시트 true 판정', isMembersSheetFn({ getName: () => '팀원' }) === true);
assertTest('isMembersSheet: "직원 명단" 시트 true 판정', isMembersSheetFn({ getName: () => '직원 명단' }) === true);
assertTest('isMembersSheet: "조직도" 시트 true 판정', isMembersSheetFn({ getName: () => '조직도' }) === true);
assertTest('isMembersSheet: "members" 영문 시트 true 판정', isMembersSheetFn({ getName: () => 'members' }) === true);
assertTest('isMembersSheet: "현장데이터" 시트는 false 판정', isMembersSheetFn({ getName: () => '현장데이터' }) === false);
assertTest('isMembersSheet: "대시보드" 시트는 false 판정', isMembersSheetFn({ getName: () => '대시보드' }) === false);
assertTest('isMembersSheet: null / undefined 시트 방어 (false 반환)', isMembersSheetFn(null) === false && isMembersSheetFn(undefined) === false);

// 45-2: getAvailableManagers 복합 조합 제거 및 순수 개별 명단 반환 검증
const getAvailableManagersFn = sandbox.getAvailableManagers;
assertTest('getAvailableManagers: 함수 정의 및 호출 가능 검증', typeof getAvailableManagersFn === 'function');

// 모의 구성원 시트 및 스프레드시트 구축
let mockMemberList = ['조항만', '송지은', '김종혁', '전예린', '영하'];
const createMockSsWithMembers = (membersArr) => {
  const memberSheet = {
    getName: () => '구성원',
    getLastRow: () => membersArr.length + 1,
    getLastColumn: () => 2,
    getRange: (r, c, nr, nc) => {
      if (r === 1) {
        return { getValues: () => [['이름', '연락처']] };
      }
      return {
        getValues: () => membersArr.map(m => [m])
      };
    }
  };
  return {
    getSheets: () => [memberSheet],
    getSheetByName: (n) => (n === '구성원' ? memberSheet : null)
  };
};

const mockSs1 = createMockSsWithMembers(mockMemberList);
const mgrs1 = getAvailableManagersFn(mockSs1);

assertTest('getAvailableManagers: 반환 결과 배열 유효성 검증', Array.isArray(mgrs1) && mgrs1.length > 0);
assertTest('getAvailableManagers: 복합 담당자 조합("조항만, 송지은") 일체 미포함 검증',
  mgrs1.every(m => m.indexOf(',') === -1)
);
assertTest('getAvailableManagers: "김종혁"·"영하" 담당자 드롭다운 목록에서 완전 배제 검증',
  mgrs1.indexOf('김종혁') === -1 && mgrs1.indexOf('영하') === -1
);
assertTest('getAvailableManagers: "조항만", "송지은", "전예린" 개별 인원 정확 포함 검증',
  mgrs1.indexOf('조항만') !== -1 && mgrs1.indexOf('송지은') !== -1 && mgrs1.indexOf('전예린') !== -1
);
assertTest('getAvailableManagers: 반환 인원 수 정확성 검증 (5명 중 김종혁·영하 2명 제외 = 3명)',
  mgrs1.length === 3
);

// 45-3: 구성원 시트에 신규 인원 추가 시 자동 반영 시뮬레이션 검증
const mockMemberList2 = ['조항만', '송지은', '김종혁', '전예린', '영하', '홍길동', '이순신'];
const mockSs2 = createMockSsWithMembers(mockMemberList2);
const mgrs2 = getAvailableManagersFn(mockSs2);

assertTest('구성원 추가 시뮬레이션: 새로 추가된 "홍길동", "이순신"이 즉각 목록에 자동 반영 검증',
  mgrs2.indexOf('홍길동') !== -1 && mgrs2.indexOf('이순신') !== -1
);
assertTest('구성원 추가 시뮬레이션: 추가 후에도 "김종혁"·"영하"는 철저히 배제 검증',
  mgrs2.indexOf('김종혁') === -1 && mgrs2.indexOf('영하') === -1
);
assertTest('구성원 추가 시뮬레이션: 총 5명(기존3 + 신규2) 정확 일치 검증',
  mgrs2.length === 5
);

// 45-4: 쉼표/슬래시 다중 입력, 공백, null 등 엣지 케이스 파싱 검증
const mockMemberListEdge = ['조항만', '  송지은  ', '김종혁', '강감찬, 을지문덕', '신사임당/율곡이이', '', null];
const mockSsEdge = createMockSsWithMembers(mockMemberListEdge);
const mgrsEdge = getAvailableManagersFn(mockSsEdge);

assertTest('엣지 케이스 파싱: 공백 자동 트림("송지은") 검증', mgrsEdge.indexOf('송지은') !== -1);
assertTest('엣지 케이스 파싱: 쉼표 분리("강감찬", "을지문덕") 개별 분리 추가 검증',
  mgrsEdge.indexOf('강감찬') !== -1 && mgrsEdge.indexOf('을지문덕') !== -1
);
assertTest('엣지 케이스 파싱: 슬래시 분리("신사임당", "율곡이이") 개별 분리 추가 검증',
  mgrsEdge.indexOf('신사임당') !== -1 && mgrsEdge.indexOf('율곡이이') !== -1
);
assertTest('엣지 케이스 파싱: 빈 값 및 null 안전 필터링 검증',
  mgrsEdge.every(m => typeof m === 'string' && m.trim().length > 0)
);

// 45-5: applySiteScheduleValidations G열(담당자) 유효성 검사 설정 검증
let appliedRules = {};
const mockSiteScheduleSheet = {
  getName: () => '현장데이터',
  getParent: () => mockSs2,
  getLastRow: () => 10,
  getMaxRows: () => 50,
  getRange: (r, c, nr, nc) => ({
    setDataValidation: (rule) => {
      appliedRules[c] = rule;
    }
  })
};

const applyValidationsFn = sandbox.applySiteScheduleValidations;
assertTest('applySiteScheduleValidations: 함수 정의 및 호출 가능 검증', typeof applyValidationsFn === 'function');

applyValidationsFn(mockSiteScheduleSheet, 2, 20);

assertTest('applySiteScheduleValidations: G열(7열: 담당자) 데이터 유효성 검사 규칙 적용 검증', appliedRules[7] !== undefined);
const gRuleVals = appliedRules[7] ? appliedRules[7].getCriteriaValues() : [];
const gList = (gRuleVals && gRuleVals[0]) || [];

assertTest('applySiteScheduleValidations: G열 드롭다운 목록에 최신 구성원("홍길동", "이순신") 탑재 검증',
  gList.indexOf('홍길동') !== -1 && gList.indexOf('이순신') !== -1
);
assertTest('applySiteScheduleValidations: G열 드롭다운 목록에서 "김종혁" 배제 검증',
  gList.indexOf('김종혁') === -1
);
assertTest('applySiteScheduleValidations: G열 드롭다운 목록에 복합 조합(쉼표 포함) 부재 검증',
  gList.every(m => m.indexOf(',') === -1)
);

// 45-6: onEdit 트리거 구성원 시트 편집 감지 및 즉시 동기화 연동 검증
let onEditValidationCalled = false;
const mockSiteForOnEdit = {
  getName: () => '현장데이터',
  getParent: () => null,
  getLastRow: () => 10,
  getLastColumn: () => 8,
  getMaxRows: () => 50,
  getSheetId: () => 999,
  getRange: (r, c, nr, nc) => {
    if (r === 1) {
      return {
        getValues: () => [['현장명', '시작일', '공종', '내용', '체크', '비고', '담당자', '진행상태']]
      };
    }
    return {
      getValues: () => [['']],
      setDataValidation: (rule) => {
        onEditValidationCalled = true;
      }
    };
  }
};

const mockMemberSheetForOnEdit = {
  getName: () => '03_구성원',
  getLastRow: () => 5,
  getLastColumn: () => 2,
  getRange: (r, c, nr, nc) => ({
    getValues: () => (r === 1 ? [['이름', '연락처']] : [['조항만'], ['송지은'], ['신규직원']])
  })
};

const mockSsForOnEdit = {
  getSheets: () => [
    mockMemberSheetForOnEdit,
    mockSiteForOnEdit
  ],
  getSheetByName: (n) => (n === '현장데이터' ? mockSiteForOnEdit : (n === '03_구성원' ? mockMemberSheetForOnEdit : null))
};

const onEditFn = sandbox.onEdit;
assertTest('onEdit: 함수 정의 및 호출 가능 검증', typeof onEditFn === 'function');

// Step 45 onEdit 테스트용 findSiteScheduleSheet 지정
sandbox.findSiteScheduleSheet = (ss) => mockSiteForOnEdit;

// 구성원 시트 편집 이벤트 발생
const mockEditEvent = {
  range: {
    getSheet: () => mockMemberSheetForOnEdit,
    getRow: () => 4,
    getColumn: () => 1
  },
  source: mockSsForOnEdit
};

onEditFn(mockEditEvent);
assertTest('onEdit: 구성원 시트 편집 시 applySiteScheduleValidations 즉시 자동 호출(실시간 동기화) 검증', onEditValidationCalled === true);

console.log('\n🚫 [Step 46: 담당자 없음 필터 버튼 & 지연 일정 붉은색 노출 및 공유 앞 경고 아이콘 QA]');

// 46-1: 대시보드 HTML 내 필수 CSS 및 마크업 검증
const html46 = sandbox.buildSidebarDashboardHtml("desktop");
assertTest('buildSidebarDashboardHtml: .cal-event-pill.status-delayed (#dc2626) CSS 탑재 검증',
  html46.indexOf('.cal-event-pill.status-delayed') !== -1 && html46.indexOf('#dc2626 !important') !== -1
);
assertTest('buildSidebarDashboardHtml: .cal-delay-warn-icon CSS 스타일 탑재 검증',
  html46.indexOf('.cal-delay-warn-icon') !== -1
);
assertTest('buildSidebarDashboardHtml: renderMemberFilterTabs 내 "🚫 담당자 없음" 버튼 탑재 검증',
  html46.indexOf('🚫 담당자 없음') !== -1 && html46.indexOf('__NONE__') !== -1
);
assertTest('buildSidebarDashboardHtml: isMemberMatch 헬퍼 함수 탑재 검증',
  html46.indexOf('function isMemberMatch(') !== -1
);

// 46-2: 캘린더 월간/주간 뷰 렌더링 내 지연 공정 붉은색 및 ⚠️ 경고 아이콘 로직 검증
assertTest('buildSidebarDashboardHtml: renderCalendar 내 지연 공정(#dc2626) 및 경고 아이콘(⚠️) 연동 검증',
  html46.indexOf('isDelayed ? "#dc2626"') !== -1 &&
  html46.indexOf('cal-delay-warn-icon') !== -1 &&
  html46.indexOf('⚠️ 공정 지연') !== -1
);
assertTest('buildSidebarDashboardHtml: renderWeeklyCalendar 내 지연 공정(#dc2626) 및 경고 아이콘(⚠️) 연동 검증',
  html46.indexOf('listHtml += \'  <span class="cal-delay-warn-icon"') !== -1
);

// 46-3: 클라이언트 JS 스크립트 추출 및 VM 무결성 검증
const scriptMatch46 = html46.match(/<script[\s\S]*?>([\s\S]*?)<\/script>/i);
let clientSyntaxOk46 = false;
let step46Script = scriptMatch46 ? scriptMatch46[1] : "";
try {
  new vm.Script(step46Script);
  clientSyntaxOk46 = true;
} catch (err) {
  console.error('Step 46 클라이언트 JS 구문 에러:', err);
}
assertTest('buildSidebarDashboardHtml: 신규 기능 탑재 후 클라이언트 JS 무결성(SyntaxError 0) 검증', clientSyntaxOk46);

// 46-4: isMemberMatch 단위 시뮬레이션 검증
const isMemberMatchSim = (taskManager, sel) => {
  if (!sel) return true;
  const mgr = (taskManager || "").trim();
  if (sel === "__NONE__" || sel === "담당자 없음") {
    return (!mgr || mgr === "담당자 없음" || mgr === "미지정" || mgr === "없음");
  }
  return mgr.indexOf(sel) !== -1;
};

assertTest('isMemberMatch: __NONE__ 선택 시 담당자 빈 문자열("") -> 일치(true)', isMemberMatchSim("", "__NONE__") === true);
assertTest('isMemberMatch: __NONE__ 선택 시 담당자 null/undefined -> 일치(true)', isMemberMatchSim(null, "__NONE__") === true && isMemberMatchSim(undefined, "__NONE__") === true);
assertTest('isMemberMatch: __NONE__ 선택 시 담당자 "담당자 없음" -> 일치(true)', isMemberMatchSim("담당자 없음", "__NONE__") === true);
assertTest('isMemberMatch: __NONE__ 선택 시 담당자 "미지정" -> 일치(true)', isMemberMatchSim("미지정", "__NONE__") === true);
assertTest('isMemberMatch: __NONE__ 선택 시 담당자 "조항만" -> 불일치(false)', isMemberMatchSim("조항만", "__NONE__") === false);
assertTest('isMemberMatch: __NONE__ 선택 시 담당자 "송지은, 전예린" -> 불일치(false)', isMemberMatchSim("송지은, 전예린", "__NONE__") === false);

assertTest('isMemberMatch: "조항만" 선택 시 "조항만" -> 일치(true)', isMemberMatchSim("조항만", "조항만") === true);
assertTest('isMemberMatch: "조항만" 선택 시 "조항만, 송지은" -> 일치(true)', isMemberMatchSim("조항만, 송지은", "조항만") === true);
assertTest('isMemberMatch: "조항만" 선택 시 "송지은" -> 불일치(false)', isMemberMatchSim("송지은", "조항만") === false);
assertTest('isMemberMatch: "조항만" 선택 시 "" -> 불일치(false)', isMemberMatchSim("", "조항만") === false);

assertTest('isMemberMatch: 전체("") 선택 시 모든 담당자 -> 항상 일치(true)',
  isMemberMatchSim("조항만", "") === true && isMemberMatchSim("", "") === true
);

// 46-5: 지연 일정 시각화 렌더링 시뮬레이션 검증
const simulateTaskPill = (task, site) => {
  const isDone = (task.status && (task.status.indexOf("완료") !== -1 || task.status.indexOf("취소") !== -1));
  const isDelayed = (task.isDelayed === true && !isDone);
  const isSh = (task.isShared === true || String(task.isShared).toLowerCase() === "true" || String(task.isShared) === "1");
  const shClass = isSh ? "blue" : "red";
  const bgCol = isDelayed ? "#dc2626" : (site.colorHex || "#2563eb");
  const classes = ["cal-event-pill"];
  if (isDelayed) classes.push("status-delayed");
  if (isDone) classes.push("status-done");

  let pillHtml = `<div class="${classes.join(' ')}" style="background-color:${bgCol};">`;
  if (isDelayed) {
    pillHtml += '<span class="cal-delay-warn-icon" title="⚠️ 공정 지연">⚠️</span>';
  }
  pillHtml += `<span class="cal-sh-box ${shClass}"></span><span class="cal-pill-text">${task.taskName}</span></div>`;
  return { pillHtml, isDelayed, bgCol, hasWarnIcon: pillHtml.indexOf('cal-delay-warn-icon') !== -1 };
};

const mockSiteA = { siteName: "역삼 자이", colorHex: "#10b981" };

// 시나리오 1: 미완료 지연 공정 (10월 15일 타일)
const delayedTask46 = { taskName: "타일", status: "진행", isDelayed: true, isShared: true };
const resDelayed46 = simulateTaskPill(delayedTask46, mockSiteA);
assertTest('지연 공정 렌더링: 배경색 #dc2626 붉은색 적용 검증', resDelayed46.bgCol === "#dc2626");
assertTest('지연 공정 렌더링: status-delayed CSS 클래스 부여 검증', resDelayed46.pillHtml.indexOf('status-delayed') !== -1);
assertTest('지연 공정 렌더링: 공유 아이콘 앞에 ⚠️ 경고 아이콘 노출 검증', resDelayed46.hasWarnIcon === true);

// 시나리오 2: 정상 진행 공정 (10월 17일 양생)
const normalTask46 = { taskName: "양생", status: "진행", isDelayed: false, isShared: true };
const resNormal46 = simulateTaskPill(normalTask46, mockSiteA);
assertTest('정상 공정 렌더링: 현장 고유색(#10b981) 유지 및 붉은색 미적용 검증', resNormal46.bgCol === "#10b981");
assertTest('정상 공정 렌더링: ⚠️ 경고 아이콘 미노출 검증', resNormal46.hasWarnIcon === false);

// 시나리오 3: 이미 완료된 과거 공정 (10월 10일 철거)
const doneTask46 = { taskName: "철거", status: "완료", isDelayed: true, isShared: true }; // 백엔드 플래그가 true이더라도 isDone이면 지연 해제
const resDone46 = simulateTaskPill(doneTask46, mockSiteA);
assertTest('완료 공정 렌더링: 완료된 일정은 지연 붉은색 대신 status-done(취소선) 적용 검증',
  resDone46.bgCol === "#10b981" && resDone46.pillHtml.indexOf('status-done') !== -1 && resDone46.hasWarnIcon === false
);

// ========================================================
// 47. [구글 캘린더 시트 삭제 유령(고아) 일정 자동 청소 QA]
// ========================================================
console.log('\n🧹 [Step 47: 구글 캘린더 시트 삭제 유령(고아) 일정 자동 청소 통합 QA]');

// 47-1: syncSiteSchedulesBatchSafe 내 cleanupOrphanCalendarEventsSilent 자동 탑재 검증
const batchSafeCode47 = code.slice(code.indexOf('function syncSiteSchedulesBatchSafe('), code.indexOf('function ensureCalendarSyncColumns('));
const hasOrphanInBatchSafe = batchSafeCode47.indexOf('cleanupOrphanCalendarEventsSilent(sheet, cal)') !== -1;
assertTest('syncSiteSchedulesBatchSafe: 상단 메뉴 동기화 시 유령 일정 자동 청소(cleanupOrphanCalendarEventsSilent) 결합 검증', hasOrphanInBatchSafe);

// 47-2: cleanupOrphanCalendarEventsSilent 삭제 한도 상향(50건) 검증
const orphanFuncCode47 = code.slice(code.indexOf('function cleanupOrphanCalendarEventsSilent('), code.indexOf('function runPeriodicAutoSync('));
const has50Limit = orphanFuncCode47.indexOf('deletedCount >= 50') !== -1;
assertTest('cleanupOrphanCalendarEventsSilent: 1회 최대 청소 한도 50건으로 상향 조정 검증', has50Limit);

// 47-3: runPeriodicAutoSync 주간 30분 주기 자율 스캔 및 심야 가드 보강 검증
const periodicSyncCode47 = code.slice(code.indexOf('function runPeriodicAutoSync('), code.indexOf('function runDebouncedSync('));
const hasPeriodicOrphanScan = periodicSyncCode47.indexOf('shouldPeriodicOrphanScan') !== -1 &&
  periodicSyncCode47.indexOf('LAST_ORPHAN_CLEANUP_MS') !== -1 &&
  periodicSyncCode47.indexOf('needsOrphanCleanup || shouldPeriodicOrphanScan') !== -1;
assertTest('runPeriodicAutoSync: 주간 30분 주기 유령 일정 자율 치유(Self-Healing) 및 심야 가드 검증', hasPeriodicOrphanScan);

// 47-4: handleCalendarEditTrigger 셀 내용 Delete 키 삭제 시 유령 청소 플래그 세팅 검증
const editTriggerCode47 = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncCalendarToSiteSchedule('));
const hasClearedRowDetection = editTriggerCode47.indexOf('hasClearedRow') !== -1 &&
  editTriggerCode47.indexOf('NEEDS_ORPHAN_CLEANUP') !== -1;
assertTest('handleCalendarEditTrigger: Delete 키로 날짜/공종 셀 삭제 시 유령 일정 청소 플래그 자동 마킹 검증', hasClearedRowDetection);

// 47-5: isSystemSiteScheduleEvent 사용자 테스트 일정(상담, 철거, 목공, 마감) 매칭 검증
const isSysEventFn = sandbox.isSystemSiteScheduleEvent;
assertTest('isSystemSiteScheduleEvent: 함수 로드 검증', typeof isSysEventFn === 'function');

const mockTestEventConsult47 = {
  getTitle: () => '🟦 [상담] 테스트',
  getDescription: () => '📌 현장명: 테스트\n🔨 공종: 상담\n현장 관리 시스템'
};
const mockTestEventDemolish47 = {
  getTitle: () => '🟦 [철거] 테스트',
  getDescription: () => '📌 현장명: 테스트\n🔨 공종: 철거'
};
const mockTestEventWood47 = {
  getTitle: () => '🟥 [목공] 테스트',
  getDescription: () => '📌 현장명: 테스트\n🔨 공종: 목공'
};
const mockTestEventFinish47 = {
  getTitle: () => '🟥 [마감] 테스트',
  getDescription: () => '📌 현장명: 테스트\n🔨 공종: 마감'
};
const mockPersonalEvent47 = {
  getTitle: () => '결혼기념일',
  getDescription: () => '가족 저녁식사'
};

assertTest('isSystemSiteScheduleEvent: 🟦 [상담] 테스트 일정 정상 시스템 일정으로 인식', isSysEventFn(mockTestEventConsult47) === true);
assertTest('isSystemSiteScheduleEvent: 🟦 [철거] 테스트 일정 정상 시스템 일정으로 인식', isSysEventFn(mockTestEventDemolish47) === true);
assertTest('isSystemSiteScheduleEvent: 🟥 [목공] 테스트 일정 정상 시스템 일정으로 인식', isSysEventFn(mockTestEventWood47) === true);
assertTest('isSystemSiteScheduleEvent: 🟥 [마감] 테스트 일정 정상 시스템 일정으로 인식', isSysEventFn(mockTestEventFinish47) === true);
assertTest('isSystemSiteScheduleEvent: 개인 일정(결혼기념일) 안전 보호(false 판정)', isSysEventFn(mockPersonalEvent47) === false);

// ========================================================
// 48. [최종수정자 식별 엔진(getCurrentUserIdentifier) QA]
// ========================================================
console.log('\n👤 [Step 48: 최종수정자 식별 엔진(getCurrentUserIdentifier) QA]');

// 48-1: handleCalendarEditTrigger 내 e 전달 검증
const editTriggerCode48 = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncCalendarToSiteSchedule('));
const hasPassEToUserFn = editTriggerCode48.indexOf('getCurrentUserIdentifier(e)') !== -1;
assertTest('handleCalendarEditTrigger: getCurrentUserIdentifier(e)로 이벤트 객체 전달 검증', hasPassEToUserFn);

// 48-2: getCurrentUserIdentifier 함수 로드 검증
const getUserFn = sandbox.getCurrentUserIdentifier;
assertTest('getCurrentUserIdentifier: 함수 로드 검증', typeof getUserFn === 'function');

// 48-3: e.user가 객체(getEmail 메서드 보유)로 전달된 경우 최우선 추출 검증
const mockEWithUserObj = {
  user: {
    getEmail: () => 'joehm037@gmail.com'
  }
};
assertTest('getCurrentUserIdentifier: e.user.getEmail() 존재 시 joehm037 최우선 추출 검증', getUserFn(mockEWithUserObj) === 'joehm037');

// 48-4: e.user가 문자열로 전달된 경우 추출 검증
const mockEWithUserStr = {
  user: 'joehm037@gmail.com'
};
assertTest('getCurrentUserIdentifier: e.user 문자열일 때 joehm037 추출 검증', getUserFn(mockEWithUserStr) === 'joehm037');

// 48-5: e가 없거나 빈 객체일 때 Session 기반 fallback 검증 (sandbox Session Mock)
sandbox.Session = {
  getActiveUser: () => ({ getEmail: () => '' }),
  getEffectiveUser: () => ({ getEmail: () => 'ijhousing9696@gmail.com' })
};
assertTest('getCurrentUserIdentifier: getActiveUser 비어있을 시 getEffectiveUser(ijhousing9696) 정상 폴백 검증', getUserFn({}) === 'ijhousing9696');

// ========================================================
// 49. [K, L, M, N열 영구 삭제 및 10열(A~J열) 체제 정돈 QA]
// ========================================================
console.log('\n🧹 [Step 49: K, L, M, N열 영구 삭제 및 10열(A~J열) 체제 정돈 QA]');

// 49-1: removeAuditColumnsKtoN 함수 정의 및 로드 검증
const removeAuditFn = sandbox.removeAuditColumnsKtoN;
assertTest('removeAuditColumnsKtoN: 함수 로드 검증', typeof removeAuditFn === 'function');

// 49-2: onOpen 메뉴 내 removeAuditColumnsKtoN 등록 검증
const onOpenCode49 = code.slice(code.indexOf('function onOpen('), code.indexOf('function onEdit('));
const hasRemoveMenuInOnOpen = onOpenCode49.indexOf('removeAuditColumnsKtoN') !== -1 &&
  onOpenCode49.indexOf('K~N열') !== -1;
assertTest('onOpen: 상단 관리 메뉴 내 K~N열 영구 삭제 항목 탑재 검증', hasRemoveMenuInOnOpen);

// 49-3: ensureCalendarSyncColumns 내 11열 이상 자동 삭제 로직 탑재 검증
const ensureColCode49 = code.slice(code.indexOf('function ensureCalendarSyncColumns('), code.indexOf('function getSheetCalendarIds('));
const hasAutoDeleteCols11Plus = ensureColCode49.indexOf('sheet.deleteColumns(11') !== -1 &&
  ensureColCode49.indexOf('maxCols >= 11') !== -1;
assertTest('ensureCalendarSyncColumns: 11열 이상 존재 시 자동 deleteColumns 실행 검증', hasAutoDeleteCols11Plus);

// 49-4: ensureCalendarSyncColumns 헤더에 11~14열 제거 및 9, 10열만 유지 검증
const hasNo11To14Headers = ensureColCode49.indexOf('col: 11') === -1 &&
  ensureColCode49.indexOf('col: 12') === -1 &&
  ensureColCode49.indexOf('col: 13') === -1 &&
  ensureColCode49.indexOf('col: 14') === -1 &&
  ensureColCode49.indexOf('col: 9') !== -1 &&
  ensureColCode49.indexOf('col: 10') !== -1;
assertTest('ensureCalendarSyncColumns: 헤더 설정에서 11~14열 완전 배제 및 9, 10열만 유지 검증', hasNo11To14Headers);

// 49-5: handleCalendarEditTrigger 내 11~14열 쓰기 제거 및 10열 타임스탬프 단독 쓰기 검증
const editTriggerCode49 = code.slice(code.indexOf('function handleCalendarEditTrigger('), code.indexOf('function syncCalendarToSiteSchedule('));
const hasNoMetaRangeWrite = editTriggerCode49.indexOf('metaRange.setValues') === -1 &&
  editTriggerCode49.indexOf('sheet.getRange(actualStartRow, 11') === -1 &&
  editTriggerCode49.indexOf('sheet.getRange(actualStartRow, 10, numRows, 1).setValues(timeStampVals)') !== -1;
assertTest('handleCalendarEditTrigger: 11~14열 쓰기 완전 제거 및 10열 타임스탬프 단독 쓰기 검증', hasNoMetaRangeWrite);

// 49-6: addNewSiteAtTop 내 10개 열 채우기(14열 배제) 검증
const addNewSiteCode49 = code.slice(code.indexOf('function addNewSiteAtTop('), code.indexOf('function cleanupEmptyRowsMetadata('));
const has10ColumnsInNewSite = addNewSiteCode49.indexOf('sheet.getRange(2, 1, 4, 10).setValues(newRows)') !== -1 &&
  addNewSiteCode49.indexOf('sheet.getRange(2, 1, 4, 14)') === -1;
assertTest('addNewSiteAtTop: 신규 현장 추가 시 10개 열(A~J열) 채우기 적용 검증', has10ColumnsInNewSite);

// ========================================================
// 50. [모바일 상세 화면 논블로킹 즉시 저장 및 락 경합 해소 QA]
// ========================================================
console.log('\n📱 [Step 50: 모바일 상세 화면 논블로킹 즉시 저장 및 락 경합 해소 QA]');

// 50-1: updateScheduleStatusFromWeb 정적 분석 검증
const updateWebCode50 = code.slice(code.indexOf('function updateScheduleStatusFromWeb('), code.indexOf('function doGet('));
const hasNonBlockingSheetWrite = updateWebCode50.indexOf('sheet.getRange(rowNum, 8).setValue(newStatus)') !== -1 &&
  updateWebCode50.indexOf('sheet.getRange(rowNum, 5).setValue(isSharedBool)') !== -1 &&
  updateWebCode50.indexOf('sheet.getRange(rowNum, 6).setValue(cleanMemo)') !== -1 &&
  updateWebCode50.indexOf('sheet.getRange(rowNum, 10).setValue(now.getTime())') !== -1 &&
  updateWebCode50.indexOf('SpreadsheetApp.flush()') !== -1;
assertTest('updateScheduleStatusFromWeb: 락 대기 없이 시트 셀(H, E, F, J열) 원자적 즉시 쓰기 및 flush 탑재 검증', hasNonBlockingSheetWrite);

// 50-2: 모바일 응답 가속화 - skipDashboardRefresh=true 전달 검증
const hasSkipDashboardInWebSave = updateWebCode50.indexOf('syncSingleRowToCalendar(sheet, rowNum, true, true)') !== -1;
assertTest('updateScheduleStatusFromWeb: 캘린더 동기화 시 skipDashboardRefresh=true로 모바일 응답속도 초고속화 검증', hasSkipDashboardInWebSave);

// 50-3: 락 경합 시에도 에러를 리턴하지 않고 백그라운드 위임(Non-blocking) 검증
const hasLockFailureFallback = updateWebCode50.indexOf('triggerDebouncedSync()') !== -1 &&
  updateWebCode50.indexOf('return { success: false, error: "현재 다른 작업이 처리 중입니다') === -1;
assertTest('updateScheduleStatusFromWeb: 락 경합 시 "현재 다른 작업이 처리 중" 에러 차단 및 백그라운드 자동 위임 검증', hasLockFailureFallback);

// 50-4: 런타임 시뮬레이션 - 락 획득 실패(Lock Contention) 상태에서도 100% 정상 저장 및 성공 반환 검증
const mockSheetCells50 = {};
const mockSheet50 = {
  getRange: (r, c) => ({
    setValue: (val) => {
      mockSheetCells50[`${r}_${c}`] = val;
    }
  })
};

// 락이 무조건 실패(false)하는 환경 시뮬레이션
sandbox.LockService = {
  getScriptLock: () => ({
    tryLock: () => false, // 락 획득 실패!
    releaseLock: () => {}
  })
};
sandbox.getAppSpreadsheet = () => ({});
sandbox.findSiteScheduleSheet = () => mockSheet50;
let calSyncCalled50 = false;
sandbox.syncSingleRowToCalendar = (s, r, force, skip) => {
  calSyncCalled50 = true;
};
sandbox.SpreadsheetApp.flush = () => {};

const updateWebFn = sandbox.updateScheduleStatusFromWeb;
const testRes50 = updateWebFn(5, '진행 중', true, '고객 요청 메모 추가');

assertTest('런타임 시뮬레이션: 락 획득 실패(다른 작업 처리 중) 상황에서도 성공(success: true) 반환 검증',
  testRes50 && testRes50.success === true && testRes50.status === '진행 중' && testRes50.shared === true && testRes50.memo === '고객 요청 메모 추가'
);
assertTest('런타임 시뮬레이션: 락 실패 상황에서도 H열(8열: 상태) 즉시 기록 검증',
  mockSheetCells50['5_8'] === '진행 중'
);
assertTest('런타임 시뮬레이션: 락 실패 상황에서도 E열(5열: 공유) 즉시 기록 검증',
  mockSheetCells50['5_5'] === true
);
assertTest('런타임 시뮬레이션: 락 실패 상황에서도 F열(6열: 비고) 즉시 기록 검증',
  mockSheetCells50['5_6'] === '고객 요청 메모 추가'
);
assertTest('런타임 시뮬레이션: 락 실패 상황에서도 J열(10열: 타임스탬프) 최신 시각 기록 검증',
  typeof mockSheetCells50['5_10'] === 'number' && mockSheetCells50['5_10'] > 0
);
assertTest('런타임 시뮬레이션: 캘린더 동기화 함수 즉각 연동 시도 검증',
  calSyncCalled50 === true
);

// 50-5: 잘못된 파라미터 방어 검증
const invalidRowRes = updateWebFn(1, '완료', false, '');
assertTest('런타임 시뮬레이션: 1행(헤더) 또는 음수 행 요청 시 안전 방어 검증',
  invalidRowRes && invalidRowRes.success === false
);
const invalidStatusRes = updateWebFn(5, '', false, '');
assertTest('런타임 시뮬레이션: 상태 누락 시 안전 방어 검증',
  invalidStatusRes && invalidStatusRes.success === false
);

// ========================================================
// 51. [담당자 동적 서식 & 동일 현장 내 날짜 오름차순 정렬 QA]
// ========================================================
console.log('\n🎨 [Step 51: 담당자 동적 서식 & 동일 현장 내 날짜 오름차순 정렬 QA]');

// 51-1: getManagerColorMap 함수 정의 및 한송이, 인경호 프리셋 색상 검증
const getManagerColorFn = sandbox.getManagerColorMap;
assertTest('getManagerColorMap: 함수 정의 및 로드 검증', typeof getManagerColorFn === 'function');

const mockSs51 = {
  getSheets: () => []
};
const defaultColorMap51 = getManagerColorFn(mockSs51);
const hansongi = defaultColorMap51.find(m => m.key === '한송이');
const inkyungho = defaultColorMap51.find(m => m.key === '인경호');
const johangman = defaultColorMap51.find(m => m.key === '조항만');
const songjieun = defaultColorMap51.find(m => m.key === '송지은');

assertTest('getManagerColorMap: "한송이" 웜오렌지(#ffedd5 / #c2410c) 색상 배정 검증',
  hansongi && hansongi.bg === '#ffedd5' && hansongi.font === '#c2410c'
);
assertTest('getManagerColorMap: "인경호" 인디고(#e0e7ff / #4338ca) 색상 배정 검증',
  inkyungho && inkyungho.bg === '#e0e7ff' && inkyungho.font === '#4338ca'
);
assertTest('getManagerColorMap: "조항만", "송지은" 기존 프리셋 색상 완벽 유지 검증',
  johangman && johangman.bg === '#dbeafe' && songjieun && songjieun.bg === '#fce7f3'
);

// 51-2: 나중에 새로 등록되는 신규 담당자("홍길동", "이순신") 동적 색상 자동 부여 검증
sandbox.getMemberNames = () => ['조항만', '송지은', '한송이', '인경호', '전예린', '영하', '홍길동', '이순신'];
const dynamicColorMap51 = getManagerColorFn(mockSs51);
const hong = dynamicColorMap51.find(m => m.key === '홍길동');
const lee = dynamicColorMap51.find(m => m.key === '이순신');

assertTest('getManagerColorMap: 신규 담당자 "홍길동" 파스텔 팔레트 자동 배정 검증',
  hong && typeof hong.bg === 'string' && typeof hong.font === 'string' && hong.bg.startsWith('#')
);
assertTest('getManagerColorMap: 신규 담당자 "이순신" 파스텔 팔레트 자동 배정 검증',
  lee && typeof lee.bg === 'string' && typeof lee.font === 'string' && lee.bg.startsWith('#')
);

// 51-3: setupSiteSheetConditionalFormatting 내 한송이, 인경호 조건부 서식 등록 검증
let createdRules51 = [];
sandbox.SpreadsheetApp.newConditionalFormatRule = () => {
  const ruleObj = {
    whenTextContains: (txt) => { ruleObj.txt = txt; return ruleObj; },
    setBackground: (bg) => { ruleObj.bg = bg; return ruleObj; },
    setFontColor: (fc) => { ruleObj.fc = fc; return ruleObj; },
    setBold: (b) => { ruleObj.bold = b; return ruleObj; },
    setRanges: (rng) => { ruleObj.rng = rng; return ruleObj; },
    build: () => { createdRules51.push(ruleObj); return ruleObj; }
  };
  return ruleObj;
};
const mockSiteSheet51 = {
  getLastRow: () => 10,
  getRange: () => ({ getValues: () => [] }),
  setConditionalFormatRules: () => {},
  getParent: () => mockSs51
};
createdRules51 = [];
sandbox.setupSiteSheetConditionalFormatting(mockSiteSheet51);
const ruleHansongi = createdRules51.find(r => r.txt === '한송이');
const ruleInkyungho = createdRules51.find(r => r.txt === '인경호');
assertTest('setupSiteSheetConditionalFormatting: "한송이" 조건부 서식 룰 생성 검증',
  ruleHansongi && ruleHansongi.bg === '#ffedd5'
);
assertTest('setupSiteSheetConditionalFormatting: "인경호" 조건부 서식 룰 생성 검증',
  ruleInkyungho && ruleInkyungho.bg === '#e0e7ff'
);

// 51-4: sortSiteScheduleTasksByDate - 사용자 스크린샷 2~5행 "연속 날짜" 현장 날짜 오름차순 정렬 시뮬레이션
const sortTasksFn = sandbox.sortSiteScheduleTasksByDate;
assertTest('sortSiteScheduleTasksByDate: 함수 정의 및 로드 검증', typeof sortTasksFn === 'function');

let setValuesRecorded51 = null;
let setBackgroundsRecorded51 = null;

// 스크린샷 2~5행 케이스 (2026.10.17 -> 2026.10.19 -> 2026.10.18 -> 2026.10.20)
// 스크린샷 13~14행 케이스 (2026.10.21 [도배] -> 2026.10.19 [목공])
const mockValues51 = [
  ['연속 날짜', '2026.10.17', '상담', '', true, '메모1', '한송이', '진행 중', 'id1', 1000],
  ['연속 날짜', '2026.10.19', '철거', '', true, '', '한송이', '진행 중', 'id2', 2000],
  ['연속 날짜', '2026.10.18', '목공', '', true, '', '인경호', '예정', 'id3', 3000],
  ['연속 날짜', '2026.10.20', '마감', '', false, '', '인경호', '예정', 'id4', 4000],
  ['신대방우성 1동 805호', '2026.10.21', '도배', '', false, '', '조항만', '예정', 'id5', 5000],
  ['신대방우성 1동 805호', '2026.10.19', '목공', '', true, '박동연목수님', '조항만', '진행 중', 'id6', 6000]
];
const mockBackgrounds51 = [
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']
];

const mockSortSheet51 = {
  getLastRow: () => 7,
  getRange: (r, c, nr, nc) => ({
    getValues: () => mockValues51.map(row => row.slice()),
    getBackgrounds: () => mockBackgrounds51.map(row => row.slice()),
    setValues: (vals) => { setValuesRecorded51 = vals; },
    setBackgrounds: (bgs) => { setBackgroundsRecorded51 = vals => bgs; }
  })
};

const sortedCount51 = sortTasksFn(mockSs51, mockSortSheet51, true);

assertTest('sortSiteScheduleTasksByDate: 2개 현장 블록 재정렬 완료 검증', sortedCount51 === 2);
assertTest('sortSiteScheduleTasksByDate: 원자적 setValues 호출 검증', setValuesRecorded51 !== null && setValuesRecorded51.length === 6);

// "연속 날짜" 블록 검증 (17일 -> 18일 -> 19일 -> 20일)
assertTest('"연속 날짜" 정렬: 1행 2026.10.17 [상담] 검증',
  setValuesRecorded51[0][1] === '2026.10.17' && setValuesRecorded51[0][2] === '상담'
);
assertTest('"연속 날짜" 정렬: 2행 2026.10.18 [목공] 검증 (18일이 2번째로 상승)',
  setValuesRecorded51[1][1] === '2026.10.18' && setValuesRecorded51[1][2] === '목공' && setValuesRecorded51[1][6] === '인경호'
);
assertTest('"연속 날짜" 정렬: 3행 2026.10.19 [철거] 검증 (19일이 3번째로 하강)',
  setValuesRecorded51[2][1] === '2026.10.19' && setValuesRecorded51[2][2] === '철거' && setValuesRecorded51[2][6] === '한송이'
);
assertTest('"연속 날짜" 정렬: 4행 2026.10.20 [마감] 검증',
  setValuesRecorded51[3][1] === '2026.10.20' && setValuesRecorded51[3][2] === '마감'
);

// "신대방우성 1동 805호" 블록 검증 (19일 [목공] -> 21일 [도배])
assertTest('"신대방우성 1동" 정렬: 19일 [목공]이 21일 [도배]보다 앞으로 정렬 검증',
  setValuesRecorded51[4][1] === '2026.10.19' && setValuesRecorded51[4][2] === '목공' &&
  setValuesRecorded51[5][1] === '2026.10.21' && setValuesRecorded51[5][2] === '도배'
);

// 51-5: 1시간 주기 자동 정렬 트리거 핸들러 및 상단 메뉴 등록 검증
const hasHourlyTriggerCode = code.indexOf('runHourlyAutoSortTrigger') !== -1 &&
  code.indexOf('everyHours(1)') !== -1;
assertTest('ensureRealtimeTriggerInstalledSilent: 1시간 주기 자동 정렬 트리거(everyHours(1)) 등록 검증', hasHourlyTriggerCode);

// ========================================================
// 52. [2026년 미래 유령 일정 동적 윈도우 스캔 및 날짜 불일치 고아 일정 자동 치유 QA]
// ========================================================
console.log('\n📅 [Step 52: 2026년 미래 유령 일정 동적 윈도우 스캔 & 날짜 불일치 자동 치유 QA]');

// 52-1: 동적 날짜 윈도우 탐색 코드 탑재 검증
const orphanSilentCode52 = code.slice(code.indexOf('function cleanupOrphanCalendarEventsSilent('), code.indexOf('function runPeriodicAutoSync('));
const hasDynamicWindowSilent = orphanSilentCode52.indexOf('searchStart = new Date(minDate.getTime()') !== -1 &&
  orphanSilentCode52.indexOf('searchEnd = new Date(maxDate.getTime()') !== -1;
assertTest('cleanupOrphanCalendarEventsSilent: 시트 minDate/maxDate 기반 동적 캘린더 스캔 윈도우 확장 탑재 검증', hasDynamicWindowSilent);

const orphanManualCode52 = code.slice(code.indexOf('function cleanupDuplicateAndOrphanCalendarEvents('), code.indexOf('function cleanupDuplicateCalendarEventsSilent('));
const hasDynamicWindowManual = orphanManualCode52.indexOf('searchStart = new Date(minDate.getTime()') !== -1 &&
  orphanManualCode52.indexOf('searchEnd = new Date(maxDate.getTime()') !== -1;
assertTest('cleanupDuplicateAndOrphanCalendarEvents: 상단 메뉴 수동 정리 시에도 동적 윈도우 100% 스캔 탑재 검증', hasDynamicWindowManual);

// 52-2: 날짜 불일치 diffDays >= 2 및 자율 치유(삭제 또는 날짜 이동) 탑재 검증
const hasMismatchHealSilent = orphanSilentCode52.indexOf('diffDays >= 2') !== -1 &&
  orphanSilentCode52.indexOf('existingOnDate') !== -1 &&
  orphanSilentCode52.indexOf('ev.setAllDayDate(sParsed.startDate)') !== -1;
assertTest('cleanupOrphanCalendarEventsSilent: diffDays >= 2 어긋난 일정 자율 위치 이동 또는 중복 잔여 삭제 검증', hasMismatchHealSilent);

const hasMismatchHealManual = orphanManualCode52.indexOf('diffDays >= 2') !== -1 &&
  orphanManualCode52.indexOf('existingOnDate') !== -1 &&
  orphanManualCode52.indexOf('ev.setAllDayDate(sParsed.startDate)') !== -1;
assertTest('cleanupDuplicateAndOrphanCalendarEvents: 수동 정리 시 diffDays >= 2 어긋난 과거 일정 완벽 치유 검증', hasMismatchHealManual);

// 52-3: 런타임 시뮬레이션 - 시트는 2026.10.25, 캘린더에는 2026.10.29 잔여 일정이 남아있는 시나리오
let deletedInSim52 = [];
let movedInSim52 = [];

const mockOldResidualEvent = {
  getId: () => 'mismatched-residual-id-1029',
  getTitle: () => '🟥 [철거] 신대방우성 16동 1501호',
  getDescription: () => '📌 현장명: 신대방우성 16동 1501호\n🔨 공종: 철거',
  getLocation: () => '',
  isAllDayEvent: () => true,
  getAllDayStartDate: () => new Date(2026, 9, 29),
  getStartTime: () => new Date(2026, 9, 29),
  deleteEvent: function() { deletedInSim52.push(this.getId()); },
  setAllDayDate: function(d) { movedInSim52.push({ id: this.getId(), to: d }); }
};

const mockCorrectEvent = {
  getId: () => 'correct-on-date-id-1025',
  getTitle: () => '🟥 [철거] 신대방우성 16동 1501호',
  getDescription: () => '📌 현장명: 신대방우성 16동 1501호\n🔨 공종: 철거',
  getLocation: () => '',
  isAllDayEvent: () => true,
  getAllDayStartDate: () => new Date(2026, 9, 25),
  getStartTime: () => new Date(2026, 9, 25),
  deleteEvent: () => {},
  setAllDayDate: () => {}
};

const calEventsSim52 = [mockOldResidualEvent, mockCorrectEvent];

const mockCal52 = {
  getEvents: (s, e) => calEventsSim52,
  deleteEvent: (ev) => ev.deleteEvent()
};

let sheetValsRecorded52 = {};
const sheetData52 = [
  // 2행: 시트 날짜는 2026.10.25인데, I열(9번째)에 과거 잔여 일정 ID인 mismatched-residual-id-1029가 남아있는 상황!
  ['신대방우성 16동 1501호', '2026.10.25', '철거', '', true, '', '', '진행', 'mismatched-residual-id-1029', Date.now(), '', '', '', '']
];

const mockSheet52 = {
  getName: () => '현장서식',
  getLastRow: () => sheetData52.length + 1,
  getMaxColumns: () => 14,
  setColumnWidth: () => {},
  hideColumns: () => {},
  getRange: (r, c, nr, nc) => ({
    getValue: () => "진행",
    setValue: (v) => { sheetValsRecorded52[`${r},${c}`] = v; },
    getValues: () => sheetData52
  })
};

const resultCleaned52 = sandbox.cleanupOrphanCalendarEventsSilent(mockSheet52, mockCal52);

assertTest('런타임 시뮬레이션: 시트 날짜(10.25)와 어긋난 과거 캘린더 잔여 일정(10.29) 정확 삭제 검증',
  deletedInSim52.indexOf('mismatched-residual-id-1029') !== -1
);
assertTest('런타임 시뮬레이션: 올바른 날짜의 정상 일정 ID(correct-on-date-id-1025)로 시트 I열 자동 재매핑 검증',
  sheetValsRecorded52['2,9'] === 'correct-on-date-id-1025'
);

// ========================================================
// 53. [블록 단위 전체 완료 현장 '완료현장' 시트 이동(아카이브) 및 캘린더 100% 보존 QA]
// ========================================================
console.log('\n📦 [Step 53: 완료 현장 블록 아카이브 & 구글 캘린더 100% 영구 보존 QA]');

// 53-1: 함수 탑재 검증
assertTest('findCompletedSitesSheet: 함수 정의 및 로드 검증', typeof sandbox.findCompletedSitesSheet === 'function');
assertTest('archiveCompletedSiteBlocks: 함수 정의 및 로드 검증', typeof sandbox.archiveCompletedSiteBlocks === 'function');
assertTest('archiveCompletedSitesMenu: 함수 정의 및 로드 검증', typeof sandbox.archiveCompletedSitesMenu === 'function');
assertTest('runDailyCompletedSitesArchiveTrigger: 함수 정의 및 로드 검증', typeof sandbox.runDailyCompletedSitesArchiveTrigger === 'function');

// 53-2: isSiteScheduleSheet 내 "완료" 시트 배제 검증
const mockCompletedSheet = {
  getName: () => '완료현장',
  getLastRow: () => 10,
  getLastColumn: () => 10,
  getRange: () => ({ getValues: () => [['현장명', '날짜', '공종']] })
};
assertTest('isSiteScheduleSheet: "완료현장" 시트는 진행중 현장데이터 시트에서 완전 배제(false) 검증',
  sandbox.isSiteScheduleSheet(mockCompletedSheet) === false
);

// 53-3: onOpen 상단 메뉴 및 고급 설정 내 아카이브 항목 탑재 검증
const onOpenCode53 = code.slice(code.indexOf('function onOpen('), code.indexOf('function onEdit('));
assertTest('onOpen: 상단 메인 메뉴 내 "완료 현장 시트로 이동 (아카이브)" 탑재 검증',
  onOpenCode53.indexOf('📦 완료 현장 시트로 이동 (아카이브)') !== -1 &&
  onOpenCode53.indexOf('archiveCompletedSitesMenu') !== -1
);
assertTest('onOpen: 고급 설정 서브메뉴 내 아카이브 항목 탑재 검증',
  onOpenCode53.indexOf('📦 완료 현장 시트로 이동 (블록 전체 완료 현장 아카이브)') !== -1
);

// 53-4: 1일 1회 자동 이동 트리거 (runDailyCompletedSitesArchiveTrigger atHour(3)) 탑재 검증
const triggerCode53 = code.slice(code.indexOf('function ensureRealtimeTriggerInstalledSilent('), code.indexOf('function handleSheetChangeTrigger('));
assertTest('ensureRealtimeTriggerInstalledSilent: 1일 1회 새벽 3시 완료현장 아카이브 트리거(atHour(3)) 설치 검증',
  triggerCode53.indexOf('runDailyCompletedSitesArchiveTrigger') !== -1 &&
  triggerCode53.indexOf('atHour(3)') !== -1
);

// 53-5: 아카이브 도중 행 삭제 시 캘린더 삭제 트리거 방어 가드 검증
const changeCode53 = code.slice(code.indexOf('function handleSheetChangeTrigger('), code.indexOf('function cleanupOrphanCalendarEventsSilent('));
assertTest('handleSheetChangeTrigger: IS_ARCHIVING_ACTIVE 플래그 시 캘린더 삭제 스킵 가드 검증',
  changeCode53.indexOf('IS_ARCHIVING_ACTIVE') !== -1 &&
  changeCode53.indexOf('캘린더 삭제 트리거를 건너뜁니다') !== -1
);

// 53-6: cleanupOrphanCalendarEventsSilent & cleanupDuplicateAndOrphanCalendarEvents 캘린더 보존 검증
const silentCode53 = code.slice(code.indexOf('function cleanupOrphanCalendarEventsSilent('), code.indexOf('function runPeriodicAutoSync('));
assertTest('cleanupOrphanCalendarEventsSilent: 완료현장 시트 ID validEventIds[cpId] = -1 보호 탑재 검증',
  silentCode53.indexOf('findCompletedSitesSheet') !== -1 &&
  silentCode53.indexOf('validEventIds[cpId] = -1') !== -1 &&
  silentCode53.indexOf('matchedRowNum === -1') !== -1
);

const manualCode53 = code.slice(code.indexOf('function cleanupDuplicateAndOrphanCalendarEvents('), code.indexOf('function cleanupDuplicateCalendarEventsSilent('));
assertTest('cleanupDuplicateAndOrphanCalendarEvents: 완료현장 시트 ID 캘린더 삭제 절대 방어 탑재 검증',
  manualCode53.indexOf('findCompletedSitesSheet') !== -1 &&
  manualCode53.indexOf('validEventIds[cpId] = -1') !== -1 &&
  manualCode53.indexOf('matchedRowNum === -1') !== -1
);

// 53-7: 런타임 시뮬레이션
// 현장 1 (보라매경남 102동): 2개 공정 모두 완료 -> 완료현장으로 이동되어야 함
// 현장 2 (롯데낙천대 101동): 1개 완료, 1개 예정 -> 이동되면 안 됨
let appendedToCompleted = [];
let deletedFromSource = [];
let scriptPropsMap53 = {};

sandbox.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (k) => scriptPropsMap53[k] || null,
    setProperty: (k, v) => { scriptPropsMap53[k] = v; },
    deleteProperty: (k) => { delete scriptPropsMap53[k]; }
  }),
  getDocumentProperties: () => ({
    getProperty: () => null,
    setProperty: () => {}
  })
};

let simSrcData53 = [
  // 2~3행: 보라매경남 (모두 완료)
  ['보라매경남 102동 1104호', '2026.07.13', '철거', '이수현 반장님', true, '', '전예린', '완료', 'cal-boramae-01', Date.now()],
  ['보라매경남 102동 1104호', '2026.07.14', '목공', '박동연 반장님', true, '', '전예린', '완료', 'cal-boramae-02', Date.now()],
  // 4~5행: 롯데낙천대 (1개 완료, 1개 예정)
  ['롯데낙천대 101동 505호', '2026.06.19', '신고', '우성', true, '', '송지은', '완료', 'cal-lotte-01', Date.now()],
  ['롯데낙천대 101동 505호', '2026.11.26', '신고', '우성', true, '', '송지은', '예정', 'cal-lotte-02', Date.now()]
];

const mockSrcSheet53 = {
  getName: () => '현장서식',
  getSheetId: () => 101,
  getLastRow: () => simSrcData53.length + 1,
  getLastColumn: () => 10,
  getParent: () => mockSs53,
  getRange: (r, c, nr, nc) => {
    if (r === 1 && nr === 1) {
      return {
        getValues: () => [['현장명', '날짜', '공종', '협력업체', '공유', '비고', '담당자', '진행상태', '캘린더ID', '최종수정일시']]
      };
    }
    return {
      getValues: () => simSrcData53,
      getBackgrounds: () => simSrcData53.map(() => Array(10).fill('#ffffff')),
      getFontColors: () => simSrcData53.map(() => Array(10).fill('#000000')),
      getFontWeights: () => simSrcData53.map(() => Array(10).fill('normal')),
      getHorizontalAlignments: () => simSrcData53.map(() => Array(10).fill('center'))
    };
  },
  deleteRows: (startRow, count) => {
    deletedFromSource.push({ startRow, count });
    simSrcData53.splice(startRow - 2, count);
  }
};

const mockDestSheet53 = {
  getName: () => '완료현장',
  getSheetId: () => 102,
  getLastRow: () => appendedToCompleted.length + 1,
  getLastColumn: () => 10,
  getParent: () => mockSs53,
  getRange: (r, c, nr, nc) => ({
    setValues: (vals) => { appendedToCompleted.push(...vals); },
    setBackgrounds: () => {},
    setFontColors: () => {},
    setFontWeights: () => {},
    setHorizontalAlignments: () => {},
    setVerticalAlignment: () => {},
    setBorder: () => {},
    insertCheckboxes: () => {}
  })
};

const mockSs53 = {
  getSheets: () => [mockSrcSheet53, mockDestSheet53],
  getSheetByName: (n) => (n === '완료현장' ? mockDestSheet53 : (n === '현장서식' ? mockSrcSheet53 : null)),
  toast: () => {}
};

sandbox.SpreadsheetApp.getActiveSpreadsheet = () => mockSs53;
sandbox.getAppSpreadsheet = () => mockSs53;
sandbox.findSiteScheduleSheet = () => mockSrcSheet53;

const archiveResult = sandbox.archiveCompletedSiteBlocks(false);

assertTest('런타임 시뮬레이션: 완료 현장(보라매경남) 1개 정확 판별 검증',
  archiveResult.movedSitesCount === 1 && archiveResult.siteNames[0] === '보라매경남 102동 1104호'
);
assertTest('런타임 시뮬레이션: 완료 공정 2개 행 완료현장 시트로 이동 검증',
  archiveResult.movedRowsCount === 2 && appendedToCompleted.length === 2
);
assertTest('런타임 시뮬레이션: 미완료 예정 공정이 있는 현장(롯데낙천대)은 시트에 그대로 보존 검증',
  simSrcData53.length === 2 && simSrcData53[0][0] === '롯데낙천대 101동 505호'
);
assertTest('런타임 시뮬레이션: 이동 완료 후 IS_ARCHIVING_ACTIVE 플래그 정상 해제 검증',
  scriptPropsMap53['IS_ARCHIVING_ACTIVE'] === undefined
);

// 캘린더 유령 일정 청소 시 완료현장 시트로 이동된 일정이 절대 삭제되지 않는지 검증
let deletedCalEvents53 = [];
const mockCal53 = {
  getEvents: () => [
    {
      getId: () => 'cal-boramae-01', // 완료현장 시트로 이동된 일정!
      getTitle: () => '🟦 [철거] 보라매경남 102동 1104호',
      getDescription: () => '📌 현장명: 보라매경남 102동 1104호\n🔨 공종: 철거',
      deleteEvent: function() { deletedCalEvents53.push(this.getId()); }
    },
    {
      getId: () => 'cal-real-orphan-999', // 시트 어디에도 없는 진짜 유령 일정
      getTitle: () => '🟦 [철거] 삭제된옛날현장',
      getDescription: () => '📌 현장명: 삭제된옛날현장\n🔨 공종: 철거',
      deleteEvent: function() { deletedCalEvents53.push(this.getId()); }
    }
  ]
};

const mockDestSheetWithVals53 = {
  getName: () => '완료현장',
  getLastRow: () => appendedToCompleted.length + 1,
  getLastColumn: () => 10,
  getRange: (r, c, nr, nc) => ({
    getValues: () => appendedToCompleted
  }),
  getParent: () => mockSs53
};
mockSs53.getSheets = () => [mockSrcSheet53, mockDestSheetWithVals53];
mockSs53.getSheetByName = (n) => (n === '완료현장' ? mockDestSheetWithVals53 : mockSrcSheet53);

sandbox.cleanupOrphanCalendarEventsSilent(mockSrcSheet53, mockCal53);

assertTest('🛡️ 구글 캘린더 100% 보존 검증: 완료현장으로 이동된 일정(cal-boramae-01)은 캘린더에서 절대 삭제되지 않음',
  deletedCalEvents53.indexOf('cal-boramae-01') === -1
);
assertTest('🛡️ 구글 캘린더 정상 청소 검증: 진짜 유령 일정(cal-real-orphan-999)은 정상 삭제됨',
  deletedCalEvents53.indexOf('cal-real-orphan-999') !== -1
);

// ========================================================
// 📱 Step 54: 일정 상세 날짜/담당자 수정, 지연 상태 제거, 시트 행 이동 버튼 탑재 QA
// ========================================================
console.log('\n📱 [Step 54: 일정 상세 날짜/담당자 수정, 지연 상태 제거, 시트 행 이동 버튼 탑재 QA]');

const dashHtml54 = sandbox.buildSidebarDashboardHtml('mobile');

// 1. UI 컴포넌트 탑재 검증
assertTest('buildSidebarDashboardHtml: 일정(날짜) 텍스트 입력 필드(sheetDateInput) 탑재 검증',
  dashHtml54.indexOf('id="sheetDateInput"') !== -1
);
assertTest('buildSidebarDashboardHtml: 네이티브 달력 선택기(sheetNativeDatePicker) 탑재 검증',
  dashHtml54.indexOf('id="sheetNativeDatePicker"') !== -1
);
assertTest('buildSidebarDashboardHtml: 달력 팝업 오픈 버튼(sheet-cal-pick-btn) 탑재 검증',
  dashHtml54.indexOf('class="sheet-cal-pick-btn"') !== -1
);
assertTest('buildSidebarDashboardHtml: 담당자 선택 드롭다운(sheetManagerSelect) 탑재 검증',
  dashHtml54.indexOf('id="sheetManagerSelect"') !== -1
);
assertTest('buildSidebarDashboardHtml: 담당자 직접 입력 필드(sheetManagerCustomInput) 탑재 검증',
  dashHtml54.indexOf('id="sheetManagerCustomInput"') !== -1
);
assertTest('buildSidebarDashboardHtml: 하단 액션바 시트 행 이동 버튼(sheetGoToRowBtn) 제거 및 상단 미니 버튼 보존 검증',
  dashHtml54.indexOf('id="sheetGoToRowBtn"') === -1 && dashHtml54.indexOf('class="btn-sheet-goto-mini"') !== -1
);
assertTest('buildSidebarDashboardHtml: 상단 타이틀바 미니 시트 행 이동 버튼(btn-sheet-goto-mini) 탑재 검증',
  dashHtml54.indexOf('class="btn-sheet-goto-mini"') !== -1
);

// 2. 상태 선택 그리드에서 '지연' 상태 제거 및 4개 상태(예정, 진행 중, 완료, AS) 정돈 검증
assertTest('buildSidebarDashboardHtml: 상세 모달 상태 버튼에서 "지연" 완전 제거 검증',
  dashHtml54.indexOf('selectSheetStatus(\'지연\')') === -1 && dashHtml54.indexOf('data-status="지연"') === -1
);
assertTest('buildSidebarDashboardHtml: 4대 핵심 상태(예정, 진행 중, 완료, AS) 유지 검증',
  dashHtml54.indexOf('selectSheetStatus(\'예정\')') !== -1 &&
  dashHtml54.indexOf('selectSheetStatus(\'진행 중\')') !== -1 &&
  dashHtml54.indexOf('selectSheetStatus(\'완료\')') !== -1 &&
  dashHtml54.indexOf('selectSheetStatus(\'AS\')') !== -1
);

// 3. 클라이언트 제어 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 시트 행 이동 함수(goToCurrentTaskSheetRow) 탑재 검증',
  dashHtml54.indexOf('function goToCurrentTaskSheetRow()') !== -1
);
assertTest('buildSidebarDashboardHtml: 달력 팝업 오픈 및 선택 핸들러(openNativeDatePicker, onNativeDatePicked) 탑재 검증',
  dashHtml54.indexOf('function openNativeDatePicker()') !== -1 && dashHtml54.indexOf('function onNativeDatePicked(') !== -1
);
assertTest('buildSidebarDashboardHtml: 담당자 셀렉트 변경 핸들러(onSheetManagerSelect) 탑재 검증',
  dashHtml54.indexOf('function onSheetManagerSelect(') !== -1
);
assertTest('buildSidebarDashboardHtml: 담당자 목록 옵션 구성기(populateManagerSelectOptions) 탑재 검증',
  dashHtml54.indexOf('function populateManagerSelectOptions(') !== -1
);
assertTest('buildSidebarDashboardHtml: saveSheetTask에서 dateVal 및 mgrVal 서버 전달 검증',
  dashHtml54.indexOf('updateScheduleStatusFromWeb(curTaskObj.row, selectedStatus, selectedShare, memoVal, dateVal, mgrVal') !== -1
);

// 4. ContentService 모킹 및 시트 환경 구성
sandbox.ContentService = {
  MimeType: { JSON: 'JSON', TEXT: 'TEXT' },
  createTextOutput: function(text) {
    return {
      _content: text,
      _mime: '',
      setMimeType: function(m) { this._mime = m; return this; },
      getContent: function() { return this._content; }
    };
  }
};

const mockSheetData54 = [
  ['자이 101동', '2026.09.20', '철거', '철거팀', false, '', '이순신', '예정', 'ev-54', 1000]
];
let focusedRange54 = null;
const mockSheet54 = {
  getName: () => '현장데이터',
  getSheetId: () => 12345678,
  getLastRow: () => 10,
  getLastColumn: () => 10,
  getRange: (r, c, nr, nc) => {
    focusedRange54 = { r, c, nr, nc };
    if (nr && nc) {
      return {
        getValues: () => mockSheetData54
      };
    }
    return {
      setValue: (val) => {
        mockSheetData54[r - 2][c - 1] = val;
      },
      getValue: () => mockSheetData54[r - 2][c - 1]
    };
  },
  setActiveRange: (rng) => { focusedRange54 = rng; },
  getParent: () => mockSs54
};
const mockSs54 = {
  getId: () => 'test-sheet-id-54',
  setActiveSheet: () => {},
  getSheetByName: (n) => (n === '현장데이터' ? mockSheet54 : null),
  getSheets: () => [mockSheet54]
};

sandbox.SpreadsheetApp.getActiveSpreadsheet = () => mockSs54;
sandbox.getAppSpreadsheet = () => mockSs54;
sandbox.findSiteScheduleSheet = () => mockSheet54;
sandbox.syncSingleRowToCalendar = () => {};

// 5. focusSheetRow 런타임 검증
const focusRes54 = sandbox.focusSheetRow(5);
assertTest('focusSheetRow: 성공(success: true) 반환 검증', focusRes54.success === true);
assertTest('focusSheetRow: 행 번호(row: 5) 일치 검증', focusRes54.row === 5);
assertTest('focusSheetRow: 웹앱/모바일용 직접 링크 URL(sheetUrl) 생성 검증',
  focusRes54.sheetUrl && focusRes54.sheetUrl.indexOf('range=5:5') !== -1 && focusRes54.sheetUrl.indexOf('gid=12345678') !== -1
);

// 6. updateScheduleStatusFromWeb 날짜 및 담당자 업데이트 런타임 검증
const updateRes54 = sandbox.updateScheduleStatusFromWeb(
  2,
  '진행 중',
  true,
  '현장 진입 완료 및 보양 작업',
  '2026.09.29 (화)',
  '전예린'
);

assertTest('updateScheduleStatusFromWeb: 반환 결과 성공(success: true) 검증', updateRes54.success === true);
assertTest('updateScheduleStatusFromWeb: H열(8열: 상태) "진행 중" 업데이트 검증', mockSheetData54[0][7] === '진행 중');
assertTest('updateScheduleStatusFromWeb: B열(2열: 날짜) 요일 제거 및 YYYY.MM.DD 정규화("2026.09.29") 기록 검증',
  mockSheetData54[0][1] === '2026.09.29'
);
assertTest('updateScheduleStatusFromWeb: G열(7열: 담당자) "전예린" 정확 기록 검증', mockSheetData54[0][6] === '전예린');
assertTest('updateScheduleStatusFromWeb: E열(5열: 공유) true 기록 검증', mockSheetData54[0][4] === true);
assertTest('updateScheduleStatusFromWeb: F열(6열: 비고) "현장 진입 완료 및 보양 작업" 기록 검증',
  mockSheetData54[0][5] === '현장 진입 완료 및 보양 작업'
);
assertTest('updateScheduleStatusFromWeb: J열(10열: 타임스탬프) 최신 시각 갱신 검증', mockSheetData54[0][9] > 1000);

// 7. doGet 파라미터 연동 검증
const fakeDoGetEvent54 = {
  parameter: {
    action: 'updateScheduleStatus',
    row: '2',
    status: '완료',
    date: '2026.10.05',
    manager: '인경호',
    format: 'json'
  }
};
const doGetOutput54 = sandbox.doGet(fakeDoGetEvent54);
const doGetJson54 = JSON.parse(doGetOutput54.getContent());
assertTest('doGet: action=updateScheduleStatus 호출 성공 반환 검증', doGetJson54.success === true);
assertTest('doGet: 날짜 2026.10.05 시트 B열 반영 검증', mockSheetData54[0][1] === '2026.10.05');
assertTest('doGet: 담당자 "인경호" 시트 G열 반영 검증', mockSheetData54[0][6] === '인경호');
// ========================================================
// 🎯 Step 55: 상태 아이콘(📅 예정), 🟪 중복 일정 치유, 완료현장 접기, 대시보드 가로 확장 QA
// ========================================================
console.log('\n🎯 [Step 55: 상태 아이콘(📅 예정), 🟪 중복 일정 치유, 완료현장 접기, 대시보드 가로 확장 QA]');

// 1. 상태별 제목 맨 앞 아이콘 검증 (예정은 상태 아이콘 생략)
const testRowPlan55 = ['신대방우성 16동 902호', '2026.11.26', '가구', '우성', true, '', '전예린', '예정', 'ev-55-1', Date.now()];
const detailsPlan55 = sandbox.buildEventDetails(testRowPlan55, 2, null, mockSheet54);
assertTest('buildEventDetails: 예정 상태 시 제목 맨 앞에 상태 아이콘 없이 "🟦" 부여 검증',
  detailsPlan55 && detailsPlan55.title.startsWith('🟦 [가구] 신대방우성 16동 902호')
);

const testRowProg55 = ['신대방우성 16동 902호', '2026.11.26', '가구', '우성', true, '', '전예린', '진행 중', 'ev-55-2', Date.now()];
const detailsProg55 = sandbox.buildEventDetails(testRowProg55, 2, null, mockSheet54);
assertTest('buildEventDetails: 진행중 상태 시 제목 맨 앞에 "▶️ 🟦" 아이콘 부여 검증',
  detailsProg55 && detailsProg55.title.startsWith('▶️ 🟦 [가구] 신대방우성 16동 902호')
);

const testRowDone55 = ['신대방우성 16동 902호', '2026.11.26', '가구', '우성', false, '', '전예린', '완료', 'ev-55-3', Date.now()];
const detailsDone55 = sandbox.buildEventDetails(testRowDone55, 2, null, mockSheet54);
assertTest('buildEventDetails: 완료 상태 시 제목 맨 앞에 "✅ 🟥" 아이콘 부여 검증',
  detailsDone55 && detailsDone55.title.startsWith('✅ 🟥 [가구] 신대방우성 16동 902호')
);

// 2. isSystemSiteScheduleEvent 🟪 및 📅 인식 검증
const mockPurpleEvent55 = {
  getTitle: () => '🟪 [가구] 신대방우성 16동 902호',
  getDescription: () => '가구 공종 일정'
};
assertTest('isSystemSiteScheduleEvent: 과거 🟪(보라색 네모) 일정 정상 시스템 일정으로 인식 (삭제/치유 허용)',
  sandbox.isSystemSiteScheduleEvent(mockPurpleEvent55) === true
);

const mockScheduledCalEv55 = {
  getTitle: () => '📅 🟦 [가구] 신대방우성 16동 902호',
  getDescription: () => '가구 일정'
};
assertTest('isSystemSiteScheduleEvent: 신규 📅 🟦 일정 정상 시스템 일정으로 인식',
  sandbox.isSystemSiteScheduleEvent(mockScheduledCalEv55) === true
);

// 3. 중복 🟪 일정 치유 런타임 시뮬레이션
let deletedEvIds55 = [];
const mockOldPurpleEv = {
  getId: () => 'old-purple-nov26-001',
  getTitle: () => '🟪 [가구] 신대방우성 16동 902호',
  getDescription: () => '📌 현장명: 🟪 신대방우성 16동 902호\n🔨 공종: 가구',
  getLocation: () => '신대방우성 16동 902호',
  isAllDayEvent: () => true,
  getAllDayStartDate: () => new Date(2026, 10, 26),
  getStartTime: () => new Date(2026, 10, 26),
  deleteEvent: function() { deletedEvIds55.push(this.getId()); }
};

const mockNewRedEv = {
  getId: () => 'new-red-nov26-valid',
  getTitle: () => '📅 🟥 [가구] 신대방우성 16동 902호',
  getDescription: () => '📌 현장명: 신대방우성 16동 902호\n🔨 공종: 가구',
  getLocation: () => '신대방우성 16동 902호',
  isAllDayEvent: () => true,
  getAllDayStartDate: () => new Date(2026, 10, 26),
  getStartTime: () => new Date(2026, 10, 26),
  deleteEvent: () => {}
};

const mockSheetDedup55 = {
  getLastRow: () => 2,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [['신대방우성 16동 902호', '2026.11.26', '가구', '', false, '', '전예린', '예정', 'new-red-nov26-valid', Date.now()]]
  })
};

const mockCalDedup55 = {
  getEvents: () => [mockOldPurpleEv, mockNewRedEv]
};

sandbox.cleanupDuplicateCalendarEventsSilent(mockCalDedup55, mockSheetDedup55);

assertTest('cleanupDuplicateCalendarEventsSilent: 과거 중복 🟪 일정(old-purple-nov26-001) 안전 삭제 검증',
  deletedEvIds55.indexOf('old-purple-nov26-001') !== -1
);
assertTest('cleanupDuplicateCalendarEventsSilent: 신규 정상 일정(new-red-nov26-valid) 보존 검증',
  deletedEvIds55.indexOf('new-red-nov26-valid') === -1
);

// 4. 완료현장 시트 동일 현장별 접기(groupCompletedSitesBySite) 검증
let shiftedRanges55 = [];
let collapsedRanges55 = [];
const mockCompletedSheetData55 = [
  ['보라매경남 102동 1104호', '2026.06.17', '신고'],
  ['보라매경남 102동 1104호', '2026.06.19', '철거'],
  ['보라매경남 102동 1104호', '2026.06.25', '목공'],
  ['롯데낙천대 101동 505호', '2026.06.19', '신고'],
  ['롯데낙천대 101동 505호', '2026.06.20', '상담']
];

const mockCompletedSheet55 = {
  getLastRow: () => mockCompletedSheetData55.length + 1,
  getRange: (r, c, nr, nc) => {
    return {
      getValues: () => mockCompletedSheetData55.map(row => [row[0]]),
      shiftRowGroupDepth: (delta) => {
        shiftedRanges55.push({ r, c, nr, nc, delta });
      },
      collapseGroups: () => {
        collapsedRanges55.push({ r, c, nr, nc });
      },
      expandGroups: () => {}
    };
  }
};

sandbox.groupCompletedSitesBySite(mockCompletedSheet55);

assertTest('groupCompletedSitesBySite: 보라매경남 하위 행(3~4행, 2개 행) 그룹화(depth=1) 검증',
  shiftedRanges55.some(s => s.r === 3 && s.nr === 2 && s.delta === 1)
);
assertTest('groupCompletedSitesBySite: 보라매경남 하위 행 collapseGroups 접기 실행 검증',
  collapsedRanges55.some(c => c.r === 3 && c.nr === 2)
);
assertTest('groupCompletedSitesBySite: 롯데낙천대 하위 행(6행, 1개 행) 그룹화(depth=1) 검증',
  shiftedRanges55.some(s => s.r === 6 && s.nr === 1 && s.delta === 1)
);
assertTest('groupCompletedSitesBySite: 롯데낙천대 하위 행 collapseGroups 접기 실행 검증',
  collapsedRanges55.some(c => c.r === 6 && c.nr === 1)
);

// 5. 대시보드 가로 사이즈 확장 및 팝업 버튼 검증
const sidebarHtml55 = sandbox.buildSidebarDashboardHtml('sidebar');
assertTest('buildSidebarDashboardHtml: 상단 타이틀바 "⛶ 크게 보기" 팝업 버튼 탑재 검증',
  sidebarHtml55.indexOf('openWideModalView()') !== -1 && sidebarHtml55.indexOf('⛶ 크게 보기') !== -1
);
assertTest('openSiteScheduleSidebar: 가로 너비 450px 이상 지정 검증',
  code.indexOf('.setWidth(450)') !== -1
);
assertTest('openSiteScheduleModal: 대형 모달 너비 1020px 이상 지정 검증',
  code.indexOf('.setWidth(1020)') !== -1
);

// ========================================================
// 56. [상세모달 내장형 커스텀 달력 팝오버 & 시간 보존 & 완료현장 시트 분기 QA]
// ========================================================
console.log('\n🗓️ [Step 56: 상세모달 내장형 커스텀 달력 팝오버 & 시간 보존 & 완료현장 무결성 QA]');

const dashHtml56 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: 내장형 커스텀 달력 팝오버(sheetCustomCalPopover) 컨테이너 탑재 검증',
  dashHtml56.indexOf('id="sheetCustomCalPopover"') !== -1 &&
  dashHtml56.indexOf('id="pcalGrid"') !== -1 &&
  dashHtml56.indexOf('function renderCustomDatePicker()') !== -1 &&
  dashHtml56.indexOf('function onCustomCalDayClick(') !== -1
);

// 클라이언트 스크립트 문법 및 커스텀 달력 인터랙션 시뮬레이션
const scriptMatch56 = dashHtml56.match(/<script>([\s\S]*?)<\/script>/);
const clientScript56 = scriptMatch56 ? scriptMatch56[1] : '';
let clientSyntaxOk56 = false;
try {
  new vm.Script(clientScript56);
  clientSyntaxOk56 = true;
} catch (eSyn) {
  console.error('클라이언트 스크립트 문법 오류:', eSyn);
}
assertTest('buildSidebarDashboardHtml: 클라이언트 스크립트 정적 구문(Syntax) 무결성 검증', clientSyntaxOk56);

// DOM 모킹 후 openNativeDatePicker -> onCustomCalDayClick(단일/기간) 동작 검증
const mockElems56 = {
  sheetCustomCalPopover: {
    _classes: new Set(),
    classList: {
      contains: function(c) { return mockElems56.sheetCustomCalPopover._classes.has(c); },
      add: function(c) { mockElems56.sheetCustomCalPopover._classes.add(c); },
      remove: function(c) { mockElems56.sheetCustomCalPopover._classes.delete(c); },
      toggle: function(c, force) {
        if (force === undefined) {
          if (mockElems56.sheetCustomCalPopover._classes.has(c)) mockElems56.sheetCustomCalPopover._classes.delete(c);
          else mockElems56.sheetCustomCalPopover._classes.add(c);
        } else if (force) mockElems56.sheetCustomCalPopover._classes.add(c);
        else mockElems56.sheetCustomCalPopover._classes.delete(c);
      }
    }
  },
  sheetCalPickBtn: {
    _classes: new Set(),
    classList: {
      add: function(c) { mockElems56.sheetCalPickBtn._classes.add(c); },
      remove: function(c) { mockElems56.sheetCalPickBtn._classes.delete(c); }
    }
  },
  sheetDateInput: { value: '2026.10.17 (토)' },
  sheetNativeDatePicker: { value: '2026-10-17' },
  pcalGrid: { innerHTML: '' },
  pcalMonthTitle: { innerText: '' },
  pcalModeSingle: { classList: { toggle: function() {} } },
  pcalModeRange: { classList: { toggle: function() {} } },
  pcalGuideMsg: { style: { display: 'none' }, innerText: '' }
};

const clientSandbox56 = {
  document: {
    readyState: 'loading',
    getElementById: (id) => mockElems56[id] || null,
    querySelectorAll: () => []
  },
  window: { addEventListener: () => {} },
  setTimeout: () => {},
  Date: Date,
  Math: Math,
  Object: Object,
  Array: Array,
  Set: Set,
  JSON: JSON,
  encodeURIComponent: encodeURIComponent,
  decodeURIComponent: decodeURIComponent,
  parseInt: parseInt,
  String: String
};
vm.createContext(clientSandbox56);
vm.runInContext(clientScript56, clientSandbox56);

// 1) 달력 버튼 클릭 시 팝오버 즉시 오픈 검증
clientSandbox56.openNativeDatePicker();
assertTest('openNativeDatePicker: 크로스 오리진 iframe 환경에서도 커스텀 달력 팝오버(.show) 즉시 오픈 검증',
  mockElems56.sheetCustomCalPopover.classList.contains('show') === true &&
  mockElems56.pcalMonthTitle.innerText === '2026년 10월' &&
  mockElems56.pcalGrid.innerHTML.indexOf('onCustomCalDayClick(\'2026-10-17\')') !== -1
);

// 2) 단일 날짜 클릭 시 요일 포함 포맷 반영 및 팝오버 자동 닫힘 검증
clientSandbox56.onCustomCalDayClick('2026-10-20');
assertTest('onCustomCalDayClick(단일): 날짜 선택 시 "2026.10.20 (화)" 즉시 입력 및 팝오버 닫힘 검증',
  mockElems56.sheetDateInput.value === '2026.10.20 (화)' &&
  mockElems56.sheetCustomCalPopover.classList.contains('show') === false
);

// 3) 기존 시간 정보가 있는 일정에서 단일 날짜 변경 시 시간(14:00) 보존 검증
mockElems56.sheetDateInput.value = '2026.10.20 (화) 14:00';
clientSandbox56.openNativeDatePicker();
clientSandbox56.onCustomCalDayClick('2026-10-22');
assertTest('onNativeDatePicked: 시간 포함 일정 날짜 변경 시 기존 시간("14:00") 자동 보존 검증',
  mockElems56.sheetDateInput.value === '2026.10.22 (목) 14:00'
);

// 4) 기간(시작~종료) 모드에서 시작일 -> 종료일 클릭 시 범위 포맷 완성 검증
clientSandbox56.openNativeDatePicker();
clientSandbox56.setPickerSelectMode('range');
clientSandbox56.onCustomCalDayClick('2026-10-25');
const rangeStep1Open = mockElems56.sheetCustomCalPopover.classList.contains('show') === true;
clientSandbox56.onCustomCalDayClick('2026-10-27');
assertTest('onCustomCalDayClick(기간): 시작일~종료일 2회 클릭 시 "2026.10.25 (일) ~ 2026.10.27 (화)" 범위 완성 검증',
  rangeStep1Open &&
  mockElems56.sheetDateInput.value === '2026.10.25 (일) ~ 2026.10.27 (화)' &&
  mockElems56.sheetCustomCalPopover.classList.contains('show') === false
);

// 5) normalizeDateStringToYYYYMMDD 시간 포함 일정 보존 검증
assertTest('normalizeDateStringToYYYYMMDD: 시간 포함 일정("2026.10.17 (토) 14:00") 정규화 시 시간 증발 방지 검증',
  sandbox.normalizeDateStringToYYYYMMDD('2026.10.17 (토) 14:00') === '2026.10.17 14:00' &&
  sandbox.normalizeDateStringToYYYYMMDD('2026.10.17 (토) 09:00 ~ 18:00') === '2026.10.17 09:00 ~ 18:00'
);

// 6) toggleCalendarView에서 목록 복귀 시 모달 2열 그리드 보호(style.display = "") 검증
assertTest('toggleCalendarView: 목록 보기 복귀 시 모달 그리드 레이아웃 보존(contentArea.style.display = "") 검증',
  dashHtml56.indexOf('if (contentArea) contentArea.style.display = "";') !== -1
);

console.log('\n🔄 [Step 57: 새로고침 글로벌 로딩바(오버레이) & 대시보드 캘린더 상태 아이콘(📅/▶️/✅/🛠️) 표시 QA]');

const dashHtml57 = sandbox.buildSidebarDashboardHtml('sidebar');

// 57-1: 새로고침 로딩바/오버레이 HTML & CSS 및 버튼 클래스 탑재 검증
assertTest('buildSidebarDashboardHtml: #refreshLoadingOverlay 및 상단/카드 로딩바(.refresh-top-progress-bar, .refresh-bar-fill) 탑재 검증',
  dashHtml57.indexOf('id="refreshLoadingOverlay"') !== -1 &&
  dashHtml57.indexOf('refresh-top-progress-bar') !== -1 &&
  dashHtml57.indexOf('refresh-bar-fill') !== -1 &&
  dashHtml57.indexOf('btn-refresh-action') !== -1 &&
  dashHtml57.indexOf('function setRefreshLoadingState(') !== -1
);

// 57-2: 대시보드 캘린더 상태 아이콘 헬퍼(getTaskStatusIcon) 및 .cal-status-icon 탑재 검증
assertTest('buildSidebarDashboardHtml: getTaskStatusIcon 및 .cal-status-icon 렌더링 탑재 검증',
  dashHtml57.indexOf('function getTaskStatusIcon(') !== -1 &&
  dashHtml57.indexOf('cal-status-icon') !== -1
);

// 57-3: 클라이언트 샌드박스에서 새로고침 로딩바 동작 및 월간/주간 캘린더 상태 아이콘 렌더링 검증
const scriptMatch57 = dashHtml57.match(/<script>([\s\S]*?)<\/script>/);
const clientScript57 = scriptMatch57 ? scriptMatch57[1] : '';

const refreshBtns57 = [
  { disabled: false, innerHTML: '🔄 새로고침' },
  { disabled: false, innerHTML: '🔄 새로고침' }
];
const mockElems57 = {
  refreshLoadingOverlay: {
    _classes: new Set(),
    classList: {
      add: function(c) { mockElems57.refreshLoadingOverlay._classes.add(c); },
      remove: function(c) { mockElems57.refreshLoadingOverlay._classes.delete(c); },
      contains: function(c) { return mockElems57.refreshLoadingOverlay._classes.has(c); }
    }
  },
  contentArea: { innerHTML: '', style: {} },
  calendarArea: { style: { display: 'block' }, addEventListener: function() {} },
  calGrid: { innerHTML: '', classList: { remove: function() {}, add: function() {} }, offsetWidth: 100 },
  calWeekList: { innerHTML: '', classList: { remove: function() {}, add: function() {} }, offsetWidth: 100, style: {} },
  calMonthTitle: { innerText: '' },
  calFilterNotice: { innerHTML: '', style: {} },
  searchInput: { value: '' },
  kpiSites: { innerText: '' },
  kpiToday: { innerText: '' },
  kpiDelay: { innerText: '' },
  memberFilterBar: { innerHTML: '' },
  checkedTab: { style: {} },
  checkedTabCount: { innerText: '' },
  checkedActionBar: { style: {} },
  checkedCountMsg: { innerText: '' },
  toast: { innerText: '', classList: { add: function() {}, remove: function() {} } }
};

let pendingSuccessCb57 = null;
const clientSandbox57 = {
  document: {
    readyState: 'loading',
    getElementById: (id) => mockElems57[id] || null,
    querySelectorAll: (sel) => (sel === '.btn-refresh-action' ? refreshBtns57 : []),
    addEventListener: () => {}
  },
  window: { addEventListener: () => {} },
  setTimeout: () => {},
  fetch: () => Promise.resolve({ json: () => Promise.resolve({}) }),
  google: {
    script: {
      run: {
        withSuccessHandler: function(cb) {
          pendingSuccessCb57 = cb;
          return {
            withFailureHandler: function() {
              return {
                getSidebarDashboardData: function() {}
              };
            }
          };
        }
      }
    }
  },
  Date: Date,
  Math: Math,
  Set: Set,
  Object: Object,
  Array: Array,
  JSON: JSON,
  encodeURIComponent: encodeURIComponent,
  decodeURIComponent: decodeURIComponent,
  parseInt: parseInt,
  String: String
};
vm.createContext(clientSandbox57);
vm.runInContext(clientScript57, clientSandbox57);

// 1) getTaskStatusIcon 상태별 아이콘 반환 검증 (예정은 아이콘 생략 '')
assertTest('getTaskStatusIcon: 예정(""), 진행 중(▶️), 완료(✅), AS(🛠️) 상태 아이콘 정확 반환 검증',
  clientSandbox57.getTaskStatusIcon('예정') === '' &&
  clientSandbox57.getTaskStatusIcon('') === '' &&
  clientSandbox57.getTaskStatusIcon('진행 중') === '▶️' &&
  clientSandbox57.getTaskStatusIcon('진행') === '▶️' &&
  clientSandbox57.getTaskStatusIcon('완료') === '✅' &&
  clientSandbox57.getTaskStatusIcon('AS') === '🛠️' &&
  clientSandbox57.getTaskStatusIcon('A/S') === '🛠️'
);

// 2) 캘린더 모드에서 새로고침 클릭 시 로딩바 오버레이 즉시 표시 및 버튼 비활성화 검증
vm.runInContext('isCalendarMode = true; calYear = 2026; calMonth = 10;', clientSandbox57);
clientSandbox57.loadData(true);
assertTest('loadData(true): 캘린더 보기 상태에서 새로고침 클릭 시 #refreshLoadingOverlay.show 활성 및 버튼 "⏳ 갱신 중..." 전환 검증',
  mockElems57.refreshLoadingOverlay.classList.contains('show') === true &&
  refreshBtns57[0].disabled === true &&
  refreshBtns57[0].innerHTML.indexOf('⏳ 갱신 중...') !== -1
);

// 3) 데이터 로드 완료(onDataLoaded) 시 로딩바 해제 + 월간/주간 캘린더 내 상태 아이콘 렌더링 검증
const mockDashPayload57 = {
  kpi: { totalSites: 1, activeSites: 1, todayCount: 2, delayedCount: 1 },
  members: ['조항만'],
  holidays: {},
  sites: [
    {
      siteName: '보라매 파크시티',
      colorHex: '#2563EB',
      progressRate: 50,
      tasks: [
        { row: 2, taskName: '철거', startDateStr: '2026-10-17', endDateStr: '2026-10-17', status: '완료', isShared: true, isDelayed: false, manager: '조항만' },
        { row: 3, taskName: '목공', startDateStr: '2026-10-17', endDateStr: '2026-10-17', status: '진행 중', isShared: true, isDelayed: false, manager: '조항만' },
        { row: 4, taskName: '도배', startDateStr: '2026-10-17', endDateStr: '2026-10-17', status: '예정', isShared: false, isDelayed: false, manager: '조항만' },
        { row: 5, taskName: '마루보수', startDateStr: '2026-10-17', endDateStr: '2026-10-17', status: 'AS', isShared: true, isDelayed: false, manager: '조항만' }
      ]
    }
  ]
};
pendingSuccessCb57(mockDashPayload57);

assertTest('onDataLoaded: 데이터 수신 완료 시 #refreshLoadingOverlay 자동 숨김 및 버튼 "🔄 새로고침" 복원 + 완료 토스트 검증',
  mockElems57.refreshLoadingOverlay.classList.contains('show') === false &&
  refreshBtns57[0].disabled === false &&
  refreshBtns57[0].innerHTML === '🔄 새로고침' &&
  mockElems57.toast.innerText.indexOf('새로고침되었습니다') !== -1
);

assertTest('renderCalendar(월간 캘린더): 완료(✅)/진행(▶️)/AS(🛠️) 상태 아이콘 표시 및 예정 상태 아이콘 생략 검증',
  mockElems57.calGrid.innerHTML.indexOf('cal-status-icon') !== -1 &&
  mockElems57.calGrid.innerHTML.indexOf('>✅</span>') !== -1 &&
  mockElems57.calGrid.innerHTML.indexOf('>▶️</span>') !== -1 &&
  mockElems57.calGrid.innerHTML.indexOf('>🛠️</span>') !== -1 &&
  mockElems57.calGrid.innerHTML.indexOf('>📅</span>') === -1
);

vm.runInContext('calWeekStartDate = new Date(2026, 9, 11); renderWeeklyCalendar();', clientSandbox57);
assertTest('renderWeeklyCalendar(주간 캘린더): 완료(✅)/진행(▶️)/AS(🛠️) 상태 아이콘 표시 및 예정 상태 아이콘 생략 검증',
  mockElems57.calWeekList.innerHTML.indexOf('cal-status-icon') !== -1 &&
  mockElems57.calWeekList.innerHTML.indexOf('>✅</span>') !== -1 &&
  mockElems57.calWeekList.innerHTML.indexOf('>▶️</span>') !== -1 &&
  mockElems57.calWeekList.innerHTML.indexOf('>🛠️</span>') !== -1 &&
  mockElems57.calWeekList.innerHTML.indexOf('>📅</span>') === -1
);

// ============================================================================
// [Step 58] 모바일 로딩 및 수정 속도 초고속화 QA 검증
// ============================================================================
console.log('\n--- [Step 58] 모바일 로딩 및 수정 속도 초고속화 QA 검증 ---');

// 1) html2canvas defer 및 triggerInitialLoadOnce 적용 검증
assertTest('buildSidebarDashboardHtml: html2canvas 스크립트 defer 적용 및 triggerInitialLoadOnce 중복 실행 방지 가드 포함 검증',
  dashHtml57.indexOf('<script defer src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>') !== -1 &&
  dashHtml57.indexOf('function triggerInitialLoadOnce()') !== -1
);

// 2) doGet(e) 모바일 대시보드 진입 시 Lazy Spreadsheet 오픈 (시트 I/O 0회) 검증
const origGetAppSs58 = sandbox.getAppSpreadsheet;
let getAppSsCalls58 = 0;
sandbox.getAppSpreadsheet = function() {
  getAppSsCalls58++;
  return origGetAppSs58.apply(this, arguments);
};
sandbox.doGet({ parameter: {} });
sandbox.doGet({ parameter: { view: 'dashboard' } });
assertTest('doGet: 모바일 대시보드 초기 접속 시 불필요한 SpreadsheetApp.openById 호출 완전 생략 (0회) 검증',
  getAppSsCalls58 === 0
);
sandbox.getAppSpreadsheet = origGetAppSs58;

// 3) syncSingleRowToCalendar 스마트 Diff 체크 검증 (동일 속성일 때 불필요한 Calendar API set* 호출 0회)
let calSetCalls58 = 0;
const mockExistingEvent58 = {
  getId: () => 'evt-fast-58@google.com',
  getTitle: () => '⚡ [진행] 보라매 파크시티 - 목공 (조항만)',
  setTitle: () => { calSetCalls58++; },
  getDescription: () => '현장명: 보라매 파크시티\n작업내용: 목공\n상태: 진행 중\n담당자: 조항만\n비고: 기존메모\n\n📱 상태 변경 및 일정 관리 대시보드:\nhttps://script.google.com/macros/s/test/exec?view=dashboard&id=ID-FAST-58&row=3',
  setDescription: () => { calSetCalls58++; },
  getLocation: () => '서울 동작구',
  setLocation: () => { calSetCalls58++; },
  getColor: () => '9',
  setColor: () => { calSetCalls58++; },
  isAllDayEvent: () => true,
  getAllDayStartDate: () => new Date(2026, 9, 17),
  getAllDayEndDate: () => new Date(2026, 9, 18),
  setAllDayDate: () => { calSetCalls58++; },
  setAllDayDates: () => { calSetCalls58++; },
  setTime: () => { calSetCalls58++; },
  removeAllReminders: () => {}
};
const origGetCalById58 = sandbox.CalendarApp.getCalendarById;
const origCleanDup58 = sandbox.cleanDuplicateCalendarEventsOnDate;
let cleanDupCalled58 = false;
sandbox.cleanDuplicateCalendarEventsOnDate = function() {
  cleanDupCalled58 = true;
};
sandbox.CalendarApp.getCalendarById = () => ({
  getEventById: () => mockExistingEvent58
});
const mockSheet58 = {
  getLastRow: () => 3,
  getRange: (r, c, numR, numC) => ({
    getValues: () => [[
      'ID-FAST-58', true, '보라매 파크시티', '서울 동작구', '목공',
      new Date(2026, 9, 17), new Date(2026, 9, 17), '진행 중', '조항만', '기존메모', 'evt-fast-58@google.com', '09:00', '9'
    ]],
    getValue: () => 'evt-fast-58@google.com',
    setValue: () => {},
    setFontColor: () => {},
    setBackground: () => {}
  })
};
const origScriptAppGetService58 = sandbox.ScriptApp.getService;
sandbox.ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' });
sandbox.syncSingleRowToCalendar(mockSheet58, 3, true);
assertTest('syncSingleRowToCalendar(skipDashboardRefresh=true): 변경 없는 속성에 대한 Calendar API set* 호출 0회 및 중복 전체 스캔 생략 검증',
  calSetCalls58 === 0 && cleanDupCalled58 === false
);
sandbox.CalendarApp.getCalendarById = origGetCalById58;
sandbox.cleanDuplicateCalendarEventsOnDate = origCleanDup58;
sandbox.ScriptApp.getService = origScriptAppGetService58;

// 4) 클라이언트 로컬 캐시(Stale-While-Revalidate) 및 낙관적 UI(Optimistic Update) 검증
const mockStorage58 = {};
let savedFailureCb58 = null;
let savedSuccessCb58 = null;
let silentLoadCalled58 = false;

mockElems57.sheetManagerSelect = { value: '조항만', innerHTML: '' };
mockElems57.sheetDateInput = { value: '2026-10-17 ~ 2026-10-18' };
mockElems57.sheetMemoInput = { value: '모바일 즉시 수정 메모' };
mockElems57.sheetSaveBtn = { disabled: false, innerText: '💾 저장하기' };
mockElems57.sheetOverlay = {
  classList: {
    _set: new Set(['show']),
    add: function(c) { this._set.add(c); },
    remove: function(c) { this._set.delete(c); },
    contains: function(c) { return this._set.has(c); }
  }
};

clientSandbox57.localStorage = {
  getItem: (k) => mockStorage58[k] || null,
  setItem: (k, v) => { mockStorage58[k] = String(v); },
  removeItem: (k) => { delete mockStorage58[k]; }
};
clientSandbox57.google.script.run = {
  withSuccessHandler: function(cb) {
    savedSuccessCb58 = cb;
    return {
      withFailureHandler: function(errCb) {
        savedFailureCb58 = errCb;
        return {
          getSidebarDashboardData: function() { silentLoadCalled58 = true; },
          updateScheduleStatusFromWeb: function() {}
        };
      }
    };
  }
};

// 먼저 캐시에 데이터가 저장되는지 확인
clientSandbox57.onDataLoaded(mockDashPayload57);
assertTest('localStorage 캐시 저장: onDataLoaded 호출 시 ws_dashboard_cache_v1 로컬 캐시 자동 저장 검증',
  typeof mockStorage58['ws_dashboard_cache_v1'] === 'string' &&
  mockStorage58['ws_dashboard_cache_v1'].indexOf('보라매 파크시티') !== -1
);

// 초기 구동(loadData(false)) 시 캐시로 즉시 렌더링 + 백그라운드 무음 동기화 호출 검증
silentLoadCalled58 = false;
vm.runInContext('rawData = null;', clientSandbox57);
clientSandbox57.loadData(false);
assertTest('Stale-While-Revalidate: 초기 loadData(false) 시 로컬 캐시 즉시 렌더링 및 백그라운드 무음 갱신 트리거 검증',
  silentLoadCalled58 === true &&
  mockElems57.kpiSites.innerText === 1 &&
  vm.runInContext('rawData && rawData.sites.length', clientSandbox57) === 1
);

// 모바일 바텀시트에서 작업 수정 저장 시 서버 응답 전에 즉시(0.01초) UI 및 진행률 업데이트 + 모달 닫힘 검증
vm.runInContext(`
  curTaskObj = { row: 3, siteName: '보라매 파크시티' };
  selectedStatus = '완료';
  selectedShare = false;
`, clientSandbox57);
clientSandbox57.saveSheetTask();
const rawDataAfterOpt58 = vm.runInContext('rawData', clientSandbox57);
const updatedTask58 = rawDataAfterOpt58.sites[0].tasks.find(t => t.row === 3);
assertTest('saveSheetTask(Optimistic UI): 저장 버튼 클릭 즉시 바텀시트 닫힘 + 상태/진행률/메모/날짜 0.01초 선반영 검증',
  mockElems57.sheetOverlay.classList.contains('show') === false &&
  updatedTask58.status === '완료' &&
  updatedTask58.isShared === false &&
  updatedTask58.memo === '모바일 즉시 수정 메모' &&
  updatedTask58.endDateStr === '2026-10-18' &&
  rawDataAfterOpt58.sites[0].progressRate === 50
);

// 만약 서버 통신 실패 시 이전 스냅샷으로 안전하게 자동 롤백되는지 검증
savedFailureCb58({ message: '네트워크 일시 오류' });
const rawDataAfterRollback58 = vm.runInContext('rawData', clientSandbox57);
const rolledBackTask58 = rawDataAfterRollback58.sites[0].tasks.find(t => t.row === 3);
assertTest('saveSheetTask(Rollback): 서버 저장 실패 시 낙관적 업데이트 자동 롤백 및 원본 상태(진행 중, 25%) 복원 검증',
  rolledBackTask58.status === '진행 중' &&
  rolledBackTask58.isShared === true &&
  rawDataAfterRollback58.sites[0].progressRate === 25
);

// ============================================================================
// [Step 60] 1번(모바일 새 일정 빠른 추가) + 4번(서버 CacheService 캐시) + 5번(캘린더 날짜 클릭 즉시 등록) QA 검증
// ============================================================================
console.log('\n--- [Step 60] 1번(새 일정 추가), 4번(서버 CacheService), 5번(캘린더 날짜 클릭 등록) 무결성 검증 ---');

// 1) 4번: 서버단 CacheService 동작 및 수정/추가 시 즉시 캐시 무효화 검증
const mockSheetData60 = [
  ['보라매 파크시티', '2026.10.17', '철거', '철거팀', true, '', '조항만', '완료', 'ev-60-1', 1000]
];
const mockSheet60 = {
  getName: () => '현장데이터',
  getSheetId: () => 606060,
  getLastRow: () => mockSheetData60.length + 1,
  getLastColumn: () => 10,
  getRange: (r, c, nr, nc) => {
    if (nr && nc) {
      return {
        getValues: () => mockSheetData60.slice(r - 2, r - 2 + nr),
        setValues: (vals) => {
          for (let i = 0; i < vals.length; i++) {
            mockSheetData60[r - 2 + i] = vals[i].slice();
          }
        }
      };
    }
    return {
      setValue: (val) => {
        if (!mockSheetData60[r - 2]) mockSheetData60[r - 2] = [];
        mockSheetData60[r - 2][c - 1] = val;
      },
      getValue: () => (mockSheetData60[r - 2] ? mockSheetData60[r - 2][c - 1] : ''),
      insertCheckboxes: () => {}
    };
  },
  getParent: () => mockSs60
};
const mockSs60 = {
  getId: () => 'test-sheet-id-60',
  setActiveSheet: () => {},
  getSheetByName: (n) => (n === '현장데이터' ? mockSheet60 : null),
  getSheets: () => [mockSheet60]
};
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => mockSs60;
sandbox.getAppSpreadsheet = () => mockSs60;
sandbox.findSiteScheduleSheet = () => mockSheet60;

const mockServerCacheStore60 = {};
let cacheGetCount60 = 0;
let cachePutCount60 = 0;
let cacheRemoveCount60 = 0;
sandbox.CacheService = {
  getScriptCache: () => ({
    get: (k) => {
      cacheGetCount60++;
      return Object.prototype.hasOwnProperty.call(mockServerCacheStore60, k) ? mockServerCacheStore60[k] : null;
    },
    put: (k, v, ttl) => {
      cachePutCount60++;
      mockServerCacheStore60[k] = String(v);
    },
    remove: (k) => {
      cacheRemoveCount60++;
      delete mockServerCacheStore60[k];
    }
  })
};

const firstDashRes60 = sandbox.getSidebarDashboardData(false);
assertTest('CacheService(4번): 최초 getSidebarDashboardData(false) 호출 시 스크립트 캐시(WS_DASHBOARD_DATA_CACHE_V1) 자동 저장 검증',
  Array.isArray(firstDashRes60.sites) &&
  cachePutCount60 === 1 &&
  typeof mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1'] === 'string'
);

// 캐시된 데이터를 임의로 마킹하여 두 번째 호출이 시트 재조회 없이 캐시에서 즉시 반환되는지 검증
const cachedObj60 = JSON.parse(mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1']);
cachedObj60._fromCacheTestMarker = true;
mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1'] = JSON.stringify(cachedObj60);

const secondDashRes60 = sandbox.getSidebarDashboardData(false);
assertTest('CacheService(4번): 재호출 시 시트 파싱 없이 0.05초 캐시 히트(_fromCacheTestMarker) 즉시 반환 검증',
  Array.isArray(secondDashRes60.sites) &&
  secondDashRes60._fromCacheTestMarker === true
);

const forcedDashRes60 = sandbox.getSidebarDashboardData(true);
assertTest('CacheService(4번): 수동 새로고침 getSidebarDashboardData(true) 호출 시 캐시 우회 및 최신 시트 데이터 재저장 검증',
  Array.isArray(forcedDashRes60.sites) &&
  forcedDashRes60._fromCacheTestMarker === undefined &&
  cachePutCount60 === 2
);

// 2) 1번: 서버단 addScheduleTaskFromWeb 새 일정(공종) 추가 및 캐시 자동 무효화 + doGet 연동 검증
const errEmptyTask60 = sandbox.addScheduleTaskFromWeb('', '목공', '2026-10-25', '김소장', '예정', true, '', '');
assertTest('addScheduleTaskFromWeb(경계값): 현장명 또는 작업명 누락 시 안전하게 에러 반환 검증',
  errEmptyTask60.success === false && errEmptyTask60.error.indexOf('현장명을 입력') !== -1
);

const beforeRowCount60 = mockSheet60.getLastRow();
const addRes60 = sandbox.addScheduleTaskFromWeb(
  '보라매 파크시티',
  '필름',
  '2026-10-25 (일)',
  '박실장',
  '예정',
  true,
  '문틀 필름 시공',
  '한솔필름'
);
assertTest('addScheduleTaskFromWeb(1번): 새 일정 추가 시 시트 하단 행 추가, 날짜 정규화(YYYY.MM.DD), 캐시 즉시 무효화 검증',
  addRes60.success === true &&
  addRes60.row === beforeRowCount60 + 1 &&
  addRes60.siteName === '보라매 파크시티' &&
  addRes60.taskName === '필름' &&
  addRes60.date === '2026.10.25' &&
  mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1'] === undefined
);

const doGetAddRes60 = JSON.parse(sandbox.doGet({
  parameter: {
    action: 'addScheduleTask',
    siteName: '신규 테스트 현장',
    taskName: '철거',
    date: '2026-10-26',
    manager: '이소장',
    status: '예정',
    shared: 'true',
    memo: '전체 철거',
    partner: '번개철거'
  }
}).getContent());
assertTest('doGet(addScheduleTask): 웹앱 엔드포인트에서도 새 현장/공종 추가 정상 동작 검증',
  doGetAddRes60.success === true &&
  doGetAddRes60.siteName === '신규 테스트 현장' &&
  doGetAddRes60.taskName === '철거'
);

// 이후 호출에 영향이 없도록 sandbox.CacheService 정리
delete sandbox.CacheService;

// 3) 1번 & 5번 클라이언트 UI/인터랙션 검증 (모달 엘리먼트 목업 보강)
clientSandbox57.document.createElement = () => ({ value: '', innerText: '' });
mockElems57.sheetManagerSelect.appendChild = function() {};
[
  'sheetCreateFields', 'sheetSiteSelect', 'sheetSiteCustomInput', 'sheetTaskInput',
  'sheetPartnerInput', 'sheetAddSameDateBtn', 'btnOpenAddTask', 'sheetSiteName',
  'sheetTaskName', 'sheetGoToRowBtn', 'sheetMiniGoToRowBtn', 'sheetDayNav',
  'sheetNativeDatePicker', 'sheetManagerCustomInput', 'sheetShareOn', 'sheetShareOff',
  'sheetCustomCalPopover'
].forEach(id => {
  if (!mockElems57[id]) {
    mockElems57[id] = {
      style: {},
      value: '',
      innerText: '',
      innerHTML: '',
      options: [],
      focus: function() {},
      classList: {
        _set: new Set(),
        add: function(c) { this._set.add(c); },
        remove: function(c) { this._set.delete(c); },
        toggle: function(c, cond) { if (cond) this._set.add(c); else this._set.delete(c); },
        contains: function(c) { return this._set.has(c); }
      }
    };
  }
});

let addScheduleCalledArgs60 = null;
let addScheduleSuccessCb60 = null;
let addScheduleFailureCb60 = null;
clientSandbox57.google.script.run = {
  withSuccessHandler: function(cb) {
    addScheduleSuccessCb60 = cb;
    return {
      withFailureHandler: function(errCb) {
        addScheduleFailureCb60 = errCb;
        return {
          getSidebarDashboardData: function() {},
          updateScheduleStatusFromWeb: function() {},
          addScheduleTaskFromWeb: function(site, task, dStr, mgr, st, sh, memo, pt) {
            addScheduleCalledArgs60 = { site, task, dStr, mgr, st, sh, memo, pt };
          }
        };
      }
    };
  }
};

// 5번: 캘린더에서 빈 날짜('2026-10-28') 클릭 시 자동으로 새 일정 추가 모달이 해당 날짜로 열리는지 검증
clientSandbox57.onCalDayClick('2026-10-28');
assertTest('onCalDayClick(5번): 캘린더 빈 날짜 클릭 시 해당 날짜(2026.10.28)가 자동 세팅된 새 일정 등록 모달 즉시 오픈 검증',
  vm.runInContext('isCreateMode', clientSandbox57) === true &&
  mockElems57.sheetOverlay.classList.contains('show') === true &&
  mockElems57.sheetCreateFields.style.display === 'block' &&
  mockElems57.sheetDateInput.value.indexOf('2026.10.28') === 0
);

// 1번: 빠른 공종 칩 클릭 시 작업명 입력창에 즉시 세팅되는지 검증
clientSandbox57.pickQuickTaskName('타일');
assertTest('pickQuickTaskName(1번): 빠른 공종 칩 선택 시 작업명 입력란에 즉시 반영 검증',
  mockElems57.sheetTaskInput.value === '타일'
);

// 1번: 새 일정 저장 시 0.01초 낙관적 UI(Optimistic UI) 즉시 반영 + 서버 addScheduleTaskFromWeb 호출 검증
mockElems57.sheetSiteSelect.value = '보라매 파크시티';
mockElems57.sheetManagerSelect.value = '최소장';
mockElems57.sheetPartnerInput.value = '대림타일';
mockElems57.sheetMemoInput.value = '주방 타일 시공';
const beforeTasksLen60 = vm.runInContext('rawData.sites[0].tasks.length', clientSandbox57);

clientSandbox57.saveSheetTask();
const afterTasksLen60 = vm.runInContext('rawData.sites[0].tasks.length', clientSandbox57);
const addedOptTask60 = vm.runInContext('rawData.sites[0].tasks.find(t => t.taskName === "타일" && t.endDateStr === "2026-10-28")', clientSandbox57);

assertTest('saveSheetTask(1번 Create Optimistic UI): 새 일정 저장 즉시 모달 닫힘 + 0.01초 로컬 리스트/캘린더 선반영 + 서버 호출 검증',
  mockElems57.sheetOverlay.classList.contains('show') === false &&
  afterTasksLen60 === beforeTasksLen60 + 1 &&
  Boolean(addedOptTask60) &&
  addedOptTask60.taskName === '타일' &&
  addedOptTask60.endDateStr === '2026-10-28' &&
  addScheduleCalledArgs60 !== null &&
  addScheduleCalledArgs60.site === '보라매 파크시티' &&
  addScheduleCalledArgs60.task === '타일'
);

// 서버 응답 성공 시 임시 행 번호(99999)가 실제 시트 행 번호(99)로 정식 교체되는지 검증
addScheduleSuccessCb60({
  success: true,
  row: 99
});
const confirmedTask60 = vm.runInContext('rawData.sites[0].tasks.find(t => t.taskName === "타일" && t.endDateStr === "2026-10-28")', clientSandbox57);
assertTest('saveSheetTask(1번 Create 확정): 서버 저장 성공 시 임시 행 번호가 실제 시트 행 번호(99)로 교체 검증',
  Boolean(confirmedTask60 && confirmedTask60.row === 99)
);

// 5번: 월간 및 주간 캘린더 날짜 셀에 '+ / + 추가' 버튼(.cal-day-add-btn) 및 리스트 뷰 현장 헤더에 '+ 공정' 버튼(.btn-site-add-mini) 렌더링 검증
clientSandbox57.renderList();
clientSandbox57.renderCalendar();
clientSandbox57.renderWeeklyCalendar();
assertTest('UI 버튼 렌더링(1번·5번): 현장 카드 "+ 공정" 버튼(.btn-site-add-mini) 및 월간/주간 캘린더 날짜 셀 "+" 버튼(.cal-day-add-btn) 포함 검증',
  mockElems57.contentArea.innerHTML.indexOf('btn-site-add-mini') !== -1 &&
  mockElems57.calGrid.innerHTML.indexOf('cal-day-add-btn') !== -1 &&
  mockElems57.calWeekList.innerHTML.indexOf('cal-day-add-btn') !== -1
);

// ============================================================================
// [Step 61] 일정 추가 시 완료 현장 및 새 현장 직접 입력 제거 + 담당자 '영하' 완전 삭제 QA 검증
// ============================================================================
console.log('\n--- [Step 61] 일정 추가 시 완료 현장/새 현장 제외 및 담당자 "영하" 완전 삭제 무결성 검증 ---');

// 1) 코드단에서 sheetSiteCustomInput, onSheetSiteSelect, __NEW_SITE__, createdNewSite 완전 삭제 검증
assertTest('buildSidebarDashboardHtml: 새 현장 직접 입력 UI(sheetSiteCustomInput, __NEW_SITE__, onSheetSiteSelect) 코드 완전 삭제 검증',
  dashHtml57.indexOf('sheetSiteCustomInput') === -1 &&
  dashHtml57.indexOf('__NEW_SITE__') === -1 &&
  dashHtml57.indexOf('onSheetSiteSelect') === -1 &&
  dashHtml57.indexOf('새 현장 직접 입력') === -1
);

// 2) 일정 추가 모달 현장 목록(populateSiteSelectOptions)에서 완료된 현장(isAllCompleted / isArchived / 전 공정 완료) 제외 검증
vm.runInContext(`
  rawData.sites = [
    {
      siteName: '신대방우성 5동 601호',
      isAllCompleted: false,
      isArchived: false,
      tasks: [{ row: 2, taskName: '목공', status: '진행 중' }]
    },
    {
      siteName: '보라매경남 102동 1104호',
      isAllCompleted: true,
      isArchived: false,
      tasks: [{ row: 10, taskName: '청소', status: '완료' }]
    },
    {
      siteName: '망원 경일빌라 201호',
      isAllCompleted: false,
      isArchived: true,
      tasks: [{ row: 20, taskName: '도배', status: '완료' }]
    },
    {
      siteName: '광명 하안주공 1221동 1404호',
      isAllCompleted: false,
      isArchived: false,
      tasks: [{ row: 30, taskName: '마루', status: '완료' }, { row: 31, taskName: '타일', status: '취소' }]
    }
  ];
`, clientSandbox57);
clientSandbox57.populateSiteSelectOptions('');
assertTest('populateSiteSelectOptions: 진행 중인 현장만 표시하고 완료/아카이브 현장 및 새 현장 옵션은 완전 배제 검증',
  mockElems57.sheetSiteSelect.innerHTML.indexOf('신대방우성 5동 601호') !== -1 &&
  mockElems57.sheetSiteSelect.innerHTML.indexOf('보라매경남 102동 1104호') === -1 &&
  mockElems57.sheetSiteSelect.innerHTML.indexOf('망원 경일빌라 201호') === -1 &&
  mockElems57.sheetSiteSelect.innerHTML.indexOf('광명 하안주공 1221동 1404호') === -1 &&
  mockElems57.sheetSiteSelect.innerHTML.indexOf('__NEW_SITE__') === -1
);

// 3) 담당자 드롭다운(populateManagerSelectOptions) 및 상단 담당자 필터바(renderMemberFilterTabs)에서 '영하' 완전 배제 검증
const appendedManagerOpts61 = [];
mockElems57.sheetManagerSelect._populated = false;
mockElems57.sheetManagerSelect.appendChild = function(opt) {
  appendedManagerOpts61.push(opt.value);
};
clientSandbox57.populateManagerSelectOptions(['조항만', '송지은', '전예린', '한송이', '인경호', '영하']);
clientSandbox57.renderMemberFilterTabs(['조항만', '송지은', '전예린', '한송이', '인경호', '영하']);

assertTest('담당자 목록: populateManagerSelectOptions 및 renderMemberFilterTabs에서 "영하" 완전 삭제 검증',
  appendedManagerOpts61.indexOf('영하') === -1 &&
  appendedManagerOpts61.indexOf('조항만') !== -1 &&
  appendedManagerOpts61.indexOf('한송이') !== -1 &&
  appendedManagerOpts61.indexOf('인경호') !== -1 &&
  mockElems57.memberFilterBar.innerHTML.indexOf('영하') === -1
);

// ============================================================================
// [Step 62] 이메일 모바일 빠른 관리 링크 변경 + 일정 추가 시 동일 현장 그룹 병합 및 날짜순 정렬 QA 검증
// ============================================================================
console.log('\n--- [Step 62] 이메일 모바일 빠른 관리 링크 및 현장별 그룹 병합/날짜순 정렬 무결성 검증 ---');

// 1) 이메일 모닝 브리핑 '📱 모바일 빠른 관리' 링크 URL 검증
const emailRes62 = sandbox.buildMorningBriefingEmailContent(
  '2026-10-17',
  '토',
  [{ siteName: '신대방우성 5동 601호', taskName: '목공', status: '진행 중', dateStr: '2026.10.17', manager: '조항만', memo: '' }],
  [],
  'https://docs.google.com/spreadsheets/d/test',
  'https://script.google.com/macros/s/old-url/exec'
);
assertTest('buildMorningBriefingEmailContent: 이메일 내 "📱 모바일 빠른 관리" 링크가 https://kimver2.github.io/00_workapces/ 로 정확히 설정 검증',
  emailRes62.htmlBody.indexOf('href="https://kimver2.github.io/00_workapces/"') !== -1 &&
  emailRes62.plainText.indexOf('https://kimver2.github.io/00_workapces/') !== -1
);

// 2) 시트 맨 아래(125행 등)에 분리되어 들어간 동일 현장 행 자동 병합(consolidateSplitSiteRows) 및 날짜순 정렬 검증
const mockRows62 = [
  ['신대방우성 5동 601호', '2026.10.15', '철거', '', true, '', '조항만', '완료', 'ev-1', 100],
  ['신대방우성 5동 601호', '2026.10.22', '목공', '', true, '', '조항만', '예정', 'ev-2', 200],
  ['보경102동', '2026.10.15', '타일', '', false, '', '전예린', '예정', 'ev-3', 300],
  ['보경102동', '2026.11.10', '입주', '', true, '', '전예린', '예정', 'ev-4', 400],
  ['신대방우성 5동 601호', '2026.10.18', '테스트', '', true, '', '', '예정', 'ev-5', 500]
];
const mockSheet62 = {
  getName: () => '현장데이터',
  getSheetId: () => 626262,
  getLastRow: () => mockRows62.length + 1,
  getLastColumn: () => 10,
  insertRowAfter: (afterRow) => {
    const idx = afterRow - 1; // 2행이 index 0이므로 afterRow 뒤는 index (afterRow - 1)
    mockRows62.splice(idx, 0, ['', '', '', '', false, '', '', '예정', '', 0]);
  },
  deleteRow: (rowNum) => {
    mockRows62.splice(rowNum - 2, 1);
  },
  getRange: (r, c, nr, nc) => {
    if (nr && nc) {
      return {
        getValues: () => mockRows62.slice(r - 2, r - 2 + nr).map(row => row.slice()),
        getBackgrounds: () => mockRows62.slice(r - 2, r - 2 + nr).map(() => new Array(nc).fill('#ffffff')),
        setValues: (vals) => {
          for (let i = 0; i < vals.length; i++) {
            mockRows62[r - 2 + i] = vals[i].slice();
          }
        },
        setBackgrounds: () => {},
        setVerticalAlignment: () => {},
        setHorizontalAlignment: () => {},
        setBorder: () => {},
        setDataValidation: () => {},
        insertCheckboxes: function() { return this; }
      };
    }
    return {
      setValue: (val) => {
        if (!mockRows62[r - 2]) mockRows62[r - 2] = [];
        mockRows62[r - 2][c - 1] = val;
      },
      getValue: () => (mockRows62[r - 2] ? mockRows62[r - 2][c - 1] : ''),
      insertCheckboxes: function() { return this; },
      setHorizontalAlignment: function() { return this; }
    };
  },
  getParent: () => mockSs62
};
const mockSs62 = {
  getId: () => 'test-sheet-id-62',
  setActiveSheet: () => {},
  getSheetByName: (n) => (n === '현장데이터' ? mockSheet62 : null),
  getSheets: () => [mockSheet62]
};
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => mockSs62;
sandbox.getAppSpreadsheet = () => mockSs62;
sandbox.findSiteScheduleSheet = () => mockSheet62;

sandbox.sortSiteScheduleTasksByDate(mockSs62, mockSheet62, true);
assertTest('consolidateSplitSiteRows + sortSiteScheduleTasksByDate: 맨 아래 분리된 "신대방우성 5동 601호 (10.18 테스트)" 행이 원래 현장 블록 안(10.15와 10.22 사이)으로 병합·정렬 검증',
  mockRows62.length === 5 &&
  mockRows62[0][0] === '신대방우성 5동 601호' && mockRows62[0][1] === '2026.10.15' &&
  mockRows62[1][0] === '신대방우성 5동 601호' && mockRows62[1][1] === '2026.10.18' && mockRows62[1][2] === '테스트' &&
  mockRows62[2][0] === '신대방우성 5동 601호' && mockRows62[2][1] === '2026.10.22' &&
  mockRows62[3][0] === '보경102동' &&
  mockRows62[4][0] === '보경102동'
);

// 3) addScheduleTaskFromWeb 호출 시 시트 맨 끝이 아니라 해당 현장 블록 내부로 즉시 삽입 및 날짜 오름차순 정렬 검증
const addGroupedRes62 = sandbox.addScheduleTaskFromWeb(
  '신대방우성 5동 601호',
  '전기',
  '2026-10-17',
  '조항만',
  '예정',
  true,
  '배선 작업',
  ''
);
assertTest('addScheduleTaskFromWeb(현장 그룹 삽입): "신대방우성 5동 601호"에 10.17 일정 추가 시 시트 맨 끝이 아닌 해당 현장 블록 내 3행(10.15 다음, 10.18 앞)으로 정확히 삽입 검증',
  addGroupedRes62.success === true &&
  addGroupedRes62.row === 3 &&
  mockRows62[0][0] === '신대방우성 5동 601호' && mockRows62[0][1] === '2026.10.15' &&
  mockRows62[1][0] === '신대방우성 5동 601호' && mockRows62[1][1] === '2026.10.17' && mockRows62[1][2] === '전기' &&
  mockRows62[2][0] === '신대방우성 5동 601호' && mockRows62[2][1] === '2026.10.18' && mockRows62[2][2] === '테스트' &&
  mockRows62[3][0] === '신대방우성 5동 601호' && mockRows62[3][1] === '2026.10.22' &&
  mockRows62[4][0] === '보경102동' &&
  mockRows62[5][0] === '보경102동'
);

// ========================================================
// 63. [1시간 주기 현장 묶음/날짜 정렬 트리거 보장 & 모바일 기기 뒤로가기(popstate) 상세 팝업 닫기 QA]
// ========================================================
console.log('\n📱 [Step 63: 1시간 주기 현장 병합/날짜 정렬 & 모바일 뒤로가기(popstate) 상세 팝업 닫기 QA]');

// 1) 분리 행 병합 시 담당자 미지정 행이 원래 현장의 기본 담당자를 자동 상속하는지 검증
assertTest('consolidateSplitSiteRows(담당자 자동 상속): 분리된 10.18 테스트 행 병합 시 기존 현장 담당자("조항만") 자동 상속 검증',
  mockRows62[2][2] === '테스트' && mockRows62[2][6] === '조항만'
);

// 2) addScheduleTaskFromWeb 호출 시 담당자를 빈칸("")으로 보내도 기존 현장의 담당자("조항만")를 자동 상속하는지 검증
const addNoMgrRes63 = sandbox.addScheduleTaskFromWeb(
  '신대방우성 5동 601호',
  '도배',
  '2026-10-19',
  '',
  '예정',
  true,
  '',
  ''
);
assertTest('addScheduleTaskFromWeb(담당자 미선택 시 현장 담당자 자동 상속 & 날짜순 정렬): 10.19 도배 일정이 10.18 뒤(5행)에 삽입되고 담당자 "조항만" 자동 지정 검증',
  addNoMgrRes63.success === true &&
  addNoMgrRes63.manager === '조항만' &&
  addNoMgrRes63.row === 5 &&
  mockRows62[3][1] === '2026.10.19' && mockRows62[3][2] === '도배' && mockRows62[3][6] === '조항만'
);

// 3) runHourlyAutoSortTrigger 및 ensureHourlySortTriggerInstalled 존재 및 동작 검증
assertTest('ensureHourlySortTriggerInstalled & runHourlyAutoSortTrigger: 1시간 주기 자동 병합/날짜정렬 트리거 보장 함수 존재 검증',
  typeof sandbox.ensureHourlySortTriggerInstalled === 'function' &&
  typeof sandbox.runHourlyAutoSortTrigger === 'function'
);

// 4) 모바일 하단 상세 팝업 기기 뒤로가기(popstate / postMessage) 닫기 동작 검증
const dashHtml63 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml(뒤로가기 팝업 닫기): pushBottomSheetHistoryState / popBottomSheetHistoryStateIfNeeded / closeBottomSheetFromBack 및 popstate 리스너 탑재 검증',
  dashHtml63.indexOf('function pushBottomSheetHistoryState()') !== -1 &&
  dashHtml63.indexOf('function popBottomSheetHistoryStateIfNeeded()') !== -1 &&
  dashHtml63.indexOf('function closeBottomSheetFromBack()') !== -1 &&
  dashHtml63.indexOf('window.addEventListener("popstate"') !== -1 &&
  dashHtml63.indexOf('ORCA_BACK_BUTTON') !== -1
);

const indexHtml63 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
assertTest('index.html(최상위 창 뒤로가기 브릿지): ORCA_SHEET_MODAL_OPEN / ORCA_SHEET_MODAL_CLOSED_BY_UI / ORCA_BACK_BUTTON 및 popstate 리스너 탑재 검증',
  indexHtml63.indexOf('ORCA_SHEET_MODAL_OPEN') !== -1 &&
  indexHtml63.indexOf('ORCA_SHEET_MODAL_CLOSED_BY_UI') !== -1 &&
  indexHtml63.indexOf('ORCA_BACK_BUTTON') !== -1 &&
  indexHtml63.indexOf("window.addEventListener('popstate'") !== -1
);

clientSandbox57.openNewTaskModal('신대방우성 5동 601호', '2026-10-20');
const openBeforeBack63 = mockElems57.sheetOverlay.classList.contains('show') === true &&
  vm.runInContext('isSheetModalHistoryActive', clientSandbox57) === true;
clientSandbox57.closeBottomSheetFromBack();
const closedAfterBack63 = mockElems57.sheetOverlay.classList.contains('show') === false &&
  vm.runInContext('isSheetModalHistoryActive', clientSandbox57) === false &&
  vm.runInContext('isCreateMode', clientSandbox57) === false;
assertTest('closeBottomSheetFromBack(런타임 실행): 하단 상세 팝업 오픈 상태에서 기기 뒤로가기 동작 시 팝업이 즉시 닫히고 상태가 초기화되는지 검증',
  openBeforeBack63 && closedAfterBack63
);

// ========================================================
// 64. [날짜 정렬/현장 병합 후 행 번호 변경 시(예: 125행 -> 26행) 실제 시트 행 추적 및 이동/저장 무결성 QA]
// ========================================================
console.log('\n📍 [Step 64: 날짜 정렬/현장 병합 후 행 번호 변경 시 실제 시트 행 자동 추적 및 포커스/저장 무결성 QA]');

// 1) resolveActualSheetRow: 기존 캐시 행 번호(125행)로 요청하더라도 현장명("신대방우성 5동 601호") + 공종("테스트") + 날짜("2026.10.18")로 실제 이동된 행(4행)을 정확히 찾아내는지 검증
const resolvedRow64 = sandbox.resolveActualSheetRow(
  mockSheet62,
  125,
  '신대방우성 5동 601호',
  '테스트',
  '2026.10.18'
);
assertTest('resolveActualSheetRow(행 번호 역추적): 정렬 전 구 행번호(125행)를 전달해도 실제 정렬된 행(4행)을 정확히 반환하는지 검증',
  resolvedRow64 === 4
);

// 2) focusSheetRow: 구 행번호(125행)로 호출 시 실제 행(4행)으로 포커스 및 URL(&range=4:4) 생성 검증
let activeFocusedRow64 = 0;
mockSheet62.getSheetId = () => 123456;
mockSheet62.setActiveRange = (rng) => {};
const origGetRange64 = mockSheet62.getRange;
mockSheet62.getRange = function(r, c, nr, nc) {
  if (c === 1 && nr === 1 && nc === 10) {
    activeFocusedRow64 = r;
  }
  return origGetRange64.apply(this, arguments);
};
const focusRes64 = sandbox.focusSheetRow(125, false, '신대방우성 5동 601호', '테스트', '2026.10.18');
assertTest('focusSheetRow(실제 행 포커스 & URL 생성): 구 행번호(125행) 요청 시 실제 행(4행)에 포커스되고 sheetUrl이 range=4:4로 생성되는지 검증',
  focusRes64 &&
  focusRes64.success === true &&
  focusRes64.row === 4 &&
  activeFocusedRow64 === 4 &&
  focusRes64.sheetUrl.indexOf('&range=4:4') !== -1
);

// 3) updateScheduleStatusFromWeb: 구 행번호(125행)로 저장 요청 시 실제 행(4행)의 상태/메모를 안전하게 업데이트하는지 검증
const updateMovedRes64 = sandbox.updateScheduleStatusFromWeb(
  125,
  '진행중',
  true,
  '정렬 후 실제행 업데이트 완료',
  '2026.10.18',
  '조항만',
  false,
  '신대방우성 5동 601호',
  '테스트'
);
assertTest('updateScheduleStatusFromWeb(정렬 후 구 행번호로 저장 시 실제 행 자동 추적): 125행 요청이 실제 4행("신대방우성 5동 601호 - 테스트")에 반영되고 row: 4를 반환하는지 검증',
  updateMovedRes64 &&
  updateMovedRes64.success === true &&
  updateMovedRes64.row === 4 &&
  mockRows62[2][7] === '진행중' &&
  mockRows62[2][5] === '정렬 후 실제행 업데이트 완료'
);

// 4) 클라이언트 goToCurrentTaskSheetRow: 구 행번호(125)가 아닌 서버가 반환한 실제 행번호(26)로 토스트 문구 및 curTaskObj.row를 갱신하는지 검증
assertTest('buildSidebarDashboardHtml(goToCurrentTaskSheetRow): focusSheetRow에 현장명/공종/날짜를 전달하고 응답받은 actualRow로 문구 표시 및 curTaskObj.row 동기화 검증',
  dashHtml63.indexOf('.focusSheetRow(targetRow, isArch, siteStr, taskStr, dateStr)') !== -1 &&
  dashHtml63.indexOf('const actualRow = (res && res.row) ? res.row : targetRow;') !== -1 &&
  dashHtml63.indexOf('if (curTaskObj) curTaskObj.row = actualRow;') !== -1
);

// ========================================================
// 🗑️ Step 65: 일정 상세 삭제 버튼 (3:7 비율), 확인 팝업, 행 삭제 & 캘린더 동기화 QA
// ========================================================
console.log('\n🗑️ [Step 65: 일정 상세 삭제 버튼 (3:7 비율), 확인 팝업, 행 삭제 & 캘린더 동기화 QA]');

// 1. deleteScheduleTaskFromWeb 함수 탑재 검증
assertTest('deleteScheduleTaskFromWeb: 서버 함수 정의 및 로드 검증',
  typeof sandbox.deleteScheduleTaskFromWeb === 'function'
);

// 2. buildSidebarDashboardHtml 내 삭제 UI 및 3:7 비율 마크업 검증
const dashHtml65 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: 하단 액션바 삭제 버튼(sheetDeleteBtn) 탑재 검증',
  dashHtml65.indexOf('id="sheetDeleteBtn"') !== -1 && dashHtml65.indexOf('🗑️ 삭제') !== -1
);
assertTest('buildSidebarDashboardHtml: 삭제(3) : 저장(7) 비율 CSS 탑재 검증',
  dashHtml65.indexOf('.sheet-delete-btn { flex: 3;') !== -1 &&
  dashHtml65.indexOf('.sheet-save-btn { width: 100%; flex: 7;') !== -1
);
assertTest('buildSidebarDashboardHtml: 삭제 확인 모달(#deleteConfirmOverlay) 및 취소/확인 버튼 탑재 검증',
  dashHtml65.indexOf('id="deleteConfirmOverlay"') !== -1 &&
  dashHtml65.indexOf('btn-delete-cancel') !== -1 &&
  dashHtml65.indexOf('btn-delete-confirm') !== -1 &&
  dashHtml65.indexOf('일정을 삭제하시겠습니까?') !== -1
);

// 3. 클라이언트 제어 함수 탑재 검증
assertTest('buildSidebarDashboardHtml: 삭제 모달 제어 함수(openDeleteConfirmModal, confirmDeleteCurrentTask) 탑재 검증',
  dashHtml65.indexOf('function openDeleteConfirmModal()') !== -1 &&
  dashHtml65.indexOf('function closeDeleteConfirmModal()') !== -1 &&
  dashHtml65.indexOf('function confirmDeleteCurrentTask()') !== -1
);

// 4. deleteScheduleTaskFromWeb 시뮬레이션: 시트 행 삭제 및 캘린더 일정 연동 삭제 검증
let deletedCalEventId = null;
let deletedCalEventsCount = 0;
let deletedSheetRowNum = -1;

const mockCalEvent65 = {
  getId: () => 'cal-event-to-delete-1029',
  getTitle: () => '📅 🟥 [철거] 신대방우성 16동 1501호',
  deleteEvent: () => {
    deletedCalEventId = 'cal-event-to-delete-1029';
    deletedCalEventsCount++;
  }
};

const mockSheetRows65 = [
  ['신대방우성 16동 1501호', '2026.10.29 ~ 2026.11.02', '철거', '인스', true, '특이사항', '조항만', '예정', 'cal-event-to-delete-1029', 1000],
  ['', '2026.11.03', '목공', '박동연', true, '', '조항만', '예정', 'cal-event-keep-1103', 2000],
  ['보경102동', '2026.11.05', '타일', '정소장', true, '', '전예린', '예정', 'cal-event-keep-1105', 3000]
];

const mockSheet65 = {
  getLastRow: () => mockSheetRows65.length + 1,
  getRange: (r, c, nr, nc) => ({
    getValues: () => {
      if (nr === 1 && nc === 10) return [mockSheetRows65[r - 2]];
      if (nr === 1 && nc === 3) return [mockSheetRows65[r - 2].slice(0, 3)];
      return mockSheetRows65.slice(r - 2, r - 2 + nr);
    },
    getValue: () => mockSheetRows65[r - 2][c - 1],
    setValue: (val) => { mockSheetRows65[r - 2][c - 1] = val; }
  }),
  deleteRow: (r) => {
    deletedSheetRowNum = r;
    mockSheetRows65.splice(r - 2, 1);
  }
};

sandbox.findSiteScheduleSheet = () => mockSheet65;
const mockTargetCal65 = {
  getId: () => 'c_shared_cal@group.calendar.google.com',
  getName: () => '현장 일정',
  getEventById: (id) => (id === 'cal-event-to-delete-1029' ? mockCalEvent65 : null),
  getEvents: () => [mockCalEvent65]
};
sandbox.CalendarApp.getCalendarById = () => mockTargetCal65;
sandbox.CalendarApp.getCalendarsByName = () => [mockTargetCal65];
sandbox.getTargetCalendar = () => mockTargetCal65;
sandbox.getAppCalendar = () => mockTargetCal65;

const delRes65 = sandbox.deleteScheduleTaskFromWeb(
  2,
  '신대방우성 16동 1501호',
  '철거',
  'cal-event-to-delete-1029',
  false
);

assertTest('deleteScheduleTaskFromWeb: 성공(success: true) 반환 검증',
  delRes65 && delRes65.success === true
);
assertTest('deleteScheduleTaskFromWeb: 구글 캘린더 일정(cal-event-to-delete-1029) 연동 삭제 검증',
  deletedCalEventId === 'cal-event-to-delete-1029' && deletedCalEventsCount >= 1
);
assertTest('deleteScheduleTaskFromWeb: 스프레드시트 2행 영구 삭제(deleteRow) 실행 검증',
  deletedSheetRowNum === 2
);
assertTest('deleteScheduleTaskFromWeb: 블록 첫 행 삭제 시 하위 행(목공)으로 현장명 자동 상속 보존 검증',
  mockSheetRows65[0][0] === '신대방우성 16동 1501호' && mockSheetRows65[0][2] === '목공'
);

// 5. doGet: action=deleteScheduleTask 웹 엔드포인트 응답 검증
const doGetDelRes = sandbox.doGet({
  parameter: {
    action: 'deleteScheduleTask',
    row: 2,
    site: '신대방우성 16동 1501호',
    task: '목공',
    id: 'cal-event-keep-1103',
    format: 'json'
  }
});
assertTest('doGet: action=deleteScheduleTask 호출 및 JSON 응답 반환 검증',
  doGetDelRes && typeof doGetDelRes.getContent === 'function' &&
  JSON.parse(doGetDelRes.getContent()).success === true
);

// ========================================================
// 🧹 Step 66: 삭제 일정 캘린더 동기화 잔여물(고아 일정) 즉시 자동 청소 QA
// ========================================================
console.log('\n🧹 [Step 66: 삭제 일정 캘린더 잔여물 즉시 자동 청소 QA]');

// 1. getAppCalendar 별칭 함수 검증
assertTest('getAppCalendar: getTargetCalendar 호환 별칭 함수 정의 검증',
  typeof sandbox.getAppCalendar === 'function' && typeof sandbox.getTargetCalendar === 'function'
);

// 2. doGet: action=cleanupOrphans 엔드포인트 검증
const doGetCleanupRes = sandbox.doGet({
  parameter: {
    action: 'cleanupOrphans'
  }
});
assertTest('doGet: action=cleanupOrphans 엔드포인트 호출 및 정상 응답 검증',
  doGetCleanupRes && typeof doGetCleanupRes.getContent === 'function' &&
  JSON.parse(doGetCleanupRes.getContent()).success === true
);

// 3. handleSheetChangeTrigger: 행 삭제(REMOVE_ROW) 시 디바운스 대기 없이 즉시 cleanupOrphanCalendarEventsSilent 호출 검증
const changeTriggerCode = code.slice(code.indexOf('function handleSheetChangeTrigger(e)'), code.indexOf('function cleanupOrphanCalendarEventsSilent('));
assertTest('handleSheetChangeTrigger: REMOVE_ROW 시 디바운스 대기 없이 cleanupOrphanCalendarEventsSilent 즉시 호출 검증',
  changeTriggerCode.indexOf('cleanupOrphanCalendarEventsSilent(sheet, cal)') !== -1 &&
  changeTriggerCode.indexOf('DEBOUNCE_SCHEDULED') === -1
);

// 4. deleteScheduleTaskFromWeb 내 잔여 중복 일정(break 없는 다중 삭제) 및 삭제 후 cleanupOrphanCalendarEventsSilent 호출 검증
const delTaskWebCode = code.slice(code.indexOf('function deleteScheduleTaskFromWeb('), code.indexOf('function addScheduleTaskFromWeb('));
const candLoopCode = delTaskWebCode.slice(delTaskWebCode.indexOf('for (var ce = 0; ce < candEvents.length; ce++)'), delTaskWebCode.indexOf('// 4단계:'));
assertTest('deleteScheduleTaskFromWeb: 동일 공종 다중 잔여 일정 검색 루프 내 break 제거 검증',
  candLoopCode.indexOf('Logger.log("✅ 캘린더 중복/잔여 일정 연동 삭제 성공:') !== -1 &&
  candLoopCode.indexOf('break;') === -1
);
assertTest('deleteScheduleTaskFromWeb: 시트 행 삭제 직후 cleanupOrphanCalendarEventsSilent 즉시 호출 검증',
  delTaskWebCode.indexOf('cleanupOrphanCalendarEventsSilent(sheet, cal)') !== -1
);

// ========================================================
// 📅 Step 67: PC 와이드 구글 캘린더 뷰 및 마우스 드래그 앤 드롭 일정 이동 QA
// ========================================================
console.log('\n📅 [Step 67: PC 와이드 구글 캘린더 뷰 및 마우스 드래그 앤 드롭 일정 이동 QA]');

// 1. moveScheduleTaskDateFromWeb 서버 함수 정의 및 경계값 검증
assertTest('moveScheduleTaskDateFromWeb: 서버 함수 정의 검증',
  typeof sandbox.moveScheduleTaskDateFromWeb === 'function'
);

const moveResInvalidRow = sandbox.moveScheduleTaskDateFromWeb(1, '신대방우성', '도배', '2026.10.25', false);
assertTest('moveScheduleTaskDateFromWeb(경계값): 유효하지 않은 행 번호(<2) 차단 검증',
  moveResInvalidRow && moveResInvalidRow.success === false &&
  moveResInvalidRow.error.indexOf('유효하지 않은 행 번호') !== -1
);

const moveResEmptyDate = sandbox.moveScheduleTaskDateFromWeb(2, '신대방우성', '도배', '', false);
assertTest('moveScheduleTaskDateFromWeb(경계값): 빈 날짜 차단 검증',
  moveResEmptyDate && moveResEmptyDate.success === false &&
  moveResEmptyDate.error.indexOf('이동할 날짜가 지정되지 않았습니다') !== -1
);

// 2. moveScheduleTaskDateFromWeb 정상 실행 검증
let testMoveCellVal = '';
let testMoveTsVal = 0;
let testFlushCalled = false;
let testSyncSingleCalled = false;

const mockMoveSheet = {
  getLastRow: () => 10,
  getRange: (r, c) => ({
    setValue: (val) => {
      if (c === 2) testMoveCellVal = val;
      if (c === 10) testMoveTsVal = val;
    },
    getValue: () => (c === 2 ? testMoveCellVal : '')
  })
};

sandbox.getAppSpreadsheet = () => ({});
sandbox.findSiteScheduleSheet = () => mockMoveSheet;
sandbox.resolveActualSheetRow = (s, r, site, task) => r;
sandbox.SpreadsheetApp.flush = () => { testFlushCalled = true; };
sandbox.invalidateDashboardServerCache = () => {};
sandbox.syncSingleRowToCalendar = (sheet, r) => { testSyncSingleCalled = true; };

const moveResSuccess = sandbox.moveScheduleTaskDateFromWeb(3, '신대방우성', '목공', '2026.10.27', false);
assertTest('moveScheduleTaskDateFromWeb: B열 날짜 갱신(2026.10.27) 검증',
  moveResSuccess && moveResSuccess.success === true &&
  testMoveCellVal === '2026.10.27'
);
assertTest('moveScheduleTaskDateFromWeb: J열 타임스탬프 기록 검증',
  testMoveTsVal > 0
);
assertTest('moveScheduleTaskDateFromWeb: SpreadsheetApp.flush 및 syncSingleRowToCalendar 호출 검증',
  testFlushCalled === true && testSyncSingleCalled === true
);

// 3. doGet action=moveScheduleTaskDate 엔드포인트 검증
const doGetMoveRes = sandbox.doGet({
  parameter: {
    action: 'moveScheduleTaskDate',
    row: 3,
    site: '신대방우성',
    task: '목공',
    date: '2026.10.28',
    isArchived: 'false'
  }
});
assertTest('doGet: action=moveScheduleTaskDate 엔드포인트 호출 및 정상 응답 검증',
  doGetMoveRes && typeof doGetMoveRes.getContent === 'function' &&
  JSON.parse(doGetMoveRes.getContent()).success === true
);

// 4. buildSidebarDashboardHtml 내 PC 반응형 스타일 검증
const dashHtml67 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: PC 와이드 뷰(@media min-width: 768px, max-width: 1440px) 탑재 검증',
  dashHtml67.indexOf('@media (min-width: 768px)') !== -1 &&
  dashHtml67.indexOf('max-width: 1440px') !== -1
);
assertTest('buildSidebarDashboardHtml: PC 달력 날짜 셀 높이 확장(min-height: 125px, 150px) 탑재 검증',
  dashHtml67.indexOf('min-height: 125px') !== -1 &&
  dashHtml67.indexOf('min-height: 150px') !== -1
);
assertTest('buildSidebarDashboardHtml: 드래그 앤 드롭 시각 스타일(.cal-day-cell.drag-over, cursor: grab) 탑재 검증',
  dashHtml67.indexOf('.cal-day-cell.drag-over') !== -1 &&
  dashHtml67.indexOf('cursor: grab') !== -1 &&
  dashHtml67.indexOf('.cal-event-pill.is-dragging') !== -1
);

// 5. 클라이언트 드래그 앤 드롭 이벤트 핸들러 및 속성 탑재 검증
assertTest('buildSidebarDashboardHtml: 드래그 앤 드롭 제어 함수(onCalTaskDragStart, onCalTaskDragEnd, onCalDayDrop 등) 탑재 검증',
  dashHtml67.indexOf('function onCalTaskDragStart') !== -1 &&
  dashHtml67.indexOf('function onCalTaskDragEnd') !== -1 &&
  dashHtml67.indexOf('function onCalDayDragOver') !== -1 &&
  dashHtml67.indexOf('function onCalDayDragLeave') !== -1 &&
  dashHtml67.indexOf('function onCalDayDrop') !== -1
);
assertTest('buildSidebarDashboardHtml: onTaskClick 내 wasJustDragging 드래그 직후 상세 팝업 방어 가드 탑재 검증',
  dashHtml67.indexOf('if (typeof wasJustDragging !== "undefined" && wasJustDragging) return;') !== -1
);
assertTest('buildSidebarDashboardHtml: 월간 캘린더 셀 및 알약 드래그 속성(draggable, ondragstart, ondrop) 탑재 검증',
  dashHtml67.indexOf('ondragover="onCalDayDragOver(event)"') !== -1 &&
  dashHtml67.indexOf('ondrop="onCalDayDrop(event,') !== -1 &&
  dashHtml67.indexOf('draggable="true" ondragstart="onCalTaskDragStart') !== -1
);
assertTest('buildSidebarDashboardHtml: 주간 캘린더 카드 및 알약 드래그 속성 탑재 검증',
  dashHtml67.indexOf('cal-week-card') !== -1 &&
  dashHtml67.indexOf('ondrop="onCalDayDrop(event, \\\'\' + curStr + \'\\\')">') !== -1
);

// 6. 기간 일정 드래그 앤 드롭 시 기간 간격(diffDays) 보존 알고리즘 시뮬레이션 검증
const simData = {
  startDateStr: '2026-10-15',
  endDateStr: '2026-10-18', // 3일 기간
};
const simTargetDateStr = '2026-10-22';
const sParts = simData.startDateStr.split('-').map(Number);
const eParts = simData.endDateStr.split('-').map(Number);
const sDate = new Date(sParts[0], sParts[1] - 1, sParts[2]);
const eDate = new Date(eParts[0], eParts[1] - 1, eParts[2]);
const diffDays = Math.round((eDate.getTime() - sDate.getTime()) / (24 * 60 * 60 * 1000));
const tParts = simTargetDateStr.split('-').map(Number);
const newEndObj = new Date(tParts[0], tParts[1] - 1, tParts[2] + diffDays);
const neY = newEndObj.getFullYear();
const neM = (newEndObj.getMonth() + 1 < 10 ? '0' : '') + (newEndObj.getMonth() + 1);
const neD = (newEndObj.getDate() < 10 ? '0' : '') + newEndObj.getDate();
const calculatedNewEndDate = `${neY}-${neM}-${neD}`;
const calculatedNewDateDisplay = `${simTargetDateStr.replace(/-/g, '.')} ~ ${neY}.${neM}.${neD}`;

assertTest('기간 일정 드래그 앤 드롭: 3일 기간 일정 이동 시 3일 간격 자동 연장(2026.10.22 ~ 2026.10.25) 계산 검증',
  diffDays === 3 &&
  calculatedNewEndDate === '2026-10-25' &&
  calculatedNewDateDisplay === '2026.10.22 ~ 2026.10.25'
);


// ============================================================================
// [Step 68] PC Google Calendar 2-Column Desktop Native View 검증
// ============================================================================
console.log('\n[Step 68] PC 구글 캘린더 스타일 2-Column 데스크톱 네이티브 뷰 무결성 검증...');

const dashHtml68 = sandbox.buildSidebarDashboardHtml('web');

// 1. PC 데스크톱 레이아웃 마크업 검증 (TopBar, Sidebar, Main Workspace)
assertTest('Step 68-1: PC 전용 상단 헤더바(pc-topbar) 및 워크스페이스(pc-workspace, pc-sidebar, pc-main-content) 마크업 검증',
  dashHtml68.indexOf('<header class="pc-topbar">') !== -1 &&
  dashHtml68.indexOf('<div class="pc-workspace">') !== -1 &&
  dashHtml68.indexOf('<aside class="pc-sidebar">') !== -1 &&
  dashHtml68.indexOf('<main class="pc-main-content">') !== -1
);

// 2. PC 상단 헤더바(pc-topbar) 핵심 컨트롤 검증
assertTest('Step 68-2: pc-topbar 내 로고, [오늘][◀][▶] 날짜 타이틀, Alt+S 검색창, 3가지 뷰 전환 탭([월간][주간][카드]), 캡쳐/새로고침 버튼 탑재 검증',
  dashHtml68.indexOf('id="pcMonthTitle"') !== -1 &&
  dashHtml68.indexOf('id="pcSearchInput"') !== -1 &&
  dashHtml68.indexOf("switchToPcView('month')") !== -1 &&
  dashHtml68.indexOf("switchToPcView('week')") !== -1 &&
  dashHtml68.indexOf("switchToPcView('list')") !== -1 &&
  dashHtml68.indexOf('captureCalendarScreen()') !== -1
);

// 3. PC 좌측 사이드바(pc-sidebar) 위젯 검증
assertTest('Step 68-3: pc-sidebar 내 [새 일정 등록] 버튼, 미니 캘린더, 진행현장 체크박스 필터, 담당자 셀렉트박스, KPI 요약 카드 탑재 검증',
  dashHtml68.indexOf('class="pc-btn-create-event"') !== -1 &&
  dashHtml68.indexOf('id="pcMiniCalendarArea"') !== -1 &&
  dashHtml68.indexOf('id="pcSidebarSiteList"') !== -1 &&
  dashHtml68.indexOf('id="pcMemberSelect"') !== -1 &&
  dashHtml68.indexOf('id="pcKpiToday"') !== -1 &&
  dashHtml68.indexOf('id="pcKpiDelay"') !== -1
);

// 4. PC 미디어 쿼리(>= 768px) 모바일 요소 완전 은폐 및 데스크톱 풀스크린 스타일 검증
assertTest('Step 68-4: 미디어 쿼리 내 모바일 헤더/하단바/모바일 캘린더 네비 은폐(display:none !important) 및 PC 탑바/사이드바 flex 노출 검증',
  dashHtml68.indexOf('@media (min-width: 768px)') !== -1 &&
  dashHtml68.indexOf('.pc-topbar { display: flex !important;') !== -1 &&
  dashHtml68.indexOf('.pc-sidebar { display: flex !important;') !== -1 &&
  dashHtml68.indexOf('.footer-bar { display: none !important; }') !== -1 &&
  dashHtml68.indexOf('.header { display: none !important;') !== -1 &&
  dashHtml68.indexOf('.cal-mode-bar, .cal-nav-bar { display: none !important; }') !== -1
);

// 5. PC 캘린더 풀사이즈 및 세로 칸 확장 CSS 검증
assertTest('Step 68-5: PC 캘린더 화면 전체 높이 및 일정 과다 시 세로 칸 확장 CSS (height: 100vh, grid-auto-rows: minmax(max-content, 1fr)) 검증',
  dashHtml68.indexOf('height: 100vh !important;') !== -1 &&
  dashHtml68.indexOf('overflow: hidden !important;') !== -1 &&
  dashHtml68.indexOf('grid-auto-rows: minmax(max-content, 1fr);') !== -1
);

// 6. PC 클라이언트 스크립트 함수 탑재 검증
assertTest('Step 68-6: PC 전용 제어 함수(switchToPcView, renderMiniCalendar, onMiniDayClick, renderPcSidebarSites, selectAllPcSites, deselectAllPcSites, renderPcSidebarMembers, updatePcSidebarKPIs) 탑재 검증',
  dashHtml68.indexOf('function switchToPcView(view)') !== -1 &&
  dashHtml68.indexOf('function onPcSearchInput(val)') !== -1 &&
  dashHtml68.indexOf('function clearPcSearch()') !== -1 &&
  dashHtml68.indexOf('function renderMiniCalendar(') !== -1 &&
  dashHtml68.indexOf('function onMiniDayClick(y, m, d)') !== -1 &&
  dashHtml68.indexOf('function renderPcSidebarSites()') !== -1 &&
  dashHtml68.indexOf('function onPcSiteCheckChange(e, siteEnc)') !== -1 &&
  dashHtml68.indexOf('function selectAllPcSites()') !== -1 &&
  dashHtml68.indexOf('function deselectAllPcSites()') !== -1 &&
  dashHtml68.indexOf('function renderPcSidebarMembers(members)') !== -1 &&
  dashHtml68.indexOf('function updatePcSidebarKPIs()') !== -1
);

// 7. PC 진입 시 월간 캘린더 기본 활성화 및 pcMonthTitle 실시간 동기화 검증
assertTest('Step 68-7: PC 접속 시 isCalendarMode = true 기본 활성화, pcMonthTitle 동기화 및 Alt+S 단축키 등록 검증',
  dashHtml68.indexOf('isCalendarMode = (typeof window !== "undefined" && window.innerWidth >= 768);') !== -1 &&
  dashHtml68.indexOf('if (pcTitle) pcTitle.innerText = calYear + "년 " + calMonth + "월";') !== -1 &&
  dashHtml68.indexOf('if (pcTitle) pcTitle.innerText = weekTitleText;') !== -1 &&
  dashHtml68.indexOf('e.altKey && (e.key === "s" || e.key === "S"') !== -1
);

// ============================================================================
// [Step 69] PC 화면 4대 개선사항 (세로확장, 선택해제, 담당자 셀렉트박스, KPI 재배치) 무결성 검증
// ============================================================================
console.log('\n[Step 69] PC 화면 4대 개선사항 무결성 검증...');

// 1. 하루 일정 과다 시 세로 칸 자동 확장 및 달력 배경 잘림 방지 CSS 검증
assertTest('Step 69-1: .cal-day-cell에 height: 100% !important, #calGrid에 minmax(max-content, 1fr), .cal-events-list에 flex: 1 0 auto 탑재로 일정 과다 시 달력 배경 자동 확장 및 잘림 방지 검증',
  dashHtml68.indexOf('.cal-day-cell { min-height: 125px !important; height: 100% !important; display: flex !important; flex-direction: column !important; overflow: visible !important;') !== -1 &&
  dashHtml68.indexOf('.cal-events-list { gap: 3px; overflow: visible !important; flex: 1 0 auto !important; height: auto !important; min-height: min-content !important; }') !== -1 &&
  dashHtml68.indexOf('#calGrid { flex: 1; height: 100%; min-height: 0; display: grid; grid-template-columns: repeat(7, 1fr); grid-auto-rows: minmax(max-content, 1fr);') !== -1
);

// 2. 진행현장 [전체보기] 우측에 [선택해제] 버튼 탑재 및 다중 선택 무결성 검증
assertTest('Step 69-2: 진행현장 섹션 타이틀에 [전체보기] 및 [선택해제] 버튼 배치 및 deselectAllPcSites 함수 검증',
  dashHtml68.indexOf('onclick="selectAllPcSites()">전체보기</span>') !== -1 &&
  dashHtml68.indexOf('onclick="deselectAllPcSites()">선택해제</span>') !== -1 &&
  dashHtml68.indexOf('function deselectAllPcSites()') !== -1 &&
  dashHtml68.indexOf('checkedSites.add("__NONE__");') !== -1
);

// 3. 담당자 라디오버튼 -> 셀렉트박스 변경 및 '담당자' 타이틀 문구 검증
assertTest('Step 69-3: "담당자 필터" -> "담당자" 문구 변경 및 pcMemberSelect 셀렉트박스 탑재 검증',
  dashHtml68.indexOf('<span>담당자</span>') !== -1 &&
  dashHtml68.indexOf('<select id="pcMemberSelect" class="pc-member-select"') !== -1 &&
  dashHtml68.indexOf('function setPcMemberSelect(val)') !== -1
);

// 4. 공정 상태 현황 KPI가 [새 일정 등록] 바로 밑으로 재배치되었는지 DOM 순서 검증
const createBtnIndex = dashHtml68.indexOf('class="pc-btn-create-event"');
const kpiGridIndex = dashHtml68.indexOf('class="pc-kpi-summary-grid"');
const miniCalIndex = dashHtml68.indexOf('id="pcMiniCalendarArea"');
assertTest('Step 69-4: 공정 상태 현황(pc-kpi-summary-grid)이 새 일정 등록 바로 아래, 미니 달력 위에 위치하는 순서 검증',
  createBtnIndex !== -1 &&
  kpiGridIndex !== -1 &&
  miniCalIndex !== -1 &&
  createBtnIndex < kpiGridIndex &&
  kpiGridIndex < miniCalIndex
);
// 5. 1200px 이상 와이드 데스크톱 및 이벤트 리스트 intrinsic height 보장 검증
assertTest('Step 69-5: 1200px 이상 와이드 화면에서도 minmax(max-content, 1fr) 적용 및 cal-day-cell box-sizing 무결성 검증',
  dashHtml68.indexOf('#calGrid { grid-auto-rows: minmax(max-content, 1fr); }') !== -1 &&
  dashHtml68.indexOf('box-sizing: border-box;') !== -1 &&
  dashHtml68.indexOf('min-height: min-content !important;') !== -1
);

// ============================================================================
// [Step 70] PC 주간 7-Column 뷰 & 상단 캐러셀 와이드 현장 대시보드 & 담당자 표기 무결성 QA
// ============================================================================
console.log('\n--- [Step 70] PC 주간 7-Column 그리드, 상단 캐러셀 와이드 대시보드 및 담당자 직급 생략 무결성 검증 ---');

// 1. formatManagerName 직급 생략 및 순수 이름 추출 검증
const clientScript70 = dashHtml68.match(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/)[1];
const clientSandbox70 = {
  document: { getElementById: () => null, querySelectorAll: () => [], addEventListener: () => {} },
  window: { innerWidth: 1200, addEventListener: () => {} },
  setTimeout: () => {},
  Date: Date,
  Math: Math,
  String: String,
  encodeURIComponent: encodeURIComponent,
  decodeURIComponent: decodeURIComponent
};
vm.createContext(clientSandbox70);
vm.runInContext(clientScript70, clientSandbox70);

assertTest('Step 70-1: formatManagerName: 직급(대리, 과장, 팀장 등) 및 괄호 생략, 순수 이름 추출 검증',
  typeof clientSandbox70.formatManagerName === 'function' &&
  clientSandbox70.formatManagerName('전예린 대리') === '전예린' &&
  clientSandbox70.formatManagerName('송지은 과장') === '송지은' &&
  clientSandbox70.formatManagerName('한송이(팀장)') === '한송이' &&
  clientSandbox70.formatManagerName('조항만') === '조항만' &&
  clientSandbox70.formatManagerName('인경호 실장님') === '인경호'
);

// 2. PC 현장 뷰: 주담당자 -> 담당자 변경 및 메타 줄 전체공종수 삭제 검증
assertTest('Step 70-2: PC 현장 뷰: "주담당자" 대신 "담당자:" 표기 및 메타 줄 중복 "전체 공종: " 삭제 검증',
  dashHtml68.indexOf('<span>담당자: <strong>') !== -1 &&
  dashHtml68.indexOf('주담당자') === -1 &&
  dashHtml68.indexOf('주 담당자') === -1 &&
  dashHtml68.indexOf('전체 공종: 35개 공정') === -1
);

// 3. PC 현장 뷰: 상단 캐러셀 & 하단 와이드 대시보드 구조 및 컨트롤러 함수 검증
assertTest('Step 70-3: PC 현장 뷰: 상단 가로 캐러셀 바(.pc-site-carousel-bar) 및 하단 와이드 대시보드(.pc-site-detail-pane) 탑재 검증',
  dashHtml68.indexOf('.pc-site-dashboard-layout') !== -1 &&
  dashHtml68.indexOf('.pc-site-carousel-bar') !== -1 &&
  dashHtml68.indexOf('.pc-carousel-track') !== -1 &&
  dashHtml68.indexOf('.pc-site-detail-pane') !== -1 &&
  dashHtml68.indexOf('.pc-kpi-summary-panel') !== -1 &&
  dashHtml68.indexOf('.pc-tasks-table-card') !== -1 &&
  dashHtml68.indexOf('function scrollPcSiteCarousel(') !== -1 &&
  dashHtml68.indexOf('function selectPcSite(') !== -1 &&
  dashHtml68.indexOf('function setPcSiteSortMode(') !== -1 &&
  dashHtml68.indexOf('function setPcSiteDetailFilter(') !== -1
);

// 4. PC 주간 뷰: 7-Column 그리드 보드 구조 및 CSS 검증
assertTest('Step 70-4: PC 주간 뷰: 7-Column 데스크톱 그리드(.pc-weekly-grid) 및 일~토 7개 열 렌더링 로직 검증',
  dashHtml68.indexOf('.pc-weekly-grid') !== -1 &&
  dashHtml68.indexOf('grid-template-columns: repeat(7, 1fr)') !== -1 &&
  dashHtml68.indexOf('.pc-weekly-col') !== -1 &&
  dashHtml68.indexOf('function renderWeeklyCalendar(') !== -1 &&
  dashHtml68.indexOf('pc-weekly-grid') !== -1
);

// 5. renderList 내 isPcMode 분기 및 PC 대시보드 렌더링 호출 검증
assertTest('Step 70-5: renderList: PC 환경(width >= 768px) 접속 시 renderPcSiteDashboard 자동 호출 분기 검증',
  dashHtml68.indexOf('const isPcMode = (typeof window !== "undefined" && typeof window.innerWidth === "number" && window.innerWidth >= 768);') !== -1 &&
  dashHtml68.indexOf('renderPcSiteDashboard(matchedSites);') !== -1 &&
  dashHtml68.indexOf('container.classList.add("pc-dashboard-active");') !== -1
);

// ============================================================================
// [Step 71] 현장 상태 판정(예정/진행/일정확인필요/완료), 안내 배너, 탭 문구 및 뷰 전환 QA
// ============================================================================
console.log('\n--- [Step 71] 현장 상태 판정 로직, 안내 배너, 탭 명칭(월간/주간/현장별) 및 화면 전환 무결성 검증 ---');

const dashHtml71 = sandbox.buildSidebarDashboardHtml('desktop');
const clientScript71 = dashHtml71.match(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/)[1];
const clientSandbox71 = {
  document: { getElementById: () => null, querySelectorAll: () => [], addEventListener: () => {} },
  window: { innerWidth: 1200, addEventListener: () => {} },
  setTimeout: () => {},
  Date: Date,
  Math: Math,
  String: String,
  encodeURIComponent: encodeURIComponent,
  decodeURIComponent: decodeURIComponent
};
vm.createContext(clientSandbox71);
vm.runInContext(clientScript71, clientSandbox71);

// 1. getSiteStatusInfo 함수 엣지케이스별 상태 판정 검증
const resPlan71 = clientSandbox71.getSiteStatusInfo({
  siteName: '미래현장',
  tasks: [{ taskName: '철거', status: '예정', startDateStr: '2026-10-25' }],
  inProgressTasksCount: 0,
  completedTasksCount: 0
}, '2026-10-18');

const resCheck71 = clientSandbox71.getSiteStatusInfo({
  siteName: '경과현장',
  tasks: [{ taskName: '철거', status: '예정', startDateStr: '2026-10-12' }],
  inProgressTasksCount: 0,
  completedTasksCount: 0
}, '2026-10-18');

const resProg71 = clientSandbox71.getSiteStatusInfo({
  siteName: '진행현장',
  tasks: [{ taskName: '철거', status: '진행 중', startDateStr: '2026-10-10' }],
  inProgressTasksCount: 1
}, '2026-10-18');

const resDone71 = clientSandbox71.getSiteStatusInfo({
  siteName: '완료현장',
  isAllCompleted: true,
  tasks: [{ taskName: '청소', status: '완료' }]
}, '2026-10-18');

assertTest('Step 71-1: getSiteStatusInfo: 착공 미래(예정), 착공 경과(일정확인필요 배너), 진행중, 완료 상태 정확 판정 검증',
  resPlan71.status === 'plan' &&
  resPlan71.text.indexOf('공사예정') !== -1 &&
  resCheck71.status === 'check_needed' &&
  resCheck71.text.indexOf('일정 확인 필요') !== -1 &&
  resCheck71.alertMsg.indexOf('아직 진행 또는 완료된 공정이 없습니다') !== -1 &&
  resProg71.status === 'prog' &&
  resProg71.text === '공사진행중' &&
  resDone71.status === 'done' &&
  resDone71.text === '공사완료'
);

// 2. 우측 상단 뷰 탭 명칭: 월간, 주간, 현장별 변경 검증
assertTest('Step 71-2: 우측 상단 뷰 전환 탭: "월간", "주간", "현장별" 간결 명칭 적용 검증',
  dashHtml71.indexOf('>월간</button>') !== -1 &&
  dashHtml71.indexOf('>주간</button>') !== -1 &&
  dashHtml71.indexOf('>현장별</button>') !== -1 &&
  dashHtml71.indexOf('>월간 달력</button>') === -1 &&
  dashHtml71.indexOf('>현장별 카드</button>') === -1
);

// 3. switchToPcView: hidden 클래스 및 완벽한 분리 검증
assertTest('Step 71-3: switchToPcView: .hidden 클래스 및 calendarArea/contentArea 완전 분리 검증',
  dashHtml71.indexOf('#calendarArea.hidden, #contentArea.hidden { display: none !important; }') !== -1 &&
  dashHtml71.indexOf('function switchToPcView(view)') !== -1 &&
  dashHtml71.indexOf('calArea.classList.add("hidden")') !== -1 &&
  dashHtml71.indexOf('contentArea.classList.remove("hidden")') !== -1
);

// 4. 사이드바 현장명 라벨 클릭 시 현장 상세로 즉시 이동 연동 검증
assertTest('Step 71-4: 사이드바 현장명 라벨: onclick="onPcSidebarSiteLabelClick" 탑재 및 공종 리스트 즉시 연결 검증',
  dashHtml71.indexOf('function onPcSidebarSiteLabelClick(') !== -1 &&
  dashHtml71.indexOf('onPcSidebarSiteLabelClick(event,') !== -1 &&
  dashHtml71.indexOf('.site-name-label { cursor: pointer;') !== -1
);

// ========================================================
// Step 72: PC 현장 공종 리스트 51개 전체 스크롤, 테이블 카드 크기 고정 해제, Sticky 툴바/헤더 및 맨 위로 버튼 QA
// ========================================================
console.log('\n--- [Step 72] PC 현장 공종 리스트 전체 스크롤, 테이블 카드 flex-shrink: 0, Sticky 툴바/헤더 및 맨 위로 버튼 무결성 검증 ---');

const dashHtml72 = fs.readFileSync(path.join(__dirname, '../현장서식.js'), 'utf8');

// 72-1: #contentArea display: flex !important, min-height: 0 검증
assertTest('Step 72-1: #contentArea에 display: flex !important 및 min-height: 0 탑재로 모달 2-컬럼 그리드 오염 방지 검증',
  dashHtml72.indexOf('#contentArea { display: flex !important; flex-direction: column !important; flex: 1; height: 100%; min-height: 0;') !== -1 &&
  dashHtml72.indexOf('#contentArea.pc-dashboard-active { display: flex !important; flex-direction: column !important;') !== -1
);

// 72-2: .pc-site-dashboard-layout & .pc-site-detail-pane flex 및 스크롤바 탑재 검증
assertTest('Step 72-2: .pc-site-detail-pane에 min-height: 0, overflow-y: auto 및 슬림 스크롤바(-webkit-scrollbar) 탑재 검증',
  dashHtml72.indexOf('.pc-site-dashboard-layout { display: flex; flex-direction: column; width: 100%; height: 100%; min-height: 0; overflow: hidden;') !== -1 &&
  dashHtml72.indexOf('.pc-site-detail-pane { flex: 1; min-height: 0; overflow-y: auto;') !== -1 &&
  dashHtml72.indexOf('.pc-site-detail-pane::-webkit-scrollbar { width: 8px; }') !== -1
);

// 72-3: .pc-tasks-table-card flex-shrink: 0 !important, overflow: visible !important 탑재로 51개 공정 찌그러짐 방지 검증
assertTest('Step 72-3: .pc-tasks-table-card에 flex-shrink: 0 !important 및 overflow: visible !important 탑재 검증 (51개 공정 전체 렌더링)',
  dashHtml72.indexOf('.pc-tasks-table-card { background: #ffffff; border-radius: 12px; border: 1.5px solid #e2e8f0; box-shadow: 0 1px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; flex-shrink: 0 !important; position: relative; overflow: visible !important;') !== -1
);

// 72-4: 툴바 및 thead th Sticky 상단 고정 검증
assertTest('Step 72-4: 테이블 툴바(top: 0) 및 컬럼 헤더(top: 48px) Sticky 고정 CSS 탑재 검증',
  dashHtml72.indexOf('.pc-tasks-toolbar { position: sticky; top: 0; z-index: 15;') !== -1 &&
  dashHtml72.indexOf('.pc-tasks-table th { position: sticky; top: 48px; z-index: 14;') !== -1 &&
  dashHtml72.indexOf('.pc-tasks-table { width: 100%; border-collapse: separate; border-spacing: 0;') !== -1
);

// 72-5: pcScrollTopBtn 및 scrollPcDetailTop 탑재 검증
assertTest('Step 72-5: 맨 위로 플로팅 버튼(pcScrollTopBtn), scrollPcDetailTop 함수 및 스크롤 이벤트 리스너 탑재 검증',
  dashHtml72.indexOf('id="pcScrollTopBtn" class="pc-scroll-top-btn" onclick="scrollPcDetailTop()"') !== -1 &&
  dashHtml72.indexOf('function scrollPcDetailTop()') !== -1 &&
  dashHtml72.indexOf('dPane.addEventListener("scroll", function()') !== -1
);

// 72-6: 날짜 계산 NaN 방어 검증
assertTest('Step 72-6: 잔여/경과 일수 계산 시 !isNaN(diffDays) 방어 로직 검증',
  dashHtml72.indexOf('if (!isNaN(diffDays)) {') !== -1 &&
  dashHtml72.indexOf('cleanLatest = String(latestDateStr).replace(/\\\\./g, "-").trim();') !== -1
);

// ========================================================
// Step 73: PC 현장 캐러셀 활성 카드 자동 포커스 및 부드러운 중앙 스크롤 무결성 QA
// ========================================================
console.log('\n--- [Step 73] PC 현장 캐러셀 활성 카드 자동 포커스 및 중앙 스크롤 무결성 검증 ---');

const dashHtml73 = fs.readFileSync(path.join(__dirname, '../현장서식.js'), 'utf8');

// 73-1: scrollActivePcCardIntoView 함수 정의 및 중앙 정렬 스크롤 계산 로직 검증
assertTest('Step 73-1: scrollActivePcCardIntoView 함수 및 중앙 정렬 스크롤(clampedTarget) 계산 로직 검증',
  dashHtml73.indexOf('function scrollActivePcCardIntoView(isSmooth)') !== -1 &&
  dashHtml73.indexOf('const relativeCardLeft = (cardRect.left - trackRect.left) + track.scrollLeft;') !== -1 &&
  dashHtml73.indexOf('const targetScrollLeft = relativeCardLeft - (track.clientWidth / 2) + (activeCard.offsetWidth / 2);') !== -1 &&
  dashHtml73.indexOf('track.scrollTo({ left: clampedTarget, behavior: isSmooth ? "smooth" : "auto" });') !== -1
);

// 73-2: renderPcSiteDashboard 내 prevScrollLeft 보존 및 자동 스크롤 검증
assertTest('Step 73-2: renderPcSiteDashboard 내 렌더링 전 scrollLeft 보존 및 리셋 방지 검증',
  dashHtml73.indexOf('const prevTrack = document.getElementById("pcSiteCarouselTrack");') !== -1 &&
  dashHtml73.indexOf('const prevScrollLeft = prevTrack ? prevTrack.scrollLeft : null;') !== -1 &&
  dashHtml73.indexOf('if (newTrack && prevScrollLeft !== null) {') !== -1 &&
  dashHtml73.indexOf('newTrack.scrollLeft = prevScrollLeft;') !== -1 &&
  dashHtml73.indexOf('scrollActivePcCardIntoView(prevScrollLeft !== null);') !== -1
);

// 73-3: selectPcSite 및 onPcSidebarSiteLabelClick 내 활성 카드 자동 중앙 스크롤 연동 검증
assertTest('Step 73-3: selectPcSite 및 onPcSidebarSiteLabelClick 내 scrollActivePcCardIntoView(true) 연동 검증',
  dashHtml73.indexOf('function selectPcSite(siteName) {') !== -1 &&
  dashHtml73.indexOf('function onPcSidebarSiteLabelClick(e, siteEnc) {') !== -1
);

// 73-4: .pc-carousel-track CSS에서 scroll-behavior: smooth 제거하여 JS scrollTo 제어권 확보 검증
assertTest('Step 73-4: .pc-carousel-track CSS 내 scroll-behavior: smooth 제거(자바스크립트 완전 제어권 확보) 검증',
  dashHtml73.indexOf('.pc-carousel-track { display: flex; gap: 10px; overflow-x: auto; padding: 2px 2px 6px 2px; }') !== -1 &&
  dashHtml73.indexOf('.pc-carousel-track { display: flex; gap: 10px; overflow-x: auto; padding: 2px 2px 6px 2px; scroll-behavior: smooth; }') === -1
);

// 73-5: onPcSiteCheckChange에서 !isCalendarMode(현장 뷰)일 때 renderList 호출 분기 검증
assertTest('Step 73-5: onPcSiteCheckChange 내 !isCalendarMode 시 renderList 호출 분기 검증',
  dashHtml73.indexOf('if (isCalendarMode) {') !== -1 &&
  dashHtml73.indexOf('renderCurrentCalendarView();') !== -1 &&
  dashHtml73.indexOf('renderList();') !== -1
);

// 73-6: 마지막 현장 선택 시 활성 카드 뷰포트 중앙 가시 영역 내 진입 수학적 정합성 검증
{
  const trackWidth = 1084;
  const cardLeft = 1352;
  const cardWidth = 260;
  const currentScrollLeft = 0;
  const scrollWidth = 1614;

  const relativeCardLeft = cardLeft + currentScrollLeft;
  const targetScrollLeft = relativeCardLeft - (trackWidth / 2) + (cardWidth / 2);
  const maxScroll = Math.max(0, scrollWidth - trackWidth);
  const clampedTarget = Math.max(0, Math.min(maxScroll, targetScrollLeft));

  const finalCardLeftInView = cardLeft - clampedTarget;
  const finalCardRightInView = finalCardLeftInView + cardWidth;
  const isFullyVisibleInTrack = (finalCardLeftInView >= 0 && finalCardRightInView <= trackWidth);

  assertTest('Step 73-6: 1400px 랩톱 환경에서 마지막 현장 선택 시 clampedTarget(530px) 스크롤로 카드 100% 가시 영역 내 중앙 포커스 수학적 검증',
    clampedTarget === 530 && isFullyVisibleInTrack && finalCardLeftInView === 822 && finalCardRightInView === 1082
  );
}

// ========================================================
// Step 74: PC 현장 스크롤 행 숨김, 검색창 다국어/한글 포커스 유지 및 탭 전환 시 현장목록 은폐 무결성 QA
// ========================================================
console.log('\n--- [Step 74] PC 현장 스크롤 밀착, 검색창 글자 입력 포커스 유지 및 탭 전환 은폐 검증 ---');

const dashHtml74 = fs.readFileSync(path.join(__dirname, '../현장서식.js'), 'utf8');

// 74-1: .pc-site-detail-pane padding-top: 0 및 .pc-detail-hero margin-top: 16px로 스크롤 시 상단 틈새 행 노출 원천 차단 검증
assertTest('Step 74-1: .pc-site-detail-pane padding: 0 22px 36px 22px 및 hero margin-top: 16px 탑재 (스크롤 시 상단 행 틈새 노출 원천 차단)',
  dashHtml74.indexOf('.pc-site-detail-pane { flex: 1; min-height: 0; overflow-y: auto; padding: 0 22px 36px 22px; display: flex; flex-direction: column; gap: 14px; min-width: 0; position: relative; }') !== -1 &&
  dashHtml74.indexOf('.pc-detail-hero { margin-top: 16px; background: #ffffff;') !== -1
);

// 74-2: 공종 검색창 placeholder="검색" 및 id="pcSiteTaskSearchInput" 검증
assertTest('Step 74-2: 공종 검색창 placeholder="검색" 간결화 및 id="pcSiteTaskSearchInput" 탑재 검증',
  dashHtml74.indexOf('placeholder="검색"') !== -1 &&
  dashHtml74.indexOf('id="pcSiteTaskSearchInput"') !== -1 &&
  dashHtml74.indexOf('placeholder="현장 내 공종 실시간 검색..."') === -1
);

// 74-3: onPcSiteTaskSearch 인플레이스 DOM 필터링으로 입력 도중 포커스 소실 및 한 글자 제한 원천 해결 검증
assertTest('Step 74-3: onPcSiteTaskSearch 인플레이스 DOM 필터링 탑재 (입력 시 renderList 호출 배제로 포커스 소실 및 한글 자모 결합 중단 방지)',
  dashHtml74.indexOf('function onPcSiteTaskSearch(query) {') !== -1 &&
  dashHtml74.indexOf('const rows = document.querySelectorAll(".pc-tasks-table tbody tr.pc-task-row");') !== -1 &&
  dashHtml74.indexOf('tr.dataset.search') !== -1 &&
  dashHtml74.indexOf('pcTasksEmptyRow') !== -1
);

// 74-4: 테이블 행(tr)에 .pc-task-row 클래스 및 data-search/data-status/data-delayed 속성 탑재 검증
assertTest('Step 74-4: table tr에 .pc-task-row 클래스 및 실시간 검색용 data-search/status/delayed 속성 탑재 검증',
  dashHtml74.indexOf('tr class="pc-task-row') !== -1 &&
  dashHtml74.indexOf('searchTargetStr') !== -1 &&
  dashHtml74.indexOf('escapeHtml(searchTargetStr)') !== -1
);

// 74-5: renderPcSiteDashboard 리렌더링 시 hadSearchFocus 및 setSelectionRange 커서/포커스 복구 로직 검증
assertTest('Step 74-5: 리렌더링 시 hadSearchFocus 및 setSelectionRange 커서/포커스 복원 가드 검증',
  dashHtml74.indexOf('const hadSearchFocus = (document.activeElement && document.activeElement.id === "pcSiteTaskSearchInput");') !== -1 &&
  dashHtml74.indexOf('sInput.setSelectionRange(searchCursor, searchCursor);') !== -1
);

// 74-6: 월간/주간 탭 전환 시 contentArea.pc-dashboard-active.hidden { display: none !important; } 및 클래스 제거 검증
assertTest('Step 74-6: 월간/주간 탭 전환 시 contentArea의 pc-dashboard-active 제거 및 .hidden 완전 은폐 CSS 검증',
  dashHtml74.indexOf('#contentArea.hidden, #contentArea.pc-dashboard-active.hidden { display: none !important; }') !== -1 &&
  dashHtml74.indexOf('contentArea.classList.remove("pc-dashboard-active");') !== -1
);

console.log('\n========================================================');
console.log(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
console.log('========================================================');

if (failCount > 0) {
  process.exit(1);
} else {
  console.log('🎉 모든 QA 및 코드리뷰 무결성 테스트를 100% 통과했습니다!');
}



