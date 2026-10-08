/**
 * ========================================================
 * 구글 스프레드시트 만년달력 자동 생성 스크립트
 * (법정공휴일만 빨간색 표시 + 비공휴 기념일 제외 + WBS 빈 행 + Today 버튼)
 * ========================================================
 */

/**
 * 만년달력 생성 메인 함수
 */
function createPerpetualCalendar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // 초기화
  
  var today = new Date();
  var curYear = today.getFullYear();
  var curMonth = today.getMonth() + 1;
  
  // 1) 동적 연도 목록 생성 (오늘 기준 -2년 ~ +2년 총 5개 연도)
  sheet.getRange("Z1").setFormula("=SEQUENCE(5, 1, YEAR(TODAY()) - 2, 1)");

  // 2) 대한민국 법정 공휴일 데이터(날짜 및 명칭) 자동 수집 (식목일, 국군의날 등 비공휴 기념일 자동 제외)
  autoSyncKoreanHolidays(curYear);
  
  // X열(공휴일 날짜), Y열(공휴일 명칭), Z열(연도 목록) 숨기기
  sheet.hideColumns(24, 3);
  
  // 3) 헤더 컨트롤 바 (B2:E2)
  sheet.getRange("B2").setValue("년도").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D2").setValue("월").setFontWeight("bold").setHorizontalAlignment("center");
  
  // 연도 드롭다운 (Z1:Z5 동적 범위 연결)
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange("Z1:Z5"), true)
    .build();
  sheet.getRange("C2")
    .setDataValidation(yearRule)
    .setValue(curYear)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  
  // 월 드롭다운 (1 ~ 12)
  var months = [];
  for (var m = 1; m <= 12; m++) months.push(m.toString());
  var monthRule = SpreadsheetApp.newDataValidation().requireValueInList(months, true).build();
  sheet.getRange("E2")
    .setDataValidation(monthRule)
    .setValue(curMonth)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
    
  // 4) G2:H2 영역에 [Today] 스마트 버튼 자동 생성
  sheet.getRange("G2").insertCheckboxes();
  sheet.getRange("G2").setValue(false).setHorizontalAlignment("center");
  sheet.getRange("H2").setValue("Today").setFontWeight("bold").setFontColor("#1d4ed8").setHorizontalAlignment("left");
  sheet.getRange("G2:H2")
    .setBackground("#eff6ff")
    .setBorder(true, true, true, true, true, true, "#bfdbfe", SpreadsheetApp.BorderStyle.SOLID);
  
  // 5) 요일 헤더 (B4:H4)
  var days = ["일", "월", "화", "수", "목", "금", "토"];
  var headerRange = sheet.getRange("B4:H4");
  headerRange.setValues([days])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#f8fafc");
  sheet.getRange("B4").setFontColor("#dc2626"); // 일요일 빨강
  sheet.getRange("H4").setFontColor("#2563eb"); // 토요일 파랑
  
  // 6) 주차별 날짜 수식 (법정공휴일일 경우에만 '일자 (기념일명)' 형태로 자동 노출)
  var weekOffsets = [0, 1, 2, 3, 4, 5];
  var formulaRows = [5, 7, 9, 11, 13, 15];
  
  for (var w = 0; w < weekOffsets.length; w++) {
    var r = formulaRows[w];
    var formula = '=LET(' +
      'start, DATE($C$2, $E$2, 1) - WEEKDAY(DATE($C$2, $E$2, 1), 1) + 1 + (' + (w * 7) + '), ' +
      'days, SEQUENCE(1, 7, start, 1), ' +
      'MAP(days, LAMBDA(d, ' +
        'IF(MONTH(d) <> $E$2, ' +
          'TEXT(DAY(d), "0"), ' +
          'IFNA(TEXT(DAY(d), "0") & " (" & VLOOKUP(d, $X$1:$Y$300, 2, FALSE) & ")", TEXT(DAY(d), "0"))' +
        ')' +
      '))' +
    ')';
    sheet.getRange("B" + r).setFormula(formula);
  }
  
  // 7) 서식 및 크기 설정 (열 너비 300px 지정)
  var dateRows = [5, 7, 9, 11, 13, 15];
  var taskRows = [6, 8, 10, 12, 14, 16];
  
  for (var col = 2; col <= 8; col++) {
    sheet.setColumnWidth(col, 300);
  }
  sheet.setRowHeight(4, 30);
  
  // 날짜 행 서식
  dateRows.forEach(function(row) {
    sheet.setRowHeight(row, 28);
    sheet.getRange(row, 2, 1, 7)
      .setFontSize(9.5)
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setBackground("#f8fafc")
      .setBorder(true, true, false, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  });
  
  // 작업/WBS 빈 행 서식 (높이 300px 지정)
  taskRows.forEach(function(row) {
    sheet.setRowHeight(row, 300);
    sheet.getRange(row, 2, 1, 7)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment("top")
      .setHorizontalAlignment("left")
      .setFontSize(9)
      .setBorder(false, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  });
  
  // 8) 조건부 서식 (우선순위 순서)
  var calDateRangeList = [];
  dateRows.forEach(function(r) {
    calDateRangeList.push(sheet.getRange("B" + r + ":H" + r));
  });
  
  var cellDateFormula = 'DATE($C$2, $E$2, 1) - WEEKDAY(DATE($C$2, $E$2, 1), 1) + 1 + (INT((ROW(B5)-5)/2)*7) + (COLUMN(B5)-2)';
  
  // 1순위: 이전/다음 달 흐리게
  var ruleOtherMonth = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + cellDateFormula + ' <> DATE($C$2, $E$2, DAY(' + cellDateFormula + '))')
    .setFontColor("#94a3b8")
    .setBackground("#f1f5f9")
    .setRanges(calDateRangeList)
    .build();

  // 2순위: 오늘 날짜 노란색 하이라이트
  var ruleToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + cellDateFormula + ' = TODAY()')
    .setBackground("#fef08a")
    .setFontColor("#854d0e")
    .setRanges(calDateRangeList)
    .build();

  // 3순위: 법정공휴일 빨간색 강조 (기념일 괄호 '(' 가 포함된 날짜)
  var ruleHoliday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISNUMBER(SEARCH("(", B5))')
    .setFontColor("#dc2626")
    .setBackground("#fef2f2")
    .setBold(true)
    .setRanges(calDateRangeList)
    .build();

  // 4순위: 일요일 빨간색 (B열)
  var ruleSunday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COLUMN(B5)=2')
    .setFontColor("#dc2626")
    .setRanges(calDateRangeList)
    .build();

  // 5순위: 토요일 파란색 (H열)
  var ruleSaturday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COLUMN(B5)=8')
    .setFontColor("#2563eb")
    .setRanges(calDateRangeList)
    .build();
    
  sheet.setConditionalFormatRules([ruleOtherMonth, ruleToday, ruleHoliday, ruleSunday, ruleSaturday]);
  
  // 9) 백그라운드 자동 갱신 트리거 등록
  registerMonthlyTrigger();
}

/**
 * G2 체크박스를 클릭했을 때 Today 동작 실행
 */
function onEdit(e) {
  if (!e) return;
  var range = e.range;
  if (range.getA1Notation() === "G2" && range.getValue() === true) {
    range.setValue(false); // 클릭 후 다시 원상복구하여 버튼처럼 동작
    goToToday();
  }
}

/**
 * 구글 공식 캘린더에서 대한민국 법정공휴일 날짜 및 명칭 수집
 * (비공휴 기념일인 식목일, 국군의날, 제헌절 등은 제외)
 */
function autoSyncKoreanHolidays(baseYear) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!baseYear) baseYear = new Date().getFullYear();
  
  // 빨간 날이 아닌 비공휴 기념일 제외 키워드 목록
  var nonHolidayKeywords = [
    "식목일", "국군의 날", "국군의날", "제헌절", "어버이날", "스승의 날", "스승의날",
    "근로자의 날", "근로자의날", "노동절", "발렌타인", "화이트데이", "할로윈", "크리스마스 이브",
    "초복", "중복", "말복", "단오", "칠석", "대보름", "정월대보름", "한식", "동지", "입춘", "춘분", "하지", "추분",
    "경찰의 날", "소방의 날", "학생독립운동기념일", "농업인의 날", "무역의 날", "금융의 날",
    "체육의 날", "문화의 날", "세계 한인의 날", "임산부의 날", "노인의 날", "청년의 날",
    "철도의 날", "정보보호의 날", "바다의 날", "발명의 날", "성년의 날", "부부의 날", "방재의 날",
    "과학의 날", "정보통신의 날", "보건의 날", "예비군의 날", "서해수호의 날", "상공의 날", "납세자의 날",
    "3·15", "4·19", "5·18", "6·10", "독도의 날"
  ];

  var holidayData = [];
  try {
    var cal = CalendarApp.getCalendarById("ko.south_korea#holiday@group.v.calendar.google.com");
    if (cal) {
      var start = new Date(baseYear - 2, 0, 1);
      var end = new Date(baseYear + 3, 0, 1);
      var events = cal.getEvents(start, end);
      for (var i = 0; i < events.length; i++) {
        var ev = events[i];
        var title = ev.getTitle();
        
        // 비공휴 기념일 여부 검사
        var isNonHoliday = false;
        for (var k = 0; k < nonHolidayKeywords.length; k++) {
          if (title.indexOf(nonHolidayKeywords[k]) !== -1) {
            isNonHoliday = true;
            break;
          }
        }
        if (isNonHoliday) continue; // 빨간 날이 아니므로 제외
        
        var rawDate = ev.getStartTime();
        var cleanDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
        var displayTitle = title.replace(/\s*\(대체공휴일\)\s*/g, "대체").trim();
        holidayData.push([cleanDate, displayTitle]);
      }
    }
  } catch (e) {
    Logger.log("Calendar sync: " + e.toString());
  }
  
  // 필수 법정공휴일 백업 데이터 (식목일, 국군의날, 제헌절 등 제외된 순수 공휴일)
  if (holidayData.length === 0) {
    var defaultHolidays = [
      ["2024-01-01", "신정"], ["2024-02-09", "설날연휴"], ["2024-02-10", "설날"], ["2024-02-11", "설날연휴"], ["2024-02-12", "대체공휴일"],
      ["2024-03-01", "삼일절"], ["2024-04-10", "총선"], ["2024-05-05", "어린이날"], ["2024-05-06", "대체공휴일"], ["2024-05-15", "부처님오신날"],
      ["2024-06-06", "현충일"], ["2024-08-15", "광복절"], ["2024-09-16", "추석연휴"], ["2024-09-17", "추석"], ["2024-09-18", "추석연휴"],
      ["2024-10-03", "개천절"], ["2024-10-09", "한글날"], ["2024-12-25", "성탄절"],
      ["2025-01-01", "신정"], ["2025-01-28", "설날연휴"], ["2025-01-29", "설날"], ["2025-01-30", "설날연휴"],
      ["2025-03-01", "삼일절"], ["2025-03-03", "대체공휴일"], ["2025-05-05", "어린이날"], ["2025-05-06", "대체공휴일"],
      ["2025-06-06", "현충일"], ["2025-08-15", "광복절"], ["2025-10-03", "개천절"], ["2025-10-05", "추석연휴"],
      ["2025-10-06", "추석"], ["2025-10-07", "추석연휴"], ["2025-10-08", "대체공휴일"], ["2025-10-09", "한글날"], ["2025-12-25", "성탄절"],
      ["2026-01-01", "신정"], ["2026-02-16", "설날연휴"], ["2026-02-17", "설날"], ["2026-02-18", "설날연휴"],
      ["2026-03-01", "삼일절"], ["2026-03-02", "대체공휴일"], ["2026-05-05", "어린이날"], ["2026-05-24", "부처님오신날"],
      ["2026-05-25", "대체공휴일"], ["2026-06-03", "지방선거"], ["2026-06-06", "현충일"], ["2026-08-15", "광복절"],
      ["2026-08-17", "대체공휴일"], ["2026-09-24", "추석연휴"], ["2026-09-25", "추석"], ["2026-09-26", "추석연휴"],
      ["2026-09-28", "대체공휴일"], ["2026-10-03", "개천절"], ["2026-10-05", "대체공휴일"], ["2026-10-09", "한글날"], ["2026-12-25", "성탄절"]
    ];
    defaultHolidays.forEach(function(item) {
      holidayData.push([new Date(item[0]), item[1]]);
    });
  }
  
  if (holidayData.length > 0) {
    sheet.getRange("X1:Y300").clearContent();
    sheet.getRange(1, 24, holidayData.length, 2).setValues(holidayData);
    sheet.getRange(1, 24, holidayData.length, 1).setNumberFormat("yyyy-mm-dd");
  }
}

/**
 * 매월 1일 새벽에 자동 실행되는 백그라운드 트리거 등록
 */
function registerMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onMonthlyTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onMonthlyTrigger')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
}

function onMonthlyTrigger() {
  autoSyncKoreanHolidays(new Date().getFullYear());
}

/**
 * Today 버튼 클릭 시 오늘 날짜로 즉시 이동
 */
function goToToday() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var today = new Date();
  var curYear = today.getFullYear();
  var curMonth = today.getMonth() + 1;
  
  sheet.getRange("C2").setValue(curYear);
  sheet.getRange("E2").setValue(curMonth);
  SpreadsheetApp.flush();
  
  // 오늘 날짜 셀의 정확한 위치 계산 및 포커스 이동
  var firstDay = new Date(curYear, curMonth - 1, 1);
  var firstSun = new Date(curYear, curMonth - 1, 1 - firstDay.getDay());
  var diffDays = Math.floor((today - firstSun) / (1000 * 60 * 60 * 24));
  var weekIdx = Math.floor(diffDays / 7);
  var dayOfWeek = diffDays % 7;
  
  var targetRow = 5 + (weekIdx * 2);
  var targetCol = 2 + dayOfWeek;
  if (targetRow >= 5 && targetRow <= 15 && targetCol >= 2 && targetCol <= 8) {
    sheet.setActiveRange(sheet.getRange(targetRow, targetCol));
  }
}
