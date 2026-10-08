/**
 * ========================================================
 * 구글 스프레드시트 만년달력 자동 생성 스크립트
 * (구글 캘린더 대한민국 공휴일 연동 + WBS 빈 행 + Today 버튼 지원)
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

  // 2) 구글 공식 캘린더에서 공휴일 자동 수집
  autoSyncKoreanHolidays(curYear);
  
  // Y열(공휴일 데이터), Z열(연도 목록) 숨기기
  sheet.hideColumns(25, 2);
  
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
  
  // 6) 주차별 날짜 수식 (WBS 빈 행 지원)
  sheet.getRange("B5").setFormula("=SEQUENCE(1, 7, DATE(C2, E2, 1) - WEEKDAY(DATE(C2, E2, 1), 1) + 1, 1)");
  sheet.getRange("B7").setFormula("=SEQUENCE(1, 7, H5 + 1, 1)");
  sheet.getRange("B9").setFormula("=SEQUENCE(1, 7, H7 + 1, 1)");
  sheet.getRange("B11").setFormula("=SEQUENCE(1, 7, H9 + 1, 1)");
  sheet.getRange("B13").setFormula("=SEQUENCE(1, 7, H11 + 1, 1)");
  sheet.getRange("B15").setFormula("=SEQUENCE(1, 7, H13 + 1, 1)");
  
  // 7) 서식 및 크기 설정
  var dateRows = [5, 7, 9, 11, 13, 15];
  var taskRows = [6, 8, 10, 12, 14, 16];
  
  for (var col = 2; col <= 8; col++) {
    sheet.setColumnWidth(col, 95);
  }
  sheet.setRowHeight(4, 30);
  
  // 날짜 행 서식
  dateRows.forEach(function(row) {
    sheet.setRowHeight(row, 26);
    sheet.getRange(row, 2, 1, 7)
      .setNumberFormat("d")
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setBackground("#f8fafc")
      .setBorder(true, true, false, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  });
  
  // 작업/WBS 빈 행 서식
  taskRows.forEach(function(row) {
    sheet.setRowHeight(row, 55);
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
  
  // 1순위: 이전/다음 달 흐리게
  var ruleOtherMonth = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=MONTH(B5)<>$E$2")
    .setFontColor("#94a3b8")
    .setBackground("#f1f5f9")
    .setRanges(calDateRangeList)
    .build();

  // 2순위: 오늘 날짜 노란색 하이라이트
  var ruleToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=B5=TODAY()")
    .setBackground("#fef08a")
    .setFontColor("#854d0e")
    .setRanges(calDateRangeList)
    .build();

  // 3순위: 대한민국 공휴일 빨간색 강조
  var ruleHoliday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=COUNTIF($Y$1:$Y$300, B5)>0")
    .setFontColor("#dc2626")
    .setBackground("#fef2f2")
    .setBold(true)
    .setRanges(calDateRangeList)
    .build();

  // 4순위: 일요일 빨간색
  var ruleSunday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=WEEKDAY(B5)=1")
    .setFontColor("#dc2626")
    .setRanges(calDateRangeList)
    .build();

  // 5순위: 토요일 파란색
  var ruleSaturday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=WEEKDAY(B5)=7")
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
 * 구글 공식 캘린더에서 대한민국 공휴일 수집 (-2년 ~ +2년)
 */
function autoSyncKoreanHolidays(baseYear) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!baseYear) baseYear = new Date().getFullYear();
  
  var holidays = [];
  try {
    var cal = CalendarApp.getCalendarById("ko.south_korea#holiday@group.v.calendar.google.com");
    if (cal) {
      var start = new Date(baseYear - 2, 0, 1);
      var end = new Date(baseYear + 3, 0, 1);
      var events = cal.getEvents(start, end);
      for (var i = 0; i < events.length; i++) {
        holidays.push([events[i].getStartTime()]);
      }
    }
  } catch (e) {
    Logger.log("Calendar sync: " + e.toString());
  }
  
  if (holidays.length > 0) {
    sheet.getRange("Y1:Y300").clearContent();
    sheet.getRange(1, 25, holidays.length, 1)
      .setValues(holidays)
      .setNumberFormat("yyyy-mm-dd");
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

/**
 * 백그라운드 트리거에 의해 실행되는 핸들러 함수
 */
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
  
  var dateRows = [5, 7, 9, 11, 13, 15];
  for (var i = 0; i < dateRows.length; i++) {
    var r = dateRows[i];
    var rowValues = sheet.getRange(r, 2, 1, 7).getValues()[0];
    for (var c = 0; c < 7; c++) {
      var val = rowValues[c];
      if (val instanceof Date) {
        if (val.getFullYear() === today.getFullYear() &&
            val.getMonth() === today.getMonth() &&
            val.getDate() === today.getDate()) {
          sheet.setActiveRange(sheet.getRange(r, c + 2));
          return;
        }
      }
    }
  }
}
