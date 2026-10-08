/**
 * ========================================================
 * 우성아이디 현장 데이터(스케줄) 자동화 서식 스크립트
 * (Notion '현장 데이터' DB -> Google Sheets)
 * ========================================================
 */

/**
 * 1. 현재 열려 있는 '현장 데이터' 시트에 노션 스타일 드롭다운 및 색상 서식 적용
 */
function applySiteScheduleFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // 현재 화면에 보고 계신 탭에 적용
  var maxRows = sheet.getMaxRows();

  // ----------------------------------------------------
  // 1. 드롭다운 (데이터 확인 / Data Validation) 설정
  // ----------------------------------------------------
  // ① 공종 (C열: 작업 공종 전체)
  var tasks = [
    "상담", "신고", "철거", "마루철거", "설비", "누수보수", "미장",
    "전기입선", "전기마감", "목공", "타일", "줄눈", "필름", "도장",
    "SMC", "도기", "도기세팅 천장", "퍼플렉스", "코킹", "도배", "바닥", "마루시공",
    "창호", "창호a/s", "도어설치", "중문", "터닝도어", "현관부속",
    "주방", "가구", "수납가구", "마감", "잔마감", "청소", "입주", "현장촬영", "S/A"
  ];
  sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(tasks, true).setAllowInvalid(true).build()
  );

  // ② 담당자 (G열)
  var managers = ["조항만", "송지은", "전예린", "영하"];
  sheet.getRange(2, 7, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(managers, true).setAllowInvalid(true).build()
  );

  // ③ 진행상태 (H열)
  var statuses = ["예정", "진행 중", "완료", "AS"];
  sheet.getRange(2, 8, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build()
  );

  // ----------------------------------------------------
  // 2. 노션 스타일 뱃지 색상 (조건부 서식) 설정
  // ----------------------------------------------------
  var rangeA = sheet.getRange(2, 1, maxRows - 1, 1); // 현장
  var rangeC = sheet.getRange(2, 3, maxRows - 1, 1); // 공종
  var rangeH = sheet.getRange(2, 8, maxRows - 1, 1); // 진행상태
  
  var rules = [];

  // [진행상태 색상 - H열]
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("예정").setBackground("#fef3c7").setFontColor("#b45309").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("진행 중").setBackground("#dbeafe").setFontColor("#1d4ed8").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("완료").setBackground("#dcfce7").setFontColor("#15803d").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("AS").setBackground("#fee2e2").setFontColor("#b91c1c").setBold(true).setRanges([rangeH]).build());

  // [공종별 뱃지 색상 - C열]
  var taskColorMap = [
    { keys: ["철거", "마루철거"], bg: "#f1f5f9", font: "#334155" },
    { keys: ["설비", "누수보수", "미장"], bg: "#cffafe", font: "#0e7490" },
    { keys: ["전기입선", "전기마감"], bg: "#fef3c7", font: "#b45309" },
    { keys: ["목공"], bg: "#ffedd5", font: "#9a3412" },
    { keys: ["타일", "줄눈", "퍼플렉스"], bg: "#ccfbf1", font: "#0f766e" },
    { keys: ["필름", "도장"], bg: "#f3e8ff", font: "#7e22ce" },
    { keys: ["도배"], bg: "#dcfce7", font: "#15803d" },
    { keys: ["창호", "중문", "도어", "터닝도어", "현관부속"], bg: "#dbeafe", font: "#1d4ed8" },
    { keys: ["바닥", "마루시공"], bg: "#fef9c3", font: "#854d0e" },
    { keys: ["주방", "가구", "수납가구"], bg: "#fce7f3", font: "#be185d" },
    { keys: ["SMC", "도기"], bg: "#e0f2fe", font: "#0284c7" },
    { keys: ["마감", "잔마감", "코킹", "청소"], bg: "#d1fae5", font: "#047857" },
    { keys: ["상담", "신고", "입주", "현장촬영"], bg: "#e0e7ff", font: "#3730a3" }
  ];

  taskColorMap.forEach(function(item) {
    item.keys.forEach(function(k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(k)
        .setBackground(item.bg)
        .setFontColor(item.font)
        .setBold(true)
        .setRanges([rangeC])
        .build());
    });
  });

  sheet.setConditionalFormatRules(rules);

  // ----------------------------------------------------
  // 3. 레이아웃 & 헤더 스타일링
  // ----------------------------------------------------
  var lastCol = Math.max(sheet.getLastColumn(), 8);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // 정렬 (A, B, C, E, G, H 가운데 / D, F 좌측)
  var lastRow = Math.max(sheet.getLastRow(), 2);
  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("left");
  });

  // 열 너비 최적화
  var colWidths = [180, 140, 110, 150, 70, 220, 100, 95];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // 필터 자동 생성
  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }
}

/**
 * 2. 노션처럼 현장별로 접었다 펴는 '행 그룹화(아코디언 토글)' 자동 세팅
 */
function groupRowsBySite() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  var siteValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // 기존 그룹 해제
  for (var r = 2; r <= lastRow; r++) {
    try {
      sheet.getRowGroup(r, 1).remove();
    } catch(e) {}
  }

  // 동일한 현장끼리 연속된 구간 찾아서 그룹화
  var startRow = 2;
  var currentSite = siteValues[0][0];

  for (var i = 1; i < siteValues.length; i++) {
    var site = siteValues[i][0];
    if (site !== currentSite) {
      if (i + 1 > startRow + 1) {
        // 2개 이상의 행이 있는 현장만 접기 그룹화
        try {
          sheet.shiftRowGroupDepth(startRow + 1, i + 1, 1);
        } catch(e) {}
      }
      startRow = i + 2;
      currentSite = site;
    }
  }
  // 마지막 현장 그룹화
  if (lastRow > startRow) {
    try {
      sheet.shiftRowGroupDepth(startRow + 1, lastRow, 1);
    } catch(e) {}
  }
}
