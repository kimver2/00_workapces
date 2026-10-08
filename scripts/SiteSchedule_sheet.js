/**
 * ========================================================
 * 우성아이디 현장 데이터 자동화 스크립트
 * (신규 현장 상단 자동 추가 + 아코디언 그룹화 + 서식 일괄 적용)
 * ========================================================
 */

/**
 * 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 전용 메뉴 자동 등록
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 현장 관리")
    .addItem("➕ 신규 현장 추가 (상단 4개 행 + 아코디언)", "addNewSiteAtTop")
    .addSeparator()
    .addItem("🔄 전체 서식/그룹 새로고침", "applySiteScheduleFormatting")
    .addToUi();
}

/**
 * 2. 1행과 2행 사이에 신규 현장 4개 행을 추가하고 아코디언 버튼을 자동 생성하는 함수
 * (버튼을 누를 때마다 상단에 새 현장이 추가되어 최신순 정렬 유지)
 */
function addNewSiteAtTop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // 현장명 입력 팝업창
  var response = ui.prompt(
    "➕ 신규 현장 추가",
    "새로운 현장명을 입력하세요 (예: 래미안 101동 202호):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // 취소 시 중단
  }

  var siteName = response.getResponseText().trim();
  if (!siteName) {
    siteName = "신규 현장";
  }

  // 1행과 2행 사이에 4개 행 삽입
  sheet.insertRowsBefore(2, 4);

  // 기본 4개 공정 템플릿 데이터 (대표 1행 + 세부공정 3행)
  var defaultTasks = ["상담", "철거", "목공", "마감"];
  var newRows = [];
  for (var i = 0; i < defaultTasks.length; i++) {
    newRows.push([
      siteName,           // A열: 현장
      "",                 // B열: 날짜
      defaultTasks[i],    // C열: 공종
      "",                 // D열: 협력업체
      false,              // E열: 공유 (체크박스 미체크)
      "",                 // F열: 비고
      "",                 // G열: 담당자
      "예정"              // H열: 진행상태
    ]);
  }

  // 데이터 일괄 입력
  sheet.getRange(2, 1, 4, 8).setValues(newRows);

  // E열 체크박스 삽입
  sheet.getRange(2, 5, 4, 1).insertCheckboxes();

  // 정렬 (A, B, C, E, G, H 가운데 / D, F 좌측)
  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("left");
  });

  // 드롭다운 유효성 검사 설정
  applyValidationsToRange(sheet, 2, 4);

  // 2행의 자식 행인 3~5행을 묶어 2행에 접기/펼치기 아코디언 버튼 생성
  try {
    sheet.getRange("3:5").shiftRowGroupDepth(1);
  } catch(e) {}

  // A2 셀로 포커스 이동 및 토스트 알림
  sheet.setActiveRange(sheet.getRange("A2"));
  sheet.toast("'" + siteName + "' 4개 행이 상단에 추가되고 아코디언이 생성되었습니다.", "추가 완료", 3);
}

/**
 * 특정 행 범위에 드롭다운 유효성 검사 적용 헬퍼
 */
function applyValidationsToRange(sheet, startRow, numRows) {
  var tasks = [
    "상담", "신고", "철거", "마루철거", "설비", "누수보수", "미장",
    "전기입선", "전기마감", "목공", "타일", "줄눈", "필름", "도장",
    "SMC", "도기", "도기세팅 천장", "퍼플렉스", "코킹", "도배", "바닥", "마루시공",
    "창호", "창호a/s", "도어설치", "중문", "터닝도어", "현관부속",
    "주방", "가구", "수납가구", "마감", "잔마감", "청소", "입주", "현장촬영", "S/A"
  ];
  sheet.getRange(startRow, 3, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(tasks, true).setAllowInvalid(true).build()
  );

  var managers = ["조항만", "송지은", "전예린", "영하"];
  sheet.getRange(startRow, 7, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(managers, true).setAllowInvalid(true).build()
  );

  var statuses = ["예정", "진행 중", "완료", "AS"];
  sheet.getRange(startRow, 8, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build()
  );
}

/**
 * 3. 전체 현장 시트 서식 및 아코디언 일괄 세팅
 */
function applySiteScheduleFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  var lastRow = Math.max(sheet.getLastRow(), 2);

  // 1. 드롭다운 설정
  applyValidationsToRange(sheet, 2, maxRows - 1);

  // 2. E열 체크박스 생성
  sheet.getRange(2, 5, maxRows - 1, 1).insertCheckboxes();

  // 3. 노션 스타일 뱃지 색상 (조건부 서식)
  var rangeC = sheet.getRange(2, 3, maxRows - 1, 1);
  var rangeH = sheet.getRange(2, 8, maxRows - 1, 1);
  var rules = [];

  // 진행상태 색상
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("예정").setBackground("#fef3c7").setFontColor("#b45309").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("진행 중").setBackground("#dbeafe").setFontColor("#1d4ed8").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("완료").setBackground("#dcfce7").setFontColor("#15803d").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("AS").setBackground("#fee2e2").setFontColor("#b91c1c").setBold(true).setRanges([rangeH]).build());

  // 공종별 색상
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

  // 4. 헤더 및 정렬
  var lastCol = Math.max(sheet.getLastColumn(), 8);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("left");
  });

  var colWidths = [180, 140, 110, 150, 70, 220, 100, 95];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }

  // 5. 기존 전체 현장별 아코디언 그룹화
  if (lastRow >= 3) {
    try {
      for (var d = 0; d < 3; d++) {
        sheet.getRange("2:" + maxRows).shiftRowGroupDepth(-1);
      }
    } catch(e) {}

    var siteValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var startRow = 2;
    var currentSite = siteValues[0][0];

    for (var i = 1; i < siteValues.length; i++) {
      var site = siteValues[i][0];
      if (site !== currentSite) {
        var endRow = i + 1;
        if (endRow > startRow) {
          try {
            sheet.getRange((startRow + 1) + ":" + endRow).shiftRowGroupDepth(1);
          } catch(e) {}
        }
        startRow = i + 2;
        currentSite = site;
      }
    }
    if (lastRow > startRow) {
      try {
        sheet.getRange((startRow + 1) + ":" + lastRow).shiftRowGroupDepth(1);
      } catch(e) {}
    }
  }
}
