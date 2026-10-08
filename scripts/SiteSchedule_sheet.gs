/**
 * ========================================================
 * 우성아이디 현장 데이터 자동화 스크립트
 * (신규 현장 상단 자동 추가 + 다중선택 칩 보존 + 아코디언)
 * ========================================================
 */

/**
 * 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 메뉴 자동 등록
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
 * 2. 1행과 2행 사이에 신규 현장 4개 행을 추가하고 아코디언 버튼 자동 생성
 */
function addNewSiteAtTop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // 현장명 입력 팝업
  var response = ui.prompt(
    "➕ 신규 현장 추가",
    "새로운 현장명을 입력하세요 (예: 래미안 101동 202호):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var siteName = response.getResponseText().trim();
  if (!siteName) {
    siteName = "신규 현장";
  }

  // 1행과 2행 사이에 4개 행 삽입
  sheet.insertRowsBefore(2, 4);

  // 기본 4개 공정 템플릿 (대표 1행 + 세부공정 3행)
  var defaultTasks = ["상담", "철거", "목공", "마감"];
  var newRows = [];
  for (var i = 0; i < defaultTasks.length; i++) {
    newRows.push([
      siteName,           // A열: 현장
      "",                 // B열: 날짜
      defaultTasks[i],    // C열: 공종
      "",                 // D열: 협력업체
      false,              // E열: 공유 (체크박스)
      "",                 // F열: 비고
      "",                 // G열: 담당자
      "예정"              // H열: 진행상태
    ]);
  }

  // 데이터 일괄 입력
  sheet.getRange(2, 1, 4, 8).setValues(newRows);

  // E열 체크박스 삽입
  sheet.getRange(2, 5, 4, 1).insertCheckboxes();

  // 정렬
  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("left");
  });

  // 기존 데이터 행에서 다중 선택 칩 드롭다운 서식을 새 행으로 복사 (다중 선택 보존)
  try {
    var sourceRow = sheet.getLastRow(); // 실제 데이터가 있는 맨 아래 행에서 복사
    sheet.getRange(sourceRow, 7).copyTo(sheet.getRange(2, 7, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); // 담당자
    sheet.getRange(sourceRow, 3).copyTo(sheet.getRange(2, 3, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); // 공종
    sheet.getRange(sourceRow, 8).copyTo(sheet.getRange(2, 8, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); // 진행상태
  } catch(e) {}

  // 2행의 자식 행인 3~5행을 묶어 2행에 접기/펼치기 아코디언 버튼 생성
  try {
    sheet.getRange("3:5").shiftRowGroupDepth(1);
  } catch(e) {}

  sheet.setActiveRange(sheet.getRange("A2"));
  ss.toast("'" + siteName + "' 4개 행이 상단에 추가되었습니다.", "추가 완료", 3);
}

/**
 * 3. 전체 현장 시트 서식 및 아코디언 일괄 세팅
 * (주의: 다중 선택 칩 설정을 지우지 않도록 드롭다운을 강제 재정의하지 않습니다)
 */
function applySiteScheduleFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  var lastRow = Math.max(sheet.getLastRow(), 2);

  // E열 체크박스 일괄 생성
  sheet.getRange(2, 5, maxRows - 1, 1).insertCheckboxes();

  // 노션 스타일 뱃지 색상 (조건부 서식)
  var rangeC = sheet.getRange(2, 3, maxRows - 1, 1); // 공종
  var rangeG = sheet.getRange(2, 7, maxRows - 1, 1); // 담당자
  var rangeH = sheet.getRange(2, 8, maxRows - 1, 1); // 진행상태
  var rules = [];

  // 진행상태 색상
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("예정").setBackground("#fef3c7").setFontColor("#b45309").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("진행 중").setBackground("#dbeafe").setFontColor("#1d4ed8").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("완료").setBackground("#dcfce7").setFontColor("#15803d").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("AS").setBackground("#fee2e2").setFontColor("#b91c1c").setBold(true).setRanges([rangeH]).build());

  // 담당자 인물 뱃지 색상 (G열)
  var managerColorMap = [
    { key: "조항만", bg: "#dbeafe", font: "#1e40af" },
    { key: "송지은", bg: "#fce7f3", font: "#be185d" },
    { key: "전예린", bg: "#f3e8ff", font: "#7e22ce" },
    { key: "영하",   bg: "#ccfbf1", font: "#0f766e" }
  ];
  managerColorMap.forEach(function(item) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(item.key)
      .setBackground(item.bg)
      .setFontColor(item.font)
      .setBold(true)
      .setRanges([rangeG])
      .build());
  });

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

  // 헤더 스타일링
  var lastCol = Math.max(sheet.getLastColumn(), 8);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // 정렬
  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("left");
  });

  // 열 너비
  var colWidths = [180, 140, 110, 150, 70, 220, 100, 95];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }

  // 기존 전체 현장 아코디언 그룹화
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
