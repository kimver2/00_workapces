/**
 * ========================================================
 * 우성아이디 거래처 리스트 서식 자동화 스크립트
 * (Notion 스타일 드롭다운 선택 메뉴 및 상태별 컬러 뱃지 적용)
 * ========================================================
 */

/**
 * 현재 활성화된 시트에 노션 스타일 드롭다운 및 색상 일괄 적용
 */
function applyNotionDropdownsAndColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // 현재 화면에 열려있는 탭에 바로 적용
  var maxRows = sheet.getMaxRows(); // 시트 맨 끝 행까지 전체 적용 (행 추가해도 양식 유지)

  // ----------------------------------------------------
  // 1. 노션 드롭다운 (선택 목록) 설정
  // ----------------------------------------------------
  // ① 대구분 (B열)
  var categories = ["시공", "외주", "자재"];
  sheet.getRange(2, 2, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(categories, true).setAllowInvalid(true).build()
  );

  // ② 업태/공종 (C열)
  var bizTypes = [
    "철거", "설비", "전기", "목공", "타일", "줄눈", "필름", "도장",
    "SMC", "도기", "도배", "청소", "마감", "욕실", "창호", "에어컨",
    "바닥", "가구", "중문", "유리", "방충망", "누수탐지", "보수", "단열시공",
    "블라인드", "커튼", "조명", "인력", "목자재", "벽지", "지물", "현관부속",
    "철물류", "인조대리석", "골재", "마루"
  ];
  sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(bizTypes, true).setAllowInvalid(true).build()
  );

  // ③ 우선순위 (D열)
  var priorities = ["A팀", "B팀", "C팀", "A 거래처", "B 거래처", "C 거래처", "본사", "LX 대리점", "미지정"];
  sheet.getRange(2, 4, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(priorities, true).setAllowInvalid(true).build()
  );

  // ④ 관리상태 (E열)
  var statuses = ["운영 중", "확인 필요", "신규 발굴", "보류"];
  sheet.getRange(2, 5, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build()
  );

  // ----------------------------------------------------
  // 2. 노션 스타일 뱃지 색상 (조건부 서식) 설정
  // ----------------------------------------------------
  var rangeB = sheet.getRange(2, 2, maxRows - 1, 1); // 대구분
  var rangeC = sheet.getRange(2, 3, maxRows - 1, 1); // 업태/공종
  var rangeD = sheet.getRange(2, 4, maxRows - 1, 1); // 우선순위
  var rangeE = sheet.getRange(2, 5, maxRows - 1, 1); // 관리상태
  
  var rules = [];

  // [관리상태 색상 - E열]
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("운영 중").setBackground("#dcfce7").setFontColor("#15803d").setBold(true).setRanges([rangeE]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("확인 필요").setBackground("#fef3c7").setFontColor("#b45309").setBold(true).setRanges([rangeE]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("신규 발굴").setBackground("#dbeafe").setFontColor("#1d4ed8").setBold(true).setRanges([rangeE]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("보류").setBackground("#fee2e2").setFontColor("#b91c1c").setBold(true).setRanges([rangeE]).build());

  // [대구분 색상 - B열]
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("시공").setBackground("#dbeafe").setFontColor("#1e40af").setBold(true).setRanges([rangeB]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("외주").setBackground("#f3e8ff").setFontColor("#6b21a8").setBold(true).setRanges([rangeB]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("자재").setBackground("#ffedd5").setFontColor("#9a3412").setBold(true).setRanges([rangeB]).build());

  // [우선순위 색상 - D열]
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("A팀").setBackground("#dcfce7").setFontColor("#166534").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("A 거래처").setBackground("#dcfce7").setFontColor("#166534").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("B팀").setBackground("#e0f2fe").setFontColor("#0369a1").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("B 거래처").setBackground("#e0f2fe").setFontColor("#0369a1").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("C팀").setBackground("#fef9c3").setFontColor("#854d0e").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("C 거래처").setBackground("#fef9c3").setFontColor("#854d0e").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("본사").setBackground("#e0e7ff").setFontColor("#3730a3").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("LX 대리점").setBackground("#e0e7ff").setFontColor("#3730a3").setBold(true).setRanges([rangeD]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("미지정").setBackground("#f1f5f9").setFontColor("#64748b").setRanges([rangeD]).build());

  // [업태/공종 색상 - C열 (전체 공종 컬러 매핑)]
  var bizColorMap = [
    { keys: ["철거", "마루"], bg: "#f1f5f9", font: "#334155" },
    { keys: ["설비", "누수"], bg: "#cffafe", font: "#0e7490" },
    { keys: ["전기", "조명"], bg: "#fef3c7", font: "#b45309" },
    { keys: ["목공", "목자재"], bg: "#ffedd5", font: "#9a3412" },
    { keys: ["타일", "줄눈", "도기"], bg: "#ccfbf1", font: "#0f766e" },
    { keys: ["필름", "도장"], bg: "#f3e8ff", font: "#7e22ce" },
    { keys: ["도배", "벽지", "지물"], bg: "#dcfce7", font: "#15803d" },
    { keys: ["창호", "중문", "유리", "방충망"], bg: "#dbeafe", font: "#1d4ed8" },
    { keys: ["바닥"], bg: "#fef9c3", font: "#854d0e" },
    { keys: ["가구", "블라인드", "커튼"], bg: "#fce7f3", font: "#be185d" },
    { keys: ["에어컨"], bg: "#e0f2fe", font: "#0369a1" },
    { keys: ["청소", "마감"], bg: "#d1fae5", font: "#047857" },
    { keys: ["욕실", "SMC"], bg: "#e0f2fe", font: "#0284c7" },
    { keys: ["인력", "철물", "골재", "단열시공", "인조대리석"], bg: "#f3f4f6", font: "#4b5563" }
  ];

  bizColorMap.forEach(function(item) {
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

  // 헤더 고정 및 필터 자동 생성
  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }
}
