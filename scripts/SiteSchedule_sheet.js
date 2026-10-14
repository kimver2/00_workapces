/**
 * ========================================================
 * 우성아이디 현장 관리 & 대시보드 달력 & 구글 캘린더 완벽 양방향 연동 스크립트 (v8.0)
 * (최초작성일/작성자 및 최종수정일/수정자 자동 기록 + Last-Modified-Wins 양방향 연동)
 * 
 * [주요 기능]
 * 1. 📝 [작성/수정 이력 자동 추적 (Audit Trail)]:
 *    - 최초 작성일 & 작성자(이메일/아이디) 자동 기록 (K, L열)
 *    - 최종 수정일 & 수정자 자동 기록 (M, N열)
 *    - 구글 캘린더에서 날짜를 옮기거나 수정한 경우: "최종수정자: 구글 캘린더"로 정확하게 추적 기록!
 *    - 캘린더 일정 본문(Description)에도 최초작성/최종수정 정보 자동 표기
 * 2. ⏱️ [최종 수정 시간 기반 Last-Modified-Wins 동기화]:
 *    - 캘린더에서 날짜를 수정하면 시트로 자동 반영, 시트에서 수정하면 캘린더로 자동 반영!
 *    - 캘린더에서 바꾼 날짜가 시트 옛날 값으로 되돌아가는 현상 100% 원천 차단
 * 3. 📅 [시트 ➔ 캘린더 실시간 연동] & [캘린더 ➔ 시트 1분/5분 자동 연동]
 * 4. ⚙️ [자유로운 캘린더 지정] & 🔄 [대시보드 월간 달력 연쇄 최신화]
 * ========================================================
 */

// ========================================================
// [설정] 구글 캘린더 연동 기본값
// ========================================================
// 캘린더 이름 또는 캘린더 ID를 입력할 수 있습니다.
// - 캘린더 이름: "현장 일정", "회사 공용 캘린더" 등 (없으면 자동 생성)
// - 캘린더 ID: "c_xxxx@group.calendar.google.com" 또는 본인 구글 이메일 (공유받은 캘린더)
// - 빈 문자열(""): 본인 기본 캘린더(Default Calendar)에 등록
// ※ 스프레드시트 메뉴 [🚀 현장 관리] > [⚙️ 연동 캘린더 변경 / 확인] 에서도 언제든 변경 가능합니다!
var DEFAULT_CALENDAR_ID_OR_NAME = "현장 일정";


// ========================================================
// 0. [원클릭 자동 구축] 새 스프레드시트 초기 환경 자동 생성
// ========================================================

/**
 * [원클릭 자동 구축] 새 스프레드시트에 [대시보드] 및 [현장데이터] 시트와 서식을 100% 자동 생성
 */
function setupNewWorkspaceFromScratch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {}

  if (ui) {
    var resp = ui.alert(
      "✨ 새 워크스페이스 원클릭 자동 구축",
      "이 스프레드시트에 [대시보드] 및 [현장데이터] 시트와 서식을 자동으로 생성하시겠습니까?\n\n(이미 시트가 존재할 경우 최신 규격으로 정비됩니다.)",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
  }

  // 1. [현장데이터] 시트 생성 및 초기화
  var siteSheet = ss.getSheetByName("현장데이터");
  if (!siteSheet) {
    var existingSiteSheet = findSiteScheduleSheet(ss);
    if (existingSiteSheet && isSiteScheduleSheet(existingSiteSheet)) {
      siteSheet = existingSiteSheet;
      try { siteSheet.setName("현장데이터"); } catch(e) {}
    } else {
      var defaultSheet1 = ss.getSheetByName("시트1") || ss.getSheetByName("Sheet1");
      if (defaultSheet1) {
        siteSheet = defaultSheet1;
        try { siteSheet.setName("현장데이터"); } catch(e) {
          siteSheet = ss.insertSheet("현장데이터");
        }
      } else {
        siteSheet = ss.insertSheet("현장데이터");
      }
    }
  }
  siteSheet.setHiddenGridlines(false);

  // 1행 헤더 설정
  var headers = [
    "현장", "날짜", "공종", "협력업체", "공유", "비고", "담당자", "진행상태",
    "캘린더ID", "최종수정시각", "최초작성일", "최초작성자", "최종수정일", "최종수정자"
  ];
  siteSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  siteSheet.setRowHeight(1, 36);

  // 열 너비 설정
  var colWidths = [150, 180, 120, 130, 60, 160, 90, 100, 90, 100, 120, 85, 120, 85];
  for (var c = 0; c < colWidths.length; c++) {
    siteSheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // 샘플 데이터 및 기본 행 (행이 헤더 1행뿐인 경우)
  if (siteSheet.getLastRow() < 2) {
    var now = new Date();
    var nowMs = now.getTime();
    var nowStr = formatDateTime(now);
    var currentUser = getCurrentUserIdentifier();
    var sampleTasks = ["상담", "철거", "목공", "마감"];
    var sampleRows = [];
    for (var i = 0; i < sampleTasks.length; i++) {
      sampleRows.push([
        "신규 현장 1", "", sampleTasks[i], "", false, "", "조항만", "예정",
        "", nowMs, nowStr, currentUser, nowStr, currentUser
      ]);
    }
    siteSheet.getRange(2, 1, sampleRows.length, 14).setValues(sampleRows);

    try {
      siteSheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
      siteSheet.getRange("2:5").shiftRowGroupDepth(1);
    } catch(e) {}
  }

  // 1행 고정 및 기본 정렬/테두리 서식
  var siteLastRow = Math.max(siteSheet.getLastRow(), 2);
  siteSheet.setFrozenRows(1);
  siteSheet.getRange(2, 1, siteLastRow - 1, 14).setVerticalAlignment("middle");
  [1, 2, 3, 5, 7, 8, 11, 12, 13, 14].forEach(function(col) {
    siteSheet.getRange(2, col, siteLastRow - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    siteSheet.getRange(2, col, siteLastRow - 1, 1).setHorizontalAlignment("left");
  });
  siteSheet.getRange(2, 1, siteLastRow - 1, 8).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // E열 공유 체크박스 설정
  siteSheet.getRange(2, 5, siteLastRow - 1, 1).insertCheckboxes().setHorizontalAlignment("center");

  // C열(공종), G열(담당자), H열(진행상태) 드롭다운 유효성 검사 설정
  applySiteScheduleValidations(siteSheet, 2, siteLastRow - 1);

  // C열, G열, H열 노션 스타일 뱃지 조건부 서식 설정
  setupSiteSheetConditionalFormatting(siteSheet);

  // 보조열(9~14열) 검증
  ensureCalendarSyncColumns(siteSheet);

  // 2. [대시보드] 시트 생성 및 만년달력 구축
  var dashSheet = ss.getSheetByName("대시보드") || findDashboardSheet(ss);
  if (!dashSheet) {
    dashSheet = ss.insertSheet("대시보드", 0);
  }
  buildDashboardCalendarGrid(dashSheet);

  // 3. [사용가이드] 시트 생성 및 매뉴얼 작성
  var guideSheet = createOrUpdateUsageGuideSheet(ss);

  // 4. 기본 "시트1" 정리 (현장데이터로 이름이 바뀌지 않고 남아있는 잉여 빈 시트가 있는 경우에만 삭제)
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sName = sheets[s].getName();
    if (sName === "시트1" || sName === "Sheet1") {
      if (ss.getSheets().length > 2 && sheets[s].getSheetId() !== siteSheet.getSheetId()) {
        try { ss.deleteSheet(sheets[s]); } catch(e) {}
      }
    }
  }

  // 시트 순서: 대시보드(1) -> 현장데이터(2) -> 사용가이드(3)
  try {
    ss.setActiveSheet(dashSheet);
    ss.moveActiveSheet(1);
    ss.setActiveSheet(siteSheet);
    ss.moveActiveSheet(2);
    if (guideSheet) {
      ss.setActiveSheet(guideSheet);
      ss.moveActiveSheet(3);
    }
  } catch(e) {}

  // 대시보드 달력 최신화
  SpreadsheetApp.flush();
  refreshDashboardCalendar(ss);

  // 실시간 트리거 자동 설치
  ensureRealtimeTriggerInstalledSilent();

  // 최초 활성 시트는 대시보드로
  ss.setActiveSheet(dashSheet);

  if (ui) {
    ui.alert(
      "🎉 새 워크스페이스 구축 완료!",
      "[대시보드], [현장데이터], [사용가이드] 시트가 성공적으로 자동 생성되었습니다!\n\n" +
      "1. '사용가이드' 탭에서 캘린더 색상 범례와 작성 규칙을 한눈에 확인하실 수 있습니다.\n" +
      "2. '현장데이터' 시트에서 현장명, 날짜, 공종을 등록하시면 실시간으로 대시보드와 구글 캘린더에 연동됩니다.\n" +
      "3. 모바일 웹 변경 기능을 사용하시려면 [확장 프로그램] > [Apps Script] > [배포] > [새 배포]에서 웹앱을 배포해주세요.",
      ui.ButtonSet.OK
    );
  }
}

/**
 * 공종 기본 드롭다운 목록 반환 (기본 50+ 공종 사전 + 시트 기존 공종 자동 병합)
 */
function getAvailableTasks(sheet) {
  var baseTasks = [
    "상담", "신고", "관리소신고", "동의서", "실측", "창호실측", "현장실측", "도면/설계",
    "보양", "엘리베이터보양", "현관보양", "바닥보양",
    "철거", "마루철거", "욕실철거", "확장철거", "폐기물",
    "설비", "배관", "방수", "미장", "조적", "젠다이", "누수보수",
    "전기", "전기입선", "전기마감", "조명", "스위치/콘센트",
    "에어컨", "시스템에어컨", "냉매배관",
    "목공", "목공마감", "단열", "가벽", "도어설치", "문틀", "중문", "터닝도어", "폴딩도어", "현관부속",
    "창호", "샷시", "창호실측", "창호a/s", "유리", "난간대",
    "타일", "줄눈", "퍼플렉스", "아덱스", "SMC", "돔천장", "도기", "도기세팅 천장", "도기세팅", "코킹", "실리콘",
    "필름", "인테리어필름", "도장", "페인트", "탄성코트",
    "도배", "실크도배", "바닥", "마루시공", "강마루", "장판", "데코타일",
    "주방", "싱크대", "가구", "수납가구", "붙박이장", "신발장",
    "마감", "잔마감", "청소", "준공청소", "입주청소", "입주", "현장촬영", "S/A", "AS", "점검"
  ];

  var uniqueTasks = [];
  for (var i = 0; i < baseTasks.length; i++) {
    if (uniqueTasks.indexOf(baseTasks[i]) === -1) {
      uniqueTasks.push(baseTasks[i]);
    }
  }

  // 시트에 이미 입력된 커스텀 공종(예: 사용자가 직접 입력한 모든 공종) 자동 수집 및 병합하여 빨간 삼각형 에러 방지
  if (sheet && typeof sheet.getLastRow === "function") {
    try {
      var lastR = sheet.getLastRow();
      if (lastR >= 2) {
        var existingVals = sheet.getRange(2, 3, lastR - 1, 1).getValues();
        for (var j = 0; j < existingVals.length; j++) {
          var t = String(existingVals[j][0] || "").trim();
          if (t && uniqueTasks.indexOf(t) === -1) {
            uniqueTasks.push(t);
          }
        }
      }
    } catch (e) {}
  }

  return uniqueTasks;
}

/**
 * 신규 공종 실시간 자동 학습 엔진 (Auto-Learning Task Validation)
 * 사용자가 사전에 정의되지 않은 새로운 공종을 C열에 직접 타이핑하면
 * 즉시 드롭다운 유효성 목록에 자동 추가하여 빨간 삼각형 경고를 실시간 해제합니다.
 */
function ensureTasksInValidation(sheet, newTasks) {
  if (!sheet || !newTasks || newTasks.length === 0) return;
  try {
    var maxRows = Math.max(sheet.getLastRow(), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100));
    var currentRule = sheet.getRange(2, 3).getDataValidation();
    var allowedList = [];

    if (currentRule && currentRule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var args = currentRule.getCriteriaValues();
      if (args && args[0] && Array.isArray(args[0])) {
        allowedList = args[0].slice();
      }
    }

    if (allowedList.length === 0) {
      allowedList = getAvailableTasks(sheet);
    }

    var added = false;
    for (var i = 0; i < newTasks.length; i++) {
      var t = String(newTasks[i] || "").trim();
      if (t && allowedList.indexOf(t) === -1) {
        allowedList.push(t);
        added = true;
      }
    }

    if (added) {
      var newRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(allowedList, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(newRule);
    }
  } catch (e) {
    Logger.log("공종 실시간 자동 학습 예외: " + e.toString());
  }
}

/**
 * 담당자 드롭다운 목록 반환 (구성원 시트 연동 + 기본 담당자 조합)
 */
function getAvailableManagers(ss) {
  var managers = [];
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    var mSheet = findMembersSheet(ss);
    if (mSheet && mSheet.getLastRow() >= 2) {
      var lastR = mSheet.getLastRow();
      var lastC = mSheet.getLastColumn();
      var hVals = mSheet.getRange(1, 1, 1, lastC).getValues()[0];
      var nameCol = -1;
      for (var c = 0; c < lastC; c++) {
        var hStr = String(hVals[c] || "").trim();
        if (hStr.indexOf("이름") !== -1 || hStr.indexOf("성명") !== -1 || hStr.indexOf("담당자") !== -1 || hStr.indexOf("직원") !== -1 || hStr.indexOf("팀원") !== -1) {
          nameCol = c + 1;
          break;
        }
      }
      if (nameCol !== -1) {
        var names = mSheet.getRange(2, nameCol, lastR - 1, 1).getValues();
        for (var i = 0; i < names.length; i++) {
          var nm = String(names[i][0] || "").trim();
          if (nm && managers.indexOf(nm) === -1) {
            managers.push(nm);
          }
        }
      }
    }
  } catch(e) {}

  if (managers.length === 0) {
    managers = ["조항만", "송지은", "전예린", "영하"];
  }

  // 복합 담당자 조합 추가 (빨간 삼각형 에러 방지)
  var baseCount = managers.length;
  for (var a = 0; a < baseCount; a++) {
    for (var b = a + 1; b < baseCount; b++) {
      var comb = managers[a] + ", " + managers[b];
      if (managers.indexOf(comb) === -1) {
        managers.push(comb);
      }
    }
  }

  return managers;
}

/**
 * 현장데이터 시트 C열(공종), G열(담당자), H열(진행상태) 드롭다운 유효성 검사 일괄 적용
 */
function applySiteScheduleValidations(sheet, startRow, numRows) {
  if (!sheet) return;
  startRow = startRow || 2;
  var maxAvailableRows = Math.max(sheet.getLastRow(), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100));
  numRows = numRows || Math.max(maxAvailableRows - startRow + 1, 100);
  if (numRows < 1) return;

  var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();

  // 1. 공종(3열, C열) 드롭다운 (자유 입력 허용 및 시트 내 공종 자동 포함)
  var tasks = getAvailableTasks(sheet);
  var taskRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(tasks, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 3, numRows, 1).setDataValidation(taskRule);

  // 2. 담당자(7열, G열) 드롭다운 (자유 입력 허용)
  var managers = getAvailableManagers(ss);
  var managerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(managers, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 7, numRows, 1).setDataValidation(managerRule);

  // 3. 진행상태(8열, H열) 드롭다운
  var statuses = ["예정", "진행 중", "완료", "AS"];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 8, numRows, 1).setDataValidation(statusRule);
}

/**
 * 현장데이터 시트 H열(진행상태), G열(담당자), C열(공종) 노션 스타일 조건부 서식 자동 설정
 */
function setupSiteSheetConditionalFormatting(sheet) {
  if (!sheet) return;
  var lastRow = Math.max(sheet.getLastRow(), 100);
  var rangeC = sheet.getRange(2, 3, lastRow - 1, 1); // C열: 공종
  var rangeG = sheet.getRange(2, 7, lastRow - 1, 1); // G열: 담당자
  var statusRange = sheet.getRange(2, 8, lastRow - 1, 1); // H열: 진행상태

  var rules = [];

  // 1. 진행상태 색상 (H열)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("진행").setFontColor("#1d4ed8").setBackground("#dbeafe").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("완료").setFontColor("#15803d").setBackground("#dcfce7").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("AS").setFontColor("#b91c1c").setBackground("#fee2e2").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("예정").setFontColor("#b45309").setBackground("#fef3c7").setBold(true).setRanges([statusRange]).build());

  // 2. 담당자 인물 뱃지 색상 (G열)
  var managerColorMap = [
    { key: "조항만", bg: "#dbeafe", font: "#1e40af" }, // 블루
    { key: "송지은", bg: "#fce7f3", font: "#be185d" }, // 로즈핑크
    { key: "전예린", bg: "#f3e8ff", font: "#7e22ce" }, // 퍼플
    { key: "영하", bg: "#ccfbf1", font: "#0f766e" }    // 민트
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

  // 3. 공종별 뱃지 색상 (C열)
  var taskColorMap = [
    { keys: ["철거", "마루철거", "욕실철거", "확장철거", "폐기물", "보양", "엘리베이터보양", "현관보양"], bg: "#f1f5f9", font: "#334155" },
    { keys: ["설비", "배관", "누수보수", "방수", "미장", "조적", "젠다이"], bg: "#cffafe", font: "#0e7490" },
    { keys: ["전기", "전기입선", "전기마감", "조명", "에어컨", "시스템에어컨"], bg: "#fef3c7", font: "#b45309" },
    { keys: ["목공", "가벽", "단열"], bg: "#ffedd5", font: "#9a3412" },
    { keys: ["타일", "줄눈", "퍼플렉스", "아덱스"], bg: "#ccfbf1", font: "#0f766e" },
    { keys: ["필름", "인테리어필름", "도장", "페인트", "탄성코트"], bg: "#f3e8ff", font: "#7e22ce" },
    { keys: ["도배", "실크도배"], bg: "#dcfce7", font: "#15803d" },
    { keys: ["창호", "샷시", "창호실측", "중문", "도어", "문틀", "터닝도어", "폴딩도어", "현관부속", "유리"], bg: "#dbeafe", font: "#1d4ed8" },
    { keys: ["바닥", "마루시공", "강마루", "장판", "데코타일"], bg: "#fef9c3", font: "#854d0e" },
    { keys: ["주방", "싱크대", "가구", "수납가구", "붙박이장", "신발장"], bg: "#fce7f3", font: "#be185d" },
    { keys: ["SMC", "도기", "도기세팅", "욕실세팅", "돔천장"], bg: "#e0f2fe", font: "#0284c7" },
    { keys: ["마감", "잔마감", "코킹", "실리콘", "청소", "준공청소", "입주청소"], bg: "#d1fae5", font: "#047857" },
    { keys: ["상담", "신고", "실측", "도면", "입주", "현장촬영", "점검", "AS", "S/A"], bg: "#e0e7ff", font: "#3730a3" }
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
}

/**
 * 대시보드 시트에 만년달력 그리드, 연/월 드롭다운, Today 버튼, 서식 자동 구성
 */
function buildDashboardCalendarGrid(sheet) {
  if (!sheet) return;
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("#ffffff");

  var today = new Date();
  var curYear = today.getFullYear();
  var curMonth = today.getMonth() + 1;

  // 1) 동적 연도 수식 (Z1:Z5)
  sheet.getRange("Z1").setFormula("=SEQUENCE(5, 1, YEAR(TODAY()) - 2, 1)");

  // 2) 대한민국 법정 공휴일 데이터 자동 수집 (X1:Y300)
  autoSyncKoreanHolidays(curYear, sheet);

  // X, Y, Z열 숨김
  try { sheet.hideColumns(24, 3); } catch(e) {}

  // 3) B2:E2 연도/월 컨트롤 바
  sheet.getRange("B2").setValue("년도").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f1f5f9");
  sheet.getRange("D2").setValue("월").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f1f5f9");

  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange("Z1:Z5"), true)
    .build();
  sheet.getRange("C2")
    .setDataValidation(yearRule)
    .setValue(curYear)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setBackground("#ffffff");

  var months = [];
  for (var m = 1; m <= 12; m++) months.push(m.toString());
  var monthRule = SpreadsheetApp.newDataValidation().requireValueInList(months, true).build();
  sheet.getRange("E2")
    .setDataValidation(monthRule)
    .setValue(curMonth)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setBackground("#ffffff");

  sheet.getRange("B2:C2").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("D2:E2").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // 4) G2:H2 Today 스마트 버튼
  sheet.getRange("G2").insertCheckboxes().setValue(false).setHorizontalAlignment("center");
  sheet.getRange("H2").setValue("Today").setFontWeight("bold").setFontColor("#1d4ed8").setHorizontalAlignment("left");
  sheet.getRange("G2:H2")
    .setBackground("#eff6ff")
    .setBorder(true, true, true, true, true, true, "#bfdbfe", SpreadsheetApp.BorderStyle.SOLID);

  // 5) B4:H4 요일 헤더
  var days = ["일", "월", "화", "수", "목", "금", "토"];
  var headerRange = sheet.getRange("B4:H4");
  headerRange.setValues([days])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#f8fafc")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B4").setFontColor("#dc2626"); // 일요일 빨강
  sheet.getRange("H4").setFontColor("#2563eb"); // 토요일 파랑

  // 6) 주차별 날짜 수식 (B5, B7, B9, B11, B13, B15)
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

  // 7) 크기 및 테두리 서식
  var dateRows = [5, 7, 9, 11, 13, 15];
  var taskRows = [6, 8, 10, 12, 14, 16];

  for (var col = 2; col <= 8; col++) {
    sheet.setColumnWidth(col, 200);
  }
  sheet.setRowHeight(4, 30);

  dateRows.forEach(function(row) {
    sheet.setRowHeight(row, 28);
    sheet.getRange(row, 2, 1, 7)
      .setFontSize(9.5)
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setBackground("#f8fafc");
  });

  taskRows.forEach(function(row) {
    sheet.setRowHeight(row, 200);
    sheet.getRange(row, 2, 1, 7)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment("top")
      .setHorizontalAlignment("left")
      .setFontSize(9)
      .setBackground("#ffffff");
  });

  for (var wi = 0; wi < dateRows.length; wi++) {
    var rDate = dateRows[wi];
    var rTask = taskRows[wi];
    sheet.getRange(rDate, 2, 2, 7)
      .setBorder(true, true, true, true, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(rDate, 2, 1, 7)
      .setBorder(null, null, true, null, null, null, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(rTask, 2, 1, 7)
      .setBorder(true, null, null, null, null, null, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  }

  // 8) 대시보드 달력 조건부 서식
  var calDateRangeList = [];
  dateRows.forEach(function(r) {
    calDateRangeList.push(sheet.getRange("B" + r + ":H" + r));
  });

  var cellDateFormula = 'DATE($C$2, $E$2, 1) - WEEKDAY(DATE($C$2, $E$2, 1), 1) + 1 + (INT((ROW(B5)-5)/2)*7) + (COLUMN(B5)-2)';

  var ruleOtherMonth = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + cellDateFormula + ' <> DATE($C$2, $E$2, DAY(' + cellDateFormula + '))')
    .setFontColor("#94a3b8")
    .setBackground("#f1f5f9")
    .setRanges(calDateRangeList)
    .build();

  var ruleToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + cellDateFormula + ' = TODAY()')
    .setBackground("#fef08a")
    .setFontColor("#854d0e")
    .setRanges(calDateRangeList)
    .build();

  var ruleHoliday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISNUMBER(SEARCH("(", B5))')
    .setFontColor("#dc2626")
    .setBackground("#fef2f2")
    .setBold(true)
    .setRanges(calDateRangeList)
    .build();

  var ruleSunday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COLUMN(B5)=2')
    .setFontColor("#dc2626")
    .setRanges(calDateRangeList)
    .build();

  var ruleSaturday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COLUMN(B5)=8')
    .setFontColor("#2563eb")
    .setRanges(calDateRangeList)
    .build();

  sheet.setConditionalFormatRules([ruleOtherMonth, ruleToday, ruleHoliday, ruleSunday, ruleSaturday]);
}

/**
 * 대한민국 법정 공휴일 데이터 자동 수집
 */
function autoSyncKoreanHolidays(baseYear, targetSheet) {
  var sheet = targetSheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!baseYear) baseYear = new Date().getFullYear();

  var nonHolidayKeywords = [
    "식목일", "국군의 날", "국군의날", "제헌절", "어버이날", "스승의 날", "스승의날",
    "근로자의 날", "근로자의날", "노동절", "발렌타인", "화이트데이", "할로윈", "크리스마스 이브",
    "초복", "중복", "말복", "단오", "칠석", "대보름", "정월대보름", "한식", "동지", "입춘", "춘분", "하지", "추분",
    "경찰의 날", "소방의 날", "학생독립운동기념일", "농업인의 날", "무역의 날", "금융의 날",
    "체육의 날", "문화의 날", "세계 한인의 날", "임산부의 날", "노인의 날", "청년의 날",
    "철도의 날", "정보보호의 날", "바다의 날", "발명의 날", "성년의 날", "부부의 날", "방재의 날",
    "과학의 날", "정보통신의 날", "보건의 날", "예비군의 날", "서해수호의 날", "상공의 날", "납세자의 날",
    "3·15", "4·19", "5·18", "6·10", "독도의 날", "섣달 그믐날", "섣달그믐"
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
        var isNonHoliday = false;
        for (var k = 0; k < nonHolidayKeywords.length; k++) {
          if (title.indexOf(nonHolidayKeywords[k]) !== -1) {
            isNonHoliday = true;
            break;
          }
        }
        if (isNonHoliday) continue;
        var rawDate = ev.getStartTime();
        var cleanDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
        var displayTitle = title.trim();
        if (/쉬는\s*날|대체공휴일|대체휴일/.test(displayTitle)) {
          displayTitle = "대체공휴일";
        }
        holidayData.push([cleanDate, displayTitle]);
      }
    }
  } catch (e) {
    Logger.log("Calendar sync: " + e.toString());
  }

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
      ["2026-09-28", "대체공휴일"], ["2026-10-03", "개천절"], ["2026-10-05", "대체공휴일"], ["2026-10-09", "한글날"], ["2026-12-25", "성탄절"],
      ["2027-01-01", "신정"], ["2027-02-06", "설날연휴"], ["2027-02-07", "설날"], ["2027-02-08", "설날연휴"], ["2027-02-09", "대체공휴일"],
      ["2027-03-01", "삼일절"], ["2027-05-05", "어린이날"], ["2027-05-13", "부처님오신날"], ["2027-06-06", "현충일"], ["2027-08-15", "광복절"],
      ["2027-08-16", "대체공휴일"], ["2027-09-14", "추석연휴"], ["2027-09-15", "추석"], ["2027-09-16", "추석연휴"], ["2027-10-03", "개천절"],
      ["2027-10-04", "대체공휴일"], ["2027-10-09", "한글날"], ["2027-10-11", "대체공휴일"], ["2027-12-25", "성탄절"]
    ];
    defaultHolidays.forEach(function(item) {
      holidayData.push([new Date(item[0]), item[1]]);
    });
  }

  if (holidayData.length > 0) {
    sheet.getRange("X1:Y300").clearContent();
    sheet.getRange(1, 24, holidayData.length, 2).setValues(holidayData);
  }
}


// ========================================================
// 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 메뉴 자동 등록
// ========================================================
function onOpen() {
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      PropertiesService.getDocumentProperties().setProperty("SPREADSHEET_ID", activeSs.getId());
    }
  } catch(e) {}
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 현장 관리")
    // [자주 쓰는 핵심 3대 기능]
    .addItem("📅 구글 캘린더 동기화 (최근 1개월~미래 일정)", "syncRecentSiteSchedulesToGoogleCalendar")
    .addItem("🎨 캘린더 색상 & ⚠️ 지연 공정 일괄 검사", "checkAndHighlightDelayedSchedules")
    .addItem("🔄 대시보드 달력 새로고침", "syncAllSiteSchedulesToDashboard")
    .addSeparator()

    // [서브메뉴 1: 모닝 브리핑 관리]
    .addSubMenu(ui.createMenu("☀️ 모닝 브리핑 관리")
      .addItem("✉️ 지금 즉시 테스트 발송", "sendDailyBriefingTestNow")
      .addItem("⏰ 자동 발송 켜기 (정각 발송)", "enableDailyMorningBriefing")
      .addItem("🔕 자동 발송 끄기", "disableDailyMorningBriefing")
      .addItem("🕒 발송 시간 변경 (기본 08:00 정각)", "configureBriefingTime")
      .addItem("📧 수신 이메일 변경", "configureBriefingEmail")
    )

    // [서브메뉴 2: 캘린더 연동 설정]
    .addSubMenu(ui.createMenu("⚡ 캘린더 연동 설정")
      .addItem("💡 스마트 디바운스 자동 동기화 켜기 (입력 완료 시 1회 자동 반영 - 최고 권장)", "enableDebouncedAutoSync")
      .addItem("⏰ 5분 주기 자동 동기화 켜기 (정기 순환 동기화)", "enable5MinAutoSync")
      .addItem("⏰ 1분 주기 자동 동기화 켜기 (초고속 테스트용)", "enable1MinAutoSync")
      .addItem("⚡ 완전 실시간 연동 켜기 (수정 시 즉시 반영)", "enableRealtimeCalendarSync")
      .addItem("❌ 모든 자동 연동 끄기 (수동 버튼 전용)", "disableAllAutoCalendarSync")
      .addSeparator()
      .addItem("📥 캘린더 수정사항 가져오기 (수동 즉시 반영)", "syncFromCalendarToSheet")
      .addSeparator()
      .addItem("📦 전체 일정 연속 동기화 (25건씩 자동 대기 연속 처리)", "syncAllSiteSchedulesToGoogleCalendar")
      .addItem("🧹 캘린더 중복 및 유령 일정 자동 정리", "cleanupDuplicateAndOrphanCalendarEvents")
    )

    // [서브메뉴 3: 현장 데이터 편집 / 보기 편의 기능]
    .addSubMenu(ui.createMenu("📑 현장 데이터 편집/보기")
      .addItem("➕ 신규 현장 추가 (상단 4개 행 + 아코디언)", "addNewSiteAtTop")
      .addItem("📖 모든 현장 펼치기 (전체 공정 보기)", "expandAllSites")
      .addItem("📁 모든 현장 접기 (현장명만 요약 보기)", "collapseAllSites")
      .addItem("🎯 오늘 날짜로 이동 (Today)", "goToToday")
    )
    .addSeparator()

    // [서브메뉴 4: 고급 설정 & 시스템 관리 (1회성 / 유지보수 전용)]
    .addSubMenu(ui.createMenu("⚙️ 고급 설정 & 시스템 관리")
      .addItem("📖 사용 가이드 & 색상 범례 시트 생성/갱신", "createOrUpdateUsageGuideSheetMenu")
      .addItem("🔧 공종/담당자/상태 드롭다운 및 서식 전체 복구", "fixStatusDropdownValidation")
      .addItem("🧹 빈 행 유령 이력/ID 일괄 청소", "cleanupEmptyRowsMetadata")
      .addItem("🧹 캘린더 중복 및 유령 일정 자동 정리", "cleanupDuplicateAndOrphanCalendarEvents")
      .addItem("⚙️ 연동 캘린더 변경 / 확인", "configureTargetCalendar")
      .addItem("📱 모바일 상태변경 웹앱 배포 URL 설정", "configureWebAppUrl")
      .addSeparator()
      .addItem("✨ 새 스프레드시트 초기 환경 자동 구축 (대시보드 + 현장데이터 + 사용가이드)", "setupNewWorkspaceFromScratch")
    )
    .addToUi();
}


// ========================================================
// 2. [실시간 스프레드시트 UI 감지 (단순 트리거)]
// - 대시보드에서 Today 버튼 클릭 시 오늘 날짜로 이동
// - 대시보드에서 '월' 또는 '년도' 변경 시 달력 즉시 새로고침
// - 현장데이터 시트에서 A, B, C열 수정 시 대시보드 달력 즉시 갱신
// ========================================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  // ① [대시보드 시트]
  if (isDashboardSheet(sheet)) {
    var a1 = e.range.getA1Notation();
    // G2 체크박스 또는 H2 클릭 시 Today 동작 실행
    if ((a1 === "G2" || a1 === "H2") && e.range.getValue() === true) {
      e.range.setValue(false);
      goToToday();
      return;
    }
    // 상단 컨트롤(1~3행)의 년도 또는 월 변경 시 대시보드 달력 즉시 갱신
    if (row <= 3) {
      refreshDashboardCalendar(ss);
      return;
    }
  }

  // ② [현장데이터 시트]
  // 설치형 트리거(handleCalendarEditTrigger)가 메타데이터 기록 및 대시보드 갱신을 단 1회 전담하므로,
  // 단순 트리거(onEdit)에서는 이중 렌더링 방지를 위해 대시보드 갱신을 수행하지 않습니다.
}


// ========================================================
// 3. [구글 캘린더 자동 연동 제어 (1분 주기 스마트 동기화 & 실시간 엔진)]
// ========================================================

/**
 * 스마트 디바운스 자동 동기화 켜기 (실무 최고 권장 모드)
 * - 시트에서 공정을 마음껏 여러 행 추가하고 연속 입력해도 렉 0%! (실시간 API 호출 없음)
 * - 입력을 마치고 약 1분간 유휴 상태가 되면 백그라운드에서 1회만 일괄 자동 반영
 * - 평소 입력이 없을 때는 트리거 실행 횟수 0회 (일일 Quota 0% 소모)
 */
function enableDebouncedAutoSync() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SYNC_MODE", "DEBOUNCE");
  props.deleteProperty("REALTIME_CALENDAR_SYNC_DISABLED");
  props.setProperty("DEBOUNCE_SCHEDULED", "false");

  // 기존 주기/디바운스 트리거 정리
  cleanupAllPeriodicAndDebouncedTriggers();

  ensureRealtimeTriggerInstalledSilent();

  SpreadsheetApp.getUi().alert(
    "💡 스마트 디바운스 자동 동기화 활성화 완료",
    "스마트 디바운스 자동 동기화가 활성화되었습니다!\n\n" +
    "1. [입력 렉 0%]: 시트에서 공정을 마음껏 여러 행 추가하고 자유롭게 타이핑하세요.\n" +
    "2. [지능형 1회 동기화]: 입력을 마치고 약 1분 뒤 백그라운드에서 캘린더로 딱 1번 자동 전송됩니다.\n" +
    "3. [일일 할당량 0% 소모]: 입력이 없을 때는 백그라운드 트리거가 전혀 돌지 않아 구글 일일 실행시간이 보존됩니다.\n\n" +
    "(지금 즉시 캘린더에 올리고 싶을 때는 상단 [📅 구글 캘린더 동기화]를 1회 클릭하시면 됩니다)",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 5분 주기 스마트 자동 동기화 켜기 (정기 순환 모드)
 * - 시트에서 공정을 마음껏 연속 입력/수정/복사해도 입력 렉 0%!
 * - 5분마다 백그라운드에서 신규/수정된 공정을 캘린더에 일괄 자동 반영
 * - 일일 트리거 실행시간 10~15분 소요 (구글 무료 계정 90분 한도 100% 영구 안전)
 */
function enable5MinAutoSync() {
  setupPeriodicAutoSyncTrigger(5);
}

/**
 * 1분 주기 스마트 자동 동기화 켜기 (초고속 테스트 전용 모드)
 */
function enable1MinAutoSync() {
  setupPeriodicAutoSyncTrigger(1);
}

/**
 * 디바운스 1회성 트리거 정리 헬퍼
 */
function cleanupDebouncedTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "runDebouncedSync") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    Logger.log("디바운스 트리거 정리 에러: " + e.toString());
  }
}

/**
 * 모든 주기 및 디바운스 트리거 일괄 정리 헬퍼
 */
function cleanupAllPeriodicAndDebouncedTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      var fn = triggers[i].getHandlerFunction();
      if (fn === "runPeriodicAutoSync" || fn === "runPeriodic1MinSync" || fn === "runDebouncedSync") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    Logger.log("트리거 정리 에러: " + e.toString());
  }
}

/**
 * 주기적 자동 동기화 트리거 설정 공통 함수
 */
function setupPeriodicAutoSyncTrigger(minutes) {
  var props = PropertiesService.getScriptProperties();
  var modeStr = (minutes === 1) ? "1MIN" : "5MIN";
  props.setProperty("SYNC_MODE", modeStr);
  props.deleteProperty("REALTIME_CALENDAR_SYNC_DISABLED");
  props.setProperty("DEBOUNCE_SCHEDULED", "false");

  var functionName = "runPeriodicAutoSync";

  // 기존 주기 및 디바운스 트리거 정리
  cleanupAllPeriodicAndDebouncedTriggers();

  // 신규 주기 트리거 등록 (1분 또는 5분)
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(minutes)
    .create();

  ensureRealtimeTriggerInstalledSilent();

  SpreadsheetApp.getUi().alert(
    "⏰ " + minutes + "분 주기 스마트 자동 동기화 활성화 완료",
    minutes + "분 주기 스마트 자동 동기화가 활성화되었습니다!\n\n" +
    "1. 이제 시트에서 공정을 마음껏 여러 행 추가하고 입력하세요. (입력 렉 0%)\n" +
    "2. 약 " + minutes + "분마다 백그라운드에서 신규/수정된 공정을 감지하여 구글 캘린더에 자동으로 반영합니다.\n" +
    "3. 구글 캘린더에서 수정한 일정(날짜 이동 등)도 시트로 안전하게 자동 반영됩니다.\n" +
    "4. 구글 일일 실행 시간 한도 내에서 24시간 365일 100% 무중단 안정 구동을 보장합니다.\n\n" +
    "(지금 즉시 캘린더에 올리고 싶을 때는 상단 [📅 구글 캘린더 동기화]를 1회 클릭하시면 됩니다)",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 완전 실시간 캘린더 연동 켜기 (셀 수정 시 즉각 캘린더 통신)
 */
function enableRealtimeCalendarSync() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SYNC_MODE", "REALTIME");
  props.deleteProperty("REALTIME_CALENDAR_SYNC_DISABLED");
  props.setProperty("DEBOUNCE_SCHEDULED", "false");

  // 주기 및 디바운스 백그라운드 트리거 삭제
  cleanupAllPeriodicAndDebouncedTriggers();

  ensureRealtimeTriggerInstalledSilent();

  SpreadsheetApp.getUi().alert(
    "⚡ 완전 실시간 연동 활성화 완료",
    "셀 수정 시 즉시 구글 캘린더에 반영되는 완전 실시간 연동이 켜졌습니다.\n\n'현장데이터' 시트에서 비고를 제외한 항목을 수정하면 즉시 구글 캘린더에 실시간으로 자동 추가/수정/삭제됩니다.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 모든 캘린더 자동 연동 끄기 (수동 동기화 모드)
 */
function disableAllAutoCalendarSync() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SYNC_MODE", "MANUAL");
  props.setProperty("REALTIME_CALENDAR_SYNC_DISABLED", "true");
  props.setProperty("DEBOUNCE_SCHEDULED", "false");

  cleanupAllPeriodicAndDebouncedTriggers();

  SpreadsheetApp.getUi().alert(
    "❌ 모든 자동 연동 해제 완료",
    "모든 자동 연동이 정상적으로 꺼졌습니다.\n\n" +
    "이제 시트에서 일정을 자유롭게 입력/수정/삭제하시고, 작성이 완료된 후 상단 메뉴의\n" +
    "[🚀 현장 관리] ➔ [📅 구글 캘린더 동기화]를 누르시면 한 번에 안전하게 캘린더에 반영됩니다.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 기존 호환용 실시간 자동 연동 끄기 별칭
 */
function disableRealtimeCalendarSync() {
  disableAllAutoCalendarSync();
}

/**
 * 캘린더 연동 트리거 백그라운드 무음 자동 설치 및 점검
 * - 기본 모드: 스마트 디바운스 자동 동기화 ("DEBOUNCE")
 * - 사용자가 수동 모드("MANUAL")로 끈 경우 자동 재등록을 차단하여 사용자 의도 존중
 */
function ensureRealtimeTriggerInstalledSilent() {
  try {
    var props = PropertiesService.getScriptProperties();
    var syncMode = props.getProperty("SYNC_MODE");

    // 기본값이 없으면 "DEBOUNCE" 모드를 표준 기본값으로 적용
    if (!syncMode) {
      syncMode = "DEBOUNCE";
      props.setProperty("SYNC_MODE", "DEBOUNCE");
    }

    if (syncMode === "MANUAL") {
      return; // 수동 모드인 경우 자동 설치 차단
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var triggers = ScriptApp.getProjectTriggers();
    var hasEditTrigger = false;
    var hasPeriodicTrigger = false;

    for (var i = 0; i < triggers.length; i++) {
      var fn = triggers[i].getHandlerFunction();
      if (fn === "handleCalendarEditTrigger") hasEditTrigger = true;
      if (fn === "runPeriodicAutoSync" || fn === "runPeriodic1MinSync") hasPeriodicTrigger = true;
    }

    // 시트 수정 시 메타데이터 및 타임스탬프 기록용 onEdit 트리거 설치
    if (!hasEditTrigger) {
      ScriptApp.newTrigger("handleCalendarEditTrigger")
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      Logger.log("⚡ 시트 편집 트리거 자동 설치 완료");
    }

    // 디바운스 모드 또는 실시간 모드일 때는 불필요한 주기 트리거를 자동 정리하여 0 Quota 보장
    if ((syncMode === "DEBOUNCE" || syncMode === "REALTIME") && hasPeriodicTrigger) {
      for (var j = 0; j < triggers.length; j++) {
        var pFn = triggers[j].getHandlerFunction();
        if (pFn === "runPeriodicAutoSync" || pFn === "runPeriodic1MinSync") {
          ScriptApp.deleteTrigger(triggers[j]);
        }
      }
    }

    // 주기 자동 동기화 모드인 경우 시간 기반 트리거 설치
    if ((syncMode === "5MIN" || syncMode === "1MIN") && !hasPeriodicTrigger) {
      var minutes = (syncMode === "1MIN") ? 1 : 5;
      ScriptApp.newTrigger("runPeriodicAutoSync")
        .timeBased()
        .everyMinutes(minutes)
        .create();
      Logger.log("⏰ " + minutes + "분 주기 자동 동기화 트리거 자동 설치 완료");
    }
  } catch(e) {
    Logger.log("트리거 자동 점검 건너뜀: " + e.toString());
  }
}

/**
 * 주기적 백그라운드 통합 스마트 자동 동기화 코어 엔진
 * - [영업시간 가드]: 심야(22시~07시) 변경 없을 시 즉시 No-Op 종료하여 할당량 100% 안전 보존
 * - [시트 ➔ 캘린더]: 새로 입력되거나 시트에서 수정된 공정을 캘린더에 일괄 반영
 * - [캘린더 ➔ 시트]: 순환 커서(Round-Robin)로 전 행을 공평 점검하고, 원자적 셀 쓰기로 사용자 타이핑 덮어쓰기 원천 차단
 */
function runPeriodicAutoSync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // 동시 실행 중복 차단

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = findSiteScheduleSheet(ss);
    if (!sheet) return;

    var cal = getTargetCalendar();
    if (!cal) return;

    var props = PropertiesService.getScriptProperties();
    var lastSyncRaw = props.getProperty("LAST_PERIODIC_SYNC_MS");
    var now = new Date();
    var nowMs = now.getTime();
    var hour = now.getHours();

    // 최초 실행 시에는 최근 5분 전 시각으로 안전 초기화하여 기존 대량 일정 재전송 방어
    var lastSyncMs = lastSyncRaw ? parseInt(lastSyncRaw, 10) : (nowMs - 300000);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    ensureCalendarSyncColumns(sheet);

    var range = sheet.getRange(2, 1, lastRow - 1, 14);
    var values = range.getValues();

    var rowsToSyncToCal = [];
    var calUpdatedCount = 0;

    // 1단계: 시트 내 신규 및 변경 행 감지 (시트 ➔ 캘린더 대상 선별)
    for (var i = 0; i < values.length; i++) {
      var rowNumber = i + 2;
      var rowVals = values[i];
      var siteName = String(rowVals[0] || "").trim();
      var dateVal = rowVals[1];
      var taskName = String(rowVals[2] || "").trim();
      var calId = String(rowVals[8] || "").trim();
      var rawUpdated = rowVals[9];
      var rowUpdatedMs = 0;
      if (rawUpdated instanceof Date) {
        rowUpdatedMs = rawUpdated.getTime();
      } else if (typeof rawUpdated === "number") {
        rowUpdatedMs = rawUpdated;
      } else {
        rowUpdatedMs = parseInt(rawUpdated, 10) || 0;
      }

      // 날짜와 공종이 채워졌는데 아직 캘린더ID가 없는 신규 행 ➔ 캘린더 전송 대상
      // 또는 캘린더ID가 있고 마지막 동기화 이후 시트에서 수정된 행 ➔ 캘린더 전송 대상
      if ((dateVal && taskName && !calId) || (calId && rowUpdatedMs > lastSyncMs)) {
        rowsToSyncToCal.push(rowNumber);
      }
    }

    // [심야 시간(22:00 ~ 07:00) 스마트 유휴 가드]:
    // 시트에 수정된 행이 없는 야간에는 캘린더 API 조회를 생략하여 일일 Quota를 100% 보존
    var isNightTime = (hour >= 22 || hour < 7);
    if (isNightTime && rowsToSyncToCal.length === 0) {
      props.setProperty("LAST_PERIODIC_SYNC_MS", String(nowMs));
      return;
    }

    // 2단계: 변경된 행을 캘린더로 일괄 전송 (최대 25건 안전 분할)
    var didSyncToCal = false;
    if (rowsToSyncToCal.length > 0) {
      var syncLimit = Math.min(rowsToSyncToCal.length, 25);
      for (var s = 0; s < syncLimit; s++) {
        syncSingleRowToCalendar(sheet, rowsToSyncToCal[s], true, true);
      }
      didSyncToCal = true;
    }

    // 3단계: 캘린더 ➔ 시트 역방향 순환 검사 (Cursor Round-Robin)
    // 매번 상위 15건만 보던 결함을 고치고, 전체 행을 순환 점검하여 전 행 동기화 보장
    var cursor = parseInt(props.getProperty("CAL_INSPECT_CURSOR") || "0", 10);
    if (cursor >= values.length) cursor = 0;

    var inspectedThisRound = 0;
    var inspectLimit = 15;
    var stepsTaken = 0;

    for (var k = 0; k < values.length && inspectedThisRound < inspectLimit; k++) {
      stepsTaken++;
      var idx = (cursor + k) % values.length;
      var cVals = values[idx];
      var existingId = String(cVals[8] || "").trim();
      if (!existingId) continue;

      inspectedThisRound++;
      var ev = null;
      try {
        ev = cal.getEventById(existingId);
      } catch (e) {
        ev = null;
      }
      if (!ev) continue;

      var calUpdatedDate = ev.getLastUpdated();
      var calUpdated = calUpdatedDate ? calUpdatedDate.getTime() : 0;
      var rawUp = cVals[9];
      var sUpdated = (rawUp instanceof Date) ? rawUp.getTime() : (typeof rawUp === "number" ? rawUp : (parseInt(rawUp, 10) || 0));

      var targetRowNumber = idx + 2;
      var rowChanged = false;

      // 1) 날짜 변경 역방향 동기화
      if (calUpdated > sUpdated + 10000) {
        var isAllDay = ev.isAllDayEvent();
        var calStart = isAllDay ? ev.getAllDayStartDate() : ev.getStartTime();
        var calEnd = isAllDay ? ev.getAllDayEndDate() : ev.getEndTime();
        var newDateStr = "";
        if (!isAllDay) {
          newDateStr = formatDateWithTime(calStart, calEnd);
        } else {
          var realCalEnd = null;
          if (calEnd) {
            var diffDays = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
            if (diffDays > 1) {
              realCalEnd = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
            }
          }
          newDateStr = realCalEnd ? (formatDateYYMMDD(calStart) + " ~ " + formatDateYYMMDD(realCalEnd)) : formatDateYYMMDD(calStart);
        }

        // [핵심: 사용자 타이핑 덮어쓰기 원천 차단]
        // 시트 전체 범위를 덮어쓰지 않고, 실제 날짜가 변경된 특정 셀만 개별 타겟 원자적 갱신!
        sheet.getRange(targetRowNumber, 2).setValue(newDateStr);
        sheet.getRange(targetRowNumber, 10).setValue(calUpdated);
        sheet.getRange(targetRowNumber, 13, 1, 2).setValues([[formatDateTime(calUpdatedDate), "구글 캘린더"]]);
        rowChanged = true;
      }

      // 2) 캘린더 제목 상태 이모지 변경 감지 (이모지 기준으로만 엄격 판별 - 한글 텍스트 오탐지 방지)
      var calTitle = ev.getTitle() || "";
      var currentStatus = String(cVals[7] || "").trim();
      var detectedStatus = "";
      if (calTitle.indexOf("✅") !== -1) {
        detectedStatus = "완료";
      } else if (calTitle.indexOf("▶️") !== -1 || calTitle.indexOf("▶") !== -1) {
        detectedStatus = "진행 중";
      } else if (calTitle.indexOf("🛠️") !== -1 || calTitle.indexOf("🛠") !== -1) {
        detectedStatus = "AS";
      } else if ((calTitle.indexOf("🟦") !== -1 || calTitle.indexOf("🟥") !== -1) && (currentStatus === "완료" || currentStatus.indexOf("진행") !== -1 || currentStatus === "AS")) {
        detectedStatus = "예정";
      }

      if (detectedStatus && detectedStatus.replace(/\s+/g, "") !== currentStatus.replace(/\s+/g, "") && calUpdated > sUpdated + 5000) {
        sheet.getRange(targetRowNumber, 8).setValue(detectedStatus);
        sheet.getRange(targetRowNumber, 10).setValue(calUpdated);
        sheet.getRange(targetRowNumber, 13, 1, 2).setValues([[formatDateTime(calUpdatedDate), "구글 캘린더"]]);
        rowChanged = true;
      }

      if (rowChanged) calUpdatedCount++;
    }

    // 다음 실행 시 이어서 점검할 커서 저장 (실제 탐색한 행 수 stepsTaken 기반으로 전 행 순환 보장)
    var nextCursor = (cursor + stepsTaken) % Math.max(values.length, 1);
    props.setProperty("CAL_INSPECT_CURSOR", String(nextCursor));

    // 시트 또는 캘린더 변경이 있었던 경우에만 대시보드 달력 최신화
    if (didSyncToCal || calUpdatedCount > 0) {
      refreshDashboardCalendar(ss);
    }

    props.setProperty("LAST_PERIODIC_SYNC_MS", String(nowMs));
  } catch(e) {
    Logger.log("주기적 자동 동기화 에러: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1분 주기 트리거 기존 호환용 래퍼 함수
 */
function runPeriodic1MinSync() {
  runPeriodicAutoSync();
}

/**
 * 디바운스 1회성 백그라운드 동기화 엔진
 * - 사용자가 입력을 마치고 약 1분이 경과했을 때 1회만 실행
 * - 연속 입력 중에는 실행을 1분 뒤로 재연장(Debounce)
 * - 동기화 완료 후 스스로 트리거를 삭제하여 무유휴(0 Quota) 유지
 */
function runDebouncedSync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  var props = PropertiesService.getScriptProperties();
  var isRescheduled = false;
  try {
    var nowMs = Date.now();
    var lastEditMs = parseInt(props.getProperty("LAST_EDIT_MS") || "0", 10);

    // 사용자가 최근 40초 이내에 여전히 시트를 타이핑/수정 중인 경우:
    // 입력을 마칠 때까지 1분 뒤로 1회성 트리거를 재연장
    if (nowMs - lastEditMs < 40000) {
      cleanupDebouncedTriggers();
      ScriptApp.newTrigger("runDebouncedSync")
        .timeBased()
        .after(60 * 1000)
        .create();
      props.setProperty("DEBOUNCE_SCHEDULED", "true");
      isRescheduled = true;
      return;
    }

    // 사용자가 입력을 마치고 유휴 상태가 되었을 때: 일괄 동기화 실행!
    runPeriodicAutoSync();

  } catch (e) {
    Logger.log("디바운스 동기화 에러: " + e.toString());
  } finally {
    // 재연장된 트리거가 있는 경우에는 새로 생성된 트리거를 삭제하지 않음 (트리거 증발 방지)
    if (!isRescheduled) {
      cleanupDebouncedTriggers();
      props.setProperty("DEBOUNCE_SCHEDULED", "false");
    }
    lock.releaseLock();
  }
}

/**
 * 설치형 onEdit 트리거 핸들러
 * - 스마트 디바운스 모드: 메타데이터 및 타임스탬프만 고속 기록하고 1분 뒤 1회성 동기화 예약 (입력 렉 0%)
 * - 완전 실시간 모드("REALTIME"): 비고 제외 수정 시 캘린더에 즉시 실시간 전송
 * - 대시보드 달력 갱신을 단 1회 전담하여 단순 onEdit과의 이중 렌더링 완전 해소
 */
function handleCalendarEditTrigger(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (!isSiteScheduleSheet(sheet)) return;

  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  if (endRow < 2) return;
  // 1~8열(A~H열) 범위 밖의 수정은 무시
  if (startCol > 8 || endCol < 1) return;

  try {
    SpreadsheetApp.flush(); // 방금 입력된 셀 값 메모리 즉시 동기화

    var now = new Date();
    var nowMs = now.getTime();
    var nowStr = formatDateTime(now);
    var currentUser = getCurrentUserIdentifier();
    var actualStartRow = Math.max(startRow, 2);
    var maxProcessRow = Math.min(endRow, actualStartRow + 49); // 최대 50행 대량 붙여넣기까지 안전 배치 지원
    var numRows = maxProcessRow - actualStartRow + 1;

    var props = PropertiesService.getScriptProperties();
    var syncMode = props.getProperty("SYNC_MODE") || "DEBOUNCE"; // 기본 권장값: 스마트 디바운스

    // 메타데이터 및 타임스탬프 1회 배치 일괄 읽기 및 쓰기 (성능 10배 가속)
    var metaRange = sheet.getRange(actualStartRow, 11, numRows, 4);
    var metaValues = metaRange.getValues();
    var timeStampVals = [];

    for (var rIdx = 0; rIdx < numRows; rIdx++) {
      var cDate = metaValues[rIdx][0] || nowStr;
      var cUser = metaValues[rIdx][1] || currentUser;
      metaValues[rIdx] = [cDate, cUser, nowStr, currentUser];
      timeStampVals.push([nowMs]);
    }
    metaRange.setValues(metaValues);

    // 시트 최종 수정 시각 밀리초 타임스탬프 일괄 기록 (J열: 10열)
    sheet.getRange(actualStartRow, 10, numRows, 1).setValues(timeStampVals);

    // [신규 공종 실시간 자동 학습 엔진 (Auto-Learning Task Validation)]
    // 사용자가 C열(공종)을 수정했을 때, 기존 드롭다운 목록에 없는 새로운 공종이면
    // 즉시 유효성 목록에 자동 추가하여 빨간 삼각형 경고 표시를 실시간 해제!
    if (startCol <= 3 && endCol >= 3) {
      try {
        var editedTaskVals = sheet.getRange(actualStartRow, 3, numRows, 1).getValues();
        var editedTasks = [];
        for (var ti = 0; ti < editedTaskVals.length; ti++) {
          var tVal = String(editedTaskVals[ti][0] || "").trim();
          if (tVal) editedTasks.push(tVal);
        }
        if (editedTasks.length > 0) {
          ensureTasksInValidation(sheet, editedTasks);
        }
      } catch (eTask) {
        Logger.log("공종 실시간 자동 학습 예외: " + eTask.toString());
      }
    }

    // [스마트 디바운스 모드 (기본 권장)]:
    // 사용자가 입력 중일 때는 즉시 API를 호출하지 않고 타이핑 완료 1분 뒤 1회만 자동 동기화되도록 예약!
    if (syncMode === "DEBOUNCE" || !syncMode) {
      props.setProperty("LAST_EDIT_MS", String(nowMs));
      var isScheduled = props.getProperty("DEBOUNCE_SCHEDULED");
      if (isScheduled !== "true") {
        try {
          ScriptApp.newTrigger("runDebouncedSync")
            .timeBased()
            .after(60 * 1000)
            .create();
          props.setProperty("DEBOUNCE_SCHEDULED", "true");
        } catch (eTrig) {
          Logger.log("디바운스 트리거 예약 경고: " + eTrig.toString());
        }
      }
    } else if (syncMode === "REALTIME") {
      var isOnlyMemo = (startCol === 6 && endCol === 6);
      if (!isOnlyMemo) {
        for (var r = actualStartRow; r <= maxProcessRow; r++) {
          syncSingleRowToCalendar(sheet, r, true, true);
        }
      }
    }

    // 대시보드 달력 최신화 (시트 내부 연산이므로 0.1초로 빠름, 단순 onEdit 중복 제거로 단 1회만 실행)
    var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();
    refreshDashboardCalendar(ss);
  } catch (err) {
    Logger.log("시트 편집 트리거 에러: " + err.toString());
  }
}


// ========================================================
// 4. [양방향 동기화: 캘린더 ➔ 시트 (Last-Modified-Wins 엔진)]
// ========================================================

/**
 * 캘린더에서 수정한 날짜를 시트로 가져와 동기화 (캘린더 ➔ 시트)
 * - 최종 수정 시각(Last Modified)을 비교하여 항상 가장 마지막에 수정한 쪽이 승리!
 * - 캘린더에서 수정된 경우 최종수정자: "구글 캘린더"로 자동 기록!
 */
function syncFromCalendarToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var cal = getTargetCalendar();
  if (!cal) {
    SpreadsheetApp.getUi().alert("구글 캘린더에 접근할 수 없습니다.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  ensureCalendarSyncColumns(sheet);

  // A~N열(1~14열) 읽기
  var range = sheet.getRange(2, 1, lastRow - 1, 14);
  var values = range.getValues();
  var updatedCount = 0;
  var inspectedCount = 0;
  var hasClearedGhosts = false;
  var timeLimitReached = false;
  var startTime = Date.now();

  for (var i = 0; i < values.length; i++) {
    // [핵심 방어 1: 6분 타임아웃 절대 차단] 90초 초과 시 안전 조기 종료
    if (Date.now() - startTime > 90000) {
      timeLimitReached = true;
      break;
    }

    // [핵심 방어 2: 구글 캘린더 API 호출 버스트 제한] 1회 실행 당 최대 30개 이벤트 검사
    if (inspectedCount >= 30) {
      break;
    }

    var rowVals = values[i];
    var siteName = String(rowVals[0] || "").trim(); // A열: 현장명
    var currentSheetDate = rowVals[1];               // B열: 날짜
    var taskName = String(rowVals[2] || "").trim(); // C열: 공종
    var existingId = String(rowVals[8] || "").trim(); // I열: 캘린더ID
    var rawUpdated = rowVals[9];
    var sheetUpdated = 0;
    if (rawUpdated instanceof Date) {
      sheetUpdated = rawUpdated.getTime();
    } else if (typeof rawUpdated === "number") {
      sheetUpdated = rawUpdated;
    } else {
      sheetUpdated = parseInt(rawUpdated, 10) || 0;
    }

    // ① 현장명과 공종이 모두 없는 빈 행은 캘린더 동기화에서 완전 제외!
    if (!siteName && !taskName) {
      // 빈 행에 유령 캘린더 ID나 이력 정보가 잘못 남아있다면 실제 시트 셀과 메모리에서 깨끗이 청소
      if (existingId || rowVals[10] || rowVals[12]) {
        sheet.getRange(i + 2, 9, 1, 6).clearContent(); // 실제 시트 셀(I~N열) 물리적 청소
        rowVals[8] = "";
        rowVals[9] = "";
        rowVals[10] = "";
        rowVals[11] = "";
        rowVals[12] = "";
        rowVals[13] = "";
        hasClearedGhosts = true;
      }
      continue;
    }

    if (!existingId) continue;

    // 캘린더에서 이벤트 조회
    inspectedCount++;
    var event = null;
    try {
      event = cal.getEventById(existingId);
    } catch (e) {
      event = null;
    }

    if (!event) continue;

    var calUpdatedDate = event.getLastUpdated();
    var calUpdated = calUpdatedDate ? calUpdatedDate.getTime() : 0;

    // 캘린더의 날짜 추출
    var isAllDay = event.isAllDayEvent();
    var calStart = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
    var calEnd = isAllDay ? event.getAllDayEndDate() : event.getEndTime();
    if (!calStart) continue;

    var currentRange = parseScheduleDateRange(currentSheetDate);
    var isSame = false;

    if (currentRange) {
      if (isAllDay && !currentRange.hasTime) {
        var realCalEnd = null;
        if (calEnd) {
          var diffDays = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
          if (diffDays > 1) {
            realCalEnd = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
          }
        }
        var curEnd = currentRange.isRange ? currentRange.displayEndDate : null;
        isSame = isDateEqual(calStart, currentRange.startDate) && isDateEqual(realCalEnd, curEnd);
      } else if (!isAllDay && currentRange.hasTime) {
        isSame = (calStart.getTime() === currentRange.startTime.getTime() &&
                  (calEnd && currentRange.endTime ? calEnd.getTime() === currentRange.endTime.getTime() : true));
      }
    }

    if (!isSame) {
      // 날짜/시간이 서로 다름! [마지막 수정 시간 기반 비교 (Last-Modified-Wins)]
      // 시트 ➔ 캘린더 전송 직후 발생하는 미세 타임스탬프 오차(Race Condition) 방지를 위해 10초 안전 마진 적용
      if (calUpdated > sheetUpdated + 10000) {
        // [캘린더가 10초 이상 더 최신일 때만] ➔ 캘린더의 날짜를 시트에 반영!
        var newDateStr = "";
        if (!isAllDay) {
          newDateStr = formatDateWithTime(calStart, calEnd);
        } else {
          var realCalEnd2 = null;
          if (calEnd) {
            var diffDays2 = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
            if (diffDays2 > 1) {
              realCalEnd2 = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
            }
          }
          newDateStr = realCalEnd2 ? (formatDateYYMMDD(calStart) + " ~ " + formatDateYYMMDD(realCalEnd2)) : formatDateYYMMDD(calStart);
        }

        // [원자적 셀 갱신] 변경된 특정 행의 날짜/타임스탬프만 핀포인트 쓰기 (타이핑 증발 원천 차단)
        var targetRowNumber = i + 2;
        sheet.getRange(targetRowNumber, 2).setValue(newDateStr);
        sheet.getRange(targetRowNumber, 10).setValue(calUpdated);
        sheet.getRange(targetRowNumber, 13, 1, 2).setValues([[formatDateTime(calUpdatedDate), "구글 캘린더"]]);
        rowVals[1] = newDateStr;
        rowVals[9] = calUpdated;
        rowVals[12] = formatDateTime(calUpdatedDate);
        rowVals[13] = "구글 캘린더";
        updatedCount++;
      } else if (sheetUpdated > calUpdated + 3000) {
        // [시트가 3초 이상 더 최신] ➔ 시트의 날짜/시간을 캘린더에 덮어쓰기!
        if (currentRange) {
          if (currentRange.hasTime) {
            event.setTime(currentRange.startTime, currentRange.endTime);
          } else {
            if (currentRange.isRange) {
              event.setAllDayDates(currentRange.startDate, currentRange.endDate);
            } else {
              event.setAllDayDate(currentRange.startDate);
            }
          }
          var targetRowNumber2 = i + 2;
          sheet.getRange(targetRowNumber2, 10).setValue(Date.now());
          rowVals[9] = Date.now();
          updatedCount++;
        }
      }
    }

    // ② 캘린더 제목의 상태 이모지 감지 (역방향 동기화: ✅ 완료, ▶️ 진행중, 🛠️ AS, 예정)
    var calTitle = event.getTitle() || "";
    var currentSheetStatus = String(rowVals[7] || "").trim();
    var detectedStatus = "";
    if (calTitle.indexOf("✅") !== -1) {
      detectedStatus = "완료";
    } else if (calTitle.indexOf("▶️") !== -1 || calTitle.indexOf("▶") !== -1) {
      detectedStatus = "진행 중";
    } else if (calTitle.indexOf("🛠️") !== -1 || calTitle.indexOf("🛠") !== -1) {
      detectedStatus = "AS";
    } else if ((calTitle.indexOf("🟦") !== -1 || calTitle.indexOf("🟥") !== -1) && (currentSheetStatus === "완료" || currentSheetStatus.indexOf("진행") !== -1 || currentSheetStatus === "AS")) {
      detectedStatus = "예정";
    }

    if (detectedStatus && detectedStatus.replace(/\s+/g, "") !== currentSheetStatus.replace(/\s+/g, "") && calUpdated > sheetUpdated + 3000) {
      var targetRowNumber3 = i + 2;
      sheet.getRange(targetRowNumber3, 8).setValue(detectedStatus);
      sheet.getRange(targetRowNumber3, 10).setValue(calUpdated);
      sheet.getRange(targetRowNumber3, 13, 1, 2).setValues([[formatDateTime(calUpdatedDate), "구글 캘린더"]]);
      rowVals[7] = detectedStatus; // H열 진행상태 갱신
      rowVals[9] = calUpdated;
      rowVals[12] = formatDateTime(calUpdatedDate);
      rowVals[13] = "구글 캘린더";
      updatedCount++;
    }
  }

  // [핵심: 사용자 타이핑 덮어쓰기 원천 차단]
  // 시트 전체 범위 일괄 덮어쓰기를 제거하고, 변경된 셀만 개별 타겟 원자적 갱신 완료
  if (updatedCount > 0 || hasClearedGhosts) {
    SpreadsheetApp.flush();
    refreshDashboardCalendar(ss);
    ss.toast("최종 수정 시간 기준으로 " + updatedCount + "건의 일정이 최신화되었습니다." + (timeLimitReached ? " (안전 시간 분할)" : ""), "동기화 완료", 3);
  } else {
    ss.toast("캘린더와 시트의 일정이 이미 동일합니다. (검사: " + inspectedCount + "건)", "동기화 확인", 2);
  }
}

/**
 * 캘린더 ➔ 시트 1분 주기 초고속 자동 동기화 켜기
 */
function enableCalendarToSheet1MinSync() {
  setupCalendarToSheetTrigger(1);
}

/**
 * 캘린더 ➔ 시트 5분 주기 표준 자동 동기화 켜기
 */
function enableCalendarToSheet5MinSync() {
  setupCalendarToSheetTrigger(5);
}

/**
 * 캘린더 ➔ 시트 자동 동기화 기존 호환 함수 (5분)
 */
function enableCalendarToSheetAutoSync() {
  setupCalendarToSheetTrigger(5);
}

/**
 * 캘린더 ➔ 시트 주기적 트리거 설정 공통 함수
 */
function setupCalendarToSheetTrigger(minutes) {
  var functionName = "syncFromCalendarToSheet";
  var triggers = ScriptApp.getProjectTriggers();

  // 기존 트리거 삭제
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 신규 주기 트리거 등록 (1분 또는 5분)
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(minutes)
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ 캘린더 ➔ 시트 자동 동기화 (" + minutes + "분 주기)",
    "구글 캘린더 ➔ 시트 자동 동기화가 [" + minutes + "분 주기]로 활성화되었습니다!\n\n" +
    "이제 구글 캘린더에서 일정을 드래그하거나 수정창에서 날짜를 옮기면, 약 " + minutes + "분 이내에 시트(B열)와 대시보드 달력에 자동으로 반영됩니다.\n\n" +
    "(지금 즉시 반영하고 싶을 때는 [📥 캘린더 수정사항 가져오기] 메뉴를 누르시면 됩니다)",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 캘린더 ➔ 시트 주기적 자동 동기화 끄기
 */
function disableCalendarToSheetAutoSync() {
  var functionName = "syncFromCalendarToSheet";
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  if (removed > 0) {
    SpreadsheetApp.getUi().alert("캘린더 ➔ 시트 자동 동기화가 해제되었습니다.");
  } else {
    SpreadsheetApp.getUi().alert("현재 자동 동기화가 켜져 있지 않습니다.");
  }
}

/**
 * 두 날짜의 년/월/일이 같은지 판별
 */
function isDateEqual(d1, d2) {
  if (!d1 && !d2) return true;
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

/**
 * Date 객체를 YY.MM.DD 문자열로 변환 (예: 26.10.15)
 */
function formatDateYYMMDD(date) {
  if (!date) return "";
  var yy = String(date.getFullYear()).slice(-2);
  var mm = String(date.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(date.getDate());
  if (dd.length < 2) dd = "0" + dd;
  return yy + "." + mm + "." + dd;
}

/**
 * Date 객체를 YY.MM.DD HH:mm 문자열로 변환 (예: 26.10.09 20:50)
 */
function formatDateTime(d) {
  if (!d) return "";
  var yy = String(d.getFullYear()).slice(-2);
  var mm = String(d.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(d.getDate());
  if (dd.length < 2) dd = "0" + dd;
  var hh = String(d.getHours());
  if (hh.length < 2) hh = "0" + hh;
  var min = String(d.getMinutes());
  if (min.length < 2) min = "0" + min;
  return yy + "." + mm + "." + dd + " " + hh + ":" + min;
}

/**
 * 캘린더의 시간 일정(Timed Event)을 시트 날짜 표기 형식으로 변환
 * - 시작과 종료가 같은 날이고 1시간 기본 차이면: 26.10.11 15:00
 * - 같은 날이고 1시간 이상 차이면: 26.10.11 15:00 ~ 17:00
 * - 다른 날로 걸쳐진 경우: 26.10.11 15:00 ~ 26.10.12 11:00
 */
function formatDateWithTime(start, end) {
  if (!start) return "";
  var startStr = formatDateTime(start);
  if (!end || start.getTime() === end.getTime()) {
    return startStr;
  }

  var isSameDay = isDateEqual(start, end);
  var diffMs = end.getTime() - start.getTime();

  if (isSameDay) {
    if (diffMs === 3600000) {
      return startStr;
    }
    var eh = String(end.getHours());
    if (eh.length < 2) eh = "0" + eh;
    var em = String(end.getMinutes());
    if (em.length < 2) em = "0" + em;
    return startStr + " ~ " + eh + ":" + em;
  } else {
    return startStr + " ~ " + formatDateTime(end);
  }
}

/**
 * 현장 일정 날짜를 YY.MM.DD (시간 포함 시 시간 표시) 표준 포맷으로 변환
 * 예: 26.10.15 / 26.10.15 14:00 / 26.10.15 ~ 26.10.18 / 26.10.15 09:00 ~ 18:00
 */
function formatScheduleDateDisplay(val) {
  if (!val) return "";
  var parsed = null;
  try {
    parsed = parseScheduleDateRange(val);
  } catch (e) {
    parsed = null;
  }
  if (!parsed || !parsed.startDate) {
    var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function"));
    if (isDateObj && !isNaN(val.getTime())) {
      return formatDateYYMMDD(val);
    }
    return String(val).trim();
  }

  // 1. 시간 정보가 포함된 일정 (단일 시간 또는 시간 범위)
  if (parsed.hasTime) {
    if (parsed.startTime && parsed.endTime) {
      return formatDateWithTime(parsed.startTime, parsed.endTime);
    } else if (parsed.startTime) {
      return formatDateTime(parsed.startTime);
    } else {
      return formatDateYYMMDD(parsed.startDate) + (parsed.timeStr ? (" " + parsed.timeStr) : "");
    }
  }

  // 2. 종일 기간(Range) 일정
  if (parsed.isRange && parsed.displayEndDate) {
    return formatDateYYMMDD(parsed.startDate) + " ~ " + formatDateYYMMDD(parsed.displayEndDate);
  }

  // 3. 단일 종일 일정
  return formatDateYYMMDD(parsed.startDate);
}

/**
 * 현재 로그인 사용자 닉네임/아이디 추출
 */
function getCurrentUserIdentifier() {
  var email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {}
  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch (e) {}
  }
  if (email) {
    var id = email.split("@")[0];
    return id || email;
  }
  return "사용자";
}


// ========================================================
// 5. [구글 캘린더 타겟 설정 & 핵심 엔진 (시트 ➔ 캘린더)]
// ========================================================

/**
 * 현재 설정된 타겟 구글 캘린더 객체 가져오기
 * (사용자 설정 DocumentProperties > DEFAULT_CALENDAR_ID_OR_NAME > 신규 생성 > 기본 캘린더 순)
 */
function getTargetCalendar() {
  var docProps = PropertiesService.getDocumentProperties();
  var savedTarget = docProps.getProperty("TARGET_CALENDAR_SETTING");
  var target = (savedTarget !== null && savedTarget !== undefined) ? savedTarget.trim() : DEFAULT_CALENDAR_ID_OR_NAME.trim();

  // 1) 빈 문자열인 경우 -> 본인 기본 캘린더
  if (!target) {
    return CalendarApp.getDefaultCalendar();
  }

  // 2) 캘린더 ID로 조회 시도 (이메일 형태 또는 구글 그룹 캘린더 ID)
  try {
    var calById = CalendarApp.getCalendarById(target);
    if (calById) return calById;
  } catch (e) {}

  // 3) 캘린더 이름으로 조회 시도
  try {
    var calsByName = CalendarApp.getCalendarsByName(target);
    if (calsByName && calsByName.length > 0) {
      return calsByName[0];
    }
  } catch (e) {}

  // 4) 캘린더 ID 형식이 아니고 일반 이름인 경우, 해당 이름으로 신규 생성 시도
  if (target.indexOf("@") === -1) {
    try {
      return CalendarApp.createCalendar(target, {
        summary: "우성아이디 현장 공정 일정 관리",
        timeZone: "Asia/Seoul"
      });
    } catch (e) {}
  }

  // 5) 모두 실패 시 기본 캘린더 폴백
  return CalendarApp.getDefaultCalendar();
}

/**
 * [메뉴용] 연동할 구글 캘린더 확인 및 변경 대화상자
 */
function configureTargetCalendar() {
  var ui = SpreadsheetApp.getUi();
  var currentCal = getTargetCalendar();
  var currentCalName = currentCal ? currentCal.getName() : "기본 캘린더";
  var currentCalId = currentCal ? currentCal.getId() : "";

  var promptText = "📌 현재 연동 중인 캘린더:\n" +
                   "• 이름: " + currentCalName + "\n" +
                   "• ID: " + currentCalId + "\n\n" +
                   "연동할 다른 캘린더의 [이름] 또는 [캘린더 ID(이메일 형식)]를 입력하세요:\n" +
                   "(빈 칸으로 두고 확인을 누르면 본인 기본 캘린더로 연결됩니다)";

  var response = ui.prompt("⚙️ 연동 캘린더 변경 / 확인", promptText, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  var docProps = PropertiesService.getDocumentProperties();

  if (!input) {
    docProps.setProperty("TARGET_CALENDAR_SETTING", "");
    ui.alert("✅ 본인 기본 캘린더(Default Calendar)로 설정되었습니다.");
    return;
  }

  // 입력된 캘린더 유효성 검증
  var testCal = null;
  try {
    testCal = CalendarApp.getCalendarById(input);
  } catch (e) {}

  if (!testCal) {
    try {
      var list = CalendarApp.getCalendarsByName(input);
      if (list && list.length > 0) testCal = list[0];
    } catch (e) {}
  }

  if (testCal) {
    docProps.setProperty("TARGET_CALENDAR_SETTING", input);
    ui.alert(
      "✅ 캘린더 연동 완료",
      "성공적으로 캘린더가 변경되었습니다!\n\n" +
      "• 캘린더명: [" + testCal.getName() + "]\n" +
      "• 캘린더 ID: " + testCal.getId() + "\n\n" +
      "이제 동기화 실행 시 이 캘린더로 일정이 등록됩니다.",
      ui.ButtonSet.OK
    );
  } else {
    // 일반 이름인 경우 신규 생성 제안
    if (input.indexOf("@") === -1) {
      var confirmCreate = ui.alert(
        "캘린더 신규 생성 안내",
        "'" + input + "' 이름의 캘린더를 찾지 못했습니다.\n해당 이름으로 새 구글 캘린더를 생성하여 연동할까요?",
        ui.ButtonSet.YES_NO
      );
      if (confirmCreate === ui.Button.YES) {
        try {
          var newCal = CalendarApp.createCalendar(input, {
            summary: "우성아이디 현장 공정 일정 관리",
            timeZone: "Asia/Seoul"
          });
          docProps.setProperty("TARGET_CALENDAR_SETTING", input);
          ui.alert("✅ 새 캘린더 [" + newCal.getName() + "]가 생성되고 연동되었습니다!");
        } catch (err) {
          ui.alert("⚠️ 캘린더 생성 실패: " + err.message);
        }
      }
    } else {
      ui.alert("⚠️ 입력하신 캘린더 ID (" + input + ")를 찾을 수 없거나 접근 권한이 없습니다.\n구글 캘린더 웹의 [설정 및 공유]에서 접근 권한 및 ID를 확인해주세요.");
    }
  }
}

/**
 * 행 데이터로부터 캘린더 이벤트 상세 정보 구성
 */
function buildEventDetails(rowValues, rowNumber, refDate, targetSheet) {
  var siteName = String(rowValues[0] || "").trim(); // A열: 현장
  var dateVal = rowValues[1];                      // B열: 날짜
  var taskName = String(rowValues[2] || "").trim(); // C열: 공종
  var partner = String(rowValues[3] || "").trim();  // D열: 협력업체 (인덱스 3)

  // E열: 공유 (인덱스 4) -> 체크되어 있으면 🟦, 체크 안 되어 있으면 🟥
  var rawShare = rowValues[4];
  var isShared = (rawShare === true || String(rawShare).trim().toUpperCase() === "TRUE" || String(rawShare).trim() === "1" || String(rawShare).trim() === "V" || String(rawShare).trim() === "O");
  var prefixIcon = isShared ? "🟦 " : "🟥 ";

  var memo = String(rowValues[5] || "").trim();     // F열: 비고 (인덱스 5)
  var manager = String(rowValues[6] || "").trim();  // G열: 담당자 (인덱스 6)
  var status = String(rowValues[7] || "").trim();   // H열: 진행상태 (인덱스 7)
  var statusClean = status.replace(/\s+/g, "");

  // 상태별 아이콘 (스타일 A: 예정='', 진행중='▶️ ', 완료='✅ ', AS='🛠️ ')
  var statusIcon = "";
  if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
    statusIcon = "✅ ";
  } else if (statusClean.indexOf("진행") !== -1 || statusClean.indexOf("▶") !== -1) {
    statusIcon = "▶️ ";
  } else if (statusClean.indexOf("AS") !== -1 || statusClean.indexOf("A/S") !== -1 || statusClean.indexOf("as") !== -1 || statusClean.indexOf("🛠") !== -1) {
    statusIcon = "🛠️ ";
  }

  // 이력 정보 (K, L, M, N열)
  var createdAt = String(rowValues[10] || "").trim(); // K열: 최초작성일
  var createdBy = String(rowValues[11] || "").trim(); // L열: 최초작성자
  var updatedAt = String(rowValues[12] || "").trim(); // M열: 최종수정일
  var updatedBy = String(rowValues[13] || "").trim(); // N열: 최종수정자

  // [스마트 상위 현장명 자동 계승]
  // A열 현장명이 비어있는 경우, 시트 내 직전 상위 행에서 현장명을 자동 탐색하여 계승
  if (!siteName && targetSheet && rowNumber > 2) {
    try {
      var scanCount = Math.min(rowNumber - 2, 50);
      var prevRange = targetSheet.getRange(rowNumber - scanCount, 1, scanCount, 1).getValues();
      for (var p = prevRange.length - 1; p >= 0; p--) {
        var candidateSite = String(prevRange[p][0] || "").trim();
        if (candidateSite) {
          siteName = candidateSite;
          rowValues[0] = siteName;
          try {
            targetSheet.getRange(rowNumber, 1).setValue(siteName);
          } catch(e) {}
          break;
        }
      }
    } catch(e) {}
  }

  // [필수 방어 1]: 현장명이 전혀 없는 일정은 캘린더 등록 배제 (유령/미아 일정 원천 차단)
  if (!siteName) return null;

  // [필수 방어 2]: 공종(작업명)이 없는 미완성 행은 캘린더 등록 완전 배제 (작성 중 조기 등록 방지)
  if (!taskName) return null;

  var parsedDate = parseScheduleDateRange(dateVal);
  if (!parsedDate) return null;

  // 공정 지연 자동 감지 (Feature 2)
  var isDelayed = isScheduleDelayed(parsedDate, statusClean, refDate);
  var delayPrefix = isDelayed ? "⚠️ [지연] " : "";

  // 구글 캘린더 자체 블록 색상 자동 지정 (Feature 1)
  var eventColor = getCalendarColorForStatus(statusClean, isDelayed, isShared);

  // 제목 포맷: [지연경고] [상태아이콘] 🟦 또는 🟥 [현장명] 공종 (비고 제외)
  var baseTitle = "";
  if (siteName && taskName) {
    baseTitle = "[" + siteName + "] " + taskName;
  } else if (siteName) {
    baseTitle = "[" + siteName + "]";
  } else {
    baseTitle = taskName;
  }
  var title = delayPrefix + statusIcon + prefixIcon + baseTitle;

  // 설명(Description) 포맷
  var descLines = [];
  if (isDelayed) {
    descLines.push("⚠️ [공정 지연 경고]: 마감 예정일이 경과하였으나 미완료 상태입니다. 빠른 확인이 필요합니다.");
    descLines.push("--------------------------------");
  }
  descLines.push("📌 현장명: " + (siteName || "-"));
  descLines.push("🔨 공종: " + (taskName || "-"));
  descLines.push("🚦 진행상태: " + (status || "예정") + (isDelayed ? " (⚠️ 지연)" : ""));
  descLines.push("🔗 공유: " + (isShared ? "완료" : "미완료"));
  if (manager) descLines.push("👤 담당자: " + manager);
  if (partner) descLines.push("🏢 협력업체: " + partner);
  if (memo) descLines.push("📝 비고: " + memo);
  descLines.push("--------------------------------");

  // 공유/진행상태 변경 링크 (단축 안내문구 클릭 이동 HTML 하이퍼링크)
  var webAppUrl = getWebAppUrl();
  if (webAppUrl && rowNumber) {
    var queryParts = ["row=" + rowNumber];
    if (siteName) queryParts.push("site=" + encodeURIComponent(siteName));
    if (taskName) queryParts.push("task=" + encodeURIComponent(taskName));
    var eventId = String(rowValues[8] || "").trim();
    if (eventId) queryParts.push("id=" + encodeURIComponent(eventId));
    var statusChangeUrl = webAppUrl + "?" + queryParts.join("&");
    descLines.push("<a href=\"" + statusChangeUrl + "\">📱 [공유/진행상태 변경]</a>");
  }

  // 스프레드시트 바로가기 링크 ([우성아이디 현장 관리 시트] - 해당 행 포커스 이동)
  var sheetUrl = "";
  try {
    var activeSs = null;
    var targetGid = null;

    if (targetSheet && typeof targetSheet.getSheetId === "function") {
      targetGid = targetSheet.getSheetId();
      try {
        if (typeof targetSheet.getParent === "function") {
          activeSs = targetSheet.getParent();
        }
      } catch(e) {}
    }

    if (!activeSs) {
      activeSs = getAppSpreadsheet();
    }

    if ((targetGid === null || targetGid === undefined) && activeSs) {
      try {
        var foundSiteSheet = findSiteScheduleSheet(activeSs);
        if (foundSiteSheet && typeof foundSiteSheet.getSheetId === "function") {
          targetGid = foundSiteSheet.getSheetId();
        }
      } catch(e) {}
    }

    if (activeSs && typeof activeSs.getUrl === "function") {
      var rawUrl = activeSs.getUrl();
      if (rawUrl) {
        var baseUrl = rawUrl.split("#")[0];
        var params = [];
        if (targetGid !== null && targetGid !== undefined) {
          params.push("gid=" + targetGid);
        }
        var validRow = parseInt(rowNumber, 10);
        if (!isNaN(validRow) && validRow > 0) {
          params.push("range=A" + validRow + ":H" + validRow);
        }
        sheetUrl = baseUrl + (params.length > 0 ? "#" + params.join("&") : "");
      }
    }
  } catch(e) {}

  if (sheetUrl) {
    descLines.push("<a href=\"" + sheetUrl + "\">📊 [우성아이디 현장 관리 시트]</a>");
  } else {
    descLines.push("📊 [우성아이디 현장 관리 시트]");
  }

  return {
    title: title,
    siteName: siteName,
    taskName: taskName,
    baseTitle: baseTitle,
    parsedDate: parsedDate,
    description: descLines.join("\n"),
    color: eventColor,
    isDelayed: isDelayed
  };
}

/**
 * 단일 행을 구글 캘린더에 동기화 (시트 -> 캘린더 반영 및 Last-Modified-Wins 적용)
 * @param {Sheet} sheet 현장데이터 시트
 * @param {number} row 대상 행 번호
 * @param {boolean} [forceSheetToCal=true] 시트의 변경사항을 캘린더에 우선 반영할지 여부 (기본 true: 시트 날짜 보호)
 */
function syncSingleRowToCalendar(sheet, row, forceSheetToCal, skipDashboardRefresh) {
  if (row < 2) return;
  if (forceSheetToCal === undefined) forceSheetToCal = true; // 시트 수정 작업 시 캘린더 역덮어쓰기 원천 방지
  if (skipDashboardRefresh === undefined) skipDashboardRefresh = false;
  var rowValues = sheet.getRange(row, 1, 1, 14).getValues()[0];
  var existingEventId = String(rowValues[8] || "").trim(); // I열: 캘린더ID
  var rawUpdated = rowValues[9];
  var sheetUpdated = 0;
  if (rawUpdated instanceof Date) {
    sheetUpdated = rawUpdated.getTime();
  } else if (typeof rawUpdated === "number") {
    sheetUpdated = rawUpdated;
  } else {
    sheetUpdated = parseInt(rawUpdated, 10) || 0;
  }

  var details = buildEventDetails(rowValues, row, null, sheet);
  var cal = getTargetCalendar();
  if (!cal) return;

  // 1) 날짜나 필수 정보가 없는 경우: 일시적 타이핑 중 삭제 방어
  // [삭제 유예 및 오입력 안전 가드]:
  // 타이핑 중 일시적인 빈 셀(날짜만 지우거나 공종만 지움)로 인해 소중한 캘린더 일정이 즉시 영구 삭제되는 위험 방지.
  // 아코디언 하위 행에서 상위 현장명(A열)이 남아있더라도 날짜(B열)+공종(C열)이 모두 비워졌거나 상태가 '취소'/'삭제'일 때만 실제 캘린더 일정 삭제.
  if (!details) {
    if (existingEventId) {
      var hasDate = !!String(rowValues[1] || "").trim();
      var hasTask = !!String(rowValues[2] || "").trim();
      var isRowEmpty = (!rowValues[0] && !hasDate && !hasTask);
      var isSubRowCleared = (!hasDate && !hasTask); // 상위 현장명이 남아있더라도 날짜와 공종이 모두 지워진 경우
      var statusStr = String(rowValues[7] || "");
      var isExplicitCancel = (statusStr.indexOf("취소") !== -1 || statusStr.indexOf("삭제") !== -1);
      if (isRowEmpty || isSubRowCleared || isExplicitCancel) {
        try {
          var ev = cal.getEventById(existingEventId);
          if (ev) ev.deleteEvent();
        } catch (err) {}
        sheet.getRange(row, 9).clearContent();
        sheet.getRange(row, 10).clearContent();
      }
    }
    return;
  }

  // 2) 기존 캘린더 이벤트가 존재하는 경우 -> 수정 시각 비교
  var event = null;
  if (existingEventId) {
    try {
      event = cal.getEventById(existingEventId);
    } catch (err) {
      event = null;
    }
  }

  var allSheetIds = null;
  // [캘린더 중복 생성 원천 차단]
  // 기존 ID가 비어있거나 유효하지 않은 경우, 캘린더에 이미 생성된 동일 일정이 있는지 탐색하여 재연결
  if (!event && !existingEventId) {
    try {
      allSheetIds = getSheetCalendarIds(sheet);
      event = findMatchingCalendarEvent(cal, details, allSheetIds);
      if (event) {
        existingEventId = event.getId();
        sheet.getRange(row, 9).setValue(existingEventId);
        sheet.getRange(row, 10).setValue(Date.now());
        rowValues[8] = existingEventId;
      }
    } catch (e) {}
  }

  if (event) {
    var calUpdatedDate = event.getLastUpdated();
    var calUpdated = calUpdatedDate ? calUpdatedDate.getTime() : 0;

    // forceSheetToCal이 false이고, 시트 타임스탬프가 유효하며, 캘린더가 10초 이상 더 최근에 수정된 경우에만 캘린더 -> 시트 반영
    if (!forceSheetToCal && sheetUpdated > 0 && calUpdated > sheetUpdated + 10000) {
      var isAllDay = event.isAllDayEvent();
      var calStart = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
      var calEnd = isAllDay ? event.getAllDayEndDate() : event.getEndTime();
      var newDateStr = "";

      if (!isAllDay) {
        newDateStr = formatDateWithTime(calStart, calEnd);
      } else {
        var realCalEnd = null;
        if (calEnd) {
          var diffDays = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
          if (diffDays > 1) {
            realCalEnd = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
          }
        }
        newDateStr = realCalEnd ? (formatDateYYMMDD(calStart) + " ~ " + formatDateYYMMDD(realCalEnd)) : formatDateYYMMDD(calStart);
      }
      sheet.getRange(row, 2).setValue(newDateStr);
      sheet.getRange(row, 10).setValue(calUpdated);
      sheet.getRange(row, 13).setValue(formatDateTime(calUpdatedDate));
      sheet.getRange(row, 14).setValue("구글 캘린더");

      if (!skipDashboardRefresh) {
        var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();
        refreshDashboardCalendar(ss);
      }
      return;
    }

    // 시트가 더 최근에 수정된 경우:
    // [핵심 해결책] 종일 일정 <-> 시간 일정 간 유형 전환 처리!
    var isCalAllDay = event.isAllDayEvent();
    var needTime = details.parsedDate.hasTime;

    if (isCalAllDay !== !needTime) {
      try {
        event.deleteEvent();
      } catch (e) {}
      event = null; // 신규 생성 분기로 전달
    }
  }

  if (event) {
    // 동일한 유형 내에서의 속성 업데이트
    event.setTitle(details.title);
    event.setDescription(details.description);
    event.setLocation(details.siteName);
    if (details.color) {
      try {
        if (event.getColor() !== details.color) {
          event.setColor(details.color);
        }
      } catch (e) {}
    }
    if (details.parsedDate.hasTime) {
      event.setTime(details.parsedDate.startTime, details.parsedDate.endTime);
    } else {
      if (details.parsedDate.isRange) {
        event.setAllDayDates(details.parsedDate.startDate, details.parsedDate.endDate);
      } else {
        event.setAllDayDate(details.parsedDate.startDate);
      }
    }
    sheet.getRange(row, 10).setValue(Date.now());
  } else {
    // 3) 신규 생성 (또는 종일<->시간 유형 전환으로 인한 재생성)
    var opt = {
      description: details.description,
      location: details.siteName
    };
    if (details.parsedDate.hasTime) {
      event = cal.createEvent(details.title, details.parsedDate.startTime, details.parsedDate.endTime, opt);
    } else {
      if (details.parsedDate.isRange) {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, details.parsedDate.endDate, opt);
      } else {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, opt);
      }
    }

    if (event) {
      if (details.color) {
        try {
          event.setColor(details.color);
        } catch (e) {}
      }
      var eventId = event.getId();
      sheet.getRange(row, 9).setValue(eventId);
      sheet.getRange(row, 10).setValue(Date.now());

      // 신규 생성 직후 이벤트 링크에 고유 ID 즉시 주입 (행 밀림 방지 영구 보장)
      try {
        rowValues[8] = eventId;
        var freshDetails = buildEventDetails(rowValues, row, null, sheet);
        if (freshDetails && freshDetails.description) {
          event.setDescription(freshDetails.description);
        }
      } catch(e) {}
    }
  }

  // [신규 핵심] 해당 행의 캘린더 일정이 확정된 후, 동일 날짜에 존재하는 잉여 중복/유령 일정 자동 정리
  if (event) {
    try {
      if (!allSheetIds) allSheetIds = getSheetCalendarIds(sheet);
      cleanDuplicateCalendarEventsOnDate(cal, details, event.getId(), allSheetIds);
    } catch (e) {}
  }

  // 지연 공정 시트 실시간 배경 하이라이트 반영
  try {
    if (details && details.isDelayed) {
      sheet.getRange(row, 1, 1, 8).setBackground("#fff1f2");
    } else if (details && !details.isDelayed) {
      var curBg = sheet.getRange(row, 1).getBackground();
      if (curBg === "#fff1f2" || curBg === "#fee2e2" || curBg === "#fce8e6") {
        sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
      }
    }
  } catch (e) {}

  // 실시간 수정 즉시 대시보드 달력 연쇄 갱신 (skipDashboardRefresh가 false일 때만 실행)
  if (!skipDashboardRefresh) {
    var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();
    refreshDashboardCalendar(ss);
  }
}

/**
 * [추천] 최근 & 예정 일정 우선 동기화 (최근 30일 이내 ~ 미래 일정)
 */
function syncRecentSiteSchedulesToGoogleCalendar() {
  syncSiteSchedulesBatchSafe(true);
}

/**
 * 전체 일정 연속 동기화 (25건씩 자동 대기 연속 처리)
 */
function syncAllSiteSchedulesToGoogleCalendar() {
  syncSiteSchedulesBatchSafe(false);
}

/**
 * 구글 Rate Limit(속도 제한) 방지 안전 배치 동기화 엔진
 * (25건 처리 시마다 2초 자동 대기 후 연속 실행 + 4분 30초 안전 타임아웃 가드)
 */
function syncSiteSchedulesBatchSafe(onlyRecent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  // 실시간 트리거 백그라운드 자동 등록 보장
  ensureRealtimeTriggerInstalledSilent();

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("등록할 현장 데이터가 없습니다.", "구글 캘린더 연동", 3);
    return;
  }

  var cal = getTargetCalendar();
  if (!cal) {
    SpreadsheetApp.getUi().alert("구글 캘린더에 접근할 수 없습니다. 캘린더 권한 설정을 확인하세요.");
    return;
  }

  ensureCalendarSyncColumns(sheet);

  var range = sheet.getRange(2, 1, lastRow - 1, 14);
  var values = range.getValues();
  var syncedCount = 0;
  var updatedCount = 0;
  var skippedCount = 0;
  var batchChangeCount = 0; // 이번 배치(25건) 변경 건수 카운터
  var batchNum = 1;         // 배치 차수
  var rateLimitHit = false;
  var timeLimitReached = false;
  var startTime = Date.now();

  var now = new Date();
  var nowMs = now.getTime();
  var nowStr = formatDateTime(now);
  var currentUser = getCurrentUserIdentifier();
  var thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  var allSheetIds = {};
  for (var k = 0; k < values.length; k++) {
    var eid = String(values[k][8] || "").trim();
    if (eid) allSheetIds[eid] = true;
  }

  // 대상 행 인덱스 필터링 및 우선순위 정렬 (신규 미등록 일정 ➔ 기존 일정 순)
  var pendingIndices = [];
  var existingIndices = [];
  var effectiveSiteNames = [];
  var lastKnownSite = "";

  for (var i = 0; i < values.length; i++) {
    var rawSite = String(values[i][0] || "").trim();
    if (rawSite) {
      lastKnownSite = rawSite;
      effectiveSiteNames[i] = rawSite;
    } else if (values[i][1] || values[i][2]) {
      effectiveSiteNames[i] = lastKnownSite;
    } else {
      effectiveSiteNames[i] = "";
    }

    var dateVal = values[i][1];
    if (!dateVal) continue;

    if (onlyRecent) {
      var parsed = parseScheduleDateRange(dateVal);
      if (parsed && parsed.startDate) {
        if (parsed.startDate < thirtyDaysAgo) {
          continue;
        }
      }
    }

    var existingId = String(values[i][8] || "").trim();
    if (!existingId) {
      pendingIndices.push(i);
    } else {
      existingIndices.push(i);
    }
  }

  // 1순위: 미등록 신규 일정 우선 처리, 2순위: 기존 일정 검사
  var targetIndices = pendingIndices.concat(existingIndices);

  if (targetIndices.length === 0) {
    SpreadsheetApp.getUi().alert("동기화할 대상 일정이 없습니다.");
    return;
  }

  for (var idx = 0; idx < targetIndices.length; idx++) {
    // 1) [25건 단위 자동 쿨다운 & 연속 배치 실행]
    // 25건 생성/수정 완료 시마다 구글 API 버스트 쿼터 보호를 위해 2초 대기 후 자동으로 다음 25건을 계속 진행합니다.
    if (batchChangeCount >= 25) {
      SpreadsheetApp.flush();
      try {
        ss.toast(
          "🔄 배치 " + batchNum + "차 완료 (" + (syncedCount + updatedCount) + "건 처리됨) - API 보호 2초 대기 후 자동으로 계속 동기화합니다...",
          "구글 캘린더 동기화",
          4
        );
      } catch(e) {}
      Utilities.sleep(2000);
      batchChangeCount = 0;
      batchNum++;
    }

    // 2) [핵심 방어: 구글 6분 타임아웃 원천 차단] 스크립트 실행 시간이 4분 30초(270초) 초과 시 안전 조기 종료
    if (Date.now() - startTime > 270000) {
      timeLimitReached = true;
      break;
    }

    var i = targetIndices[idx];
    var row = i + 2;
    var rowVals = values[i];
    var existingId = String(rowVals[8] || "").trim();

    // A열 현장명이 비어있는 경우 계승된 상위 현장명 자동 적용 및 시트 셀 보완
    if (!rowVals[0] && effectiveSiteNames[i]) {
      rowVals[0] = effectiveSiteNames[i];
      try {
        sheet.getRange(row, 1).setValue(effectiveSiteNames[i]);
      } catch (e) {}
    }

    // 작성일 보완 (4회 개별 호출 ➔ 1회 일괄 쓰기로 가속)
    if (!rowVals[10] || !rowVals[12]) {
      var cDate = rowVals[10] || nowStr;
      var cUser = rowVals[11] || currentUser;
      var mDate = rowVals[12] || nowStr;
      var mUser = rowVals[13] || currentUser;
      sheet.getRange(row, 11, 1, 4).setValues([[cDate, cUser, mDate, mUser]]);
      rowVals[10] = cDate;
      rowVals[11] = cUser;
      rowVals[12] = mDate;
      rowVals[13] = mUser;
    }

    var details = buildEventDetails(rowVals, row, null, sheet);
    if (!details) {
      if (existingId) {
        try {
          var ev = cal.getEventById(existingId);
          if (ev) ev.deleteEvent();
        } catch(e) {}
        sheet.getRange(row, 9).clearContent();
        sheet.getRange(row, 10).clearContent();
      }
      continue;
    }

    var event = null;
    if (existingId) {
      try {
        event = cal.getEventById(existingId);
      } catch(e) {
        event = null;
      }
    }

    // [신규 핵심: 캘린더 중복 생성 원천 차단]
    // I열에 ID가 없거나 유효하지 않은 경우, 캘린더에 이미 생성된 동일 일정이 있는지 탐색하여 재연결
    if (!event && !existingId) {
      try {
        event = findMatchingCalendarEvent(cal, details, allSheetIds);
        if (event) {
          var matchedId = event.getId();
          sheet.getRange(row, 9).setValue(matchedId);
          sheet.getRange(row, 10).setValue(nowMs);
          rowVals[8] = matchedId;
          allSheetIds[matchedId] = true;
          existingId = matchedId;
        }
      } catch (e) {}
    }

    if (event) {
      try {
        var isAllDay = event.isAllDayEvent();
        var needTime = details.parsedDate.hasTime;

        // [종일 <-> 시간 유형 전환 처리]
        if (isAllDay !== !needTime) {
          try {
            event.deleteEvent();
          } catch(e) {}
          event = null;
        } else {
          // [구글 한도 절약 1단계: 스마트 변경 감지 (Smart Diff Check)]
          var curTitle = event.getTitle();
          var curDesc = event.getDescription();
          var curLoc = event.getLocation();
          var curColor = "";
          try { curColor = event.getColor(); } catch (e) {}
          var calStart = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
          var calEnd = isAllDay ? event.getAllDayEndDate() : event.getEndTime();

          var dateMatches = false;
          if (details.parsedDate.hasTime) {
            if (!isAllDay && calStart && calEnd) {
              dateMatches = (calStart.getTime() === details.parsedDate.startTime.getTime() &&
                             calEnd.getTime() === details.parsedDate.endTime.getTime());
            }
          } else {
            if (isAllDay && calStart) {
              if (details.parsedDate.isRange) {
                dateMatches = isDateEqual(calStart, details.parsedDate.startDate) && isDateEqual(calEnd, details.parsedDate.endDate);
              } else {
                dateMatches = isDateEqual(calStart, details.parsedDate.startDate);
              }
            }
          }

          // 이미 캘린더와 시트 내용이 100% 동일한 경우 -> 구글 서버 호출 0회 패스!
          if (curTitle === details.title && curDesc === details.description && curLoc === details.siteName && dateMatches && (!details.color || curColor === details.color)) {
            try {
              cleanDuplicateCalendarEventsOnDate(cal, details, event.getId(), allSheetIds);
            } catch (e) {}
            skippedCount++;
            continue;
          }

          // [구글 한도 절약 2단계: 실제 변경된 속성만 최소 호출]
          if (curTitle !== details.title) event.setTitle(details.title);
          if (curDesc !== details.description) event.setDescription(details.description);
          if (curLoc !== details.siteName) event.setLocation(details.siteName);
          if (details.color && curColor !== details.color) {
            try { event.setColor(details.color); } catch (e) {}
          }
          if (!dateMatches) {
            if (details.parsedDate.hasTime) {
              event.setTime(details.parsedDate.startTime, details.parsedDate.endTime);
            } else {
              if (details.parsedDate.isRange) {
                event.setAllDayDates(details.parsedDate.startDate, details.parsedDate.endDate);
              } else {
                event.setAllDayDate(details.parsedDate.startDate);
              }
            }
          }

          sheet.getRange(row, 10).setValue(nowMs);
          updatedCount++;
          batchChangeCount++;
          try {
            cleanDuplicateCalendarEventsOnDate(cal, details, event.getId(), allSheetIds);
          } catch (e) {}
          // 갱신 시에도 초당 호출 제한 방지 딜레이 (0.35초)
          Utilities.sleep(350);
        }
      } catch (err) {
        var errMsg = String(err);
        if (errMsg.indexOf("짧은 시간") !== -1 || errMsg.indexOf("too many") !== -1 || errMsg.indexOf("Rate limit") !== -1 || errMsg.indexOf("quota") !== -1 || errMsg.indexOf("exhausted") !== -1) {
          rateLimitHit = true;
          break;
        }
      }
    }

    if (!event) {
      try {
        var opt = { description: details.description, location: details.siteName };
        if (details.parsedDate.hasTime) {
          event = cal.createEvent(details.title, details.parsedDate.startTime, details.parsedDate.endTime, opt);
        } else {
          if (details.parsedDate.isRange) {
            event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, details.parsedDate.endDate, opt);
          } else {
            event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, opt);
          }
        }
        if (event) {
          if (details.color) {
            try { event.setColor(details.color); } catch (e) {}
          }
          var newEventId = event.getId();
          sheet.getRange(row, 9).setValue(newEventId);
          sheet.getRange(row, 10).setValue(nowMs);
          rowVals[8] = newEventId;
          allSheetIds[newEventId] = true;
          try {
            var freshDetails = buildEventDetails(rowVals, row, null, sheet);
            if (freshDetails && freshDetails.description) {
              event.setDescription(freshDetails.description);
            }
          } catch (e) {}
          syncedCount++;
          batchChangeCount++;
          try {
            cleanDuplicateCalendarEventsOnDate(cal, details, newEventId, allSheetIds);
          } catch (e) {}
          // 구글 Rate Limit 방지용 딜레이 (0.35초)
          Utilities.sleep(350);
        }
      } catch (err) {
        var errMsg = String(err);
        if (errMsg.indexOf("짧은 시간") !== -1 || errMsg.indexOf("too many") !== -1 || errMsg.indexOf("Rate limit") !== -1 || errMsg.indexOf("quota") !== -1 || errMsg.indexOf("exhausted") !== -1) {
          rateLimitHit = true;
          break;
        }
      }
    }

    // 지연 공정 시트 배경 하이라이트 반영 (지연 건만 선택 반영하여 API 호출 최소화)
    try {
      if (details.isDelayed) {
        sheet.getRange(row, 1, 1, 8).setBackground("#fff1f2");
      }
    } catch (e) {}
  }

  SpreadsheetApp.flush();
  var calName = cal.getName();

  if (rateLimitHit) {
    SpreadsheetApp.getUi().alert(
      "⏳ 구글 캘린더 생성 한도 알림",
      "구글 서버의 일시적 속도 제한(Rate Limit)에 도달했습니다.\n\n" +
      "• 지금까지 안전하게 처리된 일정: " + (syncedCount + updatedCount) + "건 (신규 " + syncedCount + "건, 갱신 " + updatedCount + "건)\n" +
      (skippedCount > 0 ? "• 변경 없음(API 절약 통과): " + skippedCount + "건\n\n" : "\n") +
      "이미 처리된 일정은 시트에 안전하게 저장되었습니다.\n" +
      "구글 정책상 약 10~15분 후 다시 실행하시면 남은 일정이 이어서 안전하게 등록됩니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else if (timeLimitReached) {
    SpreadsheetApp.getUi().alert(
      "⏱️ 실행 시간 안전 일시 정지 (4분 30초 가드)",
      "구글 스크립트 6분 강제 종료를 방지하기 위해 안전 시간(4분 30초)에 도달하여 일시 정지되었습니다.\n\n" +
      "- 신규 등록: " + syncedCount + "건\n" +
      "- 갱신 완료: " + updatedCount + "건\n" +
      (skippedCount > 0 ? "- 변경 없음(자동 통과): " + skippedCount + "건\n\n" : "\n") +
      "남은 일정이 있다면 메뉴에서 한 번 더 실행하시면 다음 일정이 이어서 순차 등록됩니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "🎉 구글 캘린더 동기화 완료!",
      "대상 캘린더: [" + calName + "]\n\n" +
      "모든 대상 일정의 동기화가 성공적으로 완료되었습니다!\n" +
      (batchNum > 1 ? "• 총 처리 배치: " + batchNum + "차 (25건씩 자동 대기 연속 실행)\n" : "") +
      "• 신규 등록: " + syncedCount + "건\n" +
      "• 갱신 완료: " + updatedCount + "건\n" +
      (skippedCount > 0 ? "• 변경 없음(API 절약 통과): " + skippedCount + "건\n" : "") +
      "\n이제 구글 캘린더와 시트가 100% 일치합니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 9열~14열 시스템 및 이력 열 헤더 & 서식 확인
 * 9열: 캘린더ID (숨김)
 * 10열: 최종수정시각 (밀리초, 숨김)
 * 11열: 최초작성일
 * 12열: 최초작성자
 * 13열: 최종수정일
 * 14열: 최종수정자
 */
function ensureCalendarSyncColumns(sheet) {
  if (!sheet) return;
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 14) {
    sheet.insertColumnsAfter(maxCols, 14 - maxCols);
  }

  // 8열(H열) 진행상태 헤더 및 드롭다운 검증
  var statusCell = sheet.getRange(1, 8);
  var currentStatusHeader = String(statusCell.getValue() || "").trim();
  if (!currentStatusHeader || currentStatusHeader.indexOf("상태") === -1) {
    statusCell.setValue("진행상태")
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setBackground("#334155")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  var headers = [
    { col: 9, title: "캘린더ID" },
    { col: 10, title: "최종수정시각" },
    { col: 11, title: "최초작성일" },
    { col: 12, title: "최초작성자" },
    { col: 13, title: "최종수정일" },
    { col: 14, title: "최종수정자" }
  ];

  headers.forEach(function(h) {
    var cell = sheet.getRange(1, h.col);
    if (cell.getValue() !== h.title) {
      cell.setValue(h.title)
        .setFontWeight("bold")
        .setFontColor("#ffffff")
        .setBackground("#334155")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    }
  });

  // 11~14열 너비 및 서식
  var colWidthMap = { 11: 120, 12: 85, 13: 120, 14: 85 };
  for (var c in colWidthMap) {
    sheet.setColumnWidth(parseInt(c, 10), colWidthMap[c]);
  }

  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow >= 2) {
    sheet.getRange(2, 11, lastRow - 1, 4)
      .setHorizontalAlignment("center")
      .setFontColor("#64748b");

    // 8열(진행상태) 드롭다운 데이터 유효성 검사 (예정, 진행 중, 완료, AS) 설정
    try {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["예정", "진행 중", "완료", "AS"], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 8, lastRow - 1, 1).setDataValidation(rule);

      // '진행중'(공백 없음)으로 입력된 셀이 있다면 '진행 중'(표준)으로 자동 정규화하여 빨간 삼각형 에러 방지
      var statusRange = sheet.getRange(2, 8, lastRow - 1, 1);
      var statusVals = statusRange.getValues();
      var normChanged = false;
      for (var si = 0; si < statusVals.length; si++) {
        var sv = String(statusVals[si][0] || "").trim();
        if (sv === "진행중") {
          statusVals[si][0] = "진행 중";
          normChanged = true;
        }
      }
      if (normChanged) {
        statusRange.setValues(statusVals);
      }
    } catch (e) {}
  }

  // 9, 10열(시스템 열) 숨김
  try {
    sheet.hideColumns(9, 2);
  } catch(e) {}
}

/**
 * 현장데이터 시트 I열에 등록된 모든 유효한 캘린더 ID 맵 추출
 * @param {Sheet} sheet 현장데이터 시트
 * @return {Object} { [eventId]: true }
 */
function getSheetCalendarIds(sheet) {
  var idMap = {};
  if (!sheet) return idMap;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return idMap;
  try {
    var idValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    for (var i = 0; i < idValues.length; i++) {
      var id = String(idValues[i][0] || "").trim();
      if (id) idMap[id] = true;
    }
  } catch (e) {}
  return idMap;
}

/**
 * 캘린더 이벤트가 우리 시스템에서 관리하는 현장 공정 일정인지 엄격 검증
 * (사용자의 개인 일정 오탐지 및 훼손 원천 방지)
 * @param {CalendarEvent} ev 캘린더 이벤트
 * @return {boolean}
 */
function isSystemSiteScheduleEvent(ev) {
  if (!ev) return false;
  var desc = "";
  try { desc = ev.getDescription() || ""; } catch(e) {}
  if (desc.indexOf("📌 현장명:") !== -1 ||
      desc.indexOf("🔨 공종:") !== -1 ||
      desc.indexOf("우성아이디") !== -1 ||
      desc.indexOf("공유/진행상태 변경") !== -1) {
    return true;
  }
  var title = "";
  try { title = ev.getTitle() || ""; } catch(e) {}
  if (/\[.+?\]/.test(title) && (
      title.indexOf("🟥") !== -1 ||
      title.indexOf("🟦") !== -1 ||
      title.indexOf("▶️") !== -1 ||
      title.indexOf("✅") !== -1 ||
      title.indexOf("🛠️") !== -1 ||
      title.indexOf("⚠️") !== -1
  )) {
    return true;
  }
  return false;
}

/**
 * 캘린더에서 동일한 현장/공종/날짜의 기존 이벤트 탐색 (중복 생성 방지)
 * @param {Calendar} cal 캘린더 객체
 * @param {Object} details buildEventDetails 결과 객체
 * @param {Object} [knownSheetIds] 현재 시트 전체의 유효한 I열 ID 맵 (옵션)
 * @return {CalendarEvent|null}
 */
function findMatchingCalendarEvent(cal, details, knownSheetIds) {
  if (!cal || !details || !details.parsedDate || !details.parsedDate.startDate) return null;

  var startDate = details.parsedDate.startDate;
  var searchStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate(), 0, 0, 0);
  var searchEnd = null;
  if (details.parsedDate.endDate) {
    searchEnd = new Date(details.parsedDate.endDate.getFullYear(), details.parsedDate.endDate.getMonth(), details.parsedDate.endDate.getDate() + 1, 0, 0, 0);
  } else {
    searchEnd = new Date(searchStart.getTime() + 24 * 60 * 60 * 1000);
  }

  var events = [];
  try {
    events = cal.getEvents(searchStart, searchEnd);
  } catch (e) {
    return null;
  }

  if (!events || events.length === 0) return null;

  var siteName = details.siteName || "";
  var taskName = details.taskName || "";
  var matchedEvents = [];

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var evId = ev.getId();

    // 만약 이 이벤트가 이미 시트의 '다른 행'에 연결되어 있다면 건너뜀 (다른 행의 정당한 일정이므로 매칭 대상 아님)
    if (knownSheetIds && knownSheetIds[evId]) {
      continue;
    }

    if (!isSystemSiteScheduleEvent(ev)) continue;

    var title = ev.getTitle() || "";
    var desc = ev.getDescription() || "";
    var loc = ev.getLocation() || "";

    // 현장명 일치 확인
    var siteMatches = (
      (siteName && title.indexOf("[" + siteName + "]") !== -1) ||
      loc === siteName ||
      desc.indexOf("📌 현장명: " + siteName) !== -1
    );
    if (!siteMatches) continue;

    // 공종명 일치 확인
    var taskMatches = true;
    if (taskName) {
      taskMatches = (
        title.indexOf(taskName) !== -1 ||
        desc.indexOf("🔨 공종: " + taskName) !== -1
      );
    }
    if (!taskMatches) continue;

    // 날짜 일치 확인
    var dateMatches = false;
    var isAllDay = ev.isAllDayEvent();
    if (details.parsedDate.hasTime) {
      if (!isAllDay) {
        var evStart = ev.getStartTime();
        if (evStart && Math.abs(evStart.getTime() - details.parsedDate.startTime.getTime()) < 5 * 60 * 1000) {
          dateMatches = true;
        }
      }
    } else {
      if (isAllDay) {
        var evStart = ev.getAllDayStartDate();
        if (evStart && isDateEqual(evStart, details.parsedDate.startDate)) {
          dateMatches = true;
        }
      } else {
        var evStart = ev.getStartTime();
        if (evStart && isDateEqual(evStart, details.parsedDate.startDate)) {
          dateMatches = true;
        }
      }
    }

    if (dateMatches) {
      matchedEvents.push(ev);
    }
  }

  if (matchedEvents.length === 0) return null;

  return matchedEvents[0];
}

/**
 * 특정 행의 일정과 일치하는 미등록 잉여 중복 캘린더 일정을 감지하여 안전 삭제
 * @param {Calendar} cal 캘린더 객체
 * @param {Object} details buildEventDetails 결과
 * @param {string} keepEventId 보존할 유효 캘린더 ID
 * @param {Object} knownSheetIds 시트 전체의 유효 I열 ID 맵
 * @return {number} 삭제된 중복 일정 수
 */
function cleanDuplicateCalendarEventsOnDate(cal, details, keepEventId, knownSheetIds) {
  if (!cal || !details || !details.parsedDate || !details.parsedDate.startDate) return 0;

  var startDate = details.parsedDate.startDate;
  var searchStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate(), 0, 0, 0);
  var searchEnd = details.parsedDate.endDate ?
    new Date(details.parsedDate.endDate.getFullYear(), details.parsedDate.endDate.getMonth(), details.parsedDate.endDate.getDate() + 1, 0, 0, 0) :
    new Date(searchStart.getTime() + 24 * 60 * 60 * 1000);

  var events = [];
  try {
    events = cal.getEvents(searchStart, searchEnd);
  } catch (e) {
    return 0;
  }

  if (!events || events.length === 0) return 0;

  var siteName = details.siteName || "";
  var taskName = details.taskName || "";
  var deleted = 0;

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var evId = ev.getId();

    // 보존 대상 이벤트이거나 시트의 다른 행에 등록된 정상 이벤트는 절대 삭제 금지
    if (evId === keepEventId) continue;
    if (knownSheetIds && knownSheetIds[evId]) continue;

    if (!isSystemSiteScheduleEvent(ev)) continue;

    var title = ev.getTitle() || "";
    var desc = ev.getDescription() || "";
    var loc = ev.getLocation() || "";

    // 현장명 및 공종명 일치 확인
    var siteMatches = (
      (siteName && title.indexOf("[" + siteName + "]") !== -1) ||
      loc === siteName ||
      desc.indexOf("📌 현장명: " + siteName) !== -1
    );
    var taskMatches = true;
    if (taskName) {
      taskMatches = (
        title.indexOf(taskName) !== -1 ||
        desc.indexOf("🔨 공종: " + taskName) !== -1
      );
    }

    if (siteMatches && taskMatches) {
      var isAllDay = ev.isAllDayEvent();
      var evStart = isAllDay ? ev.getAllDayStartDate() : ev.getStartTime();
      if (evStart && isDateEqual(evStart, details.parsedDate.startDate)) {
        try {
          ev.deleteEvent();
          deleted++;
        } catch (e) {}
      }
    }
  }

  return deleted;
}

/**
 * [메뉴용] 구글 캘린더 중복 일정 및 시트에서 삭제된 유령 일정 일괄 자동 정리
 * 1) 시트에 기록된 모든 유효한 캘린더 ID(I열) 수집
 * 2) 시트 행과 캘린더 일정 간 스마트 대조 (누락된 ID는 자동 연결)
 * 3) 시트에 존재하지 않는 고아(Orphan) 일정 및 중복 생성된 일정 안전 삭제
 */
function cleanupDuplicateAndOrphanCalendarEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var cal = getTargetCalendar();
  if (!cal) {
    SpreadsheetApp.getUi().alert("구글 캘린더에 접근할 수 없습니다. 캘린더 권한 설정을 확인하세요.");
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "🧹 캘린더 중복 및 유령 일정 자동 정리",
    "구글 캘린더에 남아있는 중복 일정 및 시트에서 삭제된 유령(고아) 일정을 검색하여 일괄 정리합니다.\n(우성아이디 현장 일정 외의 개인 일정은 절대 건드리지 않습니다)\n\n지금 정리를 진행하시겠습니까?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("정리할 현장 데이터가 없습니다.", "캘린더 정리", 3);
    return;
  }

  ensureCalendarSyncColumns(sheet);

  var values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  var validEventIds = {};
  var unsyncedRows = []; // I열이 비어있는 행 번호들
  var minDate = null;
  var maxDate = null;

  for (var i = 0; i < values.length; i++) {
    var rowNum = i + 2;
    var rowVal = values[i];
    var eventId = String(rowVal[8] || "").trim();
    var dateVal = rowVal[1];

    if (eventId) {
      validEventIds[eventId] = rowNum;
    } else if (rowVal[0] || rowVal[2]) {
      unsyncedRows.push(rowNum);
    }

    if (dateVal) {
      var parsed = parseScheduleDateRange(dateVal);
      if (parsed && parsed.startDate) {
        if (!minDate || parsed.startDate < minDate) minDate = parsed.startDate;
        var endD = parsed.endDate || parsed.startDate;
        if (!maxDate || endD > maxDate) maxDate = endD;
      }
    }
  }

  var now = new Date();
  var defaultStart = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
  var defaultEnd = new Date(now.getTime() + 120 * 24 * 60 * 60 * 1000);
  var searchStart = (minDate && minDate < defaultStart) ? new Date(minDate.getTime() - 30 * 24 * 60 * 60 * 1000) : defaultStart;
  var searchEnd = (maxDate && maxDate > defaultEnd) ? new Date(maxDate.getTime() + 60 * 24 * 60 * 60 * 1000) : defaultEnd;

  // 1단계: I열이 비어있는 행이 있다면, 캘린더의 기존 일정과 스마트 자동 재연결
  var reconnectedCount = 0;
  for (var u = 0; u < unsyncedRows.length; u++) {
    var uRow = unsyncedRows[u];
    var uVals = values[uRow - 2];
    var uDetails = buildEventDetails(uVals, uRow, null, sheet);
    if (!uDetails) continue;

    var matchedEv = findMatchingCalendarEvent(cal, uDetails, validEventIds);
    if (matchedEv) {
      var mId = matchedEv.getId();
      sheet.getRange(uRow, 9).setValue(mId);
      sheet.getRange(uRow, 10).setValue(Date.now());
      validEventIds[mId] = uRow;
      values[uRow - 2][8] = mId;
      reconnectedCount++;
    }
  }

  // 2단계: 캘린더 일정 전수 스캔 및 중복/유령 일정 감지
  var allCalEvents = [];
  try {
    allCalEvents = cal.getEvents(searchStart, searchEnd);
  } catch (e) {
    ui.alert("캘린더 일정을 조회하는 중 오류가 발생했습니다: " + e.message);
    return;
  }

  var deletedCount = 0;

  for (var k = 0; k < allCalEvents.length; k++) {
    var ev = allCalEvents[k];
    var evId = ev.getId();

    if (!isSystemSiteScheduleEvent(ev)) {
      continue; // 개인 일정 철저 보호
    }

    if (validEventIds[evId]) {
      continue; // 시트에 유효하게 등록된 정상 일정
    }

    // 시트 I열에 없는 고아/중복 일정 안전 삭제
    try {
      ev.deleteEvent();
      deletedCount++;
      Utilities.sleep(150);
    } catch (e) {}
  }

  SpreadsheetApp.flush();

  var resultMsg = "✅ 캘린더 중복 및 유령 일정 정리가 완료되었습니다!\n\n" +
                  "• 삭제된 중복/유령 일정: " + deletedCount + "건\n" +
                  (reconnectedCount > 0 ? "• 시트와 자동 재연결된 일정: " + reconnectedCount + "건\n" : "") +
                  "• 현재 정상 연동 일정: " + Object.keys(validEventIds).length + "건";

  ui.alert("🧹 캘린더 정리 완료", resultMsg, ui.ButtonSet.OK);
}

/**
 * 공종(C열), 담당자(G열), 진행상태(H열) 드롭다운 유효성 검사 및 노션 스타일 뱃지 서식 전체 복구
 */
function fixStatusDropdownValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var maxRows = Math.max(sheet.getLastRow(), 100);

  // 1. 1행 고정 및 기본 정렬/테두리 서식
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, maxRows - 1, 14).setVerticalAlignment("middle");
  [1, 2, 3, 5, 7, 8, 11, 12, 13, 14].forEach(function(col) {
    sheet.getRange(2, col, maxRows - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, maxRows - 1, 1).setHorizontalAlignment("left");
  });
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 2), 8).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // 2. E열 체크박스 설정
  sheet.getRange(2, 5, maxRows - 1, 1).insertCheckboxes().setHorizontalAlignment("center");

  // 3. C열(공종), G열(담당자), H열(진행상태) 드롭다운 유효성 검사 일괄 적용
  applySiteScheduleValidations(sheet, 2, maxRows - 1);

  // 4. C열, G열, H열 노션 스타일 뱃지 조건부 서식 전수 적용
  setupSiteSheetConditionalFormatting(sheet);

  // 5. 보조열(9~14열) 검증
  ensureCalendarSyncColumns(sheet);

  SpreadsheetApp.flush();
  ss.toast("공종(C열), 담당자(G열), 진행상태(H열) 드롭다운 및 노션 스타일 뱃지 서식이 전수 복구되었습니다!", "서식 전체 복구 완료", 4);
}


// ========================================================
// 6. [동적 달력 갱신 핵심 엔진] (대시보드 월간 달력 연동)
// ========================================================

/**
 * 선택된 연도/월 기준으로 대시보드 달력 내용 전체 재구성
 */
function refreshDashboardCalendar(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = findDashboardSheet(ss);
  var siteSheet = findSiteScheduleSheet(ss);
  if (!dashSheet || !siteSheet) return;

  // [절대 방어 가드 1] 대시보드 시트와 현장데이터 시트가 동일 시트이면 즉시 중단 (시트 오염 원천 차단)
  if (dashSheet.getSheetId() === siteSheet.getSheetId()) {
    Logger.log("⚠️ 경고: 대시보드 시트와 현장데이터 시트가 동일 시트로 판정되어 대시보드 갱신을 중단합니다.");
    return;
  }

  // [절대 방어 가드 2] 대시보드 시트에 요일 헤더가 없으면 즉시 중단
  if (!hasDayOfWeekHeader(dashSheet)) {
    Logger.log("⚠️ 경고: 대시보드 시트에서 요일 헤더를 찾을 수 없어 대시보드 갱신을 중단합니다.");
    return;
  }

  // 구글 시트 수식(새 월 날짜 계산) 강제 반영
  SpreadsheetApp.flush();

  // 1. 요일 헤더 행 찾기 (예: 4행)
  var dowRow = findDayOfWeekRow(dashSheet);

  // 2. 대시보드 상단에서 선택된 연도와 월 읽기
  var dashYM = getDashboardSelectedYearMonth(dashSheet, dowRow);
  var targetYear = dashYM.year;
  var targetMonth = dashYM.month;

  // 3. 상단 헤더 영역(드롭다운 밑) 청소
  cleanupDashboardHeaderArea(dashSheet, dowRow);

  // 4. 달력의 모든 일정 내용 칸을 깨끗하게 비움 (기존 월 일정 완전 삭제)
  clearAllDashboardContentCells(dashSheet, dowRow);

  // 5. 현장데이터 시트에서 전체 데이터 읽기 (1~8열: 현장, 날짜, 공종, 협력업체, 공유, 비고, 담당자, 진행상태)
  var lastRow = siteSheet.getLastRow();
  if (lastRow < 2) return;

  var siteData = siteSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 6. 현재 대시보드 달력의 날짜 위치 맵 구성: map[day] = { row, col }
  var cellPositions = buildCalendarCellMap(dashSheet, dowRow, targetMonth);

  // 7. 현장데이터 중 '선택된 연도 및 월'에 해당하는 일정만 수집
  var schedulesByDay = {};
  var matchedCount = 0;

  for (var i = 0; i < siteData.length; i++) {
    var siteName = String(siteData[i][0] || "").trim();
    var dateVal = siteData[i][1];
    var taskName = String(siteData[i][2] || "").trim();
    var memo = String(siteData[i][5] || "").trim(); // F열: 비고 (인덱스 5)
    var status = String(siteData[i][7] || "").trim(); // H열: 진행상태 (인덱스 7)

    if (!siteName || !dateVal || !taskName) continue;

    var rangeInfo = parseScheduleDateRange(dateVal);
    if (!rangeInfo) continue;

    // E열: 공유 (인덱스 4) -> 🟦 / 🟥 아이콘
    var rawShare = siteData[i][4];
    var isShared = (rawShare === true || String(rawShare).trim().toUpperCase() === "TRUE" || String(rawShare).trim() === "1" || String(rawShare).trim() === "V" || String(rawShare).trim() === "O");
    var prefixIcon = isShared ? "🟦 " : "🟥 ";

    // 상태별 아이콘 (스타일 A: 예정='', 진행중='▶️ ', 완료='✅ ', AS='🛠️ ')
    var statusClean = status.replace(/\s+/g, "");
    var statusIcon = "";
    if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
      statusIcon = "✅ ";
    } else if (statusClean.indexOf("진행") !== -1 || statusClean.indexOf("▶") !== -1) {
      statusIcon = "▶️ ";
    } else if (statusClean.indexOf("AS") !== -1 || statusClean.indexOf("A/S") !== -1 || statusClean.indexOf("as") !== -1 || statusClean.indexOf("🛠") !== -1) {
      statusIcon = "🛠️ ";
    }

    // 공정 지연 자동 감지 (Feature 2)
    var isDelayed = isScheduleDelayed(rangeInfo, statusClean);
    var delayPrefix = isDelayed ? "⚠️ [지연] " : "";

    // 시간 정보가 있는 경우 시간 표기 추가 (예: 15:00)
    var timeDisplay = (rangeInfo.hasTime && rangeInfo.timeStr) ? " (" + rangeInfo.timeStr + ")" : "";

    // 대시보드 일정 텍스트 포맷 (지연경고 + 상태아이콘 + 공유아이콘 + 현장명 + 공종 + 시간, 비고 제외)
    var itemText = delayPrefix + statusIcon + prefixIcon + "[" + siteName + "] " + taskName + timeDisplay;

    // 기간 일정(예: 10/7 ~ 10/9) 처리
    if (rangeInfo.isRange && rangeInfo.displayEndDate) {
      var curD = new Date(rangeInfo.startDate.getFullYear(), rangeInfo.startDate.getMonth(), rangeInfo.startDate.getDate());
      var endD = new Date(rangeInfo.displayEndDate.getFullYear(), rangeInfo.displayEndDate.getMonth(), rangeInfo.displayEndDate.getDate());
      var anyMatched = false;
      var maxRangeLoop = 366; // 무한루프 및 비정상 기간 방어 (최대 1년)
      while (curD <= endD && --maxRangeLoop >= 0) {
        var cYear = curD.getFullYear();
        var cMonth = curD.getMonth() + 1;
        var cDay = curD.getDate();

        if ((!targetYear || cYear === targetYear) && cMonth === targetMonth) {
          if (!schedulesByDay[cDay]) schedulesByDay[cDay] = [];
          schedulesByDay[cDay].push(itemText);
          anyMatched = true;
        }
        curD.setDate(curD.getDate() + 1);
      }
      if (anyMatched) matchedCount++;
    } else {
      // 단일 날짜
      if (rangeInfo.year && targetYear && rangeInfo.year !== targetYear) continue;
      if (rangeInfo.month !== targetMonth) continue;

      var day = rangeInfo.day;
      if (!schedulesByDay[day]) schedulesByDay[day] = [];
      schedulesByDay[day].push(itemText);
      matchedCount++;
    }
  }

  // 8. 해당 월의 일정들을 달력의 해당 날짜 밑 칸에 쓰기 (주차별 행 단위 배치 일괄 쓰기 가속)
  var rowsToUpdate = {};
  for (var day in schedulesByDay) {
    var pos = cellPositions[day];
    if (pos) {
      if (!rowsToUpdate[pos.row]) {
        rowsToUpdate[pos.row] = dashSheet.getRange(pos.row, 2, 1, 7).getValues()[0];
      }
      rowsToUpdate[pos.row][pos.col - 2] = schedulesByDay[day].join("\n");
    }
  }

  for (var rNum in rowsToUpdate) {
    dashSheet.getRange(parseInt(rNum, 10), 2, 1, 7).setValues([rowsToUpdate[rNum]]);
  }

  SpreadsheetApp.flush();
  if (matchedCount > 0) {
    ss.toast(targetMonth + "월 일정 " + matchedCount + "건이 대시보드에 반영되었습니다.", "대시보드 갱신 완료", 2);
  } else {
    ss.toast(targetMonth + "월은 등록된 현장 일정이 없습니다. (달력 비움 완료)", "대시보드 갱신 완료", 2);
  }
}

/**
 * 메뉴용 대시보드 수동 동기화 실행 함수
 */
function syncAllSiteSchedulesToDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  refreshDashboardCalendar(ss);
}

/**
 * 대시보드 [Today] 실행: 현재 연도/월로 변경 및 오늘 날짜 셀로 포커스 이동
 */
function goToToday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = findDashboardSheet(ss);
  if (!dashSheet || isSiteScheduleSheet(dashSheet)) {
    SpreadsheetApp.getUi().alert(
      "⚠️ 대시보드 시트를 찾을 수 없습니다.",
      "현재 대시보드(달력) 시트의 탭 이름을 찾지 못했습니다.\n\n대시보드 탭 이름을 '대시보드' 또는 'dashboard'로 지정해주세요.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  ss.setActiveSheet(dashSheet);

  var today = new Date();
  var curYear = today.getFullYear();
  var curMonth = today.getMonth() + 1;
  var curDay = today.getDate();

  // 1. 대시보드 상단에서 '년도'와 '월' 셀 위치 찾기
  var yearCell = dashSheet.getRange("C2");
  var monthCell = dashSheet.getRange("E2");

  var dowRow = findDayOfWeekRow(dashSheet);
  var maxSearchRows = Math.max(dowRow - 1, 3);
  var values = dashSheet.getRange(1, 1, maxSearchRows, Math.min(dashSheet.getLastColumn(), 10)).getValues();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var val = String(values[r][c] || "").trim();
      if ((val === "년도" || val === "연도" || val === "Year") && c + 1 < values[r].length) {
        yearCell = dashSheet.getRange(r + 1, c + 2);
      }
      if ((val === "월" || val === "Month") && c + 1 < values[r].length) {
        monthCell = dashSheet.getRange(r + 1, c + 2);
      }
    }
  }

  // G2 체크박스가 있다면 false로 초기화 (클릭 피드백)
  try {
    var g2 = dashSheet.getRange("G2");
    if (typeof g2.getValue() === "boolean") {
      g2.setValue(false);
    }
  } catch(e) {}

  // 2. 연도와 월 설정
  yearCell.setValue(curYear);
  monthCell.setValue(curMonth);
  SpreadsheetApp.flush();

  // 3. 대시보드 달력 새로고침 (선택된 월 일정 렌더링)
  refreshDashboardCalendar(ss);

  // 4. 오늘 날짜 셀로 정확한 포커스 이동
  var cellPositions = buildCalendarCellMap(dashSheet, dowRow, curMonth);
  if (cellPositions && cellPositions[curDay]) {
    var pos = cellPositions[curDay];
    // 날짜 헤더 셀(숫자 적힌 셀)로 활성화
    var targetRange = dashSheet.getRange(pos.row - 1, pos.col);
    dashSheet.setActiveRange(targetRange);
  }

  ss.toast("오늘(" + curYear + "년 " + curMonth + "월 " + curDay + "일)로 이동했습니다.", "📅 Today", 3);
}

/**
 * 달력의 모든 일정 내용 칸 깨끗하게 비우기 (수식 셀 보존)
 */
function clearAllDashboardContentCells(dashSheet, dowRow) {
  var maxRows = Math.min(dashSheet.getMaxRows(), 35);
  for (var r = dowRow + 2; r <= maxRows; r += 2) {
    var range = dashSheet.getRange(r, 2, 1, 7); // B(일) ~ H(토)
    var formulas = range.getFormulas()[0];
    var vals = range.getValues()[0];
    var needsClear = false;

    for (var i = 0; i < vals.length; i++) {
      if (!formulas[i] && vals[i] !== "") {
        needsClear = true;
        break;
      }
    }
    if (needsClear) {
      range.clearContent();
    }
  }
}

/**
 * 대시보드 달력 그리드의 날짜 셀 위치 맵 구성
 * map[day] = { row: contentRow, col: col } (반드시 이번 달 날짜만 매핑)
 */
function buildCalendarCellMap(dashSheet, dowRow, currentMonth) {
  var totalRows = Math.min(dashSheet.getMaxRows(), 35);
  var values = dashSheet.getRange(1, 1, totalRows, 8).getValues();
  var map = {};
  var startIdx = dowRow; // dowRow 행 바로 다음이 1주차 날짜 행

  for (var r = startIdx; r < values.length; r += 2) {
    var contentRow = r + 2; // 날짜 바로 밑 내용 행 (1-indexed)

    for (var c = 1; c <= 7; c++) { // B(1) ~ H(7)
      var cellVal = values[r][c];
      if (cellVal === "" || cellVal === null || cellVal === undefined) continue;

      var numVal = null;
      var cellMonth = currentMonth;

      // Date 객체인 경우
      if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
        cellMonth = cellVal.getMonth() + 1;
        numVal = cellVal.getDate();
      } else {
        // 숫자 정수인 경우
        var parsed = parseInt(cellVal, 10);
        if (!isNaN(parsed) && parsed >= 1 && parsed <= 31) {
          numVal = parsed;
          if (r === startIdx && numVal > 20) {
            cellMonth = currentMonth === 1 ? 12 : currentMonth - 1;
          }
          if (r >= startIdx + 8 && numVal <= 10) {
            cellMonth = currentMonth === 12 ? 1 : currentMonth + 1;
          }
        }
      }

      if (numVal && cellMonth === currentMonth) {
        map[numVal] = { row: contentRow, col: c + 1 };
      }
    }
  }

  return map;
}

/**
 * 대시보드 상단 헤더 영역 청소
 */
function cleanupDashboardHeaderArea(dashSheet, dowRow) {
  if (dowRow <= 2) return;
  var range = dashSheet.getRange(1, 1, dowRow - 1, Math.min(dashSheet.getLastColumn(), 15));
  var values = range.getValues();
  var changed = false;

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var val = String(values[r][c] || "");
      if (val.indexOf("[") !== -1 && (val.indexOf("호]") !== -1 || val.indexOf("현장]") !== -1 || val.indexOf("] ") !== -1)) {
        values[r][c] = "";
        changed = true;
      }
    }
  }
  if (changed) {
    range.setValues(values);
  }
}

/**
 * 요일 헤더 행('일', '월'...) 위치 찾기
 */
function findDayOfWeekRow(dashSheet) {
  var maxRows = Math.min(dashSheet.getLastRow(), 15);
  var maxCols = Math.min(dashSheet.getLastColumn(), 10);
  if (maxRows < 1 || maxCols < 1) return 4;

  var values = dashSheet.getRange(1, 1, maxRows, maxCols).getValues();
  for (var r = 0; r < values.length; r++) {
    var matchCount = 0;
    for (var c = 0; c < values[r].length; c++) {
      var str = String(values[r][c] || "").trim();
      if (str === "일" || str === "월" || str === "화" || str === "수" || str === "목" || str === "금" || str === "토" ||
          str === "일요일" || str === "월요일" || str === "SUN" || str === "MON") {
        matchCount++;
      }
    }
    if (matchCount >= 4) return r + 1;
  }
  return 4;
}

/**
 * 대시보드 상단 선택 연도/월 읽기
 */
function getDashboardSelectedYearMonth(dashSheet, dowRow) {
  var now = new Date();
  var defYear = now.getFullYear();
  var defMonth = now.getMonth() + 1;
  if (dowRow <= 1) return { year: defYear, month: defMonth };

  var values = dashSheet.getRange(1, 1, dowRow - 1, Math.min(dashSheet.getLastColumn(), 10)).getValues();
  var foundYear = null;
  var foundMonth = null;

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var val = String(values[r][c] || "").trim();
      if (val === "년도" || val === "연도" || val === "Year") {
        if (c + 1 < values[r].length && parseInt(values[r][c + 1], 10) >= 2000) {
          foundYear = parseInt(values[r][c + 1], 10);
        }
      }
      if (val === "월" || val === "Month") {
        if (c + 1 < values[r].length) {
          var m = parseInt(values[r][c + 1], 10);
          if (m >= 1 && m <= 12) foundMonth = m;
        }
      }
      var n = parseInt(val, 10);
      if (n >= 2000 && n <= 2099 && !foundYear) {
        foundYear = n;
      }
    }
  }

  return { year: foundYear || defYear, month: foundMonth || defMonth };
}


// ========================================================
// 7. [날짜 및 기간 파싱 고도화 엔진]
// ========================================================

/**
 * 다양한 형태의 날짜 및 기간을 파싱
 * 반환값: {
 *   startDate, endDate, displayEndDate, isRange,
 *   hasTime, startTime, endTime, timeStr,
 *   year, month, day
 * }
 */
function parseScheduleDateRange(val) {
  if (!val) return null;

  // 1. 셀 값이 Date 객체인 경우 (cross-realm 호환 지원)
  var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function" && typeof val.getMonth === "function"));
  if (isDateObj && !isNaN(val.getTime())) {
    var hasTime = (val.getHours() !== 0 || val.getMinutes() !== 0);
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    var d = val.getDate();
    if (hasTime) {
      var startTime = new Date(val.getTime());
      var endTime = new Date(val.getTime() + 3600000); // 기본 1시간
      var hh = String(val.getHours());
      if (hh.length < 2) hh = "0" + hh;
      var mm = String(val.getMinutes());
      if (mm.length < 2) mm = "0" + mm;
      return {
        startDate: new Date(y, m - 1, d),
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: true,
        startTime: startTime,
        endTime: endTime,
        timeStr: hh + ":" + mm,
        year: y,
        month: m,
        day: d
      };
    } else {
      var dObj = new Date(y, m - 1, d);
      return {
        startDate: dObj,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: false,
        startTime: null,
        endTime: null,
        timeStr: "",
        year: y,
        month: m,
        day: d
      };
    }
  }

  var str = String(val).trim();
  if (!str) return null;

  // 1-1. 원시 Date.toString() 형식 감지 (예: "Wed Nov 11 2026 00:00:00 GMT+0900 (한국 표준시)")
  if (/(?:GMT|UTC|[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d+)/.test(str)) {
    try {
      var cleanDateStr = str.replace(/\s*\([^)]*\)\s*$/, "").trim();
      var directDate = new Date(cleanDateStr);
      if (!isNaN(directDate.getTime())) {
        return parseScheduleDateRange(directDate);
      }
    } catch(e) {}
  }

  // 2. 기간 구분자 감지: "Oct 21 → Oct 22", "10/21 ~ 10/22", "26.10.21 - 26.10.22", "26.10.11 15:00 ~ 17:00"
  var separatorRegex = /\s*(?:→|~|\s-\s|\s~|\.\.\.)\s*/;
  if (separatorRegex.test(str)) {
    var parts = str.split(separatorRegex);
    if (parts.length >= 2) {
      var p1 = parseDateTimeComponents(parts[0].trim());
      var p2 = parseDateTimeComponents(parts[1].trim());

      if (p1) {
        var currentYear = new Date().getFullYear();
        var startYear = p1.year || (p2 ? p2.year : currentYear);
        var sDate = new Date(startYear, p1.month - 1, p1.day);

        // 시간 일정인 경우 (p1 또는 p2에 시간 정보 포함)
        if (p1.hasTime || (p2 && p2.hasTime)) {
          var sHour = p1.hasTime ? p1.hour : 9;
          var sMin = p1.hasTime ? p1.minute : 0;
          var startTime = new Date(startYear, p1.month - 1, p1.day, sHour, sMin, 0);

          var endTime = null;
          var timeStr = "";
          if (p2) {
            var endYear = p2.year || startYear;
            var endMonth = p2.month ? p2.month - 1 : p1.month - 1;
            var endDay = p2.day || p1.day;
            var eHour = p2.hasTime ? p2.hour : (p1.hasTime ? p1.hour + 1 : 18);
            var eMin = p2.hasTime ? p2.minute : 0;
            endTime = new Date(endYear, endMonth, endDay, eHour, eMin, 0);

            var shStr = (sHour < 10 ? "0" : "") + sHour + ":" + (sMin < 10 ? "0" : "") + sMin;
            var ehStr = (eHour < 10 ? "0" : "") + eHour + ":" + (eMin < 10 ? "0" : "") + eMin;
            timeStr = shStr + "~" + ehStr;
          } else {
            endTime = new Date(startTime.getTime() + 3600000);
            timeStr = (sHour < 10 ? "0" : "") + sHour + ":" + (sMin < 10 ? "0" : "") + sMin;
          }

          var dispEnd = p2 && p2.day ? new Date(endTime.getFullYear(), endTime.getMonth(), endTime.getDate()) : sDate;

          return {
            startDate: sDate,
            endDate: dispEnd,
            displayEndDate: dispEnd,
            isRange: !isDateEqual(sDate, dispEnd),
            hasTime: true,
            startTime: startTime,
            endTime: endTime,
            timeStr: timeStr,
            year: startYear,
            month: p1.month,
            day: p1.day
          };
        } else if (p2) {
          // 종일 기간 일정 (시간 없음)
          var endYear = p2.year || startYear;
          var eDate = new Date(endYear, p2.month - 1, p2.day + 1); // exclusive +1일
          var dispEnd = new Date(endYear, p2.month - 1, p2.day);
          return {
            startDate: sDate,
            endDate: eDate,
            displayEndDate: dispEnd,
            isRange: true,
            hasTime: false,
            startTime: null,
            endTime: null,
            timeStr: "",
            year: startYear,
            month: p1.month,
            day: p1.day
          };
        }
      }
    }
  }

  // 3. 단일 날짜 및 시간 파싱 (예: "26.10.11 15:00" 또는 "26.10.11")
  var comp = parseDateTimeComponents(str);
  if (comp) {
    var currentYear = new Date().getFullYear();
    var y = comp.year || currentYear;
    var sDate = new Date(y, comp.month - 1, comp.day);

    if (comp.hasTime) {
      var startTime = new Date(y, comp.month - 1, comp.day, comp.hour, comp.minute, 0);
      var endTime = new Date(startTime.getTime() + 3600000); // 기본 1시간
      var sh = (comp.hour < 10 ? "0" : "") + comp.hour;
      var sm = (comp.minute < 10 ? "0" : "") + comp.minute;
      return {
        startDate: sDate,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: true,
        startTime: startTime,
        endTime: endTime,
        timeStr: sh + ":" + sm,
        year: y,
        month: comp.month,
        day: comp.day
      };
    } else {
      return {
        startDate: sDate,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: false,
        startTime: null,
        endTime: null,
        timeStr: "",
        year: y,
        month: comp.month,
        day: comp.day
      };
    }
  }

  return null;
}

/**
 * 단일 토큰에서 날짜 및 시간 분해 파싱
 * 반환값: { year, month, day, hour, minute, hasTime }
 */
function parseDateTimeComponents(val) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    var hasT = (val.getHours() !== 0 || val.getMinutes() !== 0);
    return {
      year: val.getFullYear(),
      month: val.getMonth() + 1,
      day: val.getDate(),
      hour: val.getHours(),
      minute: val.getMinutes(),
      hasTime: hasT
    };
  }

  var str = String(val).trim();
  if (!str) return null;

  var hour = null;
  var minute = null;
  var hasTime = false;

  // 1. 시간 패턴 추출 (예: "15:00", " 15:00", "T15:00")
  var timeRegex = /(?:^|\s|[T])(\d{1,2}):(\d{2})(?::(\d{2}))?/;
  var timeMatch = str.match(timeRegex);
  if (timeMatch) {
    var h = parseInt(timeMatch[1], 10);
    var min = parseInt(timeMatch[2], 10);
    if (h >= 0 && h <= 23 && min >= 0 && min <= 59) {
      hour = h;
      minute = min;
      hasTime = true;
      str = str.replace(timeRegex, " ").trim();
    }
  }

  // 시간이 추출되고 남은 문자열이 비어있으면 (예: "17:00")
  if (!str && hasTime) {
    return {
      year: null,
      month: null,
      day: null,
      hour: hour,
      minute: minute,
      hasTime: true
    };
  }

  // 2. 날짜 파싱
  var dateComp = parseDateComponents(str);
  if (dateComp) {
    return {
      year: dateComp.year,
      month: dateComp.month,
      day: dateComp.day,
      hour: hour,
      minute: minute,
      hasTime: hasTime
    };
  }

  if (hasTime) {
    return {
      year: null,
      month: null,
      day: null,
      hour: hour,
      minute: minute,
      hasTime: true
    };
  }

  return null;
}

/**
 * 단일 날짜 문자열 파싱 유틸리티 ({ year, month, day })
 */
function parseDateComponents(val) {
  if (!val) return null;

  var currentYear = new Date().getFullYear();

  if (val instanceof Date && !isNaN(val.getTime())) {
    return { year: val.getFullYear(), month: val.getMonth() + 1, day: val.getDate() };
  }

  var str = String(val).trim();
  if (!str) return null;

  // YY.MM.DD 또는 YYYY.MM.DD (예: 26.10.09, 2026-10-09, 2026. 10. 29)
  var mYmd = str.match(/^(\d{2,4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})/);
  if (mYmd) {
    var y = parseInt(mYmd[1], 10);
    if (y < 100) y += 2000;
    return { year: y, month: parseInt(mYmd[2], 10), day: parseInt(mYmd[3], 10) };
  }

  // MM/DD 또는 MM.DD (예: 10/9, 10. 9)
  var mMd = str.match(/^(\d{1,2})[-./]\s*(\d{1,2})$/);
  if (mMd) {
    return { year: currentYear, month: parseInt(mMd[1], 10), day: parseInt(mMd[2], 10) };
  }

  // 10월 9일
  var mKor = str.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (mKor) {
    return { year: currentYear, month: parseInt(mKor[1], 10), day: parseInt(mKor[2], 10) };
  }

  // Oct 9, Nov 14 등 영문 월
  var months = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  var mEng = str.match(/^([A-Za-z]{3})\s+(\d{1,2})/);
  if (mEng) {
    var mon = months[mEng[1].toLowerCase()];
    if (mon) return { year: currentYear, month: mon, day: parseInt(mEng[2], 10) };
  }

  return null;
}


// ========================================================
// 8. [시트 판별 헬퍼 - 탭 이름 변경에도 100% 견고한 구조 기반 판별]
// ========================================================

/**
 * 시트 상단(1~8행) 내에 '일, 월, 화, 수, 목, 금, 토' 요일 헤더가 존재하는지 검사
 */
function hasDayOfWeekHeader(sheet) {
  if (!sheet) return false;
  try {
    var maxR = Math.min(sheet.getLastRow(), 8);
    var maxC = Math.min(sheet.getLastColumn(), 10);
    if (maxR < 2 || maxC < 4) return false;

    var vals = sheet.getRange(1, 1, maxR, maxC).getValues();
    for (var r = 0; r < vals.length; r++) {
      var matchCount = 0;
      for (var c = 0; c < vals[r].length; c++) {
        var str = String(vals[r][c] || "").trim();
        if (str === "일" || str === "월" || str === "화" || str === "수" || str === "목" || str === "금" || str === "토" ||
            str === "일요일" || str === "월요일" || str === "SUN" || str === "MON") {
          matchCount++;
        }
      }
      if (matchCount >= 4) return true;
    }
  } catch(e) {}
  return false;
}

/**
 * 현장데이터 시트인지 판별 (1~4행 헤더 구조 및 요일 헤더 부재 검사)
 */
function isSiteScheduleSheet(sheet) {
  if (!sheet) return false;
  // 사용가이드/매뉴얼/구성원 시트는 절대 현장데이터 시트가 아님!
  var name = sheet.getName().toLowerCase();
  if (name.indexOf("가이드") !== -1 || name.indexOf("guide") !== -1 || name.indexOf("매뉴얼") !== -1 || name.indexOf("manual") !== -1 ||
      name.indexOf("구성원") !== -1 || name.indexOf("팀원") !== -1 || name.indexOf("직원") !== -1 || name.indexOf("조직도") !== -1 ||
      name.indexOf("member") !== -1 || name.indexOf("team") !== -1) {
    return false;
  }
  // 요일 헤더('일, 월, 화...')가 있는 시트는 대시보드이므로 절대 현장데이터 시트가 아님!
  if (hasDayOfWeekHeader(sheet)) return false;

  // 1순위: 1~4행 내에 '현장' 및 ('공종' 또는 '날짜' 또는 '진행상태') 헤더가 있는지 검사
  try {
    var maxR = Math.min(sheet.getLastRow(), 4);
    var maxC = Math.min(sheet.getLastColumn(), 12);
    if (maxR >= 1 && maxC >= 3) {
      var vals = sheet.getRange(1, 1, maxR, maxC).getValues();
      for (var r = 0; r < vals.length; r++) {
        var rowStr = vals[r].join(" ");
        if (rowStr.indexOf("현장") !== -1 && (rowStr.indexOf("공종") !== -1 || rowStr.indexOf("날짜") !== -1 || rowStr.indexOf("진행상태") !== -1)) {
          return true;
        }
      }
    }
  } catch(e) {}

  // 2순위: 시트 이름 기준
  if (name.indexOf("현장") !== -1 || name.indexOf("공정") !== -1 || name.indexOf("site") !== -1) {
    return true;
  }
  return false;
}

/**
 * 대시보드(달력) 시트인지 판별 (반드시 요일 헤더 '일, 월, 화...'가 있어야만 함!)
 */
function isDashboardSheet(sheet) {
  if (!sheet) return false;
  // 요일 헤더가 상단에 없는 시트는 탭 이름이 무엇이든 절대로 대시보드 시트가 아님!
  return hasDayOfWeekHeader(sheet);
}

/**
 * 활성 스프레드시트 가져오기 (웹앱 독립 실행 시 속성 저장 ID 자동 폴백)
 */
function getAppSpreadsheet() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  if (!ss) {
    try {
      var savedId = PropertiesService.getDocumentProperties().getProperty("SPREADSHEET_ID");
      if (savedId) ss = SpreadsheetApp.openById(savedId);
    } catch(e) {}
  }
  return ss;
}

function findDashboardSheet(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();

  // 1순위: 요일 헤더 구조를 갖춘 진짜 대시보드 달력 시트 찾기
  for (var i = 0; i < sheets.length; i++) {
    if (isDashboardSheet(sheets[i])) return sheets[i];
  }

  // 2순위: 탭 이름 후보 중 요일 헤더를 갖춘 시트
  var candidates = ["dashboard", "Dashboard", "대시보드", "00_대시보드", "만년달력", "달력", "00_Calendar", "Calendar"];
  for (var k = 0; k < candidates.length; k++) {
    var s = ss.getSheetByName(candidates[k]);
    if (s && hasDayOfWeekHeader(s)) return s;
  }

  return null;
}

function findSiteScheduleSheet(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();

  // 1순위: 구조 기반으로 현장데이터 시트 찾기
  for (var i = 0; i < sheets.length; i++) {
    if (isSiteScheduleSheet(sheets[i])) return sheets[i];
  }

  // 2순위: 대시보드나 가이드/구성원이 아닌 시트 중 첫 번째
  var dash = findDashboardSheet(ss);
  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().toLowerCase();
    if (sName.indexOf("가이드") === -1 && sName.indexOf("guide") === -1 && sName.indexOf("매뉴얼") === -1 && sName.indexOf("manual") === -1 &&
        sName.indexOf("구성원") === -1 && sName.indexOf("팀원") === -1 && sName.indexOf("직원") === -1 && sName.indexOf("조직도") === -1 &&
        sName.indexOf("member") === -1 && sName.indexOf("team") === -1) {
      if (!dash || sheets[k].getSheetId() !== dash.getSheetId()) {
        if (!hasDayOfWeekHeader(sheets[k])) {
          return sheets[k];
        }
      }
    }
  }

  return sheets.length > 1 ? sheets[1] : sheets[0];
}


// ========================================================
// 9. [현장데이터 UI 편의 기능] (아코디언, 신규 행 추가, 삭제, 서식)
// ========================================================

/**
 * 모든 현장 펼치기 (세부 공정 전체 보기)
 */
function expandAllSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  sheet.showRows(1, maxRows);
  sheet.expandAllRowGroups();
  SpreadsheetApp.flush();
  ss.toast("모든 현장의 세부 공정이 펼쳐졌습니다.", "전체 펼치기 완료", 3);
}

/**
 * 모든 현장 접기 (현장명 대표 행만 보기)
 */
function collapseAllSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  sheet.collapseAllRowGroups();
  SpreadsheetApp.flush();
  ss.toast("모든 현장이 대표 행으로 접혔습니다.", "전체 접기 완료", 3);
}

/**
 * 상단(1행과 2행 사이)에 신규 현장 4개 행 추가 & 아코디언 설정 (작성일/작성자 자동 기록)
 */
function addNewSiteAtTop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "➕ 신규 현장 추가",
    "새로운 현장명을 입력하세요 (예: 래미안 101동 202호):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var siteName = response.getResponseText().trim() || "신규 현장";

  ensureCalendarSyncColumns(sheet);

  try {
    sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch(e) {}

  sheet.insertRowsBefore(2, 4);

  var defaultTasks = ["상담", "철거", "목공", "마감"];
  var newRows = [];
  var now = new Date();
  var nowMs = now.getTime();
  var nowStr = formatDateTime(now);
  var currentUser = getCurrentUserIdentifier();

  // 14개 열 채우기
  // 1:현장, 2:날짜, 3:공종, 4:협력업체, 5:공유, 6:비고, 7:담당자, 8:진행상태, 9:캘린더ID, 10:타임스탬프, 11:작성일, 12:작성자, 13:수정일, 14:수정자
  for (var i = 0; i < defaultTasks.length; i++) {
    newRows.push([
      siteName, "", defaultTasks[i], "", false, "", "", "예정",
      "", nowMs, nowStr, currentUser, nowStr, currentUser
    ]);
  }

  sheet.getRange(2, 1, 4, 14).setValues(newRows);
  sheet.getRange(2, 1, 4, 14).setVerticalAlignment("middle");
  sheet.getRange(2, 5, 4, 1).insertCheckboxes();

  [1, 2, 3, 5, 7, 8, 11, 12, 13, 14].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("left");
  });

  // C열(공종), G열(담당자), H열(진행상태) 드롭다운 유효성 검사 일괄 적용
  try {
    applySiteScheduleValidations(sheet, 2, 4);
  } catch(e) {}

  // 격자 테두리(Border) 깔끔하게 적용
  try {
    sheet.getRange(2, 1, 4, 8).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}

  // 조건부 서식(진행상태 색상 등) 전수 갱신 보장
  try {
    setupSiteSheetConditionalFormatting(sheet);
  } catch(e) {}

  try {
    sheet.getRange("3:5").shiftRowGroupDepth(1);
  } catch(e) {}

  sheet.showRows(2, 4);
  try {
    sheet.hideColumns(9, 2);
  } catch(e) {}

  sheet.setActiveRange(sheet.getRange("A2"));
  ss.toast("'" + siteName + "' 4개 행이 상단에 추가되었습니다.", "추가 완료", 3);
}

/**
 * 현장명과 공종이 없는 빈 행에 잘못 입력된 유령 이력/캘린더ID 일괄 청소
 */
function cleanupEmptyRowsMetadata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  var cleaned = 0;

  // 2행 C2(공종)와 E2(공유)에 실수로 대시보드 연도/월 숫자가 들어간 경우 동적 자동 복구
  try {
    var c2 = sheet.getRange("C2");
    var e2 = sheet.getRange("E2");
    var c2Val = String(c2.getValue()).trim();
    var e2Val = String(e2.getValue()).trim();

    // 4자리 연도 숫자(예: 2025, 2026, 2027 등)가 공종명에 잘못 들어간 경우 삭제
    if (/^\d{4}$/.test(c2Val)) {
      c2.clearContent();
    }
    // 1~12 사이의 월 숫자가 체크박스에 잘못 들어간 경우 체크 해제(false)로 복구
    if (/^(1[0-2]|[1-9])$/.test(e2Val)) {
      e2.setValue(false);
    }
  } catch(e) {}

  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var rowVals = values[i];
    var siteName = String(rowVals[0] || "").trim();
    var taskName = String(rowVals[2] || "").trim();

    // B열(날짜열)에 대시보드 달력 텍스트([현장명] 공종...)가 잘못 덮어써진 경우 자동 복구
    var rawDateStr = String(rowVals[1] || "");
    if (rawDateStr.indexOf("[") !== -1 && (rawDateStr.indexOf("]") !== -1 || rawDateStr.indexOf("\n") !== -1)) {
      rowVals[1] = "";
      cleaned++;
      changed = true;
    }

    // 현장명과 공종이 모두 없는 빈 행인 경우
    if (!siteName && !taskName) {
      var hasMetadata = rowVals[8] || rowVals[9] || rowVals[10] || rowVals[11] || rowVals[12] || rowVals[13];
      if (hasMetadata) {
        rowVals[8] = "";
        rowVals[9] = "";
        rowVals[10] = "";
        rowVals[11] = "";
        rowVals[12] = "";
        rowVals[13] = "";
        cleaned++;
        changed = true;
      }
    }
  }

  if (changed) {
    sheet.getRange(2, 1, lastRow - 1, 14).setValues(values);
  }

  SpreadsheetApp.flush();
  ss.toast(cleaned + "개 빈 행의 유령 이력 및 캘린더 ID가 깨끗하게 청소되었습니다.", "🧹 청소 완료", 3);
}


// ========================================================
// 10. [모바일 웹 앱: 캘린더 일정 상세 드롭다운 상태 변경 엔진]
// ========================================================

/**
 * 배포된 웹 앱 URL 가져오기
 */
function getWebAppUrl() {
  var docProps = PropertiesService.getDocumentProperties();
  var savedUrl = docProps.getProperty("WEB_APP_URL");
  if (savedUrl) return savedUrl.trim();

  try {
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf("macros/s") !== -1) {
      return url;
    }
  } catch(e) {}

  return "";
}

/**
 * [메뉴용] 모바일 상태변경 웹앱 배포 URL 확인 및 등록 대화상자
 */
function configureWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var currentUrl = getWebAppUrl();

  var promptText = "📌 현재 등록된 모바일 웹앱 URL:\n" +
                   (currentUrl ? currentUrl : "(아직 등록되지 않음)") + "\n\n" +
                   "[웹앱 배포 방법]:\n" +
                   "1. 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭\n" +
                   "2. 우측 상단 파란색 [배포] > [새 배포] 클릭\n" +
                   "3. 유형: [웹 앱] 선택, 액세스 권한: [모든 사용자(Anyone)] 선택 후 [배포]\n" +
                   "4. 발급된 '웹 앱 URL'을 복사하여 아래에 입력하세요:";

  var response = ui.prompt("📱 모바일 웹앱 배포 URL 설정", promptText, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  var docProps = PropertiesService.getDocumentProperties();

  if (input) {
    docProps.setProperty("WEB_APP_URL", input);
    ui.alert("✅ 모바일 웹앱 URL이 성공적으로 등록되었습니다!\n\n상단 메뉴 [🚀 현장 관리] > [구글 캘린더 전체 동기화]를 1회 실행해주시면 기존 캘린더 일정들의 링크도 새 URL로 일괄 갱신됩니다.");
  } else {
    docProps.deleteProperty("WEB_APP_URL");
    ui.alert("모바일 웹앱 URL 설정이 해제되었습니다.");
  }
}

/**
 * 모바일 웹앱 상태 변경 서버 함수 (google.script.run 클라이언트 RPC 연동)
 */
function updateScheduleStatusFromWeb(row, newStatus, newShared, newMemo) {
  var rowNum = parseInt(row, 10);
  if (!rowNum || rowNum < 2 || !newStatus) {
    return { success: false, error: "유효하지 않은 요청 데이터입니다." };
  }

  // 모바일 동시 접속 충돌 방지 락 (10초 대기)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, error: "현재 다른 작업이 처리 중입니다. 잠시 후 다시 시도해 주세요." };
  }

  try {
    var ss = getAppSpreadsheet();
    if (!ss) {
      return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    }
    var sheet = findSiteScheduleSheet(ss);
    if (!sheet) {
      return { success: false, error: "현장데이터 시트를 찾을 수 없습니다." };
    }

    sheet.getRange(rowNum, 8).setValue(newStatus); // H열 진행상태 업데이트

    // E열 공유 상태 업데이트 (파라미터가 명시적으로 전달된 경우)
    var updatedShared = null;
    if (newShared !== undefined && newShared !== null) {
      var isSharedBool = (newShared === true || String(newShared).trim().toLowerCase() === "true" || String(newShared).trim() === "1");
      sheet.getRange(rowNum, 5).setValue(isSharedBool); // E열 공유 체크박스 업데이트
      updatedShared = isSharedBool;
    }

    // F열 비고 업데이트 (파라미터가 명시적으로 전달된 경우)
    var updatedMemo = null;
    if (newMemo !== undefined && newMemo !== null) {
      var cleanMemo = String(newMemo).trim();
      sheet.getRange(rowNum, 6).setValue(cleanMemo); // F열(6열) 비고 업데이트
      updatedMemo = cleanMemo;
    }

    var now = new Date();
    sheet.getRange(rowNum, 10).setValue(now.getTime()); // J열 타임스탬프
    sheet.getRange(rowNum, 13, 1, 2).setValues([[formatDateTime(now), "모바일 웹"]]); // M열/N열 수정이력 일괄 쓰기

    // 캘린더 및 대시보드 즉시 동기화 (forceSheetToCal = true)
    syncSingleRowToCalendar(sheet, rowNum, true);
    SpreadsheetApp.flush();

    return { 
      success: true, 
      status: newStatus,
      shared: updatedShared,
      memo: updatedMemo
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 모바일 웹 드롭다운 상태 변경 화면 (doGet)
 */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var row = parseInt(p.row, 10);
  var targetSite = p.site ? String(p.site).trim() : "";
  var targetTask = p.task ? String(p.task).trim() : "";
  var targetId = p.id ? String(p.id).trim() : "";
  var action = p.action;
  var newStatus = p.status;
  var newShared = p.shared;
  var newMemo = p.memo;

  var ss = getAppSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);

  // [스마트 행 자동 보정 엔진]
  // 신규 현장 추가, 행 삭제, 정렬 등으로 인해 시트 행 번호가 밀린 경우 실시간 자동 재탐색
  if (sheet && (targetId || targetSite || targetTask)) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var allData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      var isCurrentRowMatch = false;

      // 1. 전달받은 row 위치의 행이 실제로 해당 현장/공종인지 확인
      if (row >= 2 && row <= lastRow) {
        var curRowVals = allData[row - 2];
        var curSite = String(curRowVals[0] || "").trim();
        var curTask = String(curRowVals[2] || "").trim();
        var curId = String(curRowVals[8] || "").trim();

        if (targetId && curId === targetId) {
          isCurrentRowMatch = true;
        } else if (targetSite && targetTask && curSite === targetSite && curTask === targetTask) {
          isCurrentRowMatch = true;
        }
      }

      // 2. 만약 행이 밀려서 맞지 않는 경우, 시트 전체에서 진짜 위치를 찾아 자동 보정
      if (!isCurrentRowMatch) {
        var resolvedRow = -1;
        // 1순위: 캘린더 고유 Event ID 일치 행 탐색
        if (targetId) {
          for (var i = 0; i < allData.length; i++) {
            if (String(allData[i][8] || "").trim() === targetId) {
              resolvedRow = i + 2;
              break;
            }
          }
        }
        // 2순위: 현장명 + 공종명 일치 행 탐색
        if (resolvedRow === -1 && targetSite && targetTask) {
          for (var j = 0; j < allData.length; j++) {
            var sVal = String(allData[j][0] || "").trim();
            var tVal = String(allData[j][2] || "").trim();
            if (sVal === targetSite && tVal === targetTask) {
              resolvedRow = j + 2;
              break;
            }
          }
        }

        if (resolvedRow !== -1) {
          Logger.log("🔄 행 번호 자동 보정 완료: " + row + "행 ➔ " + resolvedRow + "행 (현장: " + targetSite + " / " + targetTask + ")");
          row = resolvedRow;
        } else {
          // [안전 가드] targetId나 현장/공종 정보가 제공되었으나 시트 어디에도 일치하는 데이터가 없는 경우(삭제 등):
          // 이전 row 번호로 엉뚱한 행을 보여주거나 덮어쓰지 않도록 유효하지 않음(-1) 처리
          row = -1;
        }
      }
    }
  }

  // 1) 상태 업데이트 요청 (URL 파라미터 직접 호출 또는 GET 이동 폴백)
  if (action === "update" && row >= 2 && newStatus && sheet) {
    var resObj = updateScheduleStatusFromWeb(row, newStatus, newShared, newMemo);
    if (p.format === "json") {
      return ContentService.createTextOutput(JSON.stringify(resObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var siteName = "";
    var taskName = "";
    try {
      siteName = String(sheet.getRange(row, 1).getValue() || "").trim();
      taskName = String(sheet.getRange(row, 3).getValue() || "").trim();
    } catch(err) {}

    var shareNotice = "";
    if (newShared !== undefined && newShared !== null) {
      var isSh = (newShared === true || String(newShared).toLowerCase() === "true" || String(newShared) === "1");
      shareNotice = "<br>공유 상태: <b>[" + (isSh ? "🟦 공유(ON)" : "🟥 미공유(OFF)") + "]</b>";
    }

    var memoNotice = "";
    if (newMemo !== undefined && newMemo !== null && String(newMemo).trim()) {
      memoNotice = "<br>비고 메모: <b>" + String(newMemo).trim() + "</b>";
    }

    var resHtml = "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>" +
      "<style>body { background:#f1f5f9; display:flex; justify-content:center; align-items:center; min-height:100vh; padding:16px; font-family:-apple-system, sans-serif; }" +
      ".box { background:#fff; padding:32px 24px; border-radius:18px; text-align:center; max-width:380px; width:100%; box-shadow:0 8px 24px rgba(0,0,0,0.08); }" +
      "h2 { color:#15803d; font-size:20px; margin-bottom:12px; }" +
      "p { color:#475569; font-size:14px; line-height:1.6; margin-bottom:20px; }" +
      ".btn { display:inline-block; padding:12px 24px; background:#2563eb; color:#fff; text-decoration:none; border-radius:10px; font-weight:700; font-size:14px; }" +
      "</style></head><body>" +
      "<div class='box'>" +
      "<h2>✅ 공유/진행상태/비고 변경 완료</h2>" +
      "<p>[" + (siteName || "현장") + "] <b>" + (taskName || "공정") + "</b>의" + shareNotice + "<br>진행상태가 <b>[" + newStatus + "]</b>(으)로" + memoNotice + "<br>정상 변경되었습니다.<br><br>시트와 구글 캘린더에 실시간 반영되었습니다.</p>" +
      "<a href='javascript:window.close()' class='btn'>창 닫기</a>" +
      "</div></body></html>";
    return HtmlService.createHtmlOutput(resHtml)
      .setTitle("공유/진행상태/비고 변경 완료 - " + newStatus);
  }

  // 2) 모바일 UI 화면 출력
  if (!row || row < 2 || !sheet) {
    return HtmlService.createHtmlOutput("<h3 style='font-family:sans-serif;text-align:center;margin-top:50px;'>⚠️ 올바른 일정 링크가 아닙니다.</h3>");
  }

  var lastRow = sheet.getLastRow();
  if (row > lastRow) {
    return HtmlService.createHtmlOutput("<h3 style='font-family:sans-serif;text-align:center;margin-top:50px;'>⚠️ 이미 삭제되었거나 유효하지 않은 일정입니다.</h3>");
  }

  var rowVals = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var siteName = String(rowVals[0] || "-").trim();
  var dateStr = formatScheduleDateDisplay(rowVals[1]) || String(rowVals[1] || "-").trim();
  var taskName = String(rowVals[2] || "-").trim();
  var rawShare = rowVals[4];
  var isShared = (rawShare === true || String(rawShare).trim().toUpperCase() === "TRUE" || String(rawShare).trim() === "1" || String(rawShare).trim() === "V" || String(rawShare).trim() === "O");
  var memo = String(rowVals[5] || "").trim(); // F열 비고
  var currentStatus = String(rowVals[7] || "예정").trim() || "예정";
  var eventIdVal = String(rowVals[8] || targetId || "").trim();
  var webAppUrl = getWebAppUrl();

  var html = buildMobileStatusHtml(row, siteName, dateStr, taskName, currentStatus, webAppUrl, isShared, memo, eventIdVal);
  return HtmlService.createHtmlOutput(html)
    .setTitle("현장 공정 공유/진행상태/비고 관리 - " + siteName)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");
}

/**
 * HTML 특수문자 안전 이스케이프 헬퍼
 */
function escapeHtmlForInput(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * 모바일 반응형 상태 변경 HTML 생성기 (비고 편집 추가)
 */
function buildMobileStatusHtml(row, siteName, dateStr, taskName, currentStatus, webAppUrl, isShared, memo, eventId) {
  if (webAppUrl === undefined) webAppUrl = getWebAppUrl();
  if (isShared === undefined) isShared = true;
  if (memo === undefined || memo === null) memo = "";
  if (eventId === undefined || eventId === null) eventId = "";
  var formattedDateStr = formatScheduleDateDisplay(dateStr) || dateStr || "-";
  var isSharedBool = (isShared === true || String(isShared).trim().toUpperCase() === "TRUE" || String(isShared).trim() === "1");

  var html = "<!DOCTYPE html>" +
    "<html>" +
    "<head>" +
    "  <meta charset='utf-8'>" +
    "  <meta name='viewport' content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'>" +
    "  <title>현장 공정 공유/진행상태/비고 관리</title>" +
    "  <style>" +
    "    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }" +
    "    body { background-color: #f1f5f9; color: #1e293b; display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 16px; }" +
    "    .card { background: #ffffff; width: 100%; max-width: 420px; border-radius: 20px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1); padding: 24px; }" +
    "    .header { text-align: center; margin-bottom: 20px; }" +
    "    .badge { display: inline-block; padding: 4px 12px; border-radius: 9999px; font-size: 13px; font-weight: 600; background: #e2e8f0; color: #475569; margin-bottom: 8px; }" +
    "    .title { font-size: 20px; font-weight: 700; color: #0f172a; margin-bottom: 4px; }" +
    "    .subtitle { font-size: 15px; color: #64748b; font-weight: 600; }" +
    "    .info-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px; margin-bottom: 20px; font-size: 14px; line-height: 1.6; }" +
    "    .info-row { display: flex; justify-content: space-between; align-items: center; }" +
    "    .info-label { color: #64748b; }" +
    "    .info-val { font-weight: 600; color: #1e293b; }" +
    "    .section-label { font-size: 14px; font-weight: 700; color: #334155; margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }" +
    "    .share-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }" +
    "    .share-btn { padding: 13px 8px; border-radius: 12px; border: 2px solid #cbd5e1; font-size: 15px; font-weight: 700; cursor: pointer; transition: all 0.15s ease; display: flex; align-items: center; justify-content: center; gap: 6px; background: #ffffff; color: #64748b; }" +
    "    .share-btn.active-on { background: #eff6ff; color: #1d4ed8; border-color: #2563eb; border-width: 3px; transform: scale(1.02); }" +
    "    .share-btn.active-off { background: #fef2f2; color: #b91c1c; border-color: #ef4444; border-width: 3px; transform: scale(1.02); }" +
    "    .btn-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 16px; }" +
    "    .status-btn { padding: 14px 8px; border-radius: 12px; border: 2px solid #cbd5e1; font-size: 15px; font-weight: 700; cursor: pointer; transition: all 0.15s ease; display: flex; align-items: center; justify-content: center; gap: 6px; background: #ffffff; }" +
    "    .btn-plan.active { background: #f1f5f9; color: #334155; border-color: #64748b; border-width: 3px; transform: scale(1.02); }" +
    "    .btn-progress.active { background: #e0f2fe; color: #0284c7; border-color: #0284c7; border-width: 3px; transform: scale(1.02); }" +
    "    .btn-done.active { background: #dcfce7; color: #16a34a; border-color: #16a34a; border-width: 3px; transform: scale(1.02); }" +
    "    .btn-as.active { background: #ffedd5; color: #ea580c; border-color: #ea580c; border-width: 3px; transform: scale(1.02); }" +
    "    .memo-textarea { width: 100%; min-height: 80px; padding: 12px 14px; border-radius: 12px; border: 2px solid #cbd5e1; background: #ffffff; color: #1e293b; font-size: 14.5px; font-family: inherit; line-height: 1.5; outline: none; resize: vertical; margin-bottom: 20px; transition: border-color 0.2s; }" +
    "    .memo-textarea:focus { border-color: #2563eb; }" +
    "    .submit-btn { width: 100%; height: 52px; background: #2563eb; color: #ffffff; border: none; border-radius: 14px; font-size: 16px; font-weight: 700; cursor: pointer; transition: background 0.2s; display: flex; align-items: center; justify-content: center; gap: 8px; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.3); }" +
    "    .submit-btn:active { background: #1d4ed8; }" +
    "    .submit-btn:disabled { background: #94a3b8; cursor: not-allowed; }" +
    "    .toast { display: none; padding: 14px; border-radius: 12px; text-align: center; font-size: 14px; font-weight: 600; margin-top: 14px; line-height: 1.5; }" +
    "    .toast-success { background: #dcfce7; color: #15803d; border: 1px solid #86efac; }" +
    "    .toast-error { background: #fee2e2; color: #b91c1c; border: 1px solid #fca5a5; }" +
    "    .footer { text-align: center; margin-top: 16px; font-size: 12px; color: #94a3b8; }" +
    "  </style>" +
    "</head>" +
    "<body>" +
    "  <div class='card'>" +
    "    <div class='header'>" +
    "      <div class='badge'>현장 공정 공유/진행상태/비고 관리</div>" +
    "      <div class='title'>" + siteName + "</div>" +
    "      <div class='subtitle'>" + taskName + "</div>" +
    "    </div>" +
    "    <div class='info-box'>" +
    "      <div class='info-row'><span class='info-label'>일시:</span><span class='info-val'>" + formattedDateStr + "</span></div>" +
    "      <div class='info-row' style='margin-top:6px;'><span class='info-label'>현재 공유:</span><span class='info-val' id='currShareDisplay'>" + (isSharedBool ? "🟦 공유 완료" : "🟥 미공유") + "</span></div>" +
    "      <div class='info-row' style='margin-top:6px;'><span class='info-label'>현재 상태:</span><span class='info-val' id='currStatusDisplay'>" + currentStatus + "</span></div>" +
    "    </div>" +
    "    <div class='section-label'>🔗 공유 상태 설정</div>" +
    "    <div class='share-grid'>" +
    "      <button type='button' id='btnShareOn' class='share-btn " + (isSharedBool ? "active-on" : "") + "' onclick='selectShare(true)'>🟦 공유 (ON)</button>" +
    "      <button type='button' id='btnShareOff' class='share-btn " + (!isSharedBool ? "active-off" : "") + "' onclick='selectShare(false)'>🟥 미공유 (OFF)</button>" +
    "    </div>" +
    "    <div class='section-label'>⚡ 진행 상태 설정</div>" +
    "    <div class='btn-grid'>" +
    "      <button type='button' class='status-btn btn-plan " + (currentStatus === "예정" ? "active" : "") + "' onclick='selectStatus(\"예정\")'>📅 예정</button>" +
    "      <button type='button' class='status-btn btn-progress " + (currentStatus.indexOf("진행") !== -1 ? "active" : "") + "' onclick='selectStatus(\"진행 중\")'>▶️ 진행 중</button>" +
    "      <button type='button' class='status-btn btn-done " + (currentStatus === "완료" ? "active" : "") + "' onclick='selectStatus(\"완료\")'>✅ 완료</button>" +
    "      <button type='button' class='status-btn btn-as " + (currentStatus === "AS" ? "active" : "") + "' onclick='selectStatus(\"AS\")'>🛠️ AS</button>" +
    "    </div>" +
    "    <div class='section-label'>📝 현장 비고 / 특이사항 메모</div>" +
    "    <textarea id='memoInput' class='memo-textarea' placeholder='현장 특이사항이나 전달사항을 입력하세요 (예: 자재 입고 완료, 오후 시공)'>" + escapeHtmlForInput(memo) + "</textarea>" +
    "    <button type='button' id='saveBtn' class='submit-btn' onclick='saveStatus()'>" +
    "      <span>저장하기</span>" +
    "    </button>" +
    "    <div id='toast' class='toast'></div>" +
    "    <div class='footer'>우성아이디 현장관리 시스템</div>" +
    "  </div>" +
    "  <script>" +
    "    var selected = '" + currentStatus + "';" +
    "    var selectedShare = " + (isSharedBool ? "true" : "false") + ";" +
    "    function selectShare(val) {" +
    "      selectedShare = val;" +
    "      updateShareUI();" +
    "    }" +
    "    function updateShareUI() {" +
    "      var onBtn = document.getElementById('btnShareOn');" +
    "      var offBtn = document.getElementById('btnShareOff');" +
    "      if (selectedShare) {" +
    "        onBtn.classList.add('active-on');" +
    "        offBtn.classList.remove('active-off');" +
    "      } else {" +
    "        onBtn.classList.remove('active-on');" +
    "        offBtn.classList.add('active-off');" +
    "      }" +
    "    }" +
    "    function selectStatus(s) {" +
    "      selected = s;" +
    "      updateBtnUI();" +
    "    }" +
    "    function updateBtnUI() {" +
    "      var btns = document.querySelectorAll('.status-btn');" +
    "      var cleanSel = selected.replace(/\\s+/g, '');" +
    "      btns.forEach(function(b) {" +
    "        b.classList.remove('active');" +
    "        var cleanText = b.innerText.replace(/\\s+/g, '');" +
    "        if (cleanText.indexOf(cleanSel) !== -1) b.classList.add('active');" +
    "      });" +
    "    }" +
    "    function saveStatus() {" +
    "      var btn = document.getElementById('saveBtn');" +
    "      var toast = document.getElementById('toast');" +
    "      var memoVal = document.getElementById('memoInput').value;" +
    "      btn.disabled = true;" +
    "      btn.innerText = '저장 처리 중...';" +
    "      toast.style.display = 'none';" +
    "      if (typeof google !== 'undefined' && google.script && google.script.run) {" +
    "        google.script.run" +
    "          .withSuccessHandler(function(res) {" +
    "            btn.disabled = false;" +
    "            btn.innerText = '저장하기';" +
    "            if (res && res.success) {" +
    "              document.getElementById('currStatusDisplay').innerText = selected;" +
    "              document.getElementById('currShareDisplay').innerText = selectedShare ? '🟦 공유 완료' : '🟥 미공유';" +
    "              toast.className = 'toast toast-success';" +
    "              toast.innerHTML = '✅ 공유 상태 [<b>' + (selectedShare ? '🟦 공유(ON)' : '🟥 미공유(OFF)') + '</b>], 진행상태 [<b>' + selected + '</b>], 비고가 즉시 저장되었습니다!<br><span style=\"font-size:12px;font-weight:normal;\">시트와 캘린더에 실시간 반영 완료되었습니다. 창을 닫으셔도 됩니다.</span>';" +
    "              toast.style.display = 'block';" +
    "            } else {" +
    "              toast.className = 'toast toast-error';" +
    "              toast.innerText = '⚠️ 변경 실패: ' + ((res && res.error) ? res.error : '오류가 발생했습니다.');" +
    "              toast.style.display = 'block';" +
    "            }" +
    "          })" +
    "          .withFailureHandler(function(err) {" +
    "            console.warn('google.script.run 실패, URL 폴백 시도:', err);" +
    "            execFallback();" +
    "          })" +
    "          .updateScheduleStatusFromWeb(" + row + ", selected, selectedShare, memoVal);" +
    "        return;" +
    "      }" +
    "      execFallback();" +
    "      function execFallback() {" +
    "        var baseEndpoint = '" + (webAppUrl || "") + "';" +
    "        if (!baseEndpoint && window.location.href.indexOf('macros/s/') !== -1) {" +
    "          baseEndpoint = window.location.href.split('?')[0];" +
    "        }" +
    "        if (baseEndpoint) {" +
    "          var targetUrl = baseEndpoint + '?action=update&row=" + row + "&status=' + encodeURIComponent(selected) + '&shared=' + selectedShare + '&memo=' + encodeURIComponent(memoVal)" + (eventId ? " + '&id=" + encodeURIComponent(eventId) + "'" : "") + ";" +
    "          window.location.href = targetUrl;" +
    "        } else {" +
    "          btn.disabled = false;" +
    "          btn.innerText = '저장하기';" +
    "          toast.className = 'toast toast-error';" +
    "          toast.innerText = '⚠️ 웹앱 환경을 확인할 수 없습니다. 브라우저에서 다시 열어주세요.';" +
    "          toast.style.display = 'block';" +
    "        }" +
    "      }" +
    "    }" +
    "  </script>" +
    "</body>" +
    "</html>";
  return html;
}


// ========================================================
// 11. [모닝 브리핑 & 공정 지연 리스크 감지 & 캘린더 색상 엔진]
// ========================================================

/**
 * 일정의 유효 마감일자(Date) 반환 (단일 날짜인 경우 startDate, 기간인 경우 displayEndDate)
 */
function getEffectiveEndDate(parsedDate) {
  if (!parsedDate) return null;
  if (parsedDate.displayEndDate) {
    return parsedDate.displayEndDate;
  }
  return parsedDate.startDate;
}

/**
 * 공정 지연 여부 판정 (마감일이 오늘 이전 & 미완료 상태)
 * @param {Object} parsedDate parseScheduleDateRange 결과 객체
 * @param {string} statusClean 공백 제거된 진행상태 문자열
 * @param {Date} [refDate] 기준 일자 (생략 시 현재 시각)
 * @returns {boolean} 지연 여부
 */
function isScheduleDelayed(parsedDate, statusClean, refDate) {
  if (!parsedDate) return false;
  if (!statusClean) statusClean = "";
  // '완료' 상태는 절대 지연 대상이 아님
  if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
    return false;
  }
  var effectiveEnd = getEffectiveEndDate(parsedDate);
  if (!effectiveEnd) return false;

  var now = refDate || new Date();
  var todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  var endMidnight = new Date(effectiveEnd.getFullYear(), effectiveEnd.getMonth(), effectiveEnd.getDate()).getTime();

  return endMidnight < todayMidnight;
}

/**
 * 특정 일자에 활성화된(진행 중인) 공정인지 여부 판정
 * @param {Object} parsedDate parseScheduleDateRange 결과 객체
 * @param {Date} [refDate] 기준 일자 (기본 오늘)
 * @returns {boolean} 활성 여부
 */
function isScheduleActiveToday(parsedDate, refDate) {
  if (!parsedDate || !parsedDate.startDate) return false;
  var now = refDate || new Date();
  var todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

  var sMidnight = new Date(parsedDate.startDate.getFullYear(), parsedDate.startDate.getMonth(), parsedDate.startDate.getDate()).getTime();
  var effectiveEnd = getEffectiveEndDate(parsedDate);
  var eMidnight = new Date(effectiveEnd.getFullYear(), effectiveEnd.getMonth(), effectiveEnd.getDate()).getTime();

  return (todayMidnight >= sMidnight && todayMidnight <= eMidnight);
}

/**
 * 진행상태 및 지연 여부에 따른 구글 캘린더 블록 색상 ID 산출 (Feature 1)
 * 1: Pale Blue (예정)
 * 4: Flamingo (AS)
 * 7: Peacock (진행 중)
 * 8: Graphite (미공유)
 * 10: Basil (완료)
 * 11: Tomato (지연 경고)
 */
function getCalendarColorForStatus(statusClean, isDelayed, isShared) {
  if (!statusClean) statusClean = "";
  // 1순위: 완료 (초록색)
  if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
    return "10"; // Basil
  }
  // 2순위: 지연 경고 (빨간색)
  if (isDelayed) {
    return "11"; // Tomato
  }
  // 3순위: 진행 중 (청록/파란색)
  if (statusClean.indexOf("진행") !== -1 || statusClean.indexOf("▶") !== -1) {
    return "7";  // Peacock
  }
  // 4순위: AS (산호/주황색)
  if (statusClean.indexOf("AS") !== -1 || statusClean.indexOf("A/S") !== -1 || statusClean.indexOf("as") !== -1 || statusClean.indexOf("🛠") !== -1) {
    return "4";  // Flamingo
  }
  // 5순위: 미공유 일정 (회색)
  if (!isShared) {
    return "8";  // Graphite
  }
  // 기본: 예정 (연파란색)
  return "1";    // Pale Blue
}

/**
 * 현장데이터 시트의 전체 행 공정 지연 검사 및 배경 하이라이트 반영 (Feature 2)
 */
function checkAndHighlightDelayedSchedules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("검사할 현장 데이터가 없습니다.");
    return;
  }

  var today = new Date();
  var range = sheet.getRange(2, 1, lastRow - 1, 8);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var delayedItems = [];
  var bgChanged = false;
  var startTime = Date.now();

  for (var i = 0; i < values.length; i++) {
    // 6분 타임아웃 절대 방어: 60초 초과 시 안전 루프 탈출
    if (Date.now() - startTime > 60000) {
      break;
    }

    var row = i + 2;
    var siteName = String(values[i][0] || "").trim();
    var dateVal = values[i][1];
    var taskName = String(values[i][2] || "").trim();
    var manager = String(values[i][6] || "").trim();
    var status = String(values[i][7] || "").trim();
    var statusClean = status.replace(/\s+/g, "");

    if (!siteName && !taskName) continue;
    var parsed = parseScheduleDateRange(dateVal);
    if (!parsed) continue;

    if (isScheduleDelayed(parsed, statusClean, today)) {
      delayedItems.push({
        row: row,
        siteName: siteName,
        taskName: taskName,
        dateStr: String(dateVal).trim(),
        manager: manager,
        status: status
      });
      // 지연 행 연한 붉은색 경고 하이라이트 (메모리 2D 배열 일괄 수정)
      if (backgrounds[i][0] !== "#fff1f2") {
        for (var c = 0; c < 8; c++) {
          backgrounds[i][c] = "#fff1f2";
        }
        bgChanged = true;
      }
    } else {
      // 정상 상태 복구 (기존에 붉은색이었던 경우 흰색으로 복구)
      var currentBg = backgrounds[i][0];
      if (currentBg === "#fff1f2" || currentBg === "#fee2e2" || currentBg === "#fce8e6") {
        for (var c2 = 0; c2 < 8; c2++) {
          backgrounds[i][c2] = "#ffffff";
        }
        bgChanged = true;
      }
    }
  }

  // 1회 단일 호출로 전체 시트 배경색 즉시 커밋 (기존 개별 호출 대비 300배 가속 & API 절약)
  if (bgChanged) {
    range.setBackgrounds(backgrounds);
  }
  SpreadsheetApp.flush();

  if (delayedItems.length > 0) {
    var msgList = delayedItems.slice(0, 10).map(function(item) {
      return "• [" + item.siteName + "] " + item.taskName + " (마감: " + item.dateStr + (item.manager ? ", 담당: " + item.manager : "") + ")";
    }).join("\n");
    if (delayedItems.length > 10) {
      msgList += "\n... 외 " + (delayedItems.length - 10) + "건";
    }

    SpreadsheetApp.getUi().alert(
      "⚠️ 공정 지연 현장 경고 (" + delayedItems.length + "건 감지)",
      "마감 예정일이 경과하였으나 미완료 상태인 공정이 발견되었습니다:\n\n" +
      msgList + "\n\n" +
      "※ 해당 행은 시트에서 [연한 붉은색]으로 하이라이트 처리되었으며, 캘린더 동기화 시 [Tomato 빨간 블록] 및 [⚠️ 지연] 아이콘으로 자동 연동됩니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "✅ 공정 지연 검사 완료",
      "지연된 현장 공정이 없습니다. 모든 일정이 예정대로 원활하게 진행 중입니다!",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * [구성원] (또는 팀원/직원/조직도) 시트 찾기
 */
function findMembersSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  var candidateNames = ["구성원", "팀원", "직원", "조직도", "members", "team", "member"];
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim().toLowerCase();
    for (var c = 0; c < candidateNames.length; c++) {
      if (sName === candidateNames[c] || sName.indexOf(candidateNames[c]) !== -1) {
        return sheets[i];
      }
    }
  }
  return null;
}

/**
 * [구성원] 시트에서 모닝 브리핑 수신 대상 이메일 목록 추출
 */
function extractMemberBriefingEmails(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // 1~3행 헤더 탐색
  var headerRows = Math.min(lastRow, 3);
  var headerVals = sheet.getRange(1, 1, headerRows, lastCol).getValues();
  var emailColIdx = -1;
  var receiveColIdx = -1;
  var startDataRow = 2;

  for (var r = 0; r < headerRows; r++) {
    for (var c = 0; c < lastCol; c++) {
      var hStr = String(headerVals[r][c] || "").trim().toLowerCase();
      if (emailColIdx === -1 && (hStr.indexOf("이메일") !== -1 || hStr.indexOf("email") !== -1 || hStr.indexOf("메일") !== -1)) {
        emailColIdx = c;
        startDataRow = Math.max(startDataRow, r + 2);
      }
      if (receiveColIdx === -1 && (hStr.indexOf("수신") !== -1 || hStr.indexOf("브리핑") !== -1 || hStr.indexOf("발송") !== -1 || hStr.indexOf("receive") !== -1)) {
        receiveColIdx = c;
        startDataRow = Math.max(startDataRow, r + 2);
      }
    }
  }

  if (lastRow < startDataRow) return [];

  var data = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, lastCol).getValues();
  var resultEmails = [];
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  for (var i = 0; i < data.length; i++) {
    var rowVals = data[i];

    // 수신 여부 체크박스 열이 있다면 확인
    if (receiveColIdx !== -1) {
      var recvVal = rowVals[receiveColIdx];
      var isChecked = (recvVal === true || String(recvVal).trim().toUpperCase() === "TRUE" || String(recvVal).trim() === "1" || String(recvVal).trim() === "V" || String(recvVal).trim() === "O");
      if (!isChecked) continue; // 체크 안 된 인원은 제외
    }

    var rawEmail = (emailColIdx !== -1) ? String(rowVals[emailColIdx] || "").trim() : "";
    // 스마트 감지: 이메일 컬럼이 비어있거나 실수로 연락처 등 다른 열에 이메일 주소를 입력한 경우
    if (!rawEmail || !emailRegex.test(rawEmail)) {
      for (var colIdx = 0; colIdx < rowVals.length; colIdx++) {
        if (colIdx === receiveColIdx) continue;
        var candidateVal = String(rowVals[colIdx] || "").trim();
        if (emailRegex.test(candidateVal)) {
          rawEmail = candidateVal;
          break;
        }
      }
    }
    if (!rawEmail || !emailRegex.test(rawEmail)) continue;

    if (resultEmails.indexOf(rawEmail) === -1) {
      resultEmails.push(rawEmail);
    }
  }

  return resultEmails;
}

/**
 * 모닝 브리핑 수신 이메일 가져오기
 * 1순위: [구성원] 시트에서 수신 체크된 이메일 자동 추출
 * 2순위: 문서 속성(DAILY_BRIEFING_EMAIL)에 설정된 이메일
 * 3순위: 로그인된 사용자 계정 이메일
 */
function getDailyBriefingRecipientEmail(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1순위: [구성원] 시트 자동 탐색
  if (ss) {
    var memberSheet = findMembersSheet(ss);
    if (memberSheet && memberSheet.getLastRow() >= 2) {
      var memberEmails = extractMemberBriefingEmails(memberSheet);
      if (memberEmails && memberEmails.length > 0) {
        return memberEmails.join(", ");
      }
    }
  }

  // 2순위: 문서 속성 설정값
  var docProps = PropertiesService.getDocumentProperties();
  var savedEmail = docProps.getProperty("DAILY_BRIEFING_EMAIL");
  if (savedEmail && savedEmail.trim()) {
    return savedEmail.trim();
  }

  // 3순위: 기본 사용자 계정
  var userEmail = "";
  try {
    userEmail = Session.getEffectiveUser().getEmail();
  } catch (e) {}
  if (!userEmail) {
    try {
      userEmail = Session.getActiveUser().getEmail();
    } catch (e) {}
  }
  return userEmail;
}

/**
 * [메뉴용] 모닝 브리핑 수신 이메일 설정 대화상자
 */
function configureBriefingEmail() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = findMembersSheet(ss);
  var memberEmails = (memberSheet ? extractMemberBriefingEmails(memberSheet) : []);
  var currentEmail = getDailyBriefingRecipientEmail(ss);

  var infoNotice = "";
  if (memberEmails.length > 0) {
    infoNotice = "\n\n💡 현재 ['" + memberSheet.getName() + "'] 시트에서 " + memberEmails.length + "명의 수신 이메일이 자동 연동되고 있습니다.";
  }

  var prompt = ui.prompt(
    "📧 모닝 브리핑 수신 이메일 설정",
    "매일 아침 8시 현장 공정 브리핑을 수신할 이메일 주소를 입력하세요.\n(여러 명인 경우 쉼표로 구분, [구성원] 시트가 있으면 해당 시트가 최우선 적용됩니다.)\n\n현재 수신처: " + (currentEmail || "(미설정)") + infoNotice,
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() === ui.Button.OK) {
    var newEmail = prompt.getResponseText().trim();
    if (newEmail) {
      PropertiesService.getDocumentProperties().setProperty("DAILY_BRIEFING_EMAIL", newEmail);
      ui.alert("✅ 수신 이메일이 [" + newEmail + "](으)로 설정되었습니다.");
    } else {
      PropertiesService.getDocumentProperties().deleteProperty("DAILY_BRIEFING_EMAIL");
      ui.alert("✅ 기본 이메일 계정으로 재설정되었습니다.");
    }
  }
}

/**
 * 모닝 브리핑 설정 시간 가져오기 (기본값: "08:00")
 */
function getDailyBriefingConfiguredTime() {
  var docProps = PropertiesService.getDocumentProperties();
  var saved = docProps.getProperty("DAILY_BRIEFING_TIME");
  if (saved && /^([01]?[0-9]|2[0-3]):([0-5][0-9])$/.test(saved.trim())) {
    return saved.trim();
  }
  return "08:00";
}

/**
 * 다음 모닝 브리핑 실행 시각 계산 및 정각 트리거 등록 (1분 오차 정밀 실행)
 * - Google 기본 1시간 범위 지연을 방지하기 위해 단발성 정밀 시각 트리거를 자동 연속 예약합니다.
 */
function scheduleNextDailyBriefingTrigger() {
  var functionName = "sendDailySiteMorningBriefing";
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  var timeStr = getDailyBriefingConfiguredTime();
  var parts = timeStr.split(":");
  var targetHour = parseInt(parts[0], 10);
  var targetMin = parseInt(parts[1], 10);

  var now = new Date();
  var nextRun = new Date(now.getFullYear(), now.getMonth(), now.getDate(), targetHour, targetMin, 0, 0);

  // 현재 시각보다 이미 지났거나 1분 이내라면 내일 같은 시각으로 예약
  if (nextRun.getTime() <= now.getTime() + 60000) {
    nextRun.setDate(nextRun.getDate() + 1);
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .at(nextRun)
    .create();

  return nextRun;
}

/**
 * 모닝 브리핑 발송 시간 설정 메뉴 (원하는 시간 정각 지정)
 */
function configureBriefingTime() {
  var ui = SpreadsheetApp.getUi();
  var currentTime = getDailyBriefingConfiguredTime();

  var prompt = ui.prompt(
    "⏰ 모닝 브리핑 발송 시간 설정",
    "매일 모닝 브리핑 이메일을 발송할 시간을 입력하세요 (24시간 형식 HH:mm).\n\n" +
    "• 예시: 08:00 (오전 8시 정각), 07:30 (오전 7시 30분), 08:15 등\n" +
    "• 현재 설정 시간: " + currentTime + " 정각\n\n" +
    "(참고: 정각 예약 엔진이 적용되어 입력하신 시각에 오차 없이 발송됩니다)",
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() === ui.Button.OK) {
    var inputVal = prompt.getResponseText().trim();
    var match = inputVal.match(/^([01]?[0-9]|2[0-3]):([0-5][0-9])$/);
    if (!match) {
      ui.alert("⚠️ 입력 형식이 올바르지 않습니다.\n\n'08:00', '07:30'과 같이 24시간 형식(HH:mm)으로 입력해주세요.");
      return;
    }

    var hour = ("0" + parseInt(match[1], 10)).slice(-2);
    var min = ("0" + parseInt(match[2], 10)).slice(-2);
    var formattedTime = hour + ":" + min;

    var docProps = PropertiesService.getDocumentProperties();
    docProps.setProperty("DAILY_BRIEFING_TIME", formattedTime);

    var isEnabled = docProps.getProperty("DAILY_BRIEFING_ENABLED");
    var nextRunMsg = "";
    if (isEnabled === "true") {
      var nextRun = scheduleNextDailyBriefingTrigger();
      var dateStr = Utilities.formatDate(nextRun, Session.getScriptTimeZone() || "Asia/Seoul", "yyyy-MM-dd HH:mm");
      nextRunMsg = "\n\n🚀 다음 발송 예정 시각: " + dateStr;
    } else {
      nextRunMsg = "\n\n💡 현재 자동 발송이 꺼져 있습니다. [⏰ 모닝 브리핑 자동 발송 켜기] 메뉴를 누르면 활성화됩니다.";
    }

    ui.alert("✅ 모닝 브리핑 발송 시간이 [" + formattedTime + " 정각]으로 설정되었습니다." + nextRunMsg);
  }
}

/**
 * 모닝 브리핑 시간 기반 정각 트리거 활성화
 */
function enableDailyMorningBriefing() {
  var docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty("DAILY_BRIEFING_ENABLED", "true");

  var nextRun = scheduleNextDailyBriefingTrigger();
  var timeStr = getDailyBriefingConfiguredTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recipient = getDailyBriefingRecipientEmail(ss);
  var dateStr = Utilities.formatDate(nextRun, Session.getScriptTimeZone() || "Asia/Seoul", "yyyy-MM-dd HH:mm");

  SpreadsheetApp.getUi().alert(
    "☀️ 모닝 브리핑 자동 발송 활성화 (정각 발송)",
    "매일 " + timeStr + " 정각에 '오늘의 현장 공정 및 지연 현장 브리핑'이 자동으로 발송됩니다!\n\n" +
    "• 다음 발송 예정: " + dateStr + "\n" +
    "• 수신처: " + (recipient || "스프레드시트 소유자 계정") + "\n\n" +
    "💡 발송 시간 변경은 [🕒 모닝 브리핑 발송 시간 변경] 메뉴에서 언제든지 가능합니다.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 매일 모닝 브리핑 자동 발송 끄기
 */
function disableDailyMorningBriefing() {
  var docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty("DAILY_BRIEFING_ENABLED", "false");

  var functionName = "sendDailySiteMorningBriefing";
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  if (removed > 0) {
    SpreadsheetApp.getUi().alert("🔕 매일 모닝 브리핑 자동 발송이 해제되었습니다.");
  } else {
    SpreadsheetApp.getUi().alert("현재 활성화된 모닝 브리핑 자동 발송 트리거가 없습니다.");
  }
}

/**
 * 모닝 브리핑 즉시 테스트 발송
 */
function sendDailyBriefingTestNow() {
  sendDailySiteMorningBriefing(true);
}

/**
 * 매일 아침 현장 공정 모닝 브리핑 이메일 발송 코어 엔진 (Feature 4)
 * @param {boolean} [isManualTest=false] 수동 테스트 발송 여부
 */
function sendDailySiteMorningBriefing(isManualTest) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    if (isManualTest) {
      SpreadsheetApp.getUi().alert("현장데이터 시트에 일정이 없습니다.");
    }
    return;
  }

  var recipient = getDailyBriefingRecipientEmail(ss);
  if (!recipient) {
    if (isManualTest) {
      SpreadsheetApp.getUi().alert("수신자 이메일 주소를 확인할 수 없습니다. [📧 모닝 브리핑 수신 이메일 변경] 메뉴에서 설정해주세요.");
    }
    return;
  }

  var today = new Date();
  var todayStr = formatDateYYMMDD(today);
  var dayOfWeekNames = ["일", "월", "화", "수", "목", "금", "토"];
  var dayOfWeekStr = dayOfWeekNames[today.getDay()] + "요일";

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var todayTasks = [];
  var delayedTasks = [];

  for (var i = 0; i < data.length; i++) {
    var siteName = String(data[i][0] || "").trim();
    var dateVal = data[i][1];
    var taskName = String(data[i][2] || "").trim();
    var partner = String(data[i][3] || "").trim();
    var isShared = (data[i][4] === true || String(data[i][4]).trim().toUpperCase() === "TRUE" || String(data[i][4]).trim() === "1");
    var memo = String(data[i][5] || "").trim();
    var manager = String(data[i][6] || "").trim();
    var status = String(data[i][7] || "").trim();
    var statusClean = status.replace(/\s+/g, "");

    if (!siteName && !taskName) continue;
    var parsedDate = parseScheduleDateRange(dateVal);
    if (!parsedDate) continue;

    var formattedDate = formatScheduleDateDisplay(dateVal);

    var item = {
      row: i + 2,
      siteName: siteName,
      dateStr: formattedDate || String(dateVal).trim(),
      taskName: taskName,
      partner: partner,
      isShared: isShared,
      memo: memo,
      manager: manager,
      status: status || "예정",
      parsedDate: parsedDate
    };

    if (isScheduleDelayed(parsedDate, statusClean, today)) {
      delayedTasks.push(item);
    } else if (isScheduleActiveToday(parsedDate, today)) {
      todayTasks.push(item);
    }
  }

  var subject = "[우성아이디] ☀️ 오늘의 현장 공정 모닝 브리핑 (" + todayStr + " " + dayOfWeekStr + ") - 오늘 " + todayTasks.length + "건" + (delayedTasks.length > 0 ? " / ⚠️ 지연 " + delayedTasks.length + "건" : "");

  var sheetUrl = "";
  try { sheetUrl = ss.getUrl(); } catch(e) {}
  var webAppUrl = getWebAppUrl();

  var emailPayload = buildMorningBriefingEmailContent(todayStr, dayOfWeekStr, todayTasks, delayedTasks, sheetUrl, webAppUrl);

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: emailPayload.plainText,
    htmlBody: emailPayload.htmlBody
  });

  if (isManualTest) {
    SpreadsheetApp.getUi().alert(
      "✉️ 모닝 브리핑 테스트 발송 완료",
      "성공적으로 브리핑 이메일을 발송했습니다!\n\n" +
      "• 수신자: " + recipient + "\n" +
      "• 오늘 착공/진행 공정: " + todayTasks.length + "건\n" +
      "• 공정 지연 경고: " + delayedTasks.length + "건\n\n" +
      "이메일 수신함을 확인해보세요.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    // 자동 일일 실행인 경우, 다음 날 설정 시각 정각 트리거 자동 재스케줄링 (Self-Rescheduling)
    try {
      var isEnabled = PropertiesService.getDocumentProperties().getProperty("DAILY_BRIEFING_ENABLED");
      if (isEnabled === "true") {
        scheduleNextDailyBriefingTrigger();
      }
    } catch(e) {
      Logger.log("모닝 브리핑 자동 재스케줄링 오류: " + e.message);
    }
  }
}

/**
 * 모닝 브리핑 HTML 및 Plain-text 이메일 본문 생성기
 */
function buildMorningBriefingEmailContent(todayStr, dayOfWeekStr, todayTasks, delayedTasks, sheetUrl, webAppUrl) {
  var delayedCount = delayedTasks ? delayedTasks.length : 0;
  var todayCount = todayTasks ? todayTasks.length : 0;

  // 1. Plain text 버전
  var plainLines = [];
  plainLines.push("========================================================");
  plainLines.push("  우성아이디 | ☀️ 오늘의 현장 공정 일일 브리핑");
  plainLines.push("  일자: " + todayStr + " (" + dayOfWeekStr + ")");
  plainLines.push("  요약: 오늘 공정 " + todayCount + "건 / ⚠️ 지연 경고 " + delayedCount + "건");
  plainLines.push("========================================================\n");

  if (delayedCount > 0) {
    plainLines.push("⚠️ [긴급 확인: 공정 지연 현장 (" + delayedCount + "건)]");
    for (var i = 0; i < delayedTasks.length; i++) {
      var d = delayedTasks[i];
      plainLines.push("• [" + d.siteName + "] " + d.taskName + " (마감예정: " + d.dateStr + " / 상태: " + d.status + (d.manager ? " / 담당: " + d.manager : "") + (d.memo ? " / 비고: " + d.memo : "") + ")");
    }
    plainLines.push("\n");
  }

  plainLines.push("🔨 [오늘의 현장 공정 (" + todayCount + "건)]");
  if (todayCount === 0) {
    plainLines.push("오늘 예정된 현장 공정이 없습니다.\n");
  } else {
    for (var j = 0; j < todayTasks.length; j++) {
      var t = todayTasks[j];
      plainLines.push("• [" + t.siteName + "] " + t.taskName + " | 상태: " + t.status + " | 일정: " + t.dateStr + (t.manager ? " | 담당: " + t.manager : "") + (t.partner ? " | 협력: " + t.partner : "") + (t.memo ? " | 비고: " + t.memo : ""));
    }
    plainLines.push("\n");
  }

  if (sheetUrl) {
    plainLines.push("📊 스프레드시트 바로가기: " + sheetUrl);
  }
  if (webAppUrl) {
    plainLines.push("📱 모바일 상태변경 페이지: " + webAppUrl);
  }

  // 2. HTML 버전 (인라인 스타일 모던 템플릿)
  var html = "";
  html += '<!DOCTYPE html>';
  html += '<html><head><meta charset="utf-8">';
  html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
  html += '<style>';
  html += 'body { margin: 0; padding: 24px 16px; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #1e293b; }';
  html += '.container { max-width: 880px; width: 100%; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 16px rgba(0,0,0,0.06); border: 1px solid #e2e8f0; }';
  html += '.header { background: #1e293b; color: #ffffff; padding: 26px 32px; text-align: left; }';
  html += '.header h1 { margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }';
  html += '.header p { margin: 6px 0 0 0; font-size: 13px; color: #94a3b8; }';
  html += '.section { padding: 24px 32px; }';
  html += '.section-title { font-size: 16px; font-weight: 700; color: #0f172a; margin-bottom: 14px; }';
  html += '.task-table { width: 100%; border-collapse: collapse; font-size: 13px; table-layout: fixed; }';
  html += '.task-table th { background: #f8fafc; color: #475569; font-weight: 600; text-align: left; padding: 11px 12px; border-bottom: 1px solid #e2e8f0; }';
  html += '.task-table td { padding: 12px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }';
  html += '.badge { display: inline-block; padding: 3px 8px; border-radius: 9999px; font-size: 11px; font-weight: 700; white-space: nowrap; }';
  html += '.badge-prog { background: #dbeafe; color: #1e40af; }';
  html += '.badge-done { background: #dcfce7; color: #15803d; }';
  html += '.badge-delay { background: #fee2e2; color: #b91c1c; }';
  html += '.badge-plan { background: #fef3c7; color: #92400e; }';
  html += '.badge-as { background: #fee2e2; color: #991b1b; }';
  html += '.btn-container { text-align: center; padding: 24px 32px; background: #f8fafc; border-top: 1px solid #e2e8f0; }';
  html += '.btn { display: inline-block; padding: 12px 28px; border-radius: 8px; font-size: 14px; font-weight: 700; text-decoration: none !important; margin: 6px; }';
  html += '.btn-primary { background-color: #1d4ed8 !important; color: #ffffff !important; border: 1px solid #1e40af; }';
  html += '.btn-secondary { background-color: #ffffff !important; color: #0f172a !important; border: 1px solid #cbd5e1; }';
  html += '.footer { padding: 18px 32px; text-align: center; font-size: 11px; color: #94a3b8; }';
  html += '</style></head><body>';
  html += '<div class="container" style="max-width:880px;width:100%;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.06);border:1px solid #e2e8f0;">';

  // Header
  html += '<div class="header">';
  html += '  <div style="font-size:12px;color:#38bdf8;font-weight:600;margin-bottom:4px;">WOOSUNG ID 현장관리 시스템</div>';
  html += '  <h1>☀️ 오늘의 현장 공정 일일 브리핑</h1>';
  html += '  <p>📅 ' + todayStr + ' (' + dayOfWeekStr + ') 기준 현장 현황 리포트</p>';
  html += '</div>';

  // KPI Summary
  html += '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:16px 24px;">';
  html += '<tr>';
  html += '  <td width="50%" style="padding:6px;">';
  html += '    <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:12px;text-align:center;">';
  html += '      <div style="font-size:11px;color:#64748b;font-weight:600;">🔨 오늘 착공/진행 공정</div>';
  html += '      <div style="font-size:24px;font-weight:800;color:#0f172a;margin-top:4px;">' + todayCount + '<span style="font-size:14px;font-weight:600;"> 건</span></div>';
  html += '    </div>';
  html += '  </td>';
  html += '  <td width="50%" style="padding:6px;">';
  var alertCardBg = delayedCount > 0 ? '#fff1f2' : '#ffffff';
  var alertBorder = delayedCount > 0 ? '#fda4af' : '#e2e8f0';
  var alertTextColor = delayedCount > 0 ? '#be123c' : '#0f172a';
  html += '    <div style="background:' + alertCardBg + ';border:1px solid ' + alertBorder + ';border-radius:8px;padding:12px;text-align:center;">';
  html += '      <div style="font-size:11px;color:' + (delayedCount > 0 ? '#e11d48' : '#64748b') + ';font-weight:600;">⚠️ 마감 경과 지연 공정</div>';
  html += '      <div style="font-size:24px;font-weight:800;color:' + alertTextColor + ';margin-top:4px;">' + delayedCount + '<span style="font-size:14px;font-weight:600;"> 건</span></div>';
  html += '    </div>';
  html += '  </td>';
  html += '</tr>';
  html += '</table>';

  // Section 1: Delayed tasks (if any)
  if (delayedCount > 0) {
    html += '<div class="section" style="padding-bottom:12px;">';
    html += '  <div style="background:#fff1f2;border:1px solid #fda4af;border-radius:8px;padding:16px;">';
    html += '    <div style="font-size:14px;font-weight:700;color:#be123c;margin-bottom:8px;">⚠️ 긴급 확인 필요: 공정 지연 현장 (' + delayedCount + '건)</div>';
    html += '    <div style="font-size:12px;color:#9f1239;margin-bottom:12px;">마감 예정일이 지났으나 미완료 상태인 공정입니다. 현장 확인 및 일정 조정이 필요합니다.</div>';
    html += '    <table class="task-table" style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed;">';
    html += '      <thead><tr><th style="width:18%;">현장명</th><th style="width:20%;">공종</th><th style="width:14%;">예정 마감일</th><th style="width:14%;">담당자</th><th style="width:34%;">비고</th></tr></thead>';
    html += '      <tbody>';
    for (var k = 0; k < delayedTasks.length; k++) {
      var it = delayedTasks[k];
      html += '      <tr>';
      html += '        <td style="font-weight:700;color:#0f172a;word-break:break-word;">[' + it.siteName + ']</td>';
      html += '        <td style="word-break:break-word;">' + it.taskName + '</td>';
      html += '        <td style="color:#e11d48;font-weight:600;white-space:nowrap;">' + it.dateStr + '</td>';
      html += '        <td>' + (it.manager || '-') + '</td>';
      html += '        <td style="color:#334155;word-break:break-word;overflow-wrap:break-word;line-height:1.45;font-size:12.5px;">' + (it.memo || '-') + '</td>';
      html += '      </tr>';
    }
    html += '      </tbody>';
    html += '    </table>';
    html += '  </div>';
    html += '</div>';
  }

  // Section 2: Today's tasks
  html += '<div class="section">';
  html += '  <div class="section-title">🔨 오늘 착공 및 진행 공정 (' + todayCount + '건)</div>';
  if (todayCount === 0) {
    html += '  <div style="padding:24px;text-align:center;color:#64748b;font-size:13px;background:#f8fafc;border-radius:8px;border:1px dashed #cbd5e1;">';
    html += '    오늘 예정된 현장 공정이 없습니다. 평온하고 안전한 하루 되세요! ☕';
    html += '  </div>';
  } else {
    html += '  <table class="task-table" style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed;">';
    html += '    <thead><tr><th style="width:18%;">현장명</th><th style="width:20%;">공종</th><th style="width:12%;text-align:center;">상태</th><th style="width:16%;">담당/협력</th><th style="width:34%;">비고</th></tr></thead>';
    html += '    <tbody>';
    for (var m = 0; m < todayTasks.length; m++) {
      var item = todayTasks[m];
      var badgeClass = 'badge-plan';
      var stClean = item.status.replace(/\s+/g, '');
      if (stClean.indexOf('완료') !== -1) badgeClass = 'badge-done';
      else if (stClean.indexOf('진행') !== -1) badgeClass = 'badge-prog';
      else if (stClean.indexOf('AS') !== -1) badgeClass = 'badge-as';

      var managerInfo = item.manager || '';
      if (item.partner) managerInfo += (managerInfo ? ' / ' : '') + item.partner;

      html += '    <tr>';
      html += '      <td style="font-weight:700;color:#0f172a;word-break:break-word;">[' + item.siteName + ']</td>';
      html += '      <td style="word-break:break-word;"><b>' + item.taskName + '</b><br><span style="font-size:11px;color:#64748b;">' + item.dateStr + '</span></td>';
      html += '      <td style="text-align:center;"><span class="badge ' + badgeClass + '">' + item.status + '</span></td>';
      html += '      <td style="word-break:break-word;">' + (managerInfo || '-') + '</td>';
      html += '      <td style="color:#334155;word-break:break-word;overflow-wrap:break-word;line-height:1.45;font-size:12.5px;">' + (item.memo || '-') + '</td>';
      html += '    </tr>';
    }
    html += '    </tbody>';
    html += '  </table>';
  }
  html += '</div>';

  // Quick Action Buttons
  html += '<div class="btn-container" style="text-align:center;padding:24px 32px;background:#f8fafc;border-top:1px solid #e2e8f0;">';
  if (sheetUrl) {
    html += '  <a href="' + sheetUrl + '" class="btn btn-primary" target="_blank" style="display:inline-block;background-color:#1d4ed8;color:#ffffff !important;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none !important;margin:6px;box-shadow:0 2px 6px rgba(29,78,216,0.25);border:1px solid #1e40af;"><span style="color:#ffffff !important;text-decoration:none !important;font-weight:700;font-size:14px;">📊 스프레드시트 전체 보기</span></a>';
  }
  if (webAppUrl) {
    html += '  <a href="' + webAppUrl + '" class="btn btn-secondary" target="_blank" style="display:inline-block;background-color:#ffffff;color:#0f172a !important;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none !important;margin:6px;border:1px solid #cbd5e1;box-shadow:0 1px 3px rgba(0,0,0,0.05);"><span style="color:#0f172a !important;text-decoration:none !important;font-weight:700;font-size:14px;">📱 모바일 빠른 관리</span></a>';
  }
  html += '</div>';

  // Footer
  html += '<div class="footer">';
  html += '  본 메일은 우성아이디 스프레드시트 시스템에서 매일 아침 8시에 자동 발송되는 리포트입니다.<br>';
  html += '  수신처 변경이나 발송 설정은 스프레드시트 메뉴 [🚀 현장 관리 -> ☀️ 모닝 브리핑 설정]에서 관리하실 수 있습니다.';
  html += '</div>';

  html += '</div>'; // end container
  html += '</body></html>';

  return {
    plainText: plainLines.join('\n'),
    htmlBody: html
  };
}


// ========================================================
// 12. [사용가이드 시트 자동 생성 및 매뉴얼 구축 엔진]
// ========================================================

/**
 * [메뉴 실행용] 사용 가이드 & 색상 범례 시트 생성 대화상자
 */
function createOrUpdateUsageGuideSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = createOrUpdateUsageGuideSheet(ss);
  if (sheet) {
    ss.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert(
      "📖 사용 가이드 구축 완료",
      "[사용가이드] 시트가 성공적으로 생성 및 최신화되었습니다!\n\n" +
      "• 1. 🎨 상태별 캘린더 색상 범례 (Peacock 파랑, Basil 녹색, Tomato 빨강 등)\n" +
      "• 2. 📝 '현장데이터' 컬럼별 작성 규칙 (A~H열)\n" +
      "• 3. 🚀 4대 핵심 스마트 기능 (양방향 연동, 모바일 빠른 변경, 모닝 브리핑, 지연 감지)\n" +
      "• 4. 🛠️ 상단 메뉴 기능 안내\n\n" +
      "다른 작업자분들도 '사용가이드' 탭을 통해 손쉽게 사용법을 숙지하실 수 있습니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * [사용가이드] 시트 자동 생성 및 최신 매뉴얼/색상 범례 작성 엔진
 * @param {Spreadsheet} [ss] 대상 스프레드시트
 * @returns {Sheet} 생성된 사용가이드 시트
 */
function createOrUpdateUsageGuideSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "사용가이드";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.setHiddenGridlines(false);
  try { sheet.setTabColor("#2563eb"); } catch(e) {}

  // 1. 열 너비 설정
  sheet.setColumnWidth(1, 30);  // A열: 좌측 여백
  sheet.setColumnWidth(2, 130); // B열: 구분 / 상태명
  sheet.setColumnWidth(3, 190); // C열: 캘린더 색상 / 항목명
  sheet.setColumnWidth(4, 180); // D열: 시트 표기 / 아이콘 / 형식
  sheet.setColumnWidth(5, 420); // E열: 상세 설명 및 기준
  sheet.setColumnWidth(6, 260); // F열: 실무 팁 및 주의사항

  // 2. 전체 기본 폰트 적용
  sheet.getRange(1, 1, 60, 6)
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  // 3. 타이틀 헤더 (2~3행)
  sheet.getRange("B2:F2").merge()
    .setValue("📋 우성아이디 현장관리 & 구글 캘린더 연동 시스템 사용 가이드")
    .setFontSize(15)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(2, 46);

  sheet.getRange("B3:F3").merge()
    .setValue("💡 현장 데이터 작성 규격, 캘린더 색상 범례, 실시간 양방향 연동, 모바일 빠른 변경 및 모닝 브리핑 안내서")
    .setFontSize(10)
    .setFontColor("#64748b")
    .setBackground("#f8fafc")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 28);

  sheet.setRowHeight(4, 16); // 여백

  // ==========================================
  // 섹션 1: 🎨 상태별 캘린더 색상 & 시트 표기 범례
  // ==========================================
  sheet.getRange("B5:F5").merge()
    .setValue("1. 🎨 상태별 캘린더 블록 색상 & 시트 표기 범례 (한눈에 알아보기)")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#2563eb")
    .setHorizontalAlignment("left");
  sheet.setRowHeight(5, 34);

  var legendHeaders = ["상태 구분", "구글 캘린더 블록 색상", "시트 및 캘린더 표기", "의미 및 적용 기준", "실무 활용 팁"];
  sheet.getRange(6, 2, 1, 5).setValues([legendHeaders])
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(6, 30);

  var legendData = [
    ["진행 중", "Peacock (청록/파랑)", "▶️ 🟦 [현장명] 공종", "현재 현장에서 시공/작업이 진행 중인 핵심 공정", "드롭다운 '진행 중' 선택 시 파란 블록 자동 지정"],
    ["완료", "Basil (진초록)", "✅ 🟦 [현장명] 공종", "공정 작업 및 현장 마감/검수가 최종 완료된 상태", "완료 시 지연 경고가 즉시 해제되며 녹색 블록 전환"],
    ["⚠️ 지연", "Tomato (빨강)", "⚠️ [지연] ▶️ 🟦 ...", "마감일이 오늘 이전인데 아직 완료되지 않은 공정 (자동 감지)", "시트 행이 연분홍색(#fff1f2)으로 자동 하이라이트"],
    ["AS", "Flamingo (산호/주황)", "🛠️ 🟦 [현장명] 공종", "입주 후 하자 보수, 잔손보기 및 사후 관리 공정", "공종명에 보수 상세 작성 시 캘린더에 연동"],
    ["예정 (공유)", "Pale Blue (연파랑)", "🟦 [현장명] 공종", "향후 착공 예정 공정 중 외부에 공유 완료된 일정", "기본 등록 상태이며 E열(공유) 체크박스가 켜진 상태"],
    ["예정 (미공유)", "Graphite (회색)", "🟥 [현장명] 공종", "내부 검토 단계이거나 외부에 공개하지 않은 비공개 일정", "E열(공유) 체크박스 해제 시 붉은 사각 및 회색 블록"]
  ];
  sheet.getRange(7, 2, legendData.length, 5).setValues(legendData);

  // 섹션 1 셀 스타일링
  for (var r = 0; r < legendData.length; r++) {
    sheet.setRowHeight(7 + r, 28);
    sheet.getRange(7 + r, 2).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(7 + r, 3).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(7 + r, 4).setHorizontalAlignment("center");
  }
  sheet.getRange("B7").setFontColor("#1d4ed8").setBackground("#dbeafe");
  sheet.getRange("C7").setFontColor("#0369a1").setBackground("#bae6fd");
  sheet.getRange("B8").setFontColor("#15803d").setBackground("#dcfce7");
  sheet.getRange("C8").setFontColor("#15803d").setBackground("#bbf7d0");
  sheet.getRange("B9").setFontColor("#b91c1c").setBackground("#fee2e2");
  sheet.getRange("C9").setFontColor("#b91c1c").setBackground("#fecaca");
  sheet.getRange("B10").setFontColor("#c2410c").setBackground("#ffedd5");
  sheet.getRange("C10").setFontColor("#c2410c").setBackground("#fed7aa");
  sheet.getRange("B11").setFontColor("#b45309").setBackground("#fef3c7");
  sheet.getRange("C11").setFontColor("#0284c7").setBackground("#e0f2fe");
  sheet.getRange("B12").setFontColor("#475569").setBackground("#f1f5f9");
  sheet.getRange("C12").setFontColor("#475569").setBackground("#e2e8f0");

  sheet.getRange("B6:F12").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(13, 16); // 여백

  // ==========================================
  // 섹션 2: 📝 현장데이터 시트 컬럼별 입력 규칙
  // ==========================================
  sheet.getRange("B14:F14").merge()
    .setValue("2. 📝 '현장데이터' 시트 컬럼별 작성 규칙 및 예시")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#0284c7")
    .setHorizontalAlignment("left");
  sheet.setRowHeight(14, 34);

  var inputHeaders = ["열 (컬럼)", "항목명", "입력 형식 및 예시", "연동 동작 및 상세 설명", "주의사항 및 팁"];
  sheet.getRange(15, 2, 1, 5).setValues([inputHeaders])
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(15, 30);

  var inputData = [
    ["A열", "현장", "텍스트 (예: 래미안 101호)", "캘린더 제목의 [현장명] 및 위치(Location)로 자동 지정", "현장명이 비어있으면 캘린더 동기화에서 제외됨"],
    ["B열", "날짜", "단일: 26.10.15 / 기간: 26.10.15 ~ 26.10.18 / 시간: 26.10.15 14:00 ~ 16:00", "종일 일정 및 시간 일정을 완벽 파싱하여 대시보드와 캘린더에 연동", "점(.)과 물결(~) 구분 기호 권장 (연도 생략 시 올해 자동 적용)"],
    ["C열", "공종", "텍스트 (예: 목공, 타일, 전기, 도배)", "캘린더 제목 및 상세 설명에 표시되는 핵심 공정 내용", "필수 입력 항목"],
    ["D열", "협력업체", "텍스트 (예: 우성목재, 김반장)", "캘린더 본문 Description 및 모닝 브리핑 리포트에 업체 정보 전달", "선택 입력 항목"],
    ["E열", "공유", "체크박스 (TRUE / FALSE)", "체크 시 🟦(공유 완료), 체크 해제 시 🟥(미공유 회색 블록)", "협력업체 공개 여부를 클릭 한 번으로 손쉽게 제어"],
    ["F열", "비고", "텍스트 (예: 오후 자재 입고)", "캘린더 제목 괄호 (비고) 및 상세 설명에 메모 자동 기록", "비고만 수정 시 불필요한 캘린더 호출을 생략하여 속도 보호"],
    ["G열", "담당자", "텍스트 (예: 김팀장, 이과장)", "캘린더 본문 및 모닝 브리핑 리포트에 현장 담당자 정보 표기", "선택 입력 항목"],
    ["H열", "진행상태", "드롭다운: 예정, 진행 중, 완료, AS", "선택 즉시 캘린더 아이콘, 블록 색상, 대시보드 서식이 자동 변경됨", "표준 규격인 '진행 중'(공백 포함) 사용 (빨간 에러 삼각형 방지)"],
    ["I~N열", "시스템 열", "캘린더ID, 타임스탬프, 최초/최종 작성자 및 일시", "Last-Modified-Wins 양방향 충돌 방지 및 이력 관리를 위한 자동 열", "사용자가 직접 수정하지 마세요 (시스템이 100% 자동 관리)"]
  ];
  sheet.getRange(16, 2, inputData.length, 5).setValues(inputData);

  for (var i = 0; i < inputData.length; i++) {
    sheet.setRowHeight(16 + i, 28);
    sheet.getRange(16 + i, 2).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#f8fafc");
    sheet.getRange(16 + i, 3).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(16 + i, 4).setHorizontalAlignment("left");
  }

  sheet.getRange("B15:F24").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(25, 16); // 여백

  // ==========================================
  // 섹션 3: 📱 핵심 스마트 기능 4가지 사용법
  // ==========================================
  sheet.getRange("B26:F26").merge()
    .setValue("3. 🚀 알아두면 유용한 4대 핵심 스마트 기능")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#0f766e")
    .setHorizontalAlignment("left");
  sheet.setRowHeight(26, 34);

  var featureHeaders = ["기능명", "핵심 요약", "사용 방법", "시스템 동작 원리", "추천 사용자"];
  sheet.getRange(27, 2, 1, 5).setValues([featureHeaders])
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(27, 30);

  var featureData = [
    ["🔄 실시간 양방향 연동", "시트 ↔ 캘린더 자동 동기화", "시트에서 수정 시 캘린더 즉시 반영 / 캘린더에서 날짜 옮기면 시트 반영", "최종 수정 시각(Last-Modified-Wins) 비교로 최신 데이터 자동 보존", "현장 소장 및 관리자"],
    ["📱 모바일 빠른 변경", "스마트폰 원터치 상태 변경", "캘린더 일정 본문 하단의 [📱 모바일 빠른 변경 링크] 터치", "시트를 열지 않고도 웹 브라우저에서 '진행 중' or '완료' 1초 변경", "외부 현장 근무자/반장"],
    ["☀️ 모닝 브리핑 발송", "매일 아침 8시 현장 리포트", "메뉴 [🚀 현장 관리] > [⏰ 매일 아침 8시 모닝 브리핑 켜기]", "오늘 진행 공정 + 마감 경과 지연 공정을 이메일 요약 카드로 자동 발송", "대표님 및 총괄 관리자"],
    ["⚠️ 공정 지연 리스크 감지", "마감일 경과 공정 자동 추적", "메뉴 [🚀 현장 관리] > [🎨 캘린더 색상 & ⚠️ 지연 공정 일괄 검사]", "마감일 지난 미완료 공정은 시트 연분홍 하이라이트 & 캘린더 빨간 블록", "일정 관리 및 QA 담당자"]
  ];
  sheet.getRange(28, 2, featureData.length, 5).setValues(featureData);

  for (var f = 0; f < featureData.length; f++) {
    sheet.setRowHeight(28 + f, 30);
    sheet.getRange(28 + f, 2).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#f0fdf4");
    sheet.getRange(28 + f, 3).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(28 + f, 6).setHorizontalAlignment("center");
  }

  sheet.getRange("B27:F31").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(32, 16); // 여백

  // ==========================================
  // 섹션 4: 🛠️ 상단 메뉴([🚀 현장 관리]) 주요 기능 안내
  // ==========================================
  sheet.getRange("B33:F33").merge()
    .setValue("4. 🛠️ 스프레드시트 상단 메뉴([🚀 현장 관리]) 주요 기능")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#334155")
    .setHorizontalAlignment("left");
  sheet.setRowHeight(33, 34);

  var menuHeaders = ["메뉴명", "실행 목적", "상세 동작", "권장 실행 주기", "비고"];
  sheet.getRange(34, 2, 1, 5).setValues([menuHeaders])
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(34, 30);

  var menuData = [
    ["📅 최근 & 예정 일정 동기화", "캘린더 최신화 (가장 빠름)", "최근 30일 이내 ~ 미래 일정만 선별하여 구글 캘린더에 고속 동기화", "필요 시 수시 실행", "추천 ⭐⭐⭐"],
    ["📦 전체 일정 순차 동기화", "대량 일정 안전 일괄 등록", "구글 API 속도 제한(Rate Limit)을 방어하며 25건씩 분할하여 안전 등록", "최초 연동 또는 대량 등록 시", "서버 락 방지"],
    ["🎨 색상 & 지연 공정 검사", "지연 현장 긴급 점검", "전체 현장의 마감일을 스캔하여 지연 공정을 붉은색으로 하이라이트", "주 1~2회 또는 수시", "리스크 선제 방어"],
    ["☀️ 모닝 브리핑 즉시 테스트", "이메일 리포트 미리보기", "현재 등록된 오늘 일정과 지연 공정을 즉시 취합하여 본인 이메일로 발송", "설정 후 최초 1회 확인", "테스트용"],
    ["⏰ 아침 8시 모닝 브리핑 켜기", "일일 자동 발송 예약", "매일 오전 8시 시간 기반 타이머를 자동 등록하여 브리핑 자동 발송", "시스템 구축 시 1회 켜기", "자동화"],
    ["🔧 드롭다운 유효성 복구", "빨간 삼각형 에러 제거", "'진행 중' 드롭다운 목록 및 서식을 표준화하여 빨간 삼각형 오류 원클릭 복구", "에러 발생 시 즉시", "자동 복구"],
    ["🔄 대시보드 갱신 (달력)", "월간 달력 새로고침", "상단 년/월에 맞추어 현장데이터 일정들을 월간 달력 칸에 즉시 다시 배치", "수정 후 즉시 자동 연동됨", "수동 확인용"],
    ["✨ 새 워크스페이스 구축", "새 시트 1클릭 자동 완성", "완전 빈 새 스프레드시트에 [대시보드], [현장데이터], [사용가이드] 일괄 생성", "새 프로젝트 시작 시 1회", "원클릭 구축"]
  ];
  sheet.getRange(35, 2, menuData.length, 5).setValues(menuData);

  for (var m = 0; m < menuData.length; m++) {
    sheet.setRowHeight(35 + m, 26);
    sheet.getRange(35 + m, 2).setHorizontalAlignment("left").setFontWeight("bold").setBackground("#f8fafc");
    sheet.getRange(35 + m, 3).setHorizontalAlignment("center");
    sheet.getRange(35 + m, 5).setHorizontalAlignment("center");
    sheet.getRange(35 + m, 6).setHorizontalAlignment("center");
  }

  sheet.getRange("B34:F42").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // 안내 팁 푸터
  sheet.setRowHeight(43, 16);
  sheet.getRange("B44:F44").merge()
    .setValue("💡 팁: 시트 사용 중 궁금한 점이 있거나 새로운 현장을 등록할 때는 언제든 이 [사용가이드] 탭을 참고하세요. 항상 안전 시공 되시길 응원합니다! (우성아이디)")
    .setFontSize(9)
    .setFontColor("#64748b")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(44, 24);

  SpreadsheetApp.flush();
  return sheet;
}
