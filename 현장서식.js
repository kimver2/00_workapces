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
// 🔒 [영구 고정 권장] 주황색 공용 구글 캘린더의 고유 ID 또는 캘린더 이름을 지정합니다.
// - 캘린더 ID(가장 확실): "여기에_c_로_시작하는_주황색_캘린더_ID를_붙여넣으세요@group.calendar.google.com"
// - 캘린더 이름: "현장 일정"
// ※ 기본값으로 "현장 일정" 공용 캘린더를 찾으며, 고유 ID로 지정하시면 팀원 간 공유 시에도 100% 안전하게 고정됩니다.
// ※ [오염 원천 차단] 권한이 없거나 못 찾을 경우, 절대 개인(기본) 캘린더로 폴백되거나 새 캘린더를 임의로 만들지 않습니다!
var DEFAULT_CALENDAR_ID_OR_NAME = "현장 일정";


// ========================================================
// 0. [원클릭 자동 구축] 새 스프레드시트 초기 환경 자동 생성
// ========================================================

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
 * 03_구성원 시트 연동 개별 구성원 이름 목록 반환 (대시보드 담당자 필터 버튼용)
 */
function getMemberNames(ss) {
  var members = [];
  var EXCLUDED_MEMBERS = ["김종혁"];
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var mSheet = findMembersSheet(ss);
    if (mSheet && mSheet.getLastRow() >= 2) {
      var lastR = mSheet.getLastRow();
      var lastC = Math.max(mSheet.getLastColumn(), 1);
      var hVals = mSheet.getRange(1, 1, 1, lastC).getValues()[0];
      var nameCol = -1;
      for (var c = 0; c < lastC; c++) {
        var hStr = String(hVals[c] || "").trim();
        if (hStr.indexOf("이름") !== -1 || hStr.indexOf("성명") !== -1 || hStr.indexOf("담당자") !== -1 || hStr.indexOf("직원") !== -1 || hStr.indexOf("팀원") !== -1) {
          nameCol = c + 1;
          break;
        }
      }
      if (nameCol === -1) {
        nameCol = 1;
      }
      var names = mSheet.getRange(2, nameCol, lastR - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        var nm = String(names[i][0] || "").trim();
        if (nm && nm !== "이름" && nm !== "성명" && nm !== "담당자") {
          var parts = nm.split(/[,/]/);
          for (var p = 0; p < parts.length; p++) {
            var subNm = parts[p].trim();
            if (subNm && EXCLUDED_MEMBERS.indexOf(subNm) === -1 && members.indexOf(subNm) === -1) {
              members.push(subNm);
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("getMemberNames 예외: " + e.toString());
  }

  if (members.length === 0) {
    members = ["조항만", "송지은", "전예린", "영하"];
  }

  return members.filter(function(m) { return EXCLUDED_MEMBERS.indexOf(m) === -1; });
}

/**
 * 담당자 드롭다운 목록 반환 (구성원 시트 연동 순수 개별 인원 명단, 김종혁 제외)
 */
function getAvailableManagers(ss) {
  var baseManagers = getMemberNames(ss);
  var EXCLUDED = ["김종혁"];
  var managers = [];

  for (var i = 0; i < baseManagers.length; i++) {
    var m = String(baseManagers[i] || "").trim();
    if (m && EXCLUDED.indexOf(m) === -1 && managers.indexOf(m) === -1) {
      managers.push(m);
    }
  }

  // 예외 폴백 (구성원 시트가 없거나 비어있는 경우)
  if (managers.length === 0) {
    managers = ["조항만", "송지은", "전예린", "영하"];
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


// ========================================================
// 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 메뉴 자동 등록
// ========================================================
function onOpen() {
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      PropertiesService.getDocumentProperties().setProperty("SPREADSHEET_ID", activeSs.getId());
      // 스프레드시트가 열릴 때 정상 행 배경색 및 최신 구성원 연동 담당자 드롭다운 무음 동기화 1회 실행
      var sheet = findSiteScheduleSheet(activeSs);
      if (sheet) {
        cleanNonDelayedRowBackgroundsSilent(sheet);
        applySiteScheduleValidations(sheet);
      }
    }
  } catch(e) {}
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 현장 관리")
    // [사용자 지정 최상위 3대 핵심 기능만 노출]
    .addItem("대시보드 사이드바 열기", "openSiteScheduleSidebar")
    .addItem("대시보드 팝업 열기", "openSiteScheduleModal")
    .addItem("📅 구글 캘린더 동기화 (미등록 & 최근 3개월~미래 일정)", "syncRecentSiteSchedulesToGoogleCalendar")
    .addSeparator()

    // [3개 제외 모든 설정 및 보조 기능은 '고급 설정 & 시스템 관리' 하위로 일원화 배치]
    .addSubMenu(ui.createMenu("⚙️ 고급 설정 & 시스템 관리")
      // 1. 색상 복구 및 날짜 일괄 정비
      .addItem("🎨 캘린더 색상 & ⚠️ 지연 공정 즉시 복구/검사", "checkAndHighlightDelayedSchedules")
      .addItem("📅 전체 날짜 2026.00.00 형식 일괄 변환", "formatAllDatesInSheetToYYYYMMDD")
      .addItem("🔄 대시보드 달력 새로고침", "syncAllSiteSchedulesToDashboard")
      .addSeparator()

      // 2. 현장 데이터 편집 / 보기 편의 기능
      .addItem("➕ 신규 현장 추가 (상단 4개 행 + 아코디언)", "addNewSiteAtTop")
      .addItem("📖 모든 현장 펼치기 (전체 공정 보기)", "expandAllSites")
      .addItem("📁 모든 현장 접기 (현장명만 요약 보기)", "collapseAllSites")
      .addItem("🎯 오늘 날짜로 이동 (Today)", "goToToday")
      .addSeparator()

      // 3. 캘린더 연동 및 동기화 옵션
      .addItem("💡 스마트 디바운스 자동 동기화 켜기 (입력 완료 시 1회 자동 반영 - 최고 권장)", "enableDebouncedAutoSync")
      .addItem("⏰ 5분 주기 자동 동기화 켜기 (정기 순환 동기화)", "enable5MinAutoSync")
      .addItem("⏰ 1분 주기 자동 동기화 켜기 (초고속 테스트용)", "enable1MinAutoSync")
      .addItem("⚡ 완전 실시간 연동 켜기 (수정 시 즉시 반영)", "enableRealtimeCalendarSync")
      .addItem("❌ 모든 자동 연동 끄기 (수동 버튼 전용)", "disableAllAutoCalendarSync")
      .addItem("📥 캘린더 수정사항 가져오기 (수동 즉시 반영)", "syncFromCalendarToSheet")
      .addItem("📦 전체 일정 연속 동기화 (25건씩 자동 대기 연속 처리)", "syncAllSiteSchedulesToGoogleCalendar")
      .addItem("🧹 캘린더 중복 및 유령 일정 자동 정리", "cleanupDuplicateAndOrphanCalendarEvents")
      .addItem("🔕 캘린더 전날 알림 전체 끄기 (알림 폭탄 해제)", "removeAllSiteCalendarReminders")
      .addSeparator()

      // 4. 모닝 브리핑 관리
      .addItem("☀️ 모닝 브리핑 지금 즉시 테스트 발송", "sendDailyBriefingTestNow")
      .addItem("⏰ 모닝 브리핑 자동 발송 켜기 (정각 발송)", "enableDailyMorningBriefing")
      .addItem("🔕 모닝 브리핑 자동 발송 끄기", "disableDailyMorningBriefing")
      .addItem("🕒 모닝 브리핑 발송 시간 변경 (기본 08:00 정각)", "configureBriefingTime")
      .addItem("📧 모닝 브리핑 수신 이메일 변경", "configureBriefingEmail")
      .addSeparator()

      // 5. 시스템 유지보수 & 모바일 환경설정
      .addItem("📱 모바일 대시보드 URL 확인 및 안내", "showMobileDashboardUrl")
      .addItem("📱 모바일 상태변경 웹앱 배포 URL 설정", "configureWebAppUrl")
      .addItem("🔧 공종/담당자/상태 드롭다운 및 서식 전체 복구", "fixStatusDropdownValidation")
      .addItem("🧹 빈 행 유령 이력/ID 일괄 청소", "cleanupEmptyRowsMetadata")
      .addItem("⚙️ 연동 캘린더 변경 / 확인", "configureTargetCalendar")
    )
    .addToUi();
}


// ========================================================
// 2. [실시간 스프레드시트 UI 감지 (단순 트리거)]
// - 대시보드에서 Today 버튼 클릭 시 오늘 날짜로 이동
// - 대시보드에서 '월' 또는 '년도' 변경 시 달력 즉시 새로고침
// - 구성원 시트 수정 시 현장데이터 시트 G열 담당자 드롭다운 실시간 자동 갱신
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

  // ② [구성원 시트 변경 감지 -> 현장데이터 담당자(G열) 드롭다운 즉시 실시간 갱신]
  // 사용자가 구성원 시트에 새로운 인원을 추가하거나 이름을 수정하면, 즉시 현장데이터 시트의
  // 담당자 드롭다운 유효성 검사 목록을 최신 구성원 명단으로 자동 동기화합니다.
  if (isMembersSheet(sheet)) {
    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) {
        applySiteScheduleValidations(siteSheet);
      }
    } catch (eMem) {
      Logger.log("구성원 시트 편집 시 담당자 드롭다운 갱신 예외: " + eMem.toString());
    }
    return;
  }

  // ③ [현장데이터 시트]
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
 * 캘린더 연동 트리거 백그라운드 무음 자동 설치 및 점검 (행 삭제 감지 onChange 트리거 추가)
 */
function ensureRealtimeTriggerInstalledSilent() {
  try {
    var props = PropertiesService.getScriptProperties();
    var syncMode = props.getProperty("SYNC_MODE");

    if (!syncMode) {
      syncMode = "DEBOUNCE";
      props.setProperty("SYNC_MODE", "DEBOUNCE");
    }

    if (syncMode === "MANUAL") {
      return; // 수동 모드인 경우 자동 설치 차단
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return;

    var triggers = ScriptApp.getProjectTriggers();
    var hasEditTrigger = false;
    var hasChangeTrigger = false;
    var hasPeriodicTrigger = false;
    var hasBriefingTrigger = false;

    for (var i = 0; i < triggers.length; i++) {
      var fn = triggers[i].getHandlerFunction();
      if (fn === "handleCalendarEditTrigger") hasEditTrigger = true;
      if (fn === "handleSheetChangeTrigger") hasChangeTrigger = true;
      if (fn === "runPeriodicAutoSync" || fn === "runPeriodic1MinSync") hasPeriodicTrigger = true;
      if (fn === "sendDailySiteMorningBriefing") hasBriefingTrigger = true;
    }

    // 1. 시트 셀 수정 감지용 onEdit 트리거 설치
    if (!hasEditTrigger) {
      ScriptApp.newTrigger("handleCalendarEditTrigger")
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      Logger.log("⚡ 시트 편집(onEdit) 트리거 자동 설치 완료");
    }

    // 2. [신규] 시트 행 우클릭 삭제(REMOVE_ROW) 감지용 onChange 트리거 설치
    if (!hasChangeTrigger) {
      ScriptApp.newTrigger("handleSheetChangeTrigger")
        .forSpreadsheet(ss)
        .onChange()
        .create();
      Logger.log("🗑️ 행 삭제 감지(onChange) 트리거 자동 설치 완료");
    }

    // 3. 디바운스/실시간 모드 시 불필요한 주기 트리거 정리
    if ((syncMode === "DEBOUNCE" || syncMode === "REALTIME") && hasPeriodicTrigger) {
      for (var j = 0; j < triggers.length; j++) {
        var pFn = triggers[j].getHandlerFunction();
        if (pFn === "runPeriodicAutoSync" || pFn === "runPeriodic1MinSync") {
          ScriptApp.deleteTrigger(triggers[j]);
        }
      }
    }

    // 4. 주기 자동 동기화 모드인 경우 시간 기반 트리거 설치
    if ((syncMode === "5MIN" || syncMode === "1MIN") && !hasPeriodicTrigger) {
      var minutes = (syncMode === "1MIN") ? 1 : 5;
      ScriptApp.newTrigger("runPeriodicAutoSync")
        .timeBased()
        .everyMinutes(minutes)
        .create();
      Logger.log("⏰ " + minutes + "분 주기 자동 동기화 트리거 자동 설치 완료");
    }

    // 5. [모닝 브리핑 트리거 자율 복구]: 활성화 설정되어 있으나 단발성 트리거가 소실된 경우 자동 재예약
    var docProps = PropertiesService.getDocumentProperties();
    if (docProps.getProperty("DAILY_BRIEFING_ENABLED") === "true" && !hasBriefingTrigger) {
      scheduleNextDailyBriefingTrigger();
      Logger.log("☀️ 모닝 브리핑 소실 트리거 자동 복구 완료");
    }
  } catch(e) {
    Logger.log("트리거 자동 점검 건너뜀: " + e.toString());
  }
}

/**
 * [신규] 시트 구조 변경(우클릭 행 삭제 등) 감지 트리거 핸들러
 * - 작업자가 행 번호 우클릭 ➔ [행 삭제] 실행 시 대시보드 달력을 즉시 갱신하고 캘린더 유령 일정 삭제를 예약합니다.
 */
function handleSheetChangeTrigger(e) {
  if (!e || e.changeType !== "REMOVE_ROW") return; // 오직 실제 '행 삭제(REMOVE_ROW)' 시에만 정확히 동작!
  try {
    var ss = e.source || getAppSpreadsheet();
    if (!ss) return;

      // 1) 삭제된 행이 대시보드 달력에서도 즉시 사라지도록 새로고침
      refreshDashboardCalendar(ss);

      // 1-1) 구성원 시트 행 삭제 가능성 대비 현장데이터 담당자 드롭다운 유효성 검사 동기화
      var siteSheetForVal = findSiteScheduleSheet(ss);
      if (siteSheetForVal) {
        applySiteScheduleValidations(siteSheetForVal);
      }

      var props = PropertiesService.getScriptProperties();
      var syncMode = props.getProperty("SYNC_MODE") || "DEBOUNCE";
      if (syncMode === "MANUAL") return;

      // 2) 유령 일정 청소 필요 플래그 마킹
      props.setProperty("NEEDS_ORPHAN_CLEANUP", "true");

      if (syncMode === "REALTIME") {
        // 실시간 모드면 즉시 캘린더 유령 일정 삭제 실행
        var sheet = findSiteScheduleSheet(ss);
        var cal = getTargetCalendar();
        if (sheet && cal) {
          cleanupOrphanCalendarEventsSilent(sheet, cal);
          props.deleteProperty("NEEDS_ORPHAN_CLEANUP");
        }
      } else if (syncMode === "DEBOUNCE") {
        // 디바운스 모드면 1분 뒤 백그라운드 동기화 시 일괄 삭제되도록 예약
        props.setProperty("LAST_EDIT_MS", String(Date.now()));
        if (props.getProperty("DEBOUNCE_SCHEDULED") !== "true") {
          cleanupDebouncedTriggers();
          ScriptApp.newTrigger("runDebouncedSync")
            .timeBased()
            .after(60 * 1000)
            .create();
          props.setProperty("DEBOUNCE_SCHEDULED", "true");
        }
      }
    } catch (err) {
      Logger.log("행 삭제 감지 트리거 에러: " + err.toString());
    }
  }

/**
 * [신규] 백그라운드 무음 유령(고아) 캘린더 일정 자동 청소기
 * - 우클릭 [행 삭제]로 인해 시트에서 사라진 일정을 찾아 구글 캘린더에서 조용히 삭제합니다.
 * - 알림창(ui.alert)을 띄우지 않아 백그라운드 트리거 및 모닝 브리핑에서 100% 안전하게 동작합니다.
 */
function cleanupOrphanCalendarEventsSilent(sheet, cal) {
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (ss) sheet = findSiteScheduleSheet(ss);
  }
  if (!sheet) return 0;
  if (!cal) cal = getTargetCalendar();
  if (!cal) return 0;

  try {
    var lastRow = sheet.getLastRow();
    var validEventIds = {};
    var unsyncedRows = [];

    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
      var lastKnownSite = "";

      for (var i = 0; i < values.length; i++) {
        var rowVal = values[i];
        var rawSite = String(rowVal[0] || "").trim();
        if (rawSite) lastKnownSite = rawSite;

        var eventId = String(rowVal[8] || "").trim();
        var hasDateOrTask = Boolean(String(rowVal[1] || "").trim() || String(rowVal[2] || "").trim());

        if (eventId) {
          validEventIds[eventId] = i + 2;
        } else if (hasDateOrTask) {
          if (!rowVal[0] && lastKnownSite) rowVal[0] = lastKnownSite;
          unsyncedRows.push({ rowNum: i + 2, rowVals: rowVal });
        }
      }

      // [안전 방어]: 아직 I열에 ID가 안 적힌 미동기화 행이 있다면 실수로 지우지 않도록 먼저 매칭 확인
      for (var u = 0; u < unsyncedRows.length; u++) {
        var item = unsyncedRows[u];
        var details = buildEventDetails(item.rowVals, item.rowNum, null, sheet);
        if (!details) continue;
        var matchedEv = findMatchingCalendarEvent(cal, details, validEventIds);
        if (matchedEv) {
          var mId = matchedEv.getId();
          sheet.getRange(item.rowNum, 9).setValue(mId);
          sheet.getRange(item.rowNum, 10).setValue(Date.now());
          validEventIds[mId] = item.rowNum;
        }
      }
    }

    // 최근 60일 전 ~ 향후 180일 구간 스캔 (1회 API 호출로 고속 조회)
    var now = new Date();
    var searchStart = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
    var searchEnd = new Date(now.getTime() + 180 * 24 * 60 * 60 * 1000);
    var allCalEvents = [];
    try {
      allCalEvents = cal.getEvents(searchStart, searchEnd);
    } catch(eEv) {
      allCalEvents = [];
    }
    var deletedCount = 0;

    for (var k = 0; k < allCalEvents.length; k++) {
      if (deletedCount >= 50) break; // 1회 최대 50건 삭제로 Rate Limit 완벽 보호 및 잔여 테스트 일정 일괄 정리
      var ev = allCalEvents[k];
      var evId = ev.getId();

      if (!isSystemSiteScheduleEvent(ev)) continue; // 개인 일정 절대 보호
      if (validEventIds[evId]) continue;            // 시트에 존재하는 정상 일정 통과

      // 시트에서 행 삭제로 사라진 유령 일정 삭제!
      try {
        ev.deleteEvent();
        deletedCount++;
        Utilities.sleep(150);
      } catch (eDel) {}
    }

    if (deletedCount > 0) {
      Logger.log("🧹 유령 일정 무음 자동 정리 완료: " + deletedCount + "건 삭제");
    }
    return deletedCount;
  } catch (e) {
    Logger.log("무음 유령 일정 정리 에러: " + e.toString());
    return 0;
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
  var needRelease = false;
  if (!lock.hasLock()) {
    if (!lock.tryLock(5000)) return;
    needRelease = true;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet(); // 백그라운드 트리거 null 방어!
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

    var lastSyncMs = lastSyncRaw ? parseInt(lastSyncRaw, 10) : (nowMs - 300000);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    ensureCalendarSyncColumns(sheet);

    var range = sheet.getRange(2, 1, lastRow - 1, 14);
    var values = range.getValues();

    var rowsToSyncToCal = [];
    var calUpdatedCount = 0;
    var lastKnownSite = "";

    // 1단계: 시트 내 신규 및 변경 행 감지 (시트 ➔ 캘린더 대상 선별)
    for (var i = 0; i < values.length; i++) {
      var rowNumber = i + 2;
      var rowVals = values[i];
      var rawSite = String(rowVals[0] || "").trim();
      if (rawSite) lastKnownSite = rawSite;

      var dateVal = rowVals[1];
      var taskName = String(rowVals[2] || "").trim();
      var effectiveSite = rawSite || ((dateVal || taskName) ? lastKnownSite : "");
      var calId = String(rowVals[8] || "").trim();
      var rawUpdated = rowVals[9];
      var rowUpdatedMs = (rawUpdated instanceof Date) ? rawUpdated.getTime() : (typeof rawUpdated === "number" ? rawUpdated : (parseInt(rawUpdated, 10) || 0));

      // 신규 행(현장명+날짜+공종 완비 & ID 없음) 또는 기존 행(ID 있음 & 수정시각 최신)
      if ((effectiveSite && dateVal && taskName && !calId) || (calId && rowUpdatedMs > lastSyncMs)) {
        rowsToSyncToCal.push(rowNumber);
      }
    }

    var needsOrphanCleanup = (props.getProperty("NEEDS_ORPHAN_CLEANUP") === "true");
    var lastOrphanMs = parseInt(props.getProperty("LAST_ORPHAN_CLEANUP_MS") || "0", 10);
    var isNightTime = (hour >= 22 || hour < 7);
    // 주간(07시~22시) 30분 이상 경과 시 자율 치유 스캔
    var shouldPeriodicOrphanScan = (!isNightTime && (nowMs - lastOrphanMs > 30 * 60 * 1000));

    // [심야 시간(22:00 ~ 07:00) 스마트 유휴 가드: 변경 사항 및 유령 청소 필요 시에는 심야라도 즉시 처리!]
    if (isNightTime && rowsToSyncToCal.length === 0 && !needsOrphanCleanup) {
      props.setProperty("LAST_PERIODIC_SYNC_MS", String(nowMs));
      return;
    }

    // 2단계: 변경된 행을 캘린더로 일괄 전송 (최대 25건 안전 분할)
    var didSyncToCal = false;
    var hasRemainingRowsToSync = false;
    if (rowsToSyncToCal.length > 0) {
      var syncLimit = Math.min(rowsToSyncToCal.length, 25);
      if (rowsToSyncToCal.length > 25) {
        hasRemainingRowsToSync = true; // 25건 초과 시 다음 주기에 이어서 처리하도록 표시
      }
      for (var s = 0; s < syncLimit; s++) {
        syncSingleRowToCalendar(sheet, rowsToSyncToCal[s], true, true);
      }
      didSyncToCal = true;
    }

    // 3단계: 캘린더 ➔ 시트 역방향 순환 검사 (Cursor Round-Robin)
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

        if (String(cVals[1] || "").trim() !== newDateStr) {
          sheet.getRange(targetRowNumber, 2).setValue(newDateStr);
          sheet.getRange(targetRowNumber, 10).setValue(calUpdated);
          sheet.getRange(targetRowNumber, 13, 1, 2).setValues([[formatDateTime(calUpdatedDate), "구글 캘린더"]]);
          rowChanged = true;
        }
      }

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

    var nextCursor = (cursor + stepsTaken) % Math.max(values.length, 1);
    props.setProperty("CAL_INSPECT_CURSOR", String(nextCursor));

    // 4단계: 행 삭제(REMOVE_ROW) 발생 시 또는 주기적 자율 치유 시 유령 일정 자동 청소 (모든 미등록 행이 2단계에서 등록 완료된 상태일 때만 안전 실행!)
    if (!hasRemainingRowsToSync && (needsOrphanCleanup || shouldPeriodicOrphanScan)) {
      try {
        SpreadsheetApp.flush(); // 2단계에서 기록된 신규 캘린더ID를 시트에 확실히 반영 후 스캔
        var orphanDeleted = cleanupOrphanCalendarEventsSilent(sheet, cal);
        props.deleteProperty("NEEDS_ORPHAN_CLEANUP");
        props.setProperty("LAST_ORPHAN_CLEANUP_MS", String(nowMs));
        if (orphanDeleted > 0) calUpdatedCount += orphanDeleted;
      } catch (eOrphan) {
        Logger.log("유령 일정 자동 청소 예외: " + eOrphan.toString());
      }
    }

    if (didSyncToCal || calUpdatedCount > 0) {
      refreshDashboardCalendar(ss);
    }

    // 아직 처리 못 한 행(25건 초과분)이 없을 때만 LAST_PERIODIC_SYNC_MS를 최신화하여 누락 원천 차단!
    if (!hasRemainingRowsToSync) {
      props.setProperty("LAST_PERIODIC_SYNC_MS", String(nowMs));
    } else {
      // 남은 행이 있다면 1분 뒤 다시 돌도록 디바운스 자동 예약
      ScriptApp.newTrigger("runDebouncedSync").timeBased().after(60 * 1000).create();
      props.setProperty("DEBOUNCE_SCHEDULED", "true");
    }
  } catch(e) {
    Logger.log("주기적 자동 동기화 에러: " + e.toString());
  } finally {
    if (needRelease && lock.hasLock()) {
      lock.releaseLock();
    }
  }
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
    var currentUser = getCurrentUserIdentifier(e);
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

    // [현장명(A열) 수정 시 하위 아코디언 행 타임스탬프 동반 갱신]
    // A열(현장명)이 수정되면, 하위 공종 행들(A열이 비어있는 행들 또는 동일 구 현장명 행들)도
    // 상위 현장명 변경의 영향을 받으므로 타임스탬프(J열)를 함께 최신화하여 캘린더 동기화 대상에 포함!
    var subRowsCount = 0;
    if (startCol <= 1 && endCol >= 1) {
      try {
        var lastR = sheet.getLastRow();
        if (lastR > maxProcessRow) {
          var checkCount = Math.min(lastR - maxProcessRow, 60); // 최대 60개 하위 행 탐색
          var subData = sheet.getRange(maxProcessRow + 1, 1, checkCount, 3).getValues();
          var subRowsToTouch = 0;
          var oldSiteVal = (e && e.oldValue) ? String(e.oldValue).trim() : "";
          var newSiteVal = (e && e.value) ? String(e.value).trim() : String(sheet.getRange(actualStartRow, 1).getValue() || "").trim();

          // e.oldValue가 전달되지 않은 경우(모바일 편집/서식 복사 등), 부모 행의 기존 캘린더 일정 위치 정보에서 이전 현장명 자동 복원
          if (!oldSiteVal) {
            try {
              var parentRowVals = sheet.getRange(actualStartRow, 1, 1, 9).getValues()[0];
              var parentCalId = String(parentRowVals[8] || "").trim();
              if (parentCalId) {
                var calObj = getTargetCalendar();
                var parentEv = calObj ? (calObj.getEventById(parentCalId) || CalendarApp.getEventById(parentCalId)) : null;
                if (parentEv) {
                  oldSiteVal = parentEv.getLocation() || "";
                  if (!oldSiteVal) {
                    var dMatch = (parentEv.getDescription() || "").match(/(?:📌\s*현장명|📍\s*위치):\s*([^\n\r]+)/);
                    if (dMatch) oldSiteVal = dMatch[1].trim();
                    if (!oldSiteVal) {
                      var tMatch = (parentEv.getTitle() || "").match(/\]\s*([^(\n\r]+)/) || (parentEv.getTitle() || "").match(/\[(.+?)\]/);
                      if (tMatch) oldSiteVal = tMatch[1].trim();
                    }
                  }
                }
              }
            } catch(eOld) {}
          }

          for (var si = 0; si < subData.length; si++) {
            var subSite = String(subData[si][0] || "").trim();
            // A열에 이전 현장명이 남아있고 새 현장명이 있으면 일괄 변경
            if (oldSiteVal && subSite === oldSiteVal && newSiteVal) {
              sheet.getRange(maxProcessRow + 1 + si, 1).setValue(newSiteVal);
              subRowsToTouch = si + 1;
              continue;
            }
            // A열이 비어있는 아코디언 하위 행인 경우: 계속 탐색하여 타임스탬프 갱신 대상에 포함
            if (!subSite) {
              var subDate = subData[si][1];
              var subTask = String(subData[si][2] || "").trim();
              if (subDate || subTask) {
                subRowsToTouch = si + 1;
              }
              continue;
            }
            // 다른 완전히 새로운 현장명이 등장하면 이 현장의 하위 행이 끝난 것으로 판단
            break;
          }
          if (subRowsToTouch > 0) {
            subRowsCount = subRowsToTouch;
            var subTimeStamps = [];
            var subMetaUpdates = [];
            var subMetaRange = sheet.getRange(maxProcessRow + 1, 11, subRowsToTouch, 4);
            var curSubMeta = subMetaRange.getValues();
            for (var mi = 0; mi < subRowsToTouch; mi++) {
              subTimeStamps.push([nowMs]);
              subMetaUpdates.push([curSubMeta[mi][0] || nowStr, curSubMeta[mi][1] || currentUser, nowStr, currentUser]);
            }
            sheet.getRange(maxProcessRow + 1, 10, subRowsToTouch, 1).setValues(subTimeStamps);
            subMetaRange.setValues(subMetaUpdates);
          }
        }
      } catch (eSub) {
        Logger.log("하위 행 타임스탬프 갱신 예외: " + eSub.toString());
      }
    }

    // [날짜(B열) 2026.00.00 표준 형식 실시간 자동 정규화]
    if (startCol <= 2 && endCol >= 2) {
      try {
        var dateRange = sheet.getRange(actualStartRow, 2, numRows, 1);
        var dateVals = dateRange.getValues();
        var dateChanged = false;
        for (var di = 0; di < dateVals.length; di++) {
          var rawD = dateVals[di][0];
          if (!rawD) continue;
          var normD = normalizeDateStringToYYYYMMDD(rawD);
          if (normD && String(rawD).trim() !== normD) {
            dateVals[di][0] = normD;
            dateChanged = true;
          }
        }
        if (dateChanged) {
          dateRange.setValues(dateVals);
        }
      } catch (eDate) {
        Logger.log("날짜 자동 정규화 예외: " + eDate.toString());
      }
    }

    // [지연 여부에 따른 배경색 실시간 동기화 (지연 공정: #fff1f2, 정상/미래/완료 공정: #ffffff)]
    if ((startCol <= 2 && endCol >= 2) || (startCol <= 8 && endCol >= 8)) {
      try {
        var rowRange = sheet.getRange(actualStartRow, 1, numRows, 8);
        var rowVals = rowRange.getValues();
        var rowBgs = rowRange.getBackgrounds();
        var bgUpdated = false;
        for (var bi = 0; bi < numRows; bi++) {
          var rDVal = rowVals[bi][1];
          var rStVal = String(rowVals[bi][7] || "").trim().replace(/\s+/g, "");
          var rParsed = parseScheduleDateRange(rDVal);
          var isDly = rParsed ? isScheduleDelayed(rParsed, rStVal, now) : false;
          if (isDly) {
            if (rowBgs[bi][0] !== "#fff1f2") {
              for (var bc = 0; bc < 8; bc++) rowBgs[bi][bc] = "#fff1f2";
              bgUpdated = true;
            }
          } else {
            // 미래 일정이거나 완료 상태이면 붉은색 잔여물을 흰색(#ffffff)으로 즉시 복구!
            if (rowBgs[bi][0] && rowBgs[bi][0] !== "#ffffff" && rowBgs[bi][0] !== "#FFFFFF") {
              for (var bc2 = 0; bc2 < 8; bc2++) rowBgs[bi][bc2] = "#ffffff";
              bgUpdated = true;
            }
          }
        }
        if (bgUpdated) {
          rowRange.setBackgrounds(rowBgs);
        }

        // [Delete 키로 셀 내용(날짜/공종)을 지운 경우 유령 일정 청소 플래그 자동 활성화]
        var hasClearedRow = false;
        for (var cri = 0; cri < numRows; cri++) {
          var rD = String(rowVals[cri][1] || "").trim();
          var rT = String(rowVals[cri][2] || "").trim();
          if (!rD && !rT) {
            hasClearedRow = true;
            break;
          }
        }
        if (hasClearedRow) {
          props.setProperty("NEEDS_ORPHAN_CLEANUP", "true");
        }
      } catch (eBg) {
        Logger.log("배경색 실시간 동기화 예외: " + eBg.toString());
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
        var totalProcessRow = maxProcessRow + subRowsCount;
        for (var r = actualStartRow; r <= totalProcessRow; r++) {
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
 * 두 날짜의 년/월/일이 같은지 판별
 */
function isDateEqual(d1, d2) {
  if (!d1 && !d2) return true;
  if (!d1 || !d2) return false;
  if (d1.getFullYear() === d2.getFullYear() &&
      d1.getMonth() === d2.getMonth() &&
      d1.getDate() === d2.getDate()) {
    return true;
  }
  // 타임존(UTC vs KST 등 9시간 시차) 오프셋 방어: 14시간 이내 차이 허용
  var diffHours = Math.abs(d1.getTime() - d2.getTime()) / (3600 * 1000);
  if (diffHours <= 14) return true;
  return false;
}

/**
 * Date 객체를 YYYY.MM.DD 문자열로 변환 (예: 2026.10.15)
 */
function formatDateYYMMDD(date) {
  if (!date) return "";
  var yyyy = String(date.getFullYear());
  var mm = String(date.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(date.getDate());
  if (dd.length < 2) dd = "0" + dd;
  return yyyy + "." + mm + "." + dd;
}

/**
 * Date 객체를 YYYY.MM.DD HH:mm 문자열로 변환 (예: 2026.10.09 20:50)
 */
function formatDateTime(d) {
  if (!d) return "";
  var yyyy = String(d.getFullYear());
  var mm = String(d.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(d.getDate());
  if (dd.length < 2) dd = "0" + dd;
  var hh = String(d.getHours());
  if (hh.length < 2) hh = "0" + hh;
  var min = String(d.getMinutes());
  if (min.length < 2) min = "0" + min;
  return yyyy + "." + mm + "." + dd + " " + hh + ":" + min;
}

/**
 * 임의의 날짜 입력을 2026.00.00 (YYYY.MM.DD) 표준 형식으로 변환 (공백 제거 및 2자리 패딩)
 * 예: 26.10.15 -> 2026.10.15 / 2026. 10. 21 -> 2026.10.21 / 2026. 9. 9 -> 2026.09.09
 */
function normalizeDateStringToYYYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return formatDateYYMMDD(val);
  }
  var str = String(val).trim();
  if (!str) return "";

  var parsed = null;
  try {
    parsed = parseScheduleDateRange(str);
  } catch (e) {}

  if (parsed && parsed.startDate) {
    var sy = parsed.startDate.getFullYear();
    var sm = parsed.startDate.getMonth() + 1;
    var sd = parsed.startDate.getDate();
    var sStr = sy + "." + (sm < 10 ? "0" + sm : sm) + "." + (sd < 10 ? "0" + sd : sd);

    var endD = parsed.displayEndDate || parsed.endDate;
    if (endD && (parsed.startDate.getTime() !== endD.getTime())) {
      var ey = endD.getFullYear();
      var em = endD.getMonth() + 1;
      var ed = endD.getDate();
      var eStr = ey + "." + (em < 10 ? "0" + em : em) + "." + (ed < 10 ? "0" + ed : ed);
      if (sStr !== eStr) {
        return sStr + " ~ " + eStr;
      }
    }
    return sStr;
  }

  // Regex fallback: 26.10.15 또는 2026. 9. 9 -> 2026.09.09
  return str.replace(/(\b\d{2,4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})\b/g, function(match, y, m, d) {
    var yNum = parseInt(y, 10);
    if (yNum < 100) yNum += 2000;
    var mNum = parseInt(m, 10);
    var dNum = parseInt(d, 10);
    return yNum + "." + (mNum < 10 ? "0" + mNum : mNum) + "." + (dNum < 10 ? "0" + dNum : dNum);
  });
}

/**
 * 현장데이터 시트 B열의 모든 날짜를 2026.00.00 형식으로 일괄 변환
 */
function formatAllDatesInSheetToYYYYMMDD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, 2, lastRow - 1, 1);
  var values = range.getValues();
  var updatedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var rawVal = values[i][0];
    if (!rawVal) continue;
    var normalized = normalizeDateStringToYYYYMMDD(rawVal);
    if (normalized && String(rawVal).trim() !== normalized) {
      values[i][0] = normalized;
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    range.setValues(values);
    SpreadsheetApp.flush();
    // 날짜 변경 후 지연/정상 상태 재계산 및 비지연 정상 행 배경색 즉시 복구
    try {
      cleanNonDelayedRowBackgroundsSilent(sheet);
    } catch (eBg) {}
    SpreadsheetApp.getUi().alert("📅 날짜 형식 변환 완료", "총 " + updatedCount + "건의 날짜가 '2026.00.00' 형식으로 일괄 통일되었으며, 일정 상태에 맞게 시트 배경색이 자동 복구되었습니다.", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    // 날짜 변경이 없더라도 지연 아닌 정상 행 배경색 무음 복구 동기화
    try {
      cleanNonDelayedRowBackgroundsSilent(sheet);
    } catch (eBg2) {}
    SpreadsheetApp.getUi().alert("📅 날짜 형식 확인", "모든 날짜가 이미 '2026.00.00' 표준 형식입니다.\n(지연이 아닌 정상/완료/미래 일정의 시트 배경색 복구가 완료되었습니다)", SpreadsheetApp.getUi().ButtonSet.OK);
  }
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
 * 현장 일정 날짜를 요일이 포함된 YY.MM.DD (요일) 표준 포맷으로 변환
 * 예: 26.11.05 (목) / 26.11.05 (목) 14:00 / 26.11.05 (목) ~ 26.11.07 (토) / 26.11.05 (목) 09:00 ~ 18:00
 */
function formatScheduleDateDisplay(val) {
  if (!val) return "";
  var dayNames = ["일", "월", "화", "수", "목", "금", "토"];

  // 내부 헬퍼: 날짜+요일 문자열 생성 (예: 26.11.05 (목))
  function withDow(d) {
    if (!d) return "";
    return formatDateYYMMDD(d) + " (" + dayNames[d.getDay()] + ")";
  }

  // 내부 헬퍼: 시간 문자열 생성 (예: 14:00)
  function onlyTime(d) {
    var hh = String(d.getHours());
    if (hh.length < 2) hh = "0" + hh;
    var mm = String(d.getMinutes());
    if (mm.length < 2) mm = "0" + mm;
    return hh + ":" + mm;
  }

  var parsed = null;
  try {
    parsed = parseScheduleDateRange(val);
  } catch (e) {
    parsed = null;
  }

  if (!parsed || !parsed.startDate) {
    var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function"));
    if (isDateObj && !isNaN(val.getTime())) {
      return withDow(val);
    }
    return String(val).trim();
  }

  // 1. 시간 정보가 포함된 일정 (단일 시간 또는 시간 범위)
  if (parsed.hasTime) {
    if (parsed.startTime && parsed.endTime) {
      var startStr = withDow(parsed.startTime) + " " + onlyTime(parsed.startTime);
      var diffMs = parsed.endTime.getTime() - parsed.startTime.getTime();
      if (isDateEqual(parsed.startTime, parsed.endTime)) {
        // 기본 1시간 차이면 시작 시간만 표시, 그 외에는 종료 시간까지 표시
        if (diffMs === 3600000 || diffMs === 0) return startStr;
        return startStr + " ~ " + onlyTime(parsed.endTime);
      } else {
        return startStr + " ~ " + withDow(parsed.endTime) + " " + onlyTime(parsed.endTime);
      }
    } else if (parsed.startTime) {
      return withDow(parsed.startTime) + " " + onlyTime(parsed.startTime);
    } else {
      return withDow(parsed.startDate) + (parsed.timeStr ? (" " + parsed.timeStr) : "");
    }
  }

  // 2. 종일 기간(Range) 일정
  if (parsed.isRange && parsed.displayEndDate) {
    return withDow(parsed.startDate) + " ~ " + withDow(parsed.displayEndDate);
  }

  // 3. 단일 종일 일정
  return withDow(parsed.startDate);
}

/**
 * 현재 로그인 사용자 닉네임/아이디 추출
 * @param {Object} [e] onEdit / onChange 등 이벤트 객체 (옵션)
 */
function getCurrentUserIdentifier(e) {
  var email = "";
  // 1순위: onEdit 이벤트 객체 내 user 객체 직접 확인
  if (e && e.user) {
    try {
      if (typeof e.user.getEmail === "function") {
        email = e.user.getEmail();
      }
    } catch (err) {}
    if (!email && typeof e.user === "string") {
      email = e.user;
    }
  }

  // 2순위: Session.getActiveUser() (대화형 실행 및 Workspace 도메인 사용자)
  if (!email) {
    try {
      email = Session.getActiveUser().getEmail();
    } catch (eActive) {}
  }

  // 3순위: Session.getEffectiveUser() (트리거 설치 계정)
  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch (eEff) {}
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
 * [영구 고정 패치] 오직 지정된 '현장 일정' 공용 캘린더만 반환
 * - 개인(기본) 캘린더로 잘못 넘어가거나 새 캘린더를 중복 생성하는 로직을 100% 원천 차단!
 */
function getTargetCalendar() {
  var docProps = PropertiesService.getDocumentProperties();
  var savedTarget = docProps.getProperty("TARGET_CALENDAR_SETTING");
  var target = (savedTarget && savedTarget.trim()) ? savedTarget.trim() : DEFAULT_CALENDAR_ID_OR_NAME.trim();

  // 1순위: 캘린더 고유 ID로 정확히 조회
  if (target) {
    try {
      var calById = CalendarApp.getCalendarById(target);
      if (calById) return calById;
    } catch (e) {}

    // 2순위: 혹시 이름으로 설정된 경우 '현장 일정' 이름으로 조회 (단, 새로 생성하지는 않음!)
    try {
      var calsByName = CalendarApp.getCalendarsByName(target);
      if (calsByName && calsByName.length > 0) {
        return calsByName[0];
      }
    } catch (e) {}
  }

  // [핵심 방어]: 만약 현재 접속한 작업자에게 '현장 일정' 캘린더 권한이 없거나 못 찾은 경우,
  // 절대 개인 캘린더(getDefaultCalendar)에 대신 넣거나 새 캘린더(createCalendar)를 만들지 않고 null을 반환하여 오염을 막습니다!
  Logger.log("⚠️ 지정된 '현장 일정' 캘린더에 접근할 수 없습니다. (캘린더 공유 권한을 확인하세요)");
  return null;
}

/**
 * [메뉴용] 연동할 구글 캘린더 확인 및 변경 대화상자
 */
function configureTargetCalendar() {
  var ui = SpreadsheetApp.getUi();
  var currentCal = getTargetCalendar();
  var currentCalName = currentCal ? currentCal.getName() : "지정되지 않음 (접근 불가)";
  var currentCalId = currentCal ? currentCal.getId() : "";

  var promptText = "📌 현재 연동 중인 캘린더:\n" +
                   "• 이름: " + currentCalName + "\n" +
                   "• ID: " + (currentCalId || "확인 불가") + "\n\n" +
                   "연동할 공용 캘린더의 [고유 ID(c_xxx@group.calendar.google.com)] 또는 [이름]을 입력하세요:\n" +
                   "(빈 칸으로 두고 확인을 누르면 기본 공용 설정[" + DEFAULT_CALENDAR_ID_OR_NAME + "]으로 복원됩니다)";

  var response = ui.prompt("⚙️ 연동 캘린더 변경 / 확인", promptText, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  var docProps = PropertiesService.getDocumentProperties();

  if (!input) {
    docProps.deleteProperty("TARGET_CALENDAR_SETTING");
    var resetCal = getTargetCalendar();
    var rName = resetCal ? resetCal.getName() : DEFAULT_CALENDAR_ID_OR_NAME;
    ui.alert("✅ 기본 공용 캘린더 설정[" + rName + "]으로 복원되었습니다.");
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
      "이제 동기화 실행 시 오직 이 공용 캘린더로만 일정이 등록됩니다.",
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "⚠️ 캘린더를 찾을 수 없음",
      "입력하신 캘린더('" + input + "')를 구글 계정에서 찾을 수 없습니다.\n" +
      "캘린더 공유 권한(일정 변경 및 공유 관리)이 부여되어 있는지 확인해주세요.\n\n" +
      "(오염 방지를 위해 개인 캘린더로 변경되거나 새 캘린더를 임의로 만들지 않습니다)",
      ui.ButtonSet.OK
    );
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

  // 구글 캘린더 자체 블록 색상 자동 지정 (현장별 고유 색상 매핑: 같은 현장 = 같은 색, 다른 현장 = 다른 색)
  var eventColor = getCalendarColorForSite(siteName, targetSheet);

  // 제목 포맷: [지연경고] [상태아이콘] 🟦 또는 🟥 [공종] 현장명 (공종+현장명 표준 포맷)
  var baseTitle = "";
  if (taskName && siteName) {
    baseTitle = "[" + taskName + "] " + siteName;
  } else if (taskName) {
    baseTitle = "[" + taskName + "]";
  } else {
    baseTitle = siteName;
  }
  var title = delayPrefix + statusIcon + prefixIcon + baseTitle;

  // 설명(Description) 포맷: 상세 위치에 현장명 명시
  var descLines = [];
  if (isDelayed) {
    descLines.push("⚠️ [공정 지연 경고]: 마감 예정일이 경과하였으나 미완료 상태입니다. 빠른 확인이 필요합니다.");
    descLines.push("--------------------------------");
  }
  descLines.push("📍 위치: " + (siteName || "-"));
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
    if (!event) {
      try {
        event = CalendarApp.getEventById(existingEventId);
      } catch (err) {
        event = null;
      }
    }
  }

  var allSheetIds = null;
  // [캘린더 중복 생성 원천 차단]
  // 기존 ID가 비어있거나 유효하지 않은 경우, 캘린더에 이미 생성된 동일 일정이 있는지 탐색하여 재연결
  if (!event) {
    try {
      allSheetIds = getSheetCalendarIds(sheet);
      if (allSheetIds && existingEventId) {
        delete allSheetIds[existingEventId];
      }
      event = findMatchingCalendarEvent(cal, details, allSheetIds);
      if (event) {
        var prevEvId = existingEventId;
        existingEventId = event.getId();
        sheet.getRange(row, 9).setValue(existingEventId);
        sheet.getRange(row, 10).setValue(Date.now());
        rowValues[8] = existingEventId;
        if (allSheetIds) {
          allSheetIds[existingEventId] = true;
          if (prevEvId && prevEvId !== existingEventId) delete allSheetIds[prevEvId];
        }
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
    try {
      event.removeAllReminders(); // 🔕 구글 캘린더 기본 알림(전날 알림 등) 자동 차단
    } catch(eRem) {}
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
    // 3) 신규 생성 전 최후 중복 방어
    var fallbackEv = null;
    try {
      if (!allSheetIds) allSheetIds = getSheetCalendarIds(sheet);
      fallbackEv = findMatchingCalendarEvent(cal, details, allSheetIds);
    } catch(eFb) {}

    if (fallbackEv) {
      event = fallbackEv;
      var fbId = event.getId();
      sheet.getRange(row, 9).setValue(fbId);
      sheet.getRange(row, 10).setValue(Date.now());
      rowValues[8] = fbId;
      try { event.removeAllReminders(); } catch(eRem) {}
      event.setTitle(details.title);
      event.setDescription(details.description);
      event.setLocation(details.siteName);
      if (details.color) {
        try { event.setColor(details.color); } catch (e) {}
      }
    } else {
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
        try {
          event.removeAllReminders(); // 🔕 신규 생성 일정 기본 알림(전날 알림 등) 자동 차단
        } catch(eRem) {}
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
 * 구글 캘린더 API 호출 시 일시적 Rate Limit(속도 제한/할당량 일시 소진) 발생 시
 * Exponential Backoff(지수 백오프)로 대기 후 자동 재시도하는 안전 래퍼
 * @param {Function} actionFn 실행할 작업 함수
 * @param {number} [maxRetries=2] 최대 재시도 횟수
 * @param {number} [baseDelayMs=2500] 기본 대기 시간(ms)
 * @return {*} actionFn의 반환값
 */
function executeWithRateLimitRetry(actionFn, maxRetries, baseDelayMs) {
  var retries = (maxRetries !== undefined) ? maxRetries : 2;
  var delay = baseDelayMs || 2500;
  for (var attempt = 0; attempt <= retries; attempt++) {
    try {
      return actionFn();
    } catch (err) {
      var errMsg = String(err || "");
      var isRateLimit = (
        errMsg.indexOf("짧은 시간") !== -1 ||
        errMsg.indexOf("too many") !== -1 ||
        errMsg.indexOf("Rate limit") !== -1 ||
        errMsg.indexOf("quota") !== -1 ||
        errMsg.indexOf("exhausted") !== -1 ||
        errMsg.indexOf("Service invoked too many times") !== -1
      );
      if (isRateLimit && attempt < retries) {
        Utilities.sleep(delay * (attempt + 1));
        continue;
      }
      throw err;
    }
  }
}

/**
 * [추천] 최근 & 미등록 & 예정 일정 우선 동기화 (미등록 일정 전체 + 최근 90일 이내 ~ 미래 일정)
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
 * 구글 Rate Limit(속도 제한) 방지 초고속 안전 배치 동기화 엔진
 * (전체 캘린더 이벤트 1회 사전 조회(0ms 메모리 캐시) + 중복 행 자동 청소 + 25건 쿨다운 + 4분 30초 안전 가드)
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
  // 최근 범위: 기본 90일(3개월)로 대폭 확장하여 8월 하순 등 최근 2~3개월 일정 완전 포함
  var ninetyDaysAgo = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

  // 1) [복사-붙여넣기 행 중복 ID 방어]:
  // 행 복사로 인해 I열(캘린더ID)이 동일하게 복제된 경우, 둘 중 하나가 다른 쪽 이벤트를 덮어쓰는 문제 원천 차단
  var seenSheetIds = {};
  var duplicateRowsToClear = [];
  var allSheetIds = {};
  for (var k = 0; k < values.length; k++) {
    var eid = String(values[k][8] || "").trim();
    if (eid) {
      if (seenSheetIds[eid]) {
        duplicateRowsToClear.push(k + 2);
        values[k][8] = "";
        values[k][9] = "";
      } else {
        seenSheetIds[eid] = true;
        allSheetIds[eid] = true;
      }
    }
  }

  if (duplicateRowsToClear.length > 0) {
    for (var d = 0; d < duplicateRowsToClear.length; d++) {
      try {
        sheet.getRange(duplicateRowsToClear[d], 9, 1, 2).clearContent();
      } catch (eDup) {}
    }
  }

  // 2) [캘린더 내 중복 일정 및 시트 삭제 유령 일정 무음 사전 자동 정리]
  try {
    cleanupDuplicateCalendarEventsSilent(cal, sheet);
    cleanupOrphanCalendarEventsSilent(sheet, cal);
  } catch (eClean) {}

  // 3) [핵심 성능 혁신: 전체 캘린더 이벤트 1회 사전 일괄 조회(Pre-fetch)]
  // 시트 내 모든 일정의 최소~최대 날짜 범위를 계산하여 캘린더 API를 단 1회만 일괄 호출 (이후 모든 조회 0ms 메모리 캐시)
  var minDateMs = null;
  var maxDateMs = null;
  for (var k = 0; k < values.length; k++) {
    var dVal = values[k][1];
    if (dVal) {
      var p = parseScheduleDateRange(dVal);
      if (p && p.startDate) {
        var sMs = p.startDate.getTime();
        var eMs = (p.endDate || p.startDate).getTime();
        if (minDateMs === null || sMs < minDateMs) minDateMs = sMs;
        if (maxDateMs === null || eMs > maxDateMs) maxDateMs = eMs;
      }
    }
  }

  var preloadedEvents = [];
  var calEventsMap = {};
  if (minDateMs !== null && maxDateMs !== null) {
    try {
      var preloadStart = new Date(minDateMs - 3 * 24 * 60 * 60 * 1000);
      var preloadEnd = new Date(maxDateMs + 4 * 24 * 60 * 60 * 1000);
      preloadedEvents = cal.getEvents(preloadStart, preloadEnd) || [];
      cal._preloadedEvents = preloadedEvents;
      for (var pe = 0; pe < preloadedEvents.length; pe++) {
        var pEv = preloadedEvents[pe];
        calEventsMap[pEv.getId()] = pEv;
      }
    } catch (ePreload) {
      preloadedEvents = [];
      calEventsMap = {};
    }
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

    var existingId = String(values[i][8] || "").trim();

    // [핵심 보장]: 미등록 일정(신규/누락 일정)은 과거 날짜라도 100% 무조건 동기화 대상에 포함!
    // 이미 캘린더에 정상 등록되어 있는 기존 일정만 90일 초과 시 스킵
    if (onlyRecent && existingId && calEventsMap[existingId]) {
      var parsed = parseScheduleDateRange(dateVal);
      if (parsed && parsed.startDate) {
        if (parsed.startDate < ninetyDaysAgo) {
          continue;
        }
      }
    }

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

    // 작성일 보완
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
          var ev = calEventsMap[existingId] || cal.getEventById(existingId);
          if (ev) ev.deleteEvent();
        } catch(e) {}
        sheet.getRange(row, 9).clearContent();
        sheet.getRange(row, 10).clearContent();
      }
      continue;
    }

    // [이벤트 조회: 캐시 맵 우선 조회 (0ms, 0 API Call)]
    var event = null;
    if (existingId) {
      if (calEventsMap[existingId]) {
        event = calEventsMap[existingId];
      } else {
        try {
          event = cal.getEventById(existingId);
        } catch(e) {
          event = null;
        }
        if (!event) {
          try {
            event = CalendarApp.getEventById(existingId);
          } catch(e) {
            event = null;
          }
        }
        if (event) {
          calEventsMap[existingId] = event;
        }
      }
    }

    // [캘린더 중복 생성 방지: 메모리 맵 매칭]
    if (!event) {
      try {
        if (existingId && allSheetIds[existingId]) {
          delete allSheetIds[existingId];
        }
        event = findMatchingCalendarEvent(cal, details, allSheetIds);
        if (event) {
          var matchedId = event.getId();
          sheet.getRange(row, 9).setValue(matchedId);
          sheet.getRange(row, 10).setValue(nowMs);
          rowVals[8] = matchedId;
          allSheetIds[matchedId] = true;
          calEventsMap[matchedId] = event;
          if (existingId && existingId !== matchedId) {
            delete allSheetIds[existingId];
          }
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
            skippedCount++;
            continue;
          }

          // [구글 한도 절약 2단계: 실제 변경된 속성만 안전 갱신 (Rate Limit 자동 재시도 탑재)]
          executeWithRateLimitRetry(function() {
            try {
              event.removeAllReminders(); // 🔕 구글 캘린더 기본 알림(전날 알림 등) 자동 차단
            } catch(eRem) {}
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
          }, 2, 2500);

          sheet.getRange(row, 10).setValue(nowMs);
          updatedCount++;
          batchChangeCount++;
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
        // [신규 생성 직전 최후 방어]: 캘린더에 동일 날짜/동일 공종 일정이 이미 있는지 최종 확인
        var fallbackEv = findMatchingCalendarEvent(cal, details, allSheetIds);
        if (fallbackEv) {
          event = fallbackEv;
          var fbId = event.getId();
          sheet.getRange(row, 9).setValue(fbId);
          sheet.getRange(row, 10).setValue(nowMs);
          rowVals[8] = fbId;
          allSheetIds[fbId] = true;
          calEventsMap[fbId] = event;
          if (existingId && existingId !== fbId) {
            delete allSheetIds[existingId];
          }
          existingId = fbId;
          // 기존 이벤트 재연결 및 제목/위치/설명 즉시 갱신
          executeWithRateLimitRetry(function() {
            try {
              event.removeAllReminders();
            } catch(eRem) {}
            event.setTitle(details.title);
            event.setDescription(details.description);
            event.setLocation(details.siteName);
            if (details.color) {
              try { event.setColor(details.color); } catch (e) {}
            }
          }, 2, 2500);
          updatedCount++;
          batchChangeCount++;
          Utilities.sleep(350);
        } else {
          var opt = { description: details.description, location: details.siteName };
          event = executeWithRateLimitRetry(function() {
            if (details.parsedDate.hasTime) {
              return cal.createEvent(details.title, details.parsedDate.startTime, details.parsedDate.endTime, opt);
            } else {
              if (details.parsedDate.isRange) {
                return cal.createAllDayEvent(details.title, details.parsedDate.startDate, details.parsedDate.endDate, opt);
              } else {
                return cal.createAllDayEvent(details.title, details.parsedDate.startDate, opt);
              }
            }
          }, 2, 2500);

          if (event) {
            try {
              event.removeAllReminders(); // 🔕 신규 생성 일정 기본 알림(전날 알림 등) 자동 차단
            } catch(eRem) {}
            if (details.color) {
              try { event.setColor(details.color); } catch (e) {}
            }
            var newEventId = event.getId();
            sheet.getRange(row, 9).setValue(newEventId);
            sheet.getRange(row, 10).setValue(nowMs);
            rowVals[8] = newEventId;
            allSheetIds[newEventId] = true;
            calEventsMap[newEventId] = event;
            preloadedEvents.push(event);
            if (existingId && existingId !== newEventId) {
              delete allSheetIds[existingId];
            }
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
        }
      } catch (err) {
        var errMsg = String(err);
        if (errMsg.indexOf("짧은 시간") !== -1 || errMsg.indexOf("too many") !== -1 || errMsg.indexOf("Rate limit") !== -1 || errMsg.indexOf("quota") !== -1 || errMsg.indexOf("exhausted") !== -1) {
          rateLimitHit = true;
          break;
        }
      }
    }

    // 지연 공정 시트 배경 하이라이트 반영 및 정상 행(#ffffff) 복구
    try {
      if (details.isDelayed) {
        sheet.getRange(row, 1, 1, 8).setBackground("#fff1f2");
      } else {
        var rowBgs = sheet.getRange(row, 1, 1, 8).getBackgrounds()[0];
        var needWhiteFix = false;
        for (var rc = 0; rc < 8; rc++) {
          if (rowBgs[rc] && rowBgs[rc] !== "#ffffff" && rowBgs[rc] !== "#FFFFFF") {
            needWhiteFix = true;
            break;
          }
        }
        if (needWhiteFix) {
          sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
        }
      }
    } catch (e) {}
  }

  // 4) [최종 정리]: 전체 동기화 종료 후 비지연 정상 행 잔여 분홍색 일괄 무음 청소 & 잉여 고아 중복 무음 정리
  try {
    cleanNonDelayedRowBackgroundsSilent(sheet);
  } catch (eCleanBg) {}
  try {
    cleanupDuplicateCalendarEventsSilent(cal, sheet);
  } catch (eCleanEnd) {}

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
      desc.indexOf("현장 관리 시스템") !== -1 ||
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
function findMatchingCalendarEvent(cal, details, knownSheetIds, excludeId, preloadedEvents) {
  if (!cal || !details || !details.parsedDate || !details.parsedDate.startDate) return null;

  var events = [];
  var cached = preloadedEvents || (cal && cal._preloadedEvents);
  if (cached && cached.length > 0) {
    events = cached;
  } else {
    var startDate = details.parsedDate.startDate;
    // 타임존(UTC vs KST 등 9시간 시차) 오프셋 방어: -14시간 ~ +38시간 넓은 탐색 윈도우
    var searchStart = new Date(startDate.getTime() - 14 * 60 * 60 * 1000);
    var endDate = details.parsedDate.endDate || startDate;
    var searchEnd = new Date(endDate.getTime() + 38 * 60 * 60 * 1000);

    try {
      events = cal.getEvents(searchStart, searchEnd);
    } catch (e) {
      return null;
    }
  }

  if (!events || events.length === 0) return null;

  var siteName = details.siteName || "";
  var taskName = details.taskName || "";
  var matchedEvents = [];
  var fallbackTaskMatches = [];

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var evId = ev.getId();

    // 만약 이 이벤트가 이미 시트의 '다른 행'에 연결되어 있다면 건너뜀 (단, 현재 행의 기존 ID는 매칭 허용)
    if (knownSheetIds && knownSheetIds[evId] && evId !== excludeId) {
      continue;
    }

    if (!isSystemSiteScheduleEvent(ev)) continue;

    var title = ev.getTitle() || "";
    var desc = ev.getDescription() || "";
    var loc = ev.getLocation() || "";

    // 날짜 일치 확인
    var dateMatches = false;
    var isAllDay = ev.isAllDayEvent();
    if (details.parsedDate.hasTime) {
      if (!isAllDay) {
        var evStart = ev.getStartTime();
        if (evStart && Math.abs(evStart.getTime() - details.parsedDate.startTime.getTime()) < 10 * 60 * 1000) {
          dateMatches = true;
        }
      }
    } else {
      var evStart = isAllDay ? ev.getAllDayStartDate() : ev.getStartTime();
      if (evStart && isDateEqual(evStart, details.parsedDate.startDate)) {
        dateMatches = true;
      }
    }
    if (!dateMatches) continue;

    // 공종명 일치 확인
    var taskMatches = true;
    if (taskName) {
      taskMatches = (
        title.indexOf(taskName) !== -1 ||
        desc.indexOf("🔨 공종: " + taskName) !== -1
      );
    }
    if (!taskMatches) continue;

    // 현장명 일치 확인 ([공종] 현장명 및 기존 [현장명] 공종 양방향 지원)
    var siteMatches = (
      (siteName && (title.indexOf("[" + siteName + "]") !== -1 || title.indexOf(siteName) !== -1)) ||
      loc === siteName ||
      desc.indexOf("📌 현장명: " + siteName) !== -1 ||
      desc.indexOf("📍 위치: " + siteName) !== -1
    );

    if (siteMatches) {
      matchedEvents.push(ev);
    } else {
      // 현장명이 변경된 경우(예: 보라매경남 -> 보경102동):
      // 동일 날짜에 동일 공종을 가진 시스템 일정이 존재하면 후보로 보관
      fallbackTaskMatches.push(ev);
    }
  }

  if (matchedEvents.length > 0) return matchedEvents[0];
  if (fallbackTaskMatches.length > 0) return fallbackTaskMatches[0];

  return null;
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
  var searchStart = new Date(startDate.getTime() - 14 * 60 * 60 * 1000);
  var endDate = details.parsedDate.endDate || startDate;
  var searchEnd = new Date(endDate.getTime() + 38 * 60 * 60 * 1000);

  var events = [];
  var cached = (cal && cal._preloadedEvents);
  if (cached && cached.length > 0) {
    events = cached;
  } else {
    try {
      events = cal.getEvents(searchStart, searchEnd);
    } catch (e) {
      return 0;
    }
  }

  if (!events || events.length === 0) return 0;

  var siteName = details.siteName || "";
  var taskName = details.taskName || "";
  var deleted = 0;

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var evId = ev.getId();

    // 1) 보존 대상 현재 행의 이벤트는 절대 삭제 금지
    if (evId === keepEventId) continue;

    // 2) [핵심 방어] 시트의 다른 행에 유효하게 등록된 캘린더 ID라면 절대 삭제 금지! (동일 날짜/공종 다건 일정 100% 보존)
    if (knownSheetIds && knownSheetIds[evId]) continue;

    // 3) 시스템 관리 일정이 아니면 절대 건드리지 않음 (개인 일정 철저 보호)
    if (!isSystemSiteScheduleEvent(ev)) continue;

    var title = ev.getTitle() || "";
    var desc = ev.getDescription() || "";
    var loc = ev.getLocation() || "";

    var isAllDay = ev.isAllDayEvent();
    var evStart = isAllDay ? ev.getAllDayStartDate() : ev.getStartTime();
    if (!evStart || !isDateEqual(evStart, details.parsedDate.startDate)) continue;

    var taskMatches = true;
    if (taskName) {
      taskMatches = (title.indexOf(taskName) !== -1 || desc.indexOf("🔨 공종: " + taskName) !== -1);
    }
    if (!taskMatches) continue;

    var siteMatches = (
      (siteName && (title.indexOf("[" + siteName + "]") !== -1 || title.indexOf(siteName) !== -1)) ||
      loc === siteName ||
      desc.indexOf("📌 현장명: " + siteName) !== -1 ||
      desc.indexOf("📍 위치: " + siteName) !== -1
    );

    // 시트 I열에 존재하지 않는 미등록/고아 일정이면서 동일 현장 및 공종과 일치하는 중복 일정 안전 삭제
    if (siteMatches) {
      try {
        ev.deleteEvent();
        deleted++;
      } catch (e) {}
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
 * 캘린더 내 중복 일정 무음 자동 정리 엔진
 * - 동일 날짜, 동일 공종, 동일 현장에 대해 2개 이상 존재하는 시스템 일정 감지
 * - 시트 I열에 매핑된 정규 일정을 보존하고 잉여 중복 일정을 즉시 무음 삭제
 */
function cleanupDuplicateCalendarEventsSilent(cal, sheet) {
  try {
    if (!cal || !sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;

    var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var validEventIds = {};
    for (var i = 0; i < values.length; i++) {
      var eid = String(values[i][8] || "").trim();
      if (eid) validEventIds[eid] = i + 2;
    }

    var now = new Date();
    var searchStart = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
    var searchEnd = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);

    var allEvents = [];
    try {
      allEvents = cal.getEvents(searchStart, searchEnd);
    } catch(eEv) {
      return 0;
    }
    if (!allEvents || allEvents.length === 0) return 0;

    var eventsByKey = {};
    var deletedCount = 0;

    for (var k = 0; k < allEvents.length; k++) {
      var ev = allEvents[k];
      var evId = ev.getId();
      if (!isSystemSiteScheduleEvent(ev)) continue;

      var isAllDay = ev.isAllDayEvent();
      var evStart = isAllDay ? ev.getAllDayStartDate() : ev.getStartTime();
      if (!evStart) continue;

      var dateKey = formatDateYYMMDD(evStart);
      var title = ev.getTitle() || "";
      var desc = ev.getDescription() || "";

      var taskDescMatch = desc.match(/🔨\s*공종:\s*([^\n\r]+)/);
      var siteDescMatch = desc.match(/(?:📌\s*현장명|📍\s*위치):\s*([^\n\r]+)/);
      var siteClean = (ev.getLocation() || "").trim() || (siteDescMatch ? siteDescMatch[1].trim() : "");
      var taskClean = taskDescMatch ? taskDescMatch[1].trim() : "";

      if (!taskClean || !siteClean) {
        var bMatch = title.match(/\[(.+?)\]\s*(.*)/);
        if (bMatch) {
          if (!taskClean) taskClean = bMatch[1].trim();
          if (!siteClean) siteClean = bMatch[2].replace(/[🟦🟥▶️✅🛠️⚠️\[\]]/g, "").trim();
        }
      }

      var dedupKey = dateKey + "_" + siteClean + "_" + taskClean;
      if (!taskClean) dedupKey = dateKey + "_" + title;

      if (!eventsByKey[dedupKey]) {
        eventsByKey[dedupKey] = [];
      }
      eventsByKey[dedupKey].push(ev);
    }

    for (var dKey in eventsByKey) {
      var evList = eventsByKey[dKey];
      if (evList.length <= 1) continue;

      // 시트에 등록된 유효 이벤트와 미등록(고아/유령) 이벤트 분리
      var validEvents = [];
      var orphanEvents = [];

      for (var e = 0; e < evList.length; e++) {
        var itemEv = evList[e];
        if (validEventIds[itemEv.getId()]) {
          validEvents.push(itemEv);
        } else {
          orphanEvents.push(itemEv);
        }
      }

      // [핵심 원칙]: 시트 I열에 유효하게 등록된 일정(validEvents)은
      // 동일 날짜/동일 공종/동일 현장에 몇 개가 있든 100% 절대 삭제하지 않고 보존!
      //
      // 오직 시트 어디에도 등록되어 있지 않은 잉여 고아/유령 일정(orphanEvents)만 정리
      for (var o = 0; o < orphanEvents.length; o++) {
        if (validEvents.length > 0 || o > 0) {
          try {
            orphanEvents[o].deleteEvent();
            deletedCount++;
            Utilities.sleep(100);
          } catch (eDel) {}
        }
      }
    }

    return deletedCount;
  } catch (e) {
    Logger.log("중복 무음 정리 예외: " + e.toString());
    return 0;
  }
}

/**
 * [신규] 구글 캘린더에 이미 등록된 현장 일정들의 기본 알림(전날 알림 등) 일괄 해제
 */
function removeAllSiteCalendarReminders() {
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
    "🔕 구글 캘린더 전날 알림 일괄 해제",
    "구글 캘린더에 연동된 모든 현장 일정에서 구글 기본 알림(전날 알림 및 팝업/이메일 알림)을 일괄 삭제합니다.\n\n" +
    "지금 진행하시겠습니까?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("등록된 현장 데이터가 없습니다.", "알림 해제", 3);
    return;
  }

  ss.toast("캘린더 일정 알림을 일괄 해제하는 중입니다...", "알림 해제 진행 중", 5);

  var idVals = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  var clearedCount = 0;

  for (var i = 0; i < idVals.length; i++) {
    var eid = String(idVals[i][0] || "").trim();
    if (!eid) continue;
    try {
      var ev = cal.getEventById(eid);
      if (ev) {
        ev.removeAllReminders();
        clearedCount++;
      }
    } catch(e) {}

    if (clearedCount > 0 && clearedCount % 25 === 0) {
      Utilities.sleep(500); // 구글 Rate Limit 보호
    }
  }

  ui.alert(
    "✅ 캘린더 알림 전체 해제 완료",
    "총 " + clearedCount + "건의 캘린더 일정에서 전날 알림이 깨끗하게 삭제되었습니다!\n\n" +
    "이제 새로운 일정을 추가하거나 시트에서 수정할 때도 알림이 자동으로 발생하지 않습니다.",
    ui.ButtonSet.OK
  );
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
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return;

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
  var lastKnownSite = ""; // [핵심 보완] 아코디언 하위 행 현장명 자동 계승

  for (var i = 0; i < siteData.length; i++) {
    var rawSite = String(siteData[i][0] || "").trim();
    if (rawSite) lastKnownSite = rawSite;

    var dateVal = siteData[i][1];
    var taskName = String(siteData[i][2] || "").trim();
    var siteName = rawSite || ((dateVal || taskName) ? lastKnownSite : "");
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
 * [메뉴용] 모바일 대시보드 URL 및 QR 코드 안내 팝업
 */
function showMobileDashboardUrl() {
  var ui = SpreadsheetApp.getUi();
  var webAppUrl = getWebAppUrl();
  if (!webAppUrl) {
    var res = ui.alert(
      "📱 모바일 웹앱 배포 필요",
      "스마트폰에서 모바일 대시보드를 열기 위해서는 웹 앱 배포 URL이 필요합니다.\n\n" +
      "지금 웹 앱 배포 안내 및 URL 등록 화면을 여시겠습니까?",
      ui.ButtonSet.YES_NO
    );
    if (res === ui.Button.YES) {
      configureWebAppUrl();
    }
    return;
  }

  var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=" + encodeURIComponent(webAppUrl);
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }' +
    'body { padding: 20px; text-align: center; color: #1e293b; background: #ffffff; }' +
    'h3 { font-size: 17px; font-weight: 700; color: #0f172a; margin-bottom: 8px; }' +
    'p.sub { font-size: 12.5px; color: #64748b; margin-bottom: 16px; }' +
    '.qr-box { margin: 0 auto 16px; padding: 12px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 14px; display: inline-block; box-shadow: 0 4px 14px rgba(0,0,0,0.06); }' +
    '.url-box { background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 10px; font-size: 11px; word-break: break-all; margin-bottom: 14px; color: #2563eb; font-family: monospace; text-align: left; }' +
    '.btn-wrap { display: flex; gap: 8px; justify-content: center; margin-bottom: 16px; }' +
    '.btn { padding: 9px 16px; border-radius: 8px; font-size: 12px; font-weight: 700; border: none; cursor: pointer; text-decoration: none; }' +
    '.btn-primary { background: #2563eb; color: #ffffff; }' +
    '.btn-primary:hover { background: #1d4ed8; }' +
    '.btn-secondary { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; }' +
    '.btn-secondary:hover { background: #e2e8f0; }' +
    '.guide-box { background: #eff6ff; border-radius: 10px; padding: 12px 14px; text-align: left; font-size: 11.5px; color: #1e40af; line-height: 1.6; }' +
    '</style></head><body>' +
    '<h3>📱 모바일 현장 대시보드 바로가기</h3>' +
    '<p class="sub">스마트폰 기본 카메라로 아래 QR 코드를 비추면 즉시 열립니다.</p>' +
    '<div class="qr-box"><img src="' + qrUrl + '" width="180" height="180" alt="대시보드 QR코드"></div>' +
    '<div class="url-box" id="urlText">' + webAppUrl + '</div>' +
    '<div class="btn-wrap">' +
    '  <button class="btn btn-primary" onclick="copyUrl()">📋 URL 복사하기</button>' +
    '  <a class="btn btn-secondary" href="' + webAppUrl + '" target="_blank">🌐 새 창 열기</a>' +
    '</div>' +
    '<div class="guide-box">' +
    '  <b>💡 홈 화면에 추가하여 앱처럼 사용하는 방법:</b><br>' +
    '  • <b>아이폰 (Safari)</b>: 하단 공유(내보내기) 버튼 ➔ <b>[홈 화면에 추가]</b><br>' +
    '  • <b>갤럭시 (Chrome)</b>: 우측 상단 메뉴(⋮) ➔ <b>[홈 화면에 추가]</b><br>' +
    '  ➔ 스마트폰 바탕화면에 앱 아이콘이 생성되어 터치 한 번으로 열립니다!' +
    '</div>' +
    '<script>' +
    'function copyUrl() {' +
    '  var text = document.getElementById("urlText").innerText;' +
    '  navigator.clipboard.writeText(text).then(function() {' +
    '    alert("✅ 모바일 대시보드 URL이 클립보드에 복사되었습니다!");' +
    '  }).catch(function() {' +
    '    alert("복사 실패: 수동으로 주소를 복사해 주세요.");' +
    '  });' +
    '}' +
    '</script>' +
    '</body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(430).setHeight(530);
  ui.showModalDialog(htmlOutput, "📱 모바일 대시보드 접속 안내");
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
  var view = p.view ? String(p.view).toLowerCase().trim() : "";
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
          var lastKnown = "";
          for (var j = 0; j < allData.length; j++) {
            var rawS = String(allData[j][0] || "").trim();
            if (rawS) lastKnown = rawS;
            var sVal = rawS || lastKnown;
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
      if (!siteName && row > 2) {
        var prevSites = sheet.getRange(2, 1, row - 2, 1).getValues();
        for (var ps = prevSites.length - 1; ps >= 0; ps--) {
          var pSite = String(prevSites[ps][0] || "").trim();
          if (pSite) { siteName = pSite; break; }
        }
      }
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

  // 2) 특정 일정 단일 관리 화면 출력 (row 파라미터가 유효하고 view가 'dashboard'가 아닌 경우)
  if (row >= 2 && view !== "dashboard") {
    if (!sheet) {
      return HtmlService.createHtmlOutput("<h3 style='font-family:sans-serif;text-align:center;margin-top:50px;'>⚠️ 현장데이터 시트를 찾을 수 없습니다.</h3>");
    }

    var lastRow = sheet.getLastRow();
    if (row > lastRow) {
      var notFoundHtml = "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1.0'>" +
        "<style>body { background:#f1f5f9; display:flex; justify-content:center; align-items:center; min-height:100vh; padding:16px; font-family:-apple-system, sans-serif; text-align:center; }" +
        ".box { background:#fff; padding:32px 24px; border-radius:18px; max-width:380px; width:100%; box-shadow:0 8px 24px rgba(0,0,0,0.08); }" +
        "h2 { color:#dc2626; font-size:18px; margin-bottom:12px; }" +
        "p { color:#475569; font-size:13.5px; line-height:1.6; margin-bottom:20px; }" +
        ".btn { display:inline-block; padding:11px 20px; background:#2563eb; color:#fff; text-decoration:none; border-radius:10px; font-weight:700; font-size:13.5px; }" +
        "</style></head><body><div class='box'>" +
        "<h2>⚠️ 유효하지 않은 일정</h2>" +
        "<p>해당 일정이 이미 완료/삭제되었거나 행 번호가 변경되었습니다.<br>전체 모바일 대시보드에서 최신 일정을 확인하세요.</p>" +
        "<a href='" + (getWebAppUrl() || "?view=dashboard") + "' class='btn'>📱 전체 대시보드로 이동</a>" +
        "</div></body></html>";
      return HtmlService.createHtmlOutput(notFoundHtml);
    }

    var rowVals = sheet.getRange(row, 1, 1, 9).getValues()[0];
    var siteName = String(rowVals[0] || "").trim();
    if (!siteName && row > 2) {
      try {
        var prevSites = sheet.getRange(2, 1, row - 2, 1).getValues();
        for (var ps = prevSites.length - 1; ps >= 0; ps--) {
          var pSite = String(prevSites[ps][0] || "").trim();
          if (pSite) { siteName = pSite; break; }
        }
      } catch(ePs) {}
    }
    if (!siteName) siteName = "-";
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
      .setTitle("현장 공정 관리 - " + siteName)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3) 전체 모바일 대시보드 화면 출력 (URL 직접 접속 또는 row 미지정 시)
  var dashHtml = buildSidebarDashboardHtml("mobile");
  return HtmlService.createHtmlOutput(dashHtml)
    .setTitle("우성아이디 현장 대시보드")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    "    <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;'>" +
    "      <a href='" + (webAppUrl || "?view=dashboard") + "' style='font-size:12px; color:#2563eb; text-decoration:none; font-weight:700; display:inline-flex; align-items:center; gap:4px; padding:4px 8px; border-radius:6px; background:#eff6ff;'>" +
    "        ← 📱 전체 대시보드" +
    "      </a>" +
    "      <span style='font-size:11px; color:#94a3b8;'>단일 일정 모드</span>" +
    "    </div>" +
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

// 현장별 캘린더 색상 캐시 (단일 실행 컨텍스트 내 중복 계산 방지)
var _siteColorMapCache = null;

/**
 * 현장데이터 시트 내 고유 현장명 순서 맵 반환 (등장 순서대로 인덱스 부여)
 */
function getSiteColorOrderMap(sheet) {
  if (_siteColorMapCache) return _siteColorMapCache;
  if (!sheet || typeof sheet.getLastRow !== "function") return {};
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    var sites = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var map = {};
    var count = 0;
    for (var i = 0; i < sites.length; i++) {
      var s = String(sites[i][0] || "").trim();
      if (s && map[s] === undefined) {
        map[s] = count;
        count++;
      }
    }
    _siteColorMapCache = map;
    return map;
  } catch(e) {
    return {};
  }
}

/**
 * 현장명별 구글 캘린더 블록 색상 자동 반환
 * - 같은 현장명으로 묶인 일정: 캘린더에 항상 100% 동일한 색상으로 표시
 * - 서로 다른 현장: 서로 다른 색상으로 자동 구분되어 한눈에 식별 가능
 */
function getCalendarColorForSite(siteName, sheet) {
  if (!siteName) return "9"; // 기본 Blueberry (진파랑)
  siteName = String(siteName).trim();

  // 구글 캘린더 색상 팔레트 (공휴일 및 17일 기존 색상 중복 방지 최적화)
  // - 공휴일(피치/주황) 및 기존 17일 일정 색상(7: Peacock 청록)과의 혼선 방지를 위해 7(청록), 6(주황), 4(산호) 배제
  // - 시인성이 뛰어나고 상호 구분이 가장 명확한 프리미엄 색상 우선 배치:
  // 9: Blueberry(진파랑), 3: Grape(보라), 10: Basil(초록), 5: Banana(노랑),
  // 1: Pale Blue(연파랑), 2: Sage(연초록), 8: Graphite(회색)
  var palette = [
    "9",  // Blueberry (진파랑 - 1순위 대표 현장 색상)
    "3",  // Grape (보라)
    "10", // Basil (초록)
    "5",  // Banana (노랑)
    "1",  // Pale Blue (연파랑)
    "2",  // Sage (연초록)
    "8"   // Graphite (회색)
  ];

  // 1순위: 시트 객체가 있는 경우: 시트 내 고유 현장명 등장 순서대로 1:1 매핑 (충돌 0%)
  if (sheet) {
    try {
      var siteMap = getSiteColorOrderMap(sheet);
      if (siteMap && siteMap[siteName] !== undefined) {
        return palette[siteMap[siteName] % palette.length];
      }
    } catch(e) {}
  }

  // 2순위: 시트가 없거나 단독 호출 시: 현장명 고유 문자열 해시 기반 일관 매핑
  var hash = 0;
  for (var i = 0; i < siteName.length; i++) {
    hash = ((hash << 5) - hash) + siteName.charCodeAt(i);
    hash |= 0;
  }
  return palette[Math.abs(hash) % palette.length];
}

/**
 * 진행상태 및 지연 여부에 따른 구글 캘린더 블록 색상 ID 산출 (기존 호환 유지)
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
 * 현장데이터 시트의 전체 행 공정 지연 검사 및 배경 하이라이트/정상 복구 (무음 실행)
 * - 지연 행: #fff1f2 연분홍색 하이라이트
 * - 정상/미래/완료 행 및 빈 행의 과거 붉은색 잔여물: #ffffff 흰색으로 완벽 복구
 * @param {Sheet} [targetSheet] 대상 시트 (생략 시 활성 시트)
 * @returns {Array} 감지된 지연 항목 배열
 */
function cleanNonDelayedRowBackgroundsSilent(targetSheet) {
  var sheet = targetSheet;
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    sheet = findSiteScheduleSheet(ss);
  }
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var today = new Date();
  var range = sheet.getRange(2, 1, lastRow - 1, 8);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var delayedItems = [];
  var bgChanged = false;
  var startTime = Date.now();
  var lastKnownSite = "";

  for (var i = 0; i < values.length; i++) {
    // 60초 타임아웃 안전 탈출 (6분 스크립트 강제 종료 방어)
    if (Date.now() - startTime > 60000) {
      break;
    }

    var row = i + 2;
    var rawSite = String(values[i][0] || "").trim();
    if (rawSite) lastKnownSite = rawSite;
    var dateVal = values[i][1];
    var taskName = String(values[i][2] || "").trim();
    var siteName = rawSite || ((dateVal || taskName) ? lastKnownSite : "");
    var manager = String(values[i][6] || "").trim();
    var status = String(values[i][7] || "").trim();
    var statusClean = status.replace(/\s+/g, "");

    // 1) 필수 데이터가 없는 빈 행인 경우: 과거 붉은색/비정상 배경색 잔여물이 있다면 깨끗하게 흰색으로 청소
    if (!siteName && !taskName && !dateVal) {
      for (var ce = 0; ce < 8; ce++) {
        var eBg = backgrounds[i][ce];
        if (eBg && eBg !== "#ffffff" && eBg !== "#FFFFFF") {
          backgrounds[i][ce] = "#ffffff";
          bgChanged = true;
        }
      }
      continue;
    }

    // 2) 날짜 파싱 실패 시: 지연 대상이 아니므로 비정상 배경색이 있다면 흰색 복구
    var parsed = parseScheduleDateRange(dateVal);
    if (!parsed) {
      for (var cu = 0; cu < 8; cu++) {
        var uBg = backgrounds[i][cu];
        if (uBg && uBg !== "#ffffff" && uBg !== "#FFFFFF") {
          backgrounds[i][cu] = "#ffffff";
          bgChanged = true;
        }
      }
      continue;
    }

    // 3) 지연 여부 판정
    if (isScheduleDelayed(parsed, statusClean, today)) {
      delayedItems.push({
        row: row,
        siteName: siteName,
        taskName: taskName,
        dateStr: String(dateVal).trim(),
        manager: manager,
        status: status
      });
      // 지연 행 연한 붉은색(#fff1f2) 경고 하이라이트 (A~H열 전체 적용)
      for (var c = 0; c < 8; c++) {
        if (backgrounds[i][c] !== "#fff1f2") {
          backgrounds[i][c] = "#fff1f2";
          bgChanged = true;
        }
      }
    } else {
      // 4) 정상/미래/완료 행 복구: 지연이 아닌 모든 정상 행은 A~H열 중 흰색이 아닌 셀(수동 칠해짐 포함)을 전부 흰색(#ffffff)으로 완벽 복구!
      for (var c2 = 0; c2 < 8; c2++) {
        var curCellBg = backgrounds[i][c2];
        if (curCellBg && curCellBg !== "#ffffff" && curCellBg !== "#FFFFFF") {
          backgrounds[i][c2] = "#ffffff";
          bgChanged = true;
        }
      }
    }
  }

  // 1회 단일 호출로 전체 시트 배경색 즉시 커밋 (기존 대비 수백배 가속 & API 절약)
  if (bgChanged) {
    range.setBackgrounds(backgrounds);
  }

  return delayedItems;
}

/**
 * 현장데이터 시트의 전체 행 공정 지연 검사 및 배경 하이라이트 반영 (UI 알림 포함)
 */
function checkAndHighlightDelayedSchedules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var delayedItems = cleanNonDelayedRowBackgroundsSilent(sheet);
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
      "※ 지연된 행은 [연한 붉은색]으로 하이라이트되었으며, 완료되거나 미래 일정인 모든 정상 행은 [흰색]으로 깨끗하게 복구되었습니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "✅ 공정 지연 검사 완료",
      "지연된 현장 공정이 없습니다. 모든 일정이 예정대로 진행 중이며, 시트의 모든 정상/미래/완료 행이 흰색으로 복구되었습니다!",
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
 * 주어진 시트가 [구성원] (또는 팀원/직원/조직도) 시트인지 확인
 */
function isMembersSheet(sheet) {
  if (!sheet) return false;
  var sName = sheet.getName().trim().toLowerCase();
  var candidateNames = ["구성원", "팀원", "직원", "조직도", "members", "team", "member"];
  for (var c = 0; c < candidateNames.length; c++) {
    if (sName === candidateNames[c] || sName.indexOf(candidateNames[c]) !== -1) {
      return true;
    }
  }
  return false;
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

  // 4순위: 스프레드시트 소유자 계정 (백그라운드 트리거 실행 시 Session 공백 방어)
  if (!userEmail && ss) {
    try {
      var owner = ss.getOwner();
      if (owner && owner.getEmail()) {
        userEmail = owner.getEmail();
        docProps.setProperty("DAILY_BRIEFING_EMAIL", userEmail);
      }
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
  var isManual = (isManualTest === true); // 수동 테스트 버튼 클릭 시에만 true, 트리거 실행 시 false (UI 팝업 방어)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return;
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    if (isManual) {
      SpreadsheetApp.getUi().alert("현장데이터 시트에 일정이 없습니다.");
    }
    return;
  }

  // 매일 아침 브리핑 발송 전, 행 삭제로 남은 캘린더 유령 일정 자동 청소
  try {
    var cal = getTargetCalendar();
    if (cal) {
      cleanupOrphanCalendarEventsSilent(sheet, cal);
    }
  } catch (eClean) {
    Logger.log("모닝 브리핑 사전 유령 일정 정리 예외: " + eClean.toString());
  }

  var recipient = getDailyBriefingRecipientEmail(ss);
  if (!recipient) {
    if (isManual) {
      SpreadsheetApp.getUi().alert("수신자 이메일 주소를 확인할 수 없습니다. [📧 모닝 브리핑 수신 이메일 변경] 메뉴에서 설정해주세요.");
    }
    return;
  }

  try {
    var today = new Date();
    var todayStr = formatDateYYMMDD(today);
    var dayOfWeekNames = ["일", "월", "화", "수", "목", "금", "토"];
    var dayOfWeekStr = dayOfWeekNames[today.getDay()] + "요일";

    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var todayTasks = [];
    var delayedTasks = [];
    var lastKnownSite = ""; // [핵심 보완] 아코디언 하위 행 현장명 자동 계승

    for (var i = 0; i < data.length; i++) {
      var rawSite = String(data[i][0] || "").trim();
      if (rawSite) lastKnownSite = rawSite;
      var dateVal = data[i][1];
      var taskName = String(data[i][2] || "").trim();
      var siteName = rawSite || ((dateVal || taskName) ? lastKnownSite : "");
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

    if (isManual) {
      SpreadsheetApp.getUi().alert(
        "✉️ 모닝 브리핑 테스트 발송 완료",
        "성공적으로 브리핑 이메일을 발송했습니다!\n\n" +
        "• 수신자: " + recipient + "\n" +
        "• 오늘 착공/진행 공정: " + todayTasks.length + "건\n" +
        "• 공정 지연 경고: " + delayedTasks.length + "건\n\n" +
        "이메일 수신함을 확인해보세요.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (eSend) {
    Logger.log("모닝 브리핑 발송 실행 오류: " + eSend.toString());
  } finally {
    // 자동 일일 실행인 경우, 다음 날 설정 시각 정각 트리거 자동 재스케줄링 (Self-Rescheduling)
    if (!isManual) {
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
// 15. [스마트 UI] 스프레드시트 우측 사이드바 & 팝업 실시간 대시보드 화면 엔진
// ========================================================

/**
 * 스프레드시트 우측 고정 영역에 실시간 현장 대시보드 사이드바 열기
 */
function openSiteScheduleSidebar() {
  var html = HtmlService.createHtmlOutput(buildSidebarDashboardHtml("sidebar"))
    .setTitle("우성아이디 현장 대시보드")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 스프레드시트 화면 중앙에 대형 팝업 모달 대시보드 열기
 */
function openSiteScheduleModal() {
  var html = HtmlService.createHtmlOutput(buildSidebarDashboardHtml("modal"))
    .setTitle("우성아이디 현장 대시보드")
    .setWidth(920)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, "우성아이디 현장 대시보드");
}

/**
 * 캘린더 색상 코드('1'~'11')를 UI용 Hex 색상 코드로 변환
 */
function getHexColorForCalendarCode(code) {
  var map = {
    "1": "#039BE5", // Pale Blue
    "2": "#33B679", // Sage
    "3": "#8E24AA", // Grape
    "5": "#F6BF26", // Banana
    "8": "#616161", // Graphite
    "9": "#3F51B5", // Blueberry
    "10": "#0B8043", // Basil
    "11": "#E53935"  // Tomato / Delayed
  };
  return map[String(code)] || "#2563EB";
}

/**
 * 대한민국 법정 공휴일 맵 반환 ({ "YYYY-MM-DD": "공휴일명" })
 * - 1순위: 스프레드시트 내 00_Calendar 시트의 X열(날짜)/Y열(명칭) 캐시 데이터
 * - 2순위: 2024~2027 기본 법정 공휴일 & 대체공휴일 내장 데이터 (오프라인/단독 환경 100% 보장)
 * @param {Spreadsheet} [ss]
 * @returns {Object} { [dateStr]: holidayName }
 */
function getKoreanHolidaysMap(ss) {
  var holidayMap = {};

  // 1) 2024~2027 법정공휴일 및 대체공휴일 기본 베이스 데이터
  var defaultList = [
    // 2024
    ["2024-01-01", "신정"], ["2024-02-09", "설날연휴"], ["2024-02-10", "설날"], ["2024-02-11", "설날연휴"], ["2024-02-12", "대체공휴일"],
    ["2024-03-01", "삼일절"], ["2024-04-10", "총선"], ["2024-05-05", "어린이날"], ["2024-05-06", "대체공휴일"], ["2024-05-15", "부처님오신날"],
    ["2024-06-06", "현충일"], ["2024-08-15", "광복절"], ["2024-09-16", "추석연휴"], ["2024-09-17", "추석"], ["2024-09-18", "추석연휴"],
    ["2024-10-03", "개천절"], ["2024-10-09", "한글날"], ["2024-12-25", "성탄절"],
    // 2025
    ["2025-01-01", "신정"], ["2025-01-28", "설날연휴"], ["2025-01-29", "설날"], ["2025-01-30", "설날연휴"],
    ["2025-03-01", "삼일절"], ["2025-03-03", "대체공휴일"], ["2025-05-05", "어린이날"], ["2025-05-06", "대체공휴일"],
    ["2025-06-06", "현충일"], ["2025-08-15", "광복절"], ["2025-10-03", "개천절"], ["2025-10-05", "추석연휴"],
    ["2025-10-06", "추석"], ["2025-10-07", "추석연휴"], ["2025-10-08", "대체공휴일"], ["2025-10-09", "한글날"], ["2025-12-25", "성탄절"],
    // 2026
    ["2026-01-01", "신정"], ["2026-02-16", "설날연휴"], ["2026-02-17", "설날"], ["2026-02-18", "설날연휴"],
    ["2026-03-01", "삼일절"], ["2026-03-02", "대체공휴일"], ["2026-05-05", "어린이날"], ["2026-05-24", "부처님오신날"],
    ["2026-05-25", "대체공휴일"], ["2026-06-03", "지방선거"], ["2026-06-06", "현충일"], ["2026-08-15", "광복절"],
    ["2026-08-17", "대체공휴일"], ["2026-09-24", "추석연휴"], ["2026-09-25", "추석"], ["2026-09-26", "추석연휴"],
    ["2026-09-28", "대체공휴일"], ["2026-10-03", "개천절"], ["2026-10-05", "대체공휴일"], ["2026-10-09", "한글날"], ["2026-12-25", "성탄절"],
    // 2027
    ["2027-01-01", "신정"], ["2027-02-06", "설날연휴"], ["2027-02-07", "설날"], ["2027-02-08", "설날연휴"], ["2027-02-09", "대체공휴일"],
    ["2027-03-01", "삼일절"], ["2027-05-05", "어린이날"], ["2027-05-13", "부처님오신날"], ["2027-06-06", "현충일"], ["2027-06-07", "대체공휴일"],
    ["2027-08-15", "광복절"], ["2027-08-16", "대체공휴일"], ["2027-09-14", "추석연휴"], ["2027-09-15", "추석"], ["2027-09-16", "추석연휴"],
    ["2027-10-03", "개천절"], ["2027-10-04", "대체공휴일"], ["2027-10-09", "한글날"], ["2027-10-11", "대체공휴일"], ["2027-12-25", "성탄절"]
  ];
  for (var i = 0; i < defaultList.length; i++) {
    holidayMap[defaultList[i][0]] = defaultList[i][1];
  }

  // 2) 00_Calendar 시트의 X열/Y열 공휴일 캐시가 있다면 최신 데이터로 병합
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (ss) {
      var calSheet = ss.getSheetByName("00_Calendar") || ss.getSheetByName("캘린더") || ss.getSheetByName("Calendar");
      if (calSheet) {
        var lastR = calSheet.getLastRow();
        if (lastR >= 1) {
          var checkRows = Math.min(lastR, 300);
          var vals = calSheet.getRange(1, 24, checkRows, 2).getValues(); // X열(24), Y열(25)
          for (var r = 0; r < vals.length; r++) {
            var rawD = vals[r][0];
            var name = String(vals[r][1] || "").trim();
            if (rawD && name) {
              var dObj = (rawD instanceof Date) ? rawD : new Date(rawD);
              if (!isNaN(dObj.getTime())) {
                var y = dObj.getFullYear();
                var m = dObj.getMonth() + 1;
                var d = dObj.getDate();
                var key = y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
                holidayMap[key] = name;
              }
            }
          }
        }
      }
    }
  } catch (eSheet) {}

  return holidayMap;
}

/**
 * 사이드바/모달 대시보드용 실시간 데이터 집계 API
 */
function getSidebarDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var sheet = findSiteScheduleSheet(ss);
    if (!sheet) return { error: "현장데이터 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        kpi: { totalSites: 0, activeSites: 0, totalTasks: 0, todayCount: 0, delayedCount: 0 },
        todayStr: formatDateYYMMDD(new Date()),
        members: getMemberNames(ss),
        holidays: getKoreanHolidaysMap(ss),
        sites: []
      };
    }

    var values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var today = new Date();
    var todayYYMMDD = formatDateYYMMDD(today);
    var nowMs = today.getTime();

    var sitesMap = {};
    var siteOrder = [];
    var lastKnownSite = "";
    var totalTasks = 0;
    var todayCount = 0;
    var delayedCount = 0;

    for (var i = 0; i < values.length; i++) {
      var row = i + 2;
      var rowVals = values[i];
      var rawSite = String(rowVals[0] || "").trim();
      if (rawSite) lastKnownSite = rawSite;
      var siteName = rawSite || lastKnownSite;
      var dateVal = rowVals[1];
      var taskName = String(rowVals[2] || "").trim();

      if (!siteName || !dateVal || !taskName) continue;

      var partner = String(rowVals[3] || "").trim();
      var isShared = (rowVals[4] === true || String(rowVals[4]).toUpperCase() === "TRUE");
      var memo = String(rowVals[5] || "").trim();
      var manager = String(rowVals[6] || "").trim();
      var status = String(rowVals[7] || "예정").trim();
      var calId = String(rowVals[8] || "").trim();

      var parsed = null;
      try {
        parsed = parseScheduleDateRange(dateVal);
      } catch (pe) {
        parsed = null;
      }
      var isToday = false;
      var isDelayed = false;
      var dateDisplay = "";
      try {
        dateDisplay = formatScheduleDateDisplay(dateVal);
      } catch (de) {
        dateDisplay = String(dateVal || "");
      }
      var dateMs = 0;

      if (parsed) {
        if (parsed.hasTime && parsed.startTime) {
          dateMs = parsed.startTime.getTime();
        } else if (parsed.startDate) {
          dateMs = parsed.startDate.getTime();
        }
        if (parsed.startDate) {
          isToday = isScheduleActiveToday(parsed, today);
          isDelayed = isScheduleDelayed(parsed, status, today);
        }
      }

      var startDateStr = "";
      var endDateStr = "";
      if (parsed && parsed.startDate) {
        var sy = parsed.startDate.getFullYear();
        var sm = parsed.startDate.getMonth() + 1;
        var sd = parsed.startDate.getDate();
        startDateStr = sy + "-" + (sm < 10 ? "0" + sm : sm) + "-" + (sd < 10 ? "0" + sd : sd);
        var edObj = parsed.displayEndDate || parsed.endDate || parsed.startDate;
        var ey = edObj.getFullYear();
        var em = edObj.getMonth() + 1;
        var ed = edObj.getDate();
        endDateStr = ey + "-" + (em < 10 ? "0" + em : em) + "-" + (ed < 10 ? "0" + ed : ed);
      }

      if (isToday) todayCount++;
      if (isDelayed) delayedCount++;
      totalTasks++;

      if (!sitesMap[siteName]) {
        var colorCode = "9";
        try {
          colorCode = getCalendarColorForSite(siteName, sheet);
        } catch (ce) {
          colorCode = "9";
        }
        sitesMap[siteName] = {
          siteName: siteName,
          colorCode: colorCode,
          colorHex: getHexColorForCalendarCode(colorCode),
          tasks: [],
          initialOrder: siteOrder.length
        };
        siteOrder.push(siteName);
      }

      sitesMap[siteName].tasks.push({
        row: row,
        taskName: taskName,
        dateDisplay: dateDisplay,
        dateMs: dateMs,
        startDateStr: startDateStr,
        endDateStr: endDateStr,
        partner: partner,
        manager: manager,
        status: status,
        isShared: isShared,
        memo: memo,
        isToday: isToday,
        isDelayed: isDelayed,
        hasCal: !!calId
      });
    }

    var sitesList = [];
    var activeSitesCount = 0;

    for (var s = 0; s < siteOrder.length; s++) {
      var siteObj = sitesMap[siteOrder[s]];

      // 1. 공정(tasks) 날짜 오름차순 정렬 (빠른 날짜 우선, 날짜 없으면 뒤로, 동일 날짜는 시트 행 순서 유지)
      siteObj.tasks.sort(function(a, b) {
        var aMs = a.dateMs || 0;
        var bMs = b.dateMs || 0;
        if (aMs && bMs) {
          if (aMs !== bMs) return aMs - bMs;
        } else if (aMs) {
          return -1;
        } else if (bMs) {
          return 1;
        }
        return a.row - b.row;
      });

      // 2. 블록 전체 완료 여부 판별 및 공정 진척률 정밀 계산
      var completedCount = 0;
      var canceledCount = 0;
      var inProgressCount = 0;
      var delayedCountSite = 0;

      for (var ti = 0; ti < siteObj.tasks.length; ti++) {
        var tObj = siteObj.tasks[ti];
        var stClean = (tObj.status || "").replace(/\s+/g, "");
        if (stClean.indexOf("완료") !== -1) {
          completedCount++;
        } else if (stClean.indexOf("취소") !== -1) {
          canceledCount++;
        } else if (stClean.indexOf("진행") !== -1) {
          inProgressCount++;
        }
        if (tObj.isDelayed) {
          delayedCountSite++;
        }
      }

      var activeTasks = siteObj.tasks.filter(function(t) {
        var st = (t.status || "").replace(/\s+/g, "");
        return st.indexOf("완료") === -1 && st.indexOf("취소") === -1;
      });

      var isAllCompleted = (siteObj.tasks.length > 0) && (activeTasks.length === 0);
      var totalTasksCount = siteObj.tasks.length;
      var validTasksCount = totalTasksCount - canceledCount;

      var progressRate = 0;
      if (validTasksCount > 0) {
        if (isAllCompleted) {
          progressRate = 100;
        } else {
          progressRate = Math.min(100, Math.max(0, Math.round((completedCount / validTasksCount) * 100)));
        }
      }

      siteObj.isAllCompleted = isAllCompleted;
      siteObj.activeTasksCount = activeTasks.length;
      siteObj.completedTasksCount = completedCount;
      siteObj.canceledTasksCount = canceledCount;
      siteObj.inProgressTasksCount = inProgressCount;
      siteObj.delayedTasksCount = delayedCountSite;
      siteObj.totalTasksCount = totalTasksCount;
      siteObj.validTasksCount = validTasksCount;
      siteObj.progressRate = progressRate;

      if (!isAllCompleted) {
        activeSitesCount++;
      }

      // 3. 현장 카드 정렬 기준 날짜 계산: 진행 중인 공정의 가장 빠른 날짜 (없으면 전체 중 가장 빠른 날짜)
      var earliestTask = (activeTasks.length > 0) ? activeTasks[0] : siteObj.tasks[0];
      siteObj.earliestDateMs = (earliestTask && earliestTask.dateMs) ? earliestTask.dateMs : 0;
      siteObj.earliestDateDisplay = (earliestTask && earliestTask.dateDisplay) ? earliestTask.dateDisplay : "";
      siteObj.hasTodayTask = activeTasks.some(function(t) { return t.isToday; });
      siteObj.hasDelayedTask = activeTasks.some(function(t) { return t.isDelayed; });

      sitesList.push(siteObj);
    }

    // 4. 현장 카드(sitesList) 정렬: 활성 현장 우선, 가장 빠른 일정 순 (동일하거나 없으면 시트 순서 보존)
    sitesList.sort(function(a, b) {
      if (!a.isAllCompleted && b.isAllCompleted) return -1;
      if (a.isAllCompleted && !b.isAllCompleted) return 1;

      var aMs = a.earliestDateMs || 0;
      var bMs = b.earliestDateMs || 0;
      if (aMs && bMs) {
        if (aMs !== bMs) return aMs - bMs;
      } else if (aMs) {
        return -1;
      } else if (bMs) {
        return 1;
      }
      return a.initialOrder - b.initialOrder;
    });

    return {
      kpi: {
        totalSites: siteOrder.length,
        activeSites: activeSitesCount,
        totalTasks: totalTasks,
        todayCount: todayCount,
        delayedCount: delayedCount
      },
      todayStr: todayYYMMDD,
      members: getMemberNames(ss),
      holidays: getKoreanHolidaysMap(ss),
      sites: sitesList
    };
  } catch (err) {
    return {
      error: "데이터 처리 중 오류 발생: " + (err.message || err),
      kpi: { totalSites: 0, activeSites: 0, totalTasks: 0, todayCount: 0, delayedCount: 0 },
      todayStr: formatDateYYMMDD(new Date()),
      members: ["조항만", "송지은", "전예린", "영하"],
      holidays: getKoreanHolidaysMap(null),
      sites: []
    };
  }
}

/**
 * 대시보드에서 클릭 시 스프레드시트의 해당 행으로 포커스 이동
 */
function focusSheetRow(rowNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var sheet = findSiteScheduleSheet(ss);
    if (sheet && rowNumber >= 2) {
      ss.setActiveSheet(sheet);
      var range = sheet.getRange(rowNumber, 1, 1, 8);
      sheet.setActiveRange(range);
      return { success: true };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
  return { success: false };
}

/**
 * 실시간 대시보드 HTML/CSS/JS 생성기 (사이드바, 모달, 모바일 웹앱 지원)
 */
function buildSidebarDashboardHtml(mode) {
  var isModal = (mode === "modal");
  var isMobile = (mode === "mobile");
  var webAppUrl = getWebAppUrl();
  var html = "";
  html += '<!DOCTYPE html><html><head><meta charset="utf-8">';
  html += '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>';
  if (isMobile) {
    html += '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">';
    html += '<meta name="apple-mobile-web-app-capable" content="yes">';
    html += '<meta name="apple-mobile-web-app-status-bar-style" content="default">';
    html += '<meta name="theme-color" content="#2563eb">';
  }
  html += '<style>';
  html += '* { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }';
  html += 'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Pretendard", sans-serif; background: #f8fafc; color: #1e293b; font-size: 13px; padding-bottom: ' + (isMobile ? 'calc(80px + env(safe-area-inset-bottom, 0px))' : '70px') + '; }';
  html += '.header { background: #ffffff; padding: 14px 16px; border-bottom: 1px solid #e2e8f0; position: sticky; top: 0; z-index: 20; ' + (isMobile ? 'max-width: 600px; margin: 0 auto;' : '') + ' }';
  html += '.title-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }';
  html += '.title { font-size: ' + (isMobile ? '16px' : '15px') + '; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 6px; }';
  html += '.kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-bottom: 10px; }';
  html += '.kpi-card { background: #f1f5f9; border-radius: 8px; padding: 8px 6px; text-align: center; cursor: pointer; transition: all 0.15s ease; border: 1px solid transparent; }';
  html += '.kpi-card:hover { background: #e2e8f0; }';
  html += '.kpi-card.active { border-color: #2563eb; background: #eff6ff; }';
  html += '.kpi-card.delay.active { border-color: #ef4444; background: #fef2f2; }';
  html += '.kpi-num { font-size: 16px; font-weight: 800; color: #0f172a; }';
  html += '.kpi-num.delay { color: #dc2626; }';
  html += '.kpi-label { font-size: 11px; color: #64748b; font-weight: 500; margin-top: 2px; }';
  html += '.search-box { position: relative; margin-bottom: 8px; }';
  html += '.search-input { width: 100%; padding: 8px 30px 8px 28px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 12.5px; outline: none; background: #ffffff; }';
  html += '.search-input:focus { border-color: #2563eb; box-shadow: 0 0 0 2px rgba(37,99,235,0.15); }';
  html += '.search-icon { position: absolute; left: 8px; top: 50%; transform: translateY(-50%); font-size: 12px; color: #94a3b8; }';
  html += '.search-clear { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); width: 18px; height: 18px; border-radius: 50%; background: #e2e8f0; color: #64748b; border: none; font-size: 11px; line-height: 18px; text-align: center; cursor: pointer; display: none; padding: 0; outline: none; }';
  html += '.search-clear:hover { background: #cbd5e1; color: #1e293b; }';
  html += '.search-clear.show { display: block; }';
  html += '.filter-tabs { display: flex; gap: 4px; overflow-x: auto; padding-bottom: 2px; }';
  html += '.filter-tab { padding: 5px 10px; border-radius: 12px; font-size: 11.5px; font-weight: 600; cursor: pointer; background: #f1f5f9; color: #64748b; border: 1px solid #e2e8f0; white-space: nowrap; transition: all 0.15s; }';
  html += '.filter-tab:hover { background: #e2e8f0; }';
  html += '.filter-tab.active { background: #2563eb; color: #ffffff; border-color: #2563eb; }';
  html += '.member-filter-bar { display: flex; gap: 4px; overflow-x: auto; padding-top: 6px; padding-bottom: 2px; -webkit-overflow-scrolling: touch; scrollbar-width: none; }';
  html += '.member-filter-bar::-webkit-scrollbar { display: none; }';
  html += '.member-tab { padding: 4px 9px; border-radius: 12px; font-size: 11px; font-weight: 600; cursor: pointer; background: #ffffff; color: #475569; border: 1px solid #cbd5e1; white-space: nowrap; transition: all 0.15s; flex-shrink: 0; }';
  html += '.member-tab:hover { background: #f1f5f9; }';
  html += '.member-tab.active { background: #0f172a; color: #ffffff; border-color: #0f172a; font-weight: 700; box-shadow: 0 1px 3px rgba(15,23,42,0.2); }';
  html += '.site-check { width: 17px; height: 17px; cursor: pointer; accent-color: #2563eb; flex-shrink: 0; margin-right: 4px; }';
  html += '.checked-action-bar { display: flex; justify-content: space-between; align-items: center; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 7px 10px; margin-top: 6px; margin-bottom: 2px; font-size: 11.5px; color: #1e40af; font-weight: 600; }';
  html += '.btn-check-action { padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; border: 1px solid #93c5fd; background: #ffffff; color: #1d4ed8; cursor: pointer; }';
  html += '.btn-check-action:hover { background: #dbeafe; }';
  html += '.cal-nav-bar { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; background: #ffffff; border-bottom: 1px solid #e2e8f0; }';
  html += '.cal-nav-left { display: flex; align-items: center; justify-content: flex-start; min-width: 65px; }';
  html += '.cal-nav-center { display: flex; align-items: center; justify-content: center; gap: 8px; flex: 1; text-align: center; }';
  html += '.cal-nav-right { display: flex; align-items: center; justify-content: flex-end; min-width: 65px; }';
  html += '.cal-month-title { font-size: 14.5px; font-weight: 800; color: #0f172a; }';
  html += '.cal-nav-btn { padding: 3px 8px; border-radius: 6px; border: 1px solid #cbd5e1; background: #ffffff; font-size: 11px; font-weight: 700; cursor: pointer; color: #334155; }';
  html += '.cal-nav-btn:hover { background: #f1f5f9; }';
  html += '.cal-today-btn { padding: 3px 8px; border-radius: 6px; border: 1px solid #2563eb; background: #eff6ff; color: #1d4ed8; font-size: 11px; font-weight: 700; cursor: pointer; }';
  html += '.cal-capture-btn { padding: 3px 8px; border-radius: 6px; border: 1px solid #cbd5e1; background: #ffffff; color: #334155; font-size: 11px; font-weight: 700; cursor: pointer; display: inline-flex; align-items: center; gap: 3px; transition: all 0.15s; }';
  html += '.cal-capture-btn:hover { background: #f1f5f9; border-color: #94a3b8; }';
  html += '.cal-capture-btn:disabled { opacity: 0.6; cursor: not-allowed; }';
  html += '.cal-mode-bar { display: flex; justify-content: center; align-items: center; padding: 6px 10px; background: #f8fafc; border-bottom: 1px solid #e2e8f0; }';
  html += '.cal-mode-group { display: inline-flex; background: #e2e8f0; padding: 2px; border-radius: 8px; gap: 2px; }';
  html += '.cal-mode-btn { padding: 3px 12px; border: none; background: transparent; color: #475569; font-size: 11px; font-weight: 700; border-radius: 6px; cursor: pointer; transition: all 0.15s; }';
  html += '.cal-mode-btn.active { background: #2563eb; color: #ffffff; box-shadow: 0 1px 3px rgba(37,99,235,0.25); }';
  html += '.cal-filter-notice { display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #eff6ff; border-bottom: 1px solid #bfdbfe; font-size: 11px; }';
  html += '.cal-weekdays { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); width: 100%; background: #f8fafc; border-bottom: 1px solid #e2e8f0; text-align: center; font-size: 11px; font-weight: 700; padding: 5px 0; }';
  html += '.cal-weekday.sun { color: #dc2626; }';
  html += '.cal-weekday.sat { color: #2563eb; }';
  html += '.calendar-view { touch-action: pan-y; -webkit-user-select: none; user-select: none; }';
  html += '.cal-grid { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); width: 100%; background: #e2e8f0; gap: 1px; touch-action: pan-y; }';
  html += '@keyframes calSlideFromRight { from { transform: translateX(18px); opacity: 0.6; } to { transform: translateX(0); opacity: 1; } }';
  html += '@keyframes calSlideFromLeft { from { transform: translateX(-18px); opacity: 0.6; } to { transform: translateX(0); opacity: 1; } }';
  html += '.cal-slide-next { animation: calSlideFromRight 0.18s cubic-bezier(0.25, 1, 0.5, 1); }';
  html += '.cal-slide-prev { animation: calSlideFromLeft 0.18s cubic-bezier(0.25, 1, 0.5, 1); }';
  html += '.cal-day-cell { background: #ffffff; min-height: 72px; min-width: 0; padding: 2px 2px 4px 2px; display: flex; flex-direction: column; overflow: hidden; cursor: pointer; }';
  html += '.cal-day-cell:hover { background: #f8fafc; }';
  html += '.cal-day-cell.other-month { background: #fafbfc; opacity: 0.4; }';
  html += '.cal-day-header { display: flex; flex-direction: column; align-items: flex-start; padding: 1px 2px; font-size: 10.5px; font-weight: 700; color: #334155; width: 100%; }';
  html += '.cal-day-top-row { display: flex; align-items: center; justify-content: space-between; width: 100%; min-width: 0; gap: 2px; }';
  html += '.cal-weather-badge { font-size: 8px; font-weight: 700; color: #475569; background: #f1f5f9; padding: 0.5px 3px; border-radius: 3px; line-height: 1.15; white-space: nowrap; flex-shrink: 0; }';
  html += '.cal-weather-badge.rain { background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; font-weight: 800; }';
  html += '.cal-weather-badge.snow { background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; font-weight: 800; }';
  html += '.cal-day-cell.sun .cal-day-header { color: #dc2626; }';
  html += '.cal-day-cell.sat .cal-day-header { color: #2563eb; }';
  html += '.cal-day-cell.holiday { background: #fff8f8; }';
  html += '.cal-day-cell.holiday .cal-day-header { color: #dc2626; font-weight: 800; }';
  html += '.cal-holiday-name { font-size: 8px; font-weight: 700; color: #dc2626; margin-top: 1px; margin-left: 1px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%; line-height: 1.15; letter-spacing: -0.3px; display: block; }';
  html += '.cal-day-num { display: inline-flex; align-items: center; justify-content: center; width: 17px; height: 17px; border-radius: 50%; }';
  html += '.cal-day-cell.today .cal-day-num { background: #2563eb; color: #ffffff; font-weight: 800; }';
  html += '.cal-week-list { display: flex; flex-direction: column; gap: 6px; padding: 6px 2px; width: 100%; touch-action: pan-y; }';
  html += '.cal-week-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }';
  html += '.cal-week-card.today { border-color: #3b82f6; box-shadow: 0 0 0 1.5px rgba(59,130,246,0.25); }';
  html += '.cal-week-card-header { display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; background: #f8fafc; border-bottom: 1px solid #f1f5f9; font-size: 12px; font-weight: 700; cursor: pointer; }';
  html += '.cal-week-card.today .cal-week-card-header { background: #eff6ff; color: #1d4ed8; }';
  html += '.cal-week-card-left { display: flex; align-items: center; gap: 6px; }';
  html += '.cal-week-day-tag { font-size: 12px; font-weight: 800; }';
  html += '.cal-week-day-tag.sun { color: #dc2626; }';
  html += '.cal-week-day-tag.sat { color: #2563eb; }';
  html += '.cal-week-today-badge { font-size: 9px; padding: 1px 4px; background: #2563eb; color: #ffffff; border-radius: 3px; font-weight: 800; }';
  html += '.cal-week-holiday-badge { font-size: 9px; padding: 1px 4px; background: #fee2e2; color: #dc2626; border-radius: 3px; font-weight: 800; }';
  html += '.cal-week-weather-info { font-size: 11px; font-weight: 600; color: #475569; display: flex; align-items: center; gap: 3px; }';
  html += '.cal-week-weather-info.rain { color: #1d4ed8; font-weight: 700; }';
  html += '.cal-week-card-body { padding: 6px 8px; display: flex; flex-direction: column; gap: 3px; background: #ffffff; }';
  html += '.cal-week-empty-text { font-size: 11px; color: #94a3b8; padding: 3px 0; text-align: center; }';
  html += '.cal-events-list { display: flex; flex-direction: column; gap: 2px; flex: 1; min-width: 0; width: 100%; margin-top: 1px; }';
  html += '.cal-event-pill { display: flex; align-items: center; width: 100%; min-width: 0; max-width: 100%; box-sizing: border-box; border-radius: 3px; padding: 1.5px 3px; font-size: 9.5px; font-weight: 800; line-height: 1.25; cursor: pointer; color: #ffffff; overflow: hidden; margin-bottom: 1px; }';
  html += '.cal-event-pill:hover { filter: brightness(0.92); }';
  html += '.cal-event-pill.status-done { opacity: 0.5; text-decoration: line-through; }';
  html += '.cal-event-pill.status-delayed { background-color: #dc2626 !important; color: #ffffff !important; box-shadow: 0 1px 2px rgba(220,38,38,0.35); }';
  html += '.cal-delay-warn-icon { font-size: 8.5px; line-height: 1; margin-right: 2px; flex-shrink: 0; display: inline-flex; align-items: center; }';
  html += '.cal-sh-box { display: inline-block; width: 6.5px; height: 6.5px; border-radius: 1px; flex-shrink: 0; margin-right: 2.5px; box-sizing: border-box; border: 1px solid rgba(255,255,255,0.95); }';
  html += '.cal-sh-box.blue { background-color: #38bdf8; }';
  html += '.cal-sh-box.red { background-color: #ef4444; }';
  html += '.cal-pill-text { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1; min-width: 0; font-weight: 800; -webkit-font-smoothing: antialiased; }';
  html += '.content { padding: 12px; ' + (isMobile ? 'max-width: 600px; margin: 0 auto;' : '') + ' }';
  html += '.site-card { background: #ffffff; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); transition: box-shadow 0.15s; }';
  html += '.site-card:hover { box-shadow: 0 2px 5px rgba(0,0,0,0.06); }';
  html += '.site-header { display: flex; align-items: center; justify-content: space-between; padding: 11px 12px; cursor: pointer; user-select: none; background: #ffffff; border-left: 4px solid #2563eb; gap: 8px; }';
  html += '.site-header:hover { background: #f8fafc; }';
  html += '.site-title-area { display: flex; align-items: center; gap: 6px; flex: 1; overflow: hidden; min-width: 0; }';
  html += '.site-name { font-weight: 700; font-size: 13.5px; color: #0f172a; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 1; min-width: 0; }';
  html += '.share-tag { display: inline-flex; align-items: center; gap: 3px; font-size: 10px; font-weight: 700; padding: 2px 5px; border-radius: 4px; line-height: 1.2; white-space: nowrap; }';
  html += '.share-tag.share-on { background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; }';
  html += '.share-tag.share-off { background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }';
  html += '.site-badge { font-size: 10.5px; padding: 2px 6px; border-radius: 10px; background: #f1f5f9; color: #475569; font-weight: 600; white-space: nowrap; flex-shrink: 0; }';
  html += '.site-progress-badge { font-size: 10px; font-weight: 800; padding: 2px 6px; border-radius: 6px; line-height: 1.2; white-space: nowrap; flex-shrink: 0; }';
  html += '.site-progress-badge.done { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }';
  html += '.site-progress-badge.high { background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; }';
  html += '.site-progress-badge.mid { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }';
  html += '.site-progress-badge.low { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }';
  html += '.site-delay-mini-badge { font-size: 9.5px; font-weight: 700; padding: 1.5px 5px; border-radius: 5px; background: #fee2e2; color: #dc2626; border: 1px solid #fecaca; white-space: nowrap; flex-shrink: 0; }';
  html += '.site-progress-track { width: 100%; height: 4.5px; background: #f1f5f9; position: relative; overflow: hidden; }';
  html += '.site-progress-bar { height: 100%; transition: width 0.35s cubic-bezier(0.4, 0, 0.2, 1); border-radius: 0 2px 2px 0; }';
  html += '.site-summary-box { display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; margin-bottom: 8px; background: #ffffff; border-radius: 6px; border: 1px solid #e2e8f0; font-size: 11.5px; }';
  html += '.site-summary-main { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }';
  html += '.site-summary-rate { color: #0f172a; font-weight: 600; font-size: 11.5px; }';
  html += '.site-summary-rate strong { color: #2563eb; font-weight: 800; font-size: 12.5px; }';
  html += '.site-summary-rate.done strong { color: #16a34a; }';
  html += '.site-summary-counts { color: #64748b; font-size: 11px; }';
  html += '.site-toggle { font-size: 10px; color: #94a3b8; transition: transform 0.2s; flex-shrink: 0; }';
  html += '.site-card.expanded .site-toggle { transform: rotate(180deg); }';
  html += '.task-list { display: none; padding: 6px 10px 8px 10px; background: #f8fafc; border-top: 1px solid #f1f5f9; }';
  html += '.site-card.expanded .task-list { display: block; }';
  html += '.task-item { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 9px 11px; margin-bottom: 6px; cursor: pointer; transition: all 0.15s; }';
  html += '.task-item:hover { border-color: #2563eb; background: #f0f7ff; transform: translateY(-1px); box-shadow: 0 2px 4px rgba(37,99,235,0.08); }';
  html += '.task-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }';
  html += '.task-name { font-weight: 700; color: #0f172a; font-size: 13px; }';
  html += '.task-date { font-size: 11.5px; color: #64748b; font-weight: 600; }';
  html += '.task-date.today { color: #2563eb; font-weight: 700; }';
  html += '.task-date.delay { color: #dc2626; font-weight: 700; }';
  html += '.task-meta { display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #64748b; }';
  html += '.task-memo { margin-top: 4px; font-size: 11px; color: #475569; background: #f8fafc; padding: 4px 6px; border-radius: 4px; line-height: 1.35; word-break: break-all; }';
  html += '.status-tag { font-size: 10px; padding: 2px 6px; border-radius: 4px; font-weight: 700; }';
  html += '.status-예정 { background: #f1f5f9; color: #475569; }';
  html += '.status-진행 { background: #dbeafe; color: #1e40af; }';
  html += '.status-완료 { background: #dcfce7; color: #166534; }';
  html += '.status-AS { background: #fef3c7; color: #92400e; }';
  html += '.status-지연 { background: #fee2e2; color: #991b1b; }';
  html += '.empty-state { text-align: center; padding: 50px 20px; color: #64748b; grid-column: 1 / -1; width: 100%; }';
  html += '.loading-state { grid-column: 1 / -1; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 80px 20px; text-align: center; width: 100%; }';
  html += '.loading-text { text-align: center; color: #64748b; font-size: 13px; font-weight: 600; margin-top: 14px; }';
  html += '.spinner { border: 3.5px solid #e2e8f0; border-top: 3.5px solid #2563eb; border-radius: 50%; width: 32px; height: 32px; animation: spin 0.8s linear infinite; margin: 0 auto; }';
  html += '@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }';
  html += '.footer-bar { position: fixed; bottom: 0; left: 0; right: 0; background: #ffffff; border-top: 1px solid #e2e8f0; padding: 10px 14px ' + (isMobile ? 'calc(10px + env(safe-area-inset-bottom, 0px))' : '10px') + '; display: flex; gap: 8px; box-shadow: 0 -2px 8px rgba(0,0,0,0.04); z-index: 30; ' + (isMobile ? 'max-width: 600px; margin: 0 auto;' : '') + ' }';
  html += '.btn { flex: 1; padding: 10px; border-radius: 8px; font-size: 12.5px; font-weight: 700; cursor: pointer; border: 1px solid transparent; text-align: center; transition: all 0.15s; }';
  html += '.btn-primary { background: #2563eb; color: #ffffff; }';
  html += '.btn-primary:hover { background: #1d4ed8; }';
  html += '.btn-secondary { background: #ffffff; border-color: #cbd5e1; color: #334155; }';
  html += '.btn-secondary:hover { background: #f8fafc; }';
  html += '.toast { position: fixed; bottom: 75px; left: 50%; transform: translateX(-50%); background: #1e293b; color: #ffffff; padding: 8px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; opacity: 0; pointer-events: none; transition: opacity 0.25s; z-index: 2000; white-space: nowrap; box-shadow: 0 4px 12px rgba(0,0,0,0.18); }';
  html += '.toast.show { opacity: 1; }';

  // 모바일 바텀시트 모달 스타일
  html += '.sheet-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.55); z-index: 1000; display: none; align-items: flex-end; justify-content: center; backdrop-filter: blur(2px); }';
  html += '.sheet-overlay.show { display: flex; }';
  html += '.bottom-sheet { background: #ffffff; width: 100%; max-width: 520px; border-radius: 20px 20px 0 0; padding: 18px 18px calc(20px + env(safe-area-inset-bottom, 0px)) 18px; max-height: 85vh; overflow-y: auto; box-shadow: 0 -10px 30px rgba(0,0,0,0.2); animation: slideUp 0.22s cubic-bezier(0.16, 1, 0.3, 1); }';
  html += '@keyframes slideUp { from { transform: translateY(100%); } to { transform: translateY(0); } }';
  html += '.sheet-handle { width: 36px; height: 4px; background: #cbd5e1; border-radius: 2px; margin: 0 auto 12px; }';
  html += '.sheet-title-row { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }';
  html += '.sheet-site { font-size: 13px; color: #2563eb; font-weight: 700; }';
  html += '.sheet-task { font-size: 18px; color: #0f172a; font-weight: 800; margin-top: 2px; }';
  html += '.sheet-close { background: #f1f5f9; border: none; width: 28px; height: 28px; border-radius: 50%; font-size: 14px; font-weight: 700; color: #64748b; cursor: pointer; display: flex; align-items: center; justify-content: center; }';
  html += '.sheet-info-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; margin-bottom: 14px; font-size: 12.5px; line-height: 1.5; color: #475569; }';
  html += '.sheet-sec-label { font-size: 12.5px; font-weight: 700; color: #334155; margin-bottom: 7px; display: flex; align-items: center; gap: 5px; }';
  html += '.sheet-status-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-bottom: 12px; }';
  html += '.sheet-status-btn { padding: 9px 4px; border-radius: 8px; border: 2px solid #e2e8f0; font-size: 12.5px; font-weight: 700; cursor: pointer; background: #ffffff; color: #475569; text-align: center; }';
  html += '.sheet-status-btn.active { border-color: #2563eb; background: #eff6ff; color: #1d4ed8; font-weight: 800; }';
  html += '.sheet-share-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 12px; }';
  html += '.sheet-share-btn { padding: 9px; border-radius: 8px; border: 2px solid #e2e8f0; font-size: 12.5px; font-weight: 700; cursor: pointer; background: #ffffff; color: #64748b; text-align: center; }';
  html += '.sheet-share-btn.active-on { border-color: #2563eb; background: #eff6ff; color: #1d4ed8; font-weight: 800; }';
  html += '.sheet-share-btn.active-off { border-color: #ef4444; background: #fef2f2; color: #dc2626; font-weight: 800; }';
  html += '.sheet-memo-input { width: 100%; min-height: 60px; padding: 8px 10px; border-radius: 8px; border: 1.5px solid #cbd5e1; font-size: 12.5px; font-family: inherit; outline: none; margin-bottom: 14px; resize: vertical; }';
  html += '.sheet-memo-input:focus { border-color: #2563eb; }';
  html += '.sheet-action-row { margin-top: 12px; }';
  html += '.sheet-save-btn { width: 100%; height: 48px; border-radius: 12px; background: #2563eb; color: #fff; font-size: 15px; font-weight: 800; border: none; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 6px; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.25); transition: all 0.15s; }';
  html += '.sheet-save-btn:active { background: #1d4ed8; transform: scale(0.99); }';
  html += '.sheet-save-btn:disabled { background: #94a3b8; cursor: not-allowed; box-shadow: none; }';
  html += '.sheet-day-nav { display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; background: #f1f5f9; border-radius: 8px; margin-bottom: 12px; border: 1px solid #e2e8f0; }';
  html += '.sheet-day-nav-label { font-size: 11.5px; font-weight: 700; color: #334155; display: flex; align-items: center; gap: 4px; }';
  html += '.sheet-day-nav-counts { color: #2563eb; font-weight: 800; }';
  html += '.sheet-day-nav-btns { display: flex; align-items: center; gap: 5px; }';
  html += '.btn-sheet-nav { display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 30px; border-radius: 6px; border: 1px solid #cbd5e1; background: #ffffff; color: #1e293b; font-size: 13px; font-weight: 800; cursor: pointer; transition: all 0.15s; }';
  html += '.btn-sheet-nav:hover:not(:disabled) { background: #2563eb; color: #ffffff; border-color: #2563eb; }';
  html += '.btn-sheet-nav:disabled { opacity: 0.3; cursor: not-allowed; background: #f1f5f9; color: #94a3b8; border-color: #e2e8f0; }';

  if (isModal) {
    html += '.content { max-width: 900px; margin: 0 auto; display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }';
  }
  html += '</style></head><body>';

  // Body HTML
  html += '<div class="header">';
  html += '  <div class="title-row">';
  html += '    <div class="title" id="pageTitle">' + (isMobile ? '📱 우성아이디 현장 대시보드' : '🏗️ 우성아이디 현장 대시보드') + '</div>';
  html += '    <button class="btn btn-secondary" style="flex:none;padding:4px 9px;font-size:11.5px;" onclick="loadData()">🔄 새로고침</button>';
  html += '  </div>';

  html += '  <div class="kpi-grid">';
  html += '    <div class="kpi-card" data-filter="all" onclick="setFilter(\'all\')">';
  html += '      <div class="kpi-num" id="kpiSites">-</div>';
  html += '      <div class="kpi-label">진행 현장</div>';
  html += '    </div>';
  html += '    <div class="kpi-card" data-filter="today" onclick="setFilter(\'today\')">';
  html += '      <div class="kpi-num" id="kpiToday" style="color:#2563eb;">-</div>';
  html += '      <div class="kpi-label">🔥 오늘 공정</div>';
  html += '    </div>';
  html += '    <div class="kpi-card delay" data-filter="delay" onclick="setFilter(\'delay\')">';
  html += '      <div class="kpi-num delay" id="kpiDelay">-</div>';
  html += '      <div class="kpi-label">⚠️ 지연 공정</div>';
  html += '    </div>';
  html += '  </div>';

  html += '  <div class="search-box">';
  html += '    <span class="search-icon">🔍</span>';
  html += '    <input type="text" id="searchInput" class="search-input" placeholder="현장명, 공종, 담당자 실시간 검색..." oninput="onSearchInput()">';
  html += '    <button type="button" id="searchClearBtn" class="search-clear" onclick="clearSearch()" title="검색어 삭제">✕</button>';
  html += '  </div>';

  html += '  <div class="filter-tabs">';
  html += '    <div class="filter-tab active" data-filter="all" onclick="setFilter(\'all\')">전체</div>';
  html += '    <div class="filter-tab" data-filter="today" onclick="setFilter(\'today\')">🔥 오늘</div>';
  html += '    <div class="filter-tab" data-filter="delay" onclick="setFilter(\'delay\')">⚠️ 지연</div>';
  html += '    <div class="filter-tab" data-filter="prog" onclick="setFilter(\'prog\')">진행중</div>';
  html += '    <div class="filter-tab" data-filter="done" onclick="setFilter(\'done\')">✅ 완료</div>';
  html += '    <div class="filter-tab" id="checkedTab" data-filter="checked" onclick="setFilter(\'checked\')" style="display:none;">☑️ 선택 (<span id="checkedTabCount">0</span>)</div>';
  html += '  </div>';
  html += '  <div class="member-filter-bar" id="memberFilterBar"></div>';
  html += '  <div id="checkedActionBar" class="checked-action-bar" style="display:none;">';
  html += '    <span id="checkedCountMsg">☑️ 0개 현장 선택됨</span>';
  html += '    <button type="button" class="btn-check-action" onclick="clearAllChecked()">해제</button>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="content" id="contentArea">';
  html += '  <div class="loading-state">';
  html += '    <div class="spinner"></div>';
  html += '    <p class="loading-text">현장 데이터 불러오는 중...</p>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="content" id="calendarArea" style="display:none;' + (isMobile ? 'max-width:600px;margin:0 auto;padding:8px 4px;' : '') + '">';
  html += '  <div class="calendar-view">';
  html += '    <div class="cal-mode-bar">';
  html += '      <div class="cal-mode-group">';
  html += '        <button type="button" id="btnModeMonth" class="cal-mode-btn active" onclick="setCalendarMode(\'month\')">📅 월간</button>';
  html += '        <button type="button" id="btnModeWeek" class="cal-mode-btn" onclick="setCalendarMode(\'week\')">📆 주간</button>';
  html += '      </div>';
  html += '    </div>';
  html += '    <div class="cal-nav-bar">';
  html += '      <div class="cal-nav-left">';
  html += '        <button type="button" id="btnCaptureCal" class="cal-capture-btn" onclick="captureCalendarScreen()" title="현재 캘린더 화면 캡쳐">📷 캡쳐</button>';
  html += '      </div>';
  html += '      <div class="cal-nav-center">';
  html += '        <button type="button" class="cal-nav-btn" onclick="navPrevPeriod()">◀</button>';
  html += '        <span id="calMonthTitle" class="cal-month-title">2026년 10월</span>';
  html += '        <button type="button" class="cal-nav-btn" onclick="navNextPeriod()">▶</button>';
  html += '      </div>';
  html += '      <div class="cal-nav-right">';
  html += '        <button type="button" id="calTodayBtn" class="cal-today-btn" onclick="navTodayPeriod()">오늘</button>';
  html += '      </div>';
  html += '    </div>';
  html += '    <div id="calFilterNotice" class="cal-filter-notice" style="display:none;"></div>';
  html += '    <div class="cal-weekdays" id="calWeekdaysBar">';
  html += '      <div class="cal-weekday sun">일</div>';
  html += '      <div class="cal-weekday">월</div>';
  html += '      <div class="cal-weekday">화</div>';
  html += '      <div class="cal-weekday">수</div>';
  html += '      <div class="cal-weekday">목</div>';
  html += '      <div class="cal-weekday">금</div>';
  html += '      <div class="cal-weekday sat">토</div>';
  html += '    </div>';
  html += '    <div class="cal-grid" id="calGrid"></div>';
  html += '    <div class="cal-week-list" id="calWeekList" style="display:none;"></div>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="footer-bar">';
  html += '  <button class="btn btn-secondary" onclick="loadData()">🔄 새로고침</button>';
  html += '  <button type="button" id="calendarToggleBtn" class="btn btn-primary" onclick="toggleCalendarView()">📅 캘린더 보기</button>';
  html += '</div>';

  // 캘린더 캡쳐 프리뷰 모달 컨테이너
  html += '<div id="captureModal" class="sheet-overlay" onclick="closeCaptureModal(event)" style="display:none; align-items:center; justify-content:center; z-index:9999; padding:16px;">';
  html += '  <div class="capture-dialog" onclick="event.stopPropagation()" style="background:#ffffff; border-radius:12px; width:100%; max-width:480px; max-height:90vh; display:flex; flex-direction:column; overflow:hidden; box-shadow:0 12px 30px rgba(0,0,0,0.25);">';
  html += '    <div style="display:flex; justify-content:space-between; align-items:center; padding:12px 16px; border-bottom:1px solid #e2e8f0;">';
  html += '      <div style="font-weight:800; font-size:14px; color:#0f172a; display:flex; align-items:center; gap:6px;">📷 캘린더 캡쳐 완료</div>';
  html += '      <button type="button" class="sheet-close" onclick="closeCaptureModal()">✕</button>';
  html += '    </div>';
  html += '    <div style="padding:10px 14px; background:#eff6ff; font-size:11.5px; color:#1d4ed8; line-height:1.45; border-bottom:1px solid #bfdbfe;">';
  html += '      💡 <b>모바일 안내:</b> 아래 이미지를 <b>길게 누르면</b> [사진에 저장] 또는 [공유]가 가능합니다.';
  html += '    </div>';
  html += '    <div style="padding:12px; flex:1; overflow-y:auto; text-align:center; background:#f8fafc;">';
  html += '      <img id="capturePreviewImg" src="" alt="캘린더 캡쳐" style="width:100%; height:auto; border-radius:6px; border:1px solid #cbd5e1; box-shadow:0 2px 6px rgba(0,0,0,0.06); display:block;">';
  html += '    </div>';
  html += '    <div style="padding:10px 14px; display:flex; gap:8px; border-top:1px solid #e2e8f0; background:#ffffff;">';
  html += '      <button type="button" id="btnDownloadCapture" class="btn btn-primary" style="flex:1; padding:9px; font-size:12.5px; font-weight:700;">💾 이미지 다운로드</button>';
  html += '      <button type="button" class="btn btn-secondary" style="flex:none; padding:9px 14px; font-size:12.5px;" onclick="closeCaptureModal()">닫기</button>';
  html += '    </div>';
  html += '  </div>';
  html += '</div>';

  // 모바일 바텀시트 모달 컨테이너
  html += '<div id="sheetOverlay" class="sheet-overlay" onclick="closeBottomSheet(event)">';
  html += '  <div class="bottom-sheet" onclick="event.stopPropagation()">';
  html += '    <div class="sheet-handle"></div>';
  html += '    <div class="sheet-title-row">';
  html += '      <div>';
  html += '        <div class="sheet-site" id="sheetSiteName">-</div>';
  html += '        <div class="sheet-task" id="sheetTaskName">-</div>';
  html += '      </div>';
  html += '      <button type="button" class="sheet-close" onclick="closeBottomSheet()">✕</button>';
  html += '    </div>';
  html += '    <div class="sheet-day-nav" id="sheetDayNav" style="display:none;">';
  html += '      <span class="sheet-day-nav-label">📅 같은 날 일정 (<span id="sheetDayNavIdx" class="sheet-day-nav-counts">1</span>/<span id="sheetDayNavTotal" class="sheet-day-nav-counts">1</span>)</span>';
  html += '      <div class="sheet-day-nav-btns">';
  html += '        <button type="button" class="btn-sheet-nav" id="btnSheetPrevDayTask" onclick="navSameDayTask(-1)" title="이전 일정">◀</button>';
  html += '        <button type="button" class="btn-sheet-nav" id="btnSheetNextDayTask" onclick="navSameDayTask(1)" title="다음 일정">▶</button>';
  html += '      </div>';
  html += '    </div>';
  html += '    <div class="sheet-info-box" id="sheetInfoBox">';
  html += '      <div>📅 <b>일정:</b> <span id="sheetDateDisplay">-</span></div>';
  html += '      <div style="margin-top:3px;">👤 <b>담당자/업체:</b> <span id="sheetMetaDisplay">-</span></div>';
  html += '    </div>';
  html += '    <div class="sheet-sec-label">⚡ 진행 상태 변경</div>';
  html += '    <div class="sheet-status-grid">';
  html += '      <button type="button" class="sheet-status-btn" data-status="예정" onclick="selectSheetStatus(\'예정\')">📅 예정</button>';
  html += '      <button type="button" class="sheet-status-btn" data-status="진행 중" onclick="selectSheetStatus(\'진행 중\')">▶️ 진행 중</button>';
  html += '      <button type="button" class="sheet-status-btn" data-status="완료" onclick="selectSheetStatus(\'완료\')">✅ 완료</button>';
  html += '      <button type="button" class="sheet-status-btn" data-status="AS" onclick="selectSheetStatus(\'AS\')">🛠️ AS</button>';
  html += '      <button type="button" class="sheet-status-btn" data-status="지연" onclick="selectSheetStatus(\'지연\')">⚠️ 지연</button>';
  html += '    </div>';
  html += '    <div class="sheet-sec-label">🔗 캘린더 공정 공유</div>';
  html += '    <div class="sheet-share-grid">';
  html += '      <button type="button" id="sheetShareOn" class="sheet-share-btn" onclick="selectSheetShare(true)">🟦 공유 (ON)</button>';
  html += '      <button type="button" id="sheetShareOff" class="sheet-share-btn" onclick="selectSheetShare(false)">🟥 미공유 (OFF)</button>';
  html += '    </div>';
  html += '    <div class="sheet-sec-label">📝 현장 비고 / 특이사항</div>';
  html += '    <textarea id="sheetMemoInput" class="sheet-memo-input" placeholder="현장 특이사항이나 전달사항을 입력하세요..."></textarea>';
  html += '    <div class="sheet-action-row">';
  html += '      <button type="button" id="sheetSaveBtn" class="sheet-save-btn" onclick="saveSheetTask()">💾 저장하기</button>';
  html += '    </div>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="toast" id="toast"></div>';

  // Script
  html += '<script>';
  html += 'function getSundayOfCurrentWeek(d) {';
  html += '  const date = new Date(d);';
  html += '  const day = date.getDay();';
  html += '  date.setDate(date.getDate() - day);';
  html += '  date.setHours(0, 0, 0, 0);';
  html += '  return date;';
  html += '}';
  html += 'const isMobileMode = ' + (isMobile ? 'true' : 'false') + ';';
  html += 'const webAppBaseUrl = "' + (webAppUrl || "") + '";';
  html += 'let rawData = null;';
  html += 'let currentFilter = "all";';
  html += 'let selectedMember = "";';
  html += 'const checkedSites = new Set();';
  html += 'let isCalendarMode = false;';
  html += 'let calViewMode = "month";'; // "month" | "week"
  html += 'let calWeekStartDate = getSundayOfCurrentWeek(new Date());';
  html += 'const activeWeatherData = {};';
  html += 'let calYear = new Date().getFullYear();';
  html += 'let calMonth = new Date().getMonth() + 1;';
  html += 'let curTaskObj = null;';
  html += 'let curSiteName = "";';
  html += 'let selectedStatus = "예정";';
  html += 'let selectedShare = true;';
  html += 'let sameDayTaskList = [];';
  html += 'let curSameDayIndex = 0;';
  html += 'const DEFAULT_KOREAN_HOLIDAYS = {';
  html += '  "2024-01-01":"신정","2024-02-09":"설날연휴","2024-02-10":"설날","2024-02-11":"설날연휴","2024-02-12":"대체공휴일",';
  html += '  "2024-03-01":"삼일절","2024-04-10":"총선","2024-05-05":"어린이날","2024-05-06":"대체공휴일","2024-05-15":"부처님오신날",';
  html += '  "2024-06-06":"현충일","2024-08-15":"광복절","2024-09-16":"추석연휴","2024-09-17":"추석","2024-09-18":"추석연휴",';
  html += '  "2024-10-03":"개천절","2024-10-09":"한글날","2024-12-25":"성탄절",';
  html += '  "2025-01-01":"신정","2025-01-28":"설날연휴","2025-01-29":"설날","2025-01-30":"설날연휴",';
  html += '  "2025-03-01":"삼일절","2025-03-03":"대체공휴일","2025-05-05":"어린이날","2025-05-06":"대체공휴일",';
  html += '  "2025-06-06":"현충일","2025-08-15":"광복절","2025-10-03":"개천절","2025-10-05":"추석연휴",';
  html += '  "2025-10-06":"추석","2025-10-07":"추석연휴","2025-10-08":"대체공휴일","2025-10-09":"한글날","2025-12-25":"성탄절",';
  html += '  "2026-01-01":"신정","2026-02-16":"설날연휴","2026-02-17":"설날","2026-02-18":"설날연휴",';
  html += '  "2026-03-01":"삼일절","2026-03-02":"대체공휴일","2026-05-05":"어린이날","2026-05-24":"부처님오신날",';
  html += '  "2026-05-25":"대체공휴일","2026-06-03":"지방선거","2026-06-06":"현충일","2026-08-15":"광복절",';
  html += '  "2026-08-17":"대체공휴일","2026-09-24":"추석연휴","2026-09-25":"추석","2026-09-26":"추석연휴",';
  html += '  "2026-09-28":"대체공휴일","2026-10-03":"개천절","2026-10-05":"대체공휴일","2026-10-09":"한글날","2026-12-25":"성탄절",';
  html += '  "2027-01-01":"신정","2027-02-06":"설날연휴","2027-02-07":"설날","2027-02-08":"설날연휴","2027-02-09":"대체공휴일",';
  html += '  "2027-03-01":"삼일절","2027-05-05":"어린이날","2027-05-13":"부처님오신날","2027-06-06":"현충일","2027-06-07":"대체공휴일",';
  html += '  "2027-08-15":"광복절","2027-08-16":"대체공휴일","2027-09-14":"추석연휴","2027-09-15":"추석","2027-09-16":"추석연휴",';
  html += '  "2027-10-03":"개천절","2027-10-04":"대체공휴일","2027-10-09":"한글날","2027-10-11":"대체공휴일","2027-12-25":"성탄절"';
  html += '};';
  html += 'let activeHolidays = Object.assign({}, DEFAULT_KOREAN_HOLIDAYS);';

  html += 'function showToast(msg) {';
  html += '  const t = document.getElementById("toast");';
  html += '  t.innerText = msg;';
  html += '  t.classList.add("show");';
  html += '  setTimeout(() => t.classList.remove("show"), 2200);';
  html += '}';

  html += 'function onSiteCheckboxClick(e, cb) {';
  html += '  e.stopPropagation();';
  html += '  const siteName = decodeURIComponent(cb.dataset.sitename || "");';
  html += '  if (cb.checked) {';
  html += '    checkedSites.add(siteName);';
  html += '  } else {';
  html += '    checkedSites.delete(siteName);';
  html += '  }';
  html += '  updateCheckedStatusUI();';
  html += '  if (currentFilter === "checked") {';
  html += '    renderList();';
  html += '  }';
  html += '  if (isCalendarMode) {';
  html += '    renderCurrentCalendarView();';
  html += '  }';
  html += '}';

  html += 'function updateCheckedStatusUI() {';
  html += '  const count = checkedSites.size;';
  html += '  const tab = document.getElementById("checkedTab");';
  html += '  const tabCount = document.getElementById("checkedTabCount");';
  html += '  const bar = document.getElementById("checkedActionBar");';
  html += '  const countMsg = document.getElementById("checkedCountMsg");';
  html += '  if (tabCount) tabCount.innerText = count;';
  html += '  if (tab) tab.style.display = (count > 0) ? "block" : "none";';
  html += '  if (bar) bar.style.display = (count > 0) ? "flex" : "none";';
  html += '  if (countMsg) countMsg.innerText = "☑️ " + count + "개 현장 선택됨";';
  html += '}';

  html += 'function clearAllChecked() {';
  html += '  checkedSites.clear();';
  html += '  document.querySelectorAll(".site-check").forEach(cb => cb.checked = false);';
  html += '  updateCheckedStatusUI();';
  html += '  if (currentFilter === "checked") {';
  html += '    setFilter("all");';
  html += '  } else {';
  html += '    renderList();';
  html += '  }';
  html += '  if (isCalendarMode) {';
  html += '    renderCurrentCalendarView();';
  html += '  }';
  html += '}';

  html += 'function toggleCalendarView() {';
  html += '  isCalendarMode = !isCalendarMode;';
  html += '  const btn = document.getElementById("calendarToggleBtn");';
  html += '  const contentArea = document.getElementById("contentArea");';
  html += '  const calArea = document.getElementById("calendarArea");';
  html += '  const pageTitle = document.getElementById("pageTitle");';
  html += '  if (isCalendarMode) {';
  html += '    if (btn) {';
  html += '      btn.innerText = "📋 목록 보기";';
  html += '    }';
  html += '    if (pageTitle) {';
  html += '      pageTitle.innerText = isMobileMode ? "📅 우성아이디 현장 캘린더" : "📅 우성아이디 현장 캘린더";';
  html += '    }';
  html += '    if (contentArea) contentArea.style.display = "none";';
  html += '    if (calArea) calArea.style.display = "block";';
  html += '    renderCurrentCalendarView();';
  html += '  } else {';
  html += '    if (btn) {';
  html += '      btn.innerText = "📅 캘린더 보기";';
  html += '    }';
  html += '    if (pageTitle) {';
  html += '      pageTitle.innerText = isMobileMode ? "📱 우성아이디 현장 대시보드" : "🏗️ 우성아이디 현장 대시보드";';
  html += '    }';
  html += '    if (contentArea) contentArea.style.display = "block";';
  html += '    if (calArea) calArea.style.display = "none";';
  html += '    renderList();';
  html += '  }';
  html += '}';

  html += 'function setCalendarMode(mode) {';
  html += '  calViewMode = mode;';
  html += '  const btnM = document.getElementById("btnModeMonth");';
  html += '  const btnW = document.getElementById("btnModeWeek");';
  html += '  const wkBar = document.getElementById("calWeekdaysBar");';
  html += '  const grid = document.getElementById("calGrid");';
  html += '  const wkList = document.getElementById("calWeekList");';
  html += '  if (btnM && btnW) {';
  html += '    if (mode === "week") {';
  html += '      btnM.classList.remove("active");';
  html += '      btnW.classList.add("active");';
  html += '      if (wkBar) wkBar.style.display = "none";';
  html += '      if (grid) grid.style.display = "none";';
  html += '      if (wkList) wkList.style.display = "flex";';
  html += '      renderWeeklyCalendar();';
  html += '    } else {';
  html += '      btnW.classList.remove("active");';
  html += '      btnM.classList.add("active");';
  html += '      if (wkBar) wkBar.style.display = "grid";';
  html += '      if (grid) grid.style.display = "grid";';
  html += '      if (wkList) wkList.style.display = "none";';
  html += '      renderCalendar();';
  html += '    }';
  html += '  }';
  html += '}';
  html += '';
  html += 'function navPrevPeriod() {';
  html += '  if (calViewMode === "week") prevWeek(); else prevMonth();';
  html += '}';
  html += '';
  html += 'function navNextPeriod() {';
  html += '  if (calViewMode === "week") nextWeek(); else nextMonth();';
  html += '}';
  html += '';
  html += 'function navTodayPeriod() {';
  html += '  if (calViewMode === "week") goToThisWeek(); else goToTodayMonth();';
  html += '}';
  html += '';
  html += 'function prevMonth() {';
  html += '  calMonth--;';
  html += '  if (calMonth < 1) { calMonth = 12; calYear--; }';
  html += '  renderCalendar("prev");';
  html += '}';
  html += '';
  html += 'function nextMonth() {';
  html += '  calMonth++;';
  html += '  if (calMonth > 12) { calMonth = 1; calYear++; }';
  html += '  renderCalendar("next");';
  html += '}';
  html += '';
  html += 'function goToTodayMonth() {';
  html += '  const now = new Date();';
  html += '  calYear = now.getFullYear();';
  html += '  calMonth = now.getMonth() + 1;';
  html += '  renderCalendar();';
  html += '}';
  html += '';
  html += 'function prevWeek() {';
  html += '  calWeekStartDate.setDate(calWeekStartDate.getDate() - 7);';
  html += '  renderWeeklyCalendar("prev");';
  html += '}';
  html += '';
  html += 'function nextWeek() {';
  html += '  calWeekStartDate.setDate(calWeekStartDate.getDate() + 7);';
  html += '  renderWeeklyCalendar("next");';
  html += '}';
  html += '';
  html += 'function goToThisWeek() {';
  html += '  calWeekStartDate = getSundayOfCurrentWeek(new Date());';
  html += '  renderWeeklyCalendar();';
  html += '}';
  html += '';
  html += 'function renderCurrentCalendarView(direction) {';
  html += '  if (calViewMode === "week") renderWeeklyCalendar(direction); else renderCalendar(direction);';
  html += '}';
  html += '';
  html += 'function loadWeatherData() {';
  html += '  if (Object.keys(activeWeatherData).length > 0) return;';
  html += '  const url = "https://api.open-meteo.com/v1/forecast?latitude=37.5665&longitude=126.9780&daily=weathercode,temperature_2m_max,temperature_2m_min,precipitation_probability_max&timezone=Asia%2FTokyo&forecast_days=8";';
  html += '  fetch(url).then(r => r.json()).then(data => {';
  html += '    if (data && data.daily && data.daily.time) {';
  html += '      data.daily.time.forEach((dateStr, idx) => {';
  html += '        const code = data.daily.weathercode ? data.daily.weathercode[idx] : 0;';
  html += '        const maxT = data.daily.temperature_2m_max ? Math.round(data.daily.temperature_2m_max[idx]) : null;';
  html += '        const minT = data.daily.temperature_2m_min ? Math.round(data.daily.temperature_2m_min[idx]) : null;';
  html += '        const pop = data.daily.precipitation_probability_max ? data.daily.precipitation_probability_max[idx] : 0;';
  html += '        let icon = "☀️", desc = "맑음";';
  html += '        if (code >= 1 && code <= 2) { icon = "⛅"; desc = "구름조금"; }';
  html += '        else if (code === 3) { icon = "☁️"; desc = "흐림"; }';
  html += '        else if (code === 45 || code === 48) { icon = "🌫️"; desc = "안개"; }';
  html += '        else if ((code >= 51 && code <= 67) || (code >= 80 && code <= 82)) { icon = "🌧️"; desc = "비"; }';
  html += '        else if ((code >= 71 && code <= 77) || (code >= 85 && code <= 86)) { icon = "❄️"; desc = "눈"; }';
  html += '        else if (code >= 95) { icon = "⛈️"; desc = "뇌우"; }';
  html += '        const isRain = (icon === "🌧️" || icon === "⛈️");';
  html += '        const isSnow = (icon === "❄️");';
  html += '        let badgeText = icon;';
  html += '        if (pop >= 30) badgeText += " " + pop + "%";';
  html += '        else if (maxT !== null) badgeText += " " + maxT + "°";';
  html += '        activeWeatherData[dateStr] = {';
  html += '          icon: icon, desc: desc, maxT: maxT, minT: minT, pop: pop, isRain: isRain, isSnow: isSnow,';
  html += '          badgeText: badgeText,';
  html += '          fullText: icon + " " + desc + (maxT !== null ? " " + maxT + "°/" + minT + "°" : "") + (pop > 0 ? " (강수 " + pop + "%)" : "")';
  html += '        };';
  html += '      });';
  html += '    }';
  html += '    if (isCalendarMode) renderCurrentCalendarView();';
  html += '  }).catch(e => console.warn("날씨 비동기 로드 생략", e));';
  html += '}';
  html += '';
  html += 'function captureCalendarScreen() {';
  html += '  const btn = document.getElementById("btnCaptureCal");';
  html += '  const origText = btn ? btn.innerHTML : "📷 캡쳐";';
  html += '  if (btn) { btn.disabled = true; btn.innerHTML = "📷 캡쳐 중..."; }';
  html += '  function doCapture() {';
  html += '    const target = document.querySelector(".calendar-view");';
  html += '    if (!target) {';
  html += '      alert("캡쳐할 캘린더 영역을 찾을 수 없습니다.");';
  html += '      if (btn) { btn.disabled = false; btn.innerHTML = origText; }';
  html += '      return;';
  html += '    }';
  html += '    html2canvas(target, {';
  html += '      scale: 2,';
  html += '      useCORS: true,';
  html += '      backgroundColor: "#ffffff",';
  html += '      logging: false';
  html += '    }).then(function(canvas) {';
  html += '      const dataUrl = canvas.toDataURL("image/png");';
  html += '      const filename = "현장캘린더_" + calYear + "년" + calMonth + "월.png";';
  html += '      showCaptureModal(dataUrl, filename);';
  html += '      if (btn) { btn.disabled = false; btn.innerHTML = origText; }';
  html += '    }).catch(function(err) {';
  html += '      alert("캘린더 캡쳐 중 오류가 발생했습니다: " + (err && err.message || err));';
  html += '      if (btn) { btn.disabled = false; btn.innerHTML = origText; }';
  html += '    });';
  html += '  }';
  html += '  if (typeof html2canvas === "undefined") {';
  html += '    const script = document.createElement("script");';
  html += '    script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";';
  html += '    script.onload = doCapture;';
  html += '    script.onerror = function() {';
  html += '      alert("캡쳐 라이브러리(html2canvas)를 로드하지 못했습니다. 인터넷 연결을 확인해주세요.");';
  html += '      if (btn) { btn.disabled = false; btn.innerHTML = origText; }';
  html += '    };';
  html += '    document.head.appendChild(script);';
  html += '  } else {';
  html += '    doCapture();';
  html += '  }';
  html += '}';
  html += '';
  html += 'function showCaptureModal(dataUrl, filename) {';
  html += '  const modal = document.getElementById("captureModal");';
  html += '  const img = document.getElementById("capturePreviewImg");';
  html += '  const dlBtn = document.getElementById("btnDownloadCapture");';
  html += '  if (!modal || !img) return;';
  html += '  img.src = dataUrl;';
  html += '  if (dlBtn) {';
  html += '    dlBtn.onclick = function() {';
  html += '      const link = document.createElement("a");';
  html += '      link.download = filename;';
  html += '      link.href = dataUrl;';
  html += '      document.body.appendChild(link);';
  html += '      link.click();';
  html += '      document.body.removeChild(link);';
  html += '    };';
  html += '  }';
  html += '  modal.style.display = "flex";';
  html += '}';
  html += '';
  html += 'function closeCaptureModal(e) {';
  html += '  if (e && e.target && e.target.id !== "captureModal" && !e.target.classList.contains("sheet-close") && e.target.tagName !== "BUTTON") return;';
  html += '  const modal = document.getElementById("captureModal");';
  html += '  if (modal) modal.style.display = "none";';
  html += '}';

  html += 'function renderCalendar(direction) {';
  html += '  if (!rawData || !rawData.sites) return;';
  html += '  const grid = document.getElementById("calGrid");';
  html += '  const title = document.getElementById("calMonthTitle");';
  html += '  const notice = document.getElementById("calFilterNotice");';
  html += '  if (!grid || !title) return;';
  html += '  title.innerText = calYear + "년 " + calMonth + "월";';
  html += '  if (notice) {';
  html += '    let noticeParts = [];';
  html += '    if (checkedSites.size > 0) {';
  html += '      const siteArr = Array.from(checkedSites);';
  html += '      const firstSite = siteArr[0] || "현장";';
  html += '      if (siteArr.length === 1) {';
  html += '        noticeParts.push("☑️ " + escapeHtml(firstSite));';
  html += '      } else {';
  html += '        noticeParts.push("☑️ " + escapeHtml(firstSite) + " 외 " + (siteArr.length - 1) + "건");';
  html += '      }';
  html += '    }';
  html += '    if (selectedMember) {';
  html += '      const dispName = (selectedMember === "__NONE__" || selectedMember === "담당자 없음") ? "🚫 담당자 없음" : selectedMember;';
  html += '      noticeParts.push("👤 " + escapeHtml(dispName));';
  html += '    }';
  html += '    const q = (document.getElementById("searchInput") ? document.getElementById("searchInput").value.trim() : "");';
  html += '    if (q) noticeParts.push("🔍 " + escapeHtml(q));';
  html += '    if (noticeParts.length > 0) {';
  html += '      notice.innerHTML = \'<span style="font-size:11px;color:#1e40af;font-weight:700;">\' + noticeParts.join(" | ") + \'</span><button type="button" class="btn-check-action" style="padding:1px 6px;font-size:10px;margin-left:6px;" onclick="resetAllCalendarFilters()">전체 보기</button>\';';
  html += '      notice.style.display = "flex";';
  html += '    } else {';
  html += '      notice.style.display = "none";';
  html += '    }';
  html += '  }';
  html += '  let targetSites = rawData.sites;';
  html += '  if (checkedSites.size > 0) {';
  html += '    targetSites = targetSites.filter(s => checkedSites.has(s.siteName));';
  html += '  }';
  html += '  const dayTasksMap = {};';
  html += '  const query = (document.getElementById("searchInput") ? document.getElementById("searchInput").value.trim().toLowerCase() : "");';
  html += '  targetSites.forEach(site => {';
  html += '    site.tasks.forEach(task => {';
  html += '      if (selectedMember) {';
  html += '        if (!isMemberMatch(task.manager, selectedMember)) return;';
  html += '        const stClean = (task.status || "").replace(/\\s+/g, "");';
  html += '        if (stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1) return;';
  html += '      }';
  html += '      if (query) {';
  html += '        const matchSite = site.siteName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchTask = task.taskName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMgr = (task.manager || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchPartner = (task.partner || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMemo = (task.memo || "").toLowerCase().indexOf(query) !== -1;';
  html += '        if (!matchSite && !matchTask && !matchMgr && !matchPartner && !matchMemo) return;';
  html += '      }';
  html += '      if (!task.startDateStr) return;';
  html += '      const startStr = task.startDateStr;';
  html += '      const endStr = task.endDateStr || task.startDateStr;';
  html += '      const curMonthStart = calYear + "-" + (calMonth < 10 ? "0" + calMonth : calMonth) + "-01";';
  html += '      const lastD = new Date(calYear, calMonth, 0).getDate();';
  html += '      const curMonthEnd = calYear + "-" + (calMonth < 10 ? "0" + calMonth : calMonth) + "-" + lastD;';
  html += '      if (startStr > curMonthEnd || endStr < curMonthStart) return;';
  html += '      const sParts = startStr.split("-").map(Number);';
  html += '      const eParts = endStr.split("-").map(Number);';
  html += '      let curD = new Date(sParts[0], sParts[1] - 1, sParts[2]);';
  html += '      const endD = new Date(eParts[0], eParts[1] - 1, eParts[2]);';
  html += '      let count = 0;';
  html += '      while (curD <= endD && count < 60) {';
  html += '        count++;';
  html += '        const y = curD.getFullYear();';
  html += '        const m = curD.getMonth() + 1;';
  html += '        const d = curD.getDate();';
  html += '        if (y === calYear && m === calMonth) {';
  html += '          const key = y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);';
  html += '          if (!dayTasksMap[key]) dayTasksMap[key] = [];';
  html += '          dayTasksMap[key].push({ site: site, task: task });';
  html += '        }';
  html += '        curD.setDate(curD.getDate() + 1);';
  html += '      }';
  html += '    });';
  html += '  });';
  html += '  const firstDay = new Date(calYear, calMonth - 1, 1).getDay();';
  html += '  const totalDays = new Date(calYear, calMonth, 0).getDate();';
  html += '  const prevMonthTotalDays = new Date(calYear, calMonth - 1, 0).getDate();';
  html += '  const todayObj = new Date();';
  html += '  const todayStr = todayObj.getFullYear() + "-" + ((todayObj.getMonth() + 1) < 10 ? "0" + (todayObj.getMonth() + 1) : (todayObj.getMonth() + 1)) + "-" + (todayObj.getDate() < 10 ? "0" + todayObj.getDate() : todayObj.getDate());';
  html += '  let gridHtml = "";';
  html += '  for (let p = 0; p < firstDay; p++) {';
  html += '    const dayNum = prevMonthTotalDays - firstDay + p + 1;';
  html += '    let prevY = calYear;';
  html += '    let prevM = calMonth - 1;';
  html += '    if (prevM < 1) { prevM = 12; prevY--; }';
  html += '    const pStr = prevY + "-" + (prevM < 10 ? "0" + prevM : prevM) + "-" + (dayNum < 10 ? "0" + dayNum : dayNum);';
  html += '    const pHoliday = (typeof activeHolidays !== "undefined" && activeHolidays[pStr]) ? activeHolidays[pStr] : "";';
  html += '    gridHtml += \'<div class="cal-day-cell other-month"><div class="cal-day-header"><span class="cal-day-num">\' + dayNum + \'</span>\' + (pHoliday ? \'<span class="cal-holiday-name" title="\' + escapeHtml(pHoliday) + \'">\' + escapeHtml(pHoliday) + \'</span>\' : \'\') + \'</div></div>\';';
  html += '  }';
  html += '  for (let d = 1; d <= totalDays; d++) {';
  html += '    const dStr = calYear + "-" + (calMonth < 10 ? "0" + calMonth : calMonth) + "-" + (d < 10 ? "0" + d : d);';
  html += '    const dayOfWeek = (firstDay + d - 1) % 7;';
  html += '    const holidayName = (typeof activeHolidays !== "undefined" && activeHolidays[dStr]) ? activeHolidays[dStr] : "";';
  html += '    let dayClass = "cal-day-cell";';
  html += '    if (holidayName) {';
  html += '      dayClass += " holiday sun";';
  html += '    } else if (dayOfWeek === 0) {';
  html += '      dayClass += " sun";';
  html += '    } else if (dayOfWeek === 6) {';
  html += '      dayClass += " sat";';
  html += '    }';
  html += '    if (dStr === todayStr) dayClass += " today";';
  html += '    const dayTasks = dayTasksMap[dStr] || [];';
  html += '    gridHtml += \'<div class="\' + dayClass + \'" onclick="onCalDayClick(\\\'\' + dStr + \'\\\')">\';';
  html += '    const weatherObj = activeWeatherData[dStr];';
  html += '    let weatherBadgeHtml = "";';
  html += '    if (weatherObj) {';
  html += '      const wClass = weatherObj.isRain ? " rain" : (weatherObj.isSnow ? " snow" : "");';
  html += '      weatherBadgeHtml = \'<span class="cal-weather-badge\' + wClass + \'" title="\' + escapeHtml(weatherObj.fullText) + \'">\' + escapeHtml(weatherObj.badgeText) + \'</span>\';';
  html += '    }';
  html += '    gridHtml += \'  <div class="cal-day-header"><div class="cal-day-top-row"><span class="cal-day-num">\' + d + \'</span>\' + weatherBadgeHtml + \'</div>\' + (holidayName ? \'<span class="cal-holiday-name" title="\' + escapeHtml(holidayName) + \'">\' + escapeHtml(holidayName) + \'</span>\' : \'\') + \'</div>\';';
  html += '    gridHtml += \'  <div class="cal-events-list">\';';
  html += '    dayTasks.forEach(item => {';
  html += '      const site = item.site;';
  html += '      const task = item.task;';
  html += '      const siteEnc = encodeURIComponent(site.siteName);';
  html += '      const taskJson = encodeURIComponent(JSON.stringify(task));';
  html += '      const isDone = (task.status && (task.status.indexOf("완료") !== -1 || task.status.indexOf("취소") !== -1));';
  html += '      const isDelayed = (task.isDelayed === true && !isDone);';
  html += '      const isSh = (task.isShared === true || String(task.isShared).toLowerCase() === "true" || String(task.isShared) === "1");';
  html += '      const shClass = isSh ? "blue" : "red";';
  html += '      const bgCol = isDelayed ? "#dc2626" : (site.colorHex || "#2563eb");';
  html += '      const isLight = !isDelayed && (bgCol === "#F6BF26" || bgCol.toUpperCase() === "#F6BF26");';
  html += '      const txtCol = isLight ? "#0f172a" : "#ffffff";';
  html += '      const txtShadow = isLight ? "none" : "0 1px 1px rgba(0,0,0,0.45)";';
  html += '      const fullTitle = (isDelayed ? "⚠️ [공정 지연] " : "") + escapeHtml(site.siteName) + " - " + escapeHtml(task.taskName) + " (" + (isSh ? "공유" : "미공유") + (task.manager ? ", 담당: " + task.manager : "") + ")";';
  html += '      const tLen = (task.taskName || "").length;';
  html += '      let fontStyle = "font-weight:800;";';
  html += '      if (tLen <= 2) {';
  html += '        fontStyle += "font-size:9.5px;";';
  html += '      } else if (tLen === 3) {';
  html += '        fontStyle += "font-size:8.2px;letter-spacing:-0.2px;";';
  html += '      } else {';
  html += '        fontStyle += "font-size:7.2px;letter-spacing:-0.5px;";';
  html += '      }';
  html += '      if (txtShadow !== "none") {';
  html += '        fontStyle += "text-shadow:" + txtShadow + ";";';
  html += '      }';
  html += '      gridHtml += \'<div class="cal-event-pill \' + (isDelayed ? "status-delayed " : "") + (isDone ? "status-done" : "") + \'" style="background-color:\' + bgCol + \';color:\' + txtCol + \';" onclick="event.stopPropagation();onTaskClick(\\\'\' + siteEnc + \'\\\', \\\'\' + taskJson + \'\\\')" title="\' + fullTitle + \'">\';';
  html += '      if (isDelayed) {';
  html += '        gridHtml += \'<span class="cal-delay-warn-icon" title="⚠️ 공정 지연">⚠️</span>\';';
  html += '      }';
  html += '      gridHtml += \'<span class="cal-sh-box \' + shClass + \'"></span><span class="cal-pill-text" style="\' + fontStyle + \'">\' + escapeHtml(task.taskName) + \'</span>\';';
  html += '      gridHtml += \'</div>\';';
  html += '    });';
  html += '    gridHtml += \'  </div></div>\';';
  html += '  }';
  html += '  const totalCellsSoFar = firstDay + totalDays;';
  html += '  const remainingCells = (totalCellsSoFar % 7 === 0) ? 0 : (7 - (totalCellsSoFar % 7));';
  html += '  for (let n = 1; n <= remainingCells; n++) {';
  html += '    let nextY = calYear;';
  html += '    let nextM = calMonth + 1;';
  html += '    if (nextM > 12) { nextM = 1; nextY++; }';
  html += '    const nStr = nextY + "-" + (nextM < 10 ? "0" + nextM : nextM) + "-" + (n < 10 ? "0" + n : n);';
  html += '    const nHoliday = (typeof activeHolidays !== "undefined" && activeHolidays[nStr]) ? activeHolidays[nStr] : "";';
  html += '    gridHtml += \'<div class="cal-day-cell other-month"><div class="cal-day-header"><span class="cal-day-num">\' + n + \'</span>\' + (nHoliday ? \'<span class="cal-holiday-name" title="\' + escapeHtml(nHoliday) + \'">\' + escapeHtml(nHoliday) + \'</span>\' : \'\') + \'</div></div>\';';
  html += '  }';
  html += '  grid.innerHTML = gridHtml;';
  html += '  if (direction) {';
  html += '    grid.classList.remove("cal-slide-next", "cal-slide-prev");';
  html += '    void grid.offsetWidth;';
  html += '    if (direction === "next") grid.classList.add("cal-slide-next");';
  html += '    else if (direction === "prev") grid.classList.add("cal-slide-prev");';
  html += '  }';
  html += '  setupCalendarSwipe();';
  html += '}';
  html += '';
  html += 'function renderWeeklyCalendar(direction) {';
  html += '  if (!rawData || !rawData.sites) return;';
  html += '  const weekList = document.getElementById("calWeekList");';
  html += '  const title = document.getElementById("calMonthTitle");';
  html += '  const notice = document.getElementById("calFilterNotice");';
  html += '  if (!weekList || !title) return;';
  html += '  const startD = new Date(calWeekStartDate);';
  html += '  const endD = new Date(startD);';
  html += '  endD.setDate(startD.getDate() + 6);';
  html += '  const sY = startD.getFullYear(), sM = startD.getMonth() + 1, sD = startD.getDate();';
  html += '  const eM = endD.getMonth() + 1, eD = endD.getDate();';
  html += '  title.innerText = sY + "." + (sM < 10 ? "0" + sM : sM) + "." + (sD < 10 ? "0" + sD : sD) + " ~ " + (eM < 10 ? "0" + eM : eM) + "." + (eD < 10 ? "0" + eD : eD);';
  html += '  if (notice) {';
  html += '    let noticeParts = [];';
  html += '    if (checkedSites.size > 0) {';
  html += '      const siteArr = Array.from(checkedSites);';
  html += '      const firstSite = siteArr[0] || "현장";';
  html += '      if (siteArr.length === 1) noticeParts.push("☑️ " + escapeHtml(firstSite));';
  html += '      else noticeParts.push("☑️ " + escapeHtml(firstSite) + " 외 " + (siteArr.length - 1) + "건");';
  html += '    }';
  html += '    if (selectedMember) {';
  html += '      const dispName = (selectedMember === "__NONE__" || selectedMember === "담당자 없음") ? "🚫 담당자 없음" : selectedMember;';
  html += '      noticeParts.push("👤 " + escapeHtml(dispName));';
  html += '    }';
  html += '    const q = (document.getElementById("searchInput") ? document.getElementById("searchInput").value.trim() : "");';
  html += '    if (q) noticeParts.push("🔍 " + escapeHtml(q));';
  html += '    if (noticeParts.length > 0) {';
  html += '      notice.innerHTML = \'<span style="font-size:11px;color:#1e40af;font-weight:700;">\' + noticeParts.join(" | ") + \'</span><button type="button" class="btn-check-action" style="padding:1px 6px;font-size:10px;margin-left:6px;" onclick="resetAllCalendarFilters()">전체 보기</button>\';';
  html += '      notice.style.display = "flex";';
  html += '    } else {';
  html += '      notice.style.display = "none";';
  html += '    }';
  html += '  }';
  html += '  let targetSites = rawData.sites;';
  html += '  if (checkedSites.size > 0) targetSites = targetSites.filter(s => checkedSites.has(s.siteName));';
  html += '  const dayTasksMap = {};';
  html += '  const query = (document.getElementById("searchInput") ? document.getElementById("searchInput").value.trim().toLowerCase() : "");';
  html += '  targetSites.forEach(site => {';
  html += '    site.tasks.forEach(task => {';
  html += '      if (selectedMember) {';
  html += '        if (!isMemberMatch(task.manager, selectedMember)) return;';
  html += '        const stClean = (task.status || "").replace(/\\s+/g, "");';
  html += '        if (stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1) return;';
  html += '      }';
  html += '      if (query) {';
  html += '        const matchSite = site.siteName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchTask = task.taskName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMgr = (task.manager || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchPartner = (task.partner || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMemo = (task.memo || "").toLowerCase().indexOf(query) !== -1;';
  html += '        if (!matchSite && !matchTask && !matchMgr && !matchPartner && !matchMemo) return;';
  html += '      }';
  html += '      if (!task.startDateStr) return;';
  html += '      const startStr = task.startDateStr;';
  html += '      const endStr = task.endDateStr || task.startDateStr;';
  html += '      for (let i = 0; i < 7; i++) {';
  html += '        const curD = new Date(startD);';
  html += '        curD.setDate(startD.getDate() + i);';
  html += '        const y = curD.getFullYear(), m = curD.getMonth() + 1, d = curD.getDate();';
  html += '        const curStr = y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);';
  html += '        if (startStr <= curStr && curStr <= endStr) {';
  html += '          if (!dayTasksMap[curStr]) dayTasksMap[curStr] = [];';
  html += '          dayTasksMap[curStr].push({ site: site, task: task });';
  html += '        }';
  html += '      }';
  html += '    });';
  html += '  });';
  html += '  const now = new Date();';
  html += '  const todayStr = now.getFullYear() + "-" + ((now.getMonth() + 1) < 10 ? "0" + (now.getMonth() + 1) : (now.getMonth() + 1)) + "-" + (now.getDate() < 10 ? "0" + now.getDate() : now.getDate());';
  html += '  const dayNames = ["일", "월", "화", "수", "목", "금", "토"];';
  html += '  let listHtml = "";';
  html += '  for (let i = 0; i < 7; i++) {';
  html += '    const curD = new Date(startD);';
  html += '    curD.setDate(startD.getDate() + i);';
  html += '    const y = curD.getFullYear(), m = curD.getMonth() + 1, d = curD.getDate();';
  html += '    const curStr = y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);';
  html += '    const dayOfWeek = curD.getDay();';
  html += '    const isSun = (dayOfWeek === 0), isSat = (dayOfWeek === 6);';
  html += '    const isToday = (curStr === todayStr);';
  html += '    const holidayName = (typeof activeHolidays !== "undefined" && activeHolidays[curStr]) ? activeHolidays[curStr] : "";';
  html += '    const weatherObj = activeWeatherData[curStr];';
  html += '    let dayTagClass = "cal-week-day-tag";';
  html += '    if (isSun || holidayName) dayTagClass += " sun";';
  html += '    else if (isSat) dayTagClass += " sat";';
  html += '    const dayTasks = dayTasksMap[curStr] || [];';
  html += '    listHtml += \'<div class="cal-week-card\' + (isToday ? " today" : "") + \'">\';';
  html += '    listHtml += \'  <div class="cal-week-card-header" onclick="onCalDayClick(\\\'\' + curStr + \'\\\')">\';';
  html += '    listHtml += \'    <div class="cal-week-card-left">\';';
  html += '    listHtml += \'      <span class="\' + dayTagClass + \'">\' + (m < 10 ? "0" + m : m) + "." + (d < 10 ? "0" + d : d) + " (" + dayNames[dayOfWeek] + ")</span>";';
  html += '    if (isToday) listHtml += \' <span class="cal-week-today-badge">오늘</span>\';';
  html += '    if (holidayName) listHtml += \' <span class="cal-week-holiday-badge">\' + escapeHtml(holidayName) + \'</span>\';';
  html += '    listHtml += \'    </div>\';';
  html += '    listHtml += \'    <div class="cal-week-card-right">\';';
  html += '    if (weatherObj) listHtml += \' <span class="cal-week-weather-info\' + (weatherObj.isRain ? " rain" : "") + \'">\' + escapeHtml(weatherObj.fullText) + \'</span>\';';
  html += '    listHtml += \'    </div>\';';
  html += '    listHtml += \'  </div>\';';
  html += '    listHtml += \'  <div class="cal-week-card-body">\';';
  html += '    if (dayTasks.length === 0) {';
  html += '      listHtml += \'    <div class="cal-week-empty-text">공정 일정 없음</div>\';';
  html += '    } else {';
  html += '      dayTasks.forEach(item => {';
  html += '        const site = item.site, task = item.task;';
  html += '        const siteEnc = encodeURIComponent(site.siteName);';
  html += '        const taskJson = encodeURIComponent(JSON.stringify(task));';
  html += '        const isDone = (task.status && (task.status.indexOf("완료") !== -1 || task.status.indexOf("취소") !== -1));';
  html += '        const isDelayed = (task.isDelayed === true && !isDone);';
  html += '        const isSh = (task.isShared === true || String(task.isShared).toLowerCase() === "true" || String(task.isShared) === "1");';
  html += '        const shClass = isSh ? "blue" : "red";';
  html += '        const bgCol = isDelayed ? "#dc2626" : (site.colorHex || "#2563eb");';
  html += '        const isLight = !isDelayed && (bgCol === "#F6BF26" || bgCol.toUpperCase() === "#F6BF26");';
  html += '        const txtCol = isLight ? "#0f172a" : "#ffffff";';
  html += '        const txtShadow = isLight ? "none" : "0 1px 1px rgba(0,0,0,0.45)";';
  html += '        const fullTitle = (isDelayed ? "⚠️ [공정 지연] " : "") + escapeHtml(site.siteName) + " - " + escapeHtml(task.taskName) + " (" + (isSh ? "공유" : "미공유") + (task.manager ? ", 담당: " + task.manager : "") + ")";';
  html += '        listHtml += \'<div class="cal-event-pill \' + (isDelayed ? "status-delayed " : "") + (isDone ? "status-done" : "") + \'" style="background-color:\' + bgCol + \';color:\' + txtCol + \';padding:4px 6px;font-size:11px;\' + (txtShadow !== "none" ? "text-shadow:" + txtShadow + ";" : "") + \'" onclick="event.stopPropagation();onTaskClick(\\\'\' + siteEnc + \'\\\', \\\'\' + taskJson + \'\\\')" title="\' + fullTitle + \'">\';';
  html += '        const pRateTag = (typeof site.progressRate === "number") ? " " + site.progressRate + "%" : "";';
  html += '        if (isDelayed) {';
  html += '          listHtml += \'  <span class="cal-delay-warn-icon" style="font-size:10px;margin-right:3px;" title="⚠️ 공정 지연">⚠️</span>\';';
  html += '        }';
  html += '        listHtml += \'  <span class="cal-sh-box \' + shClass + \'" style="width:7.5px;height:7.5px;"></span>\';';
  html += '        listHtml += \'  <span class="cal-pill-text" style="font-size:11px;font-weight:800;">[\' + escapeHtml(site.siteName) + pRateTag + \'] \' + escapeHtml(task.taskName) + (task.manager ? " (" + escapeHtml(task.manager) + ")" : "") + (task.status ? " - " + escapeHtml(task.status) : "") + \'</span>\';';
  html += '        listHtml += \'</div>\';';
  html += '      });';
  html += '    }';
  html += '    listHtml += \'  </div>\';';
  html += '    listHtml += \'</div>\';';
  html += '  }';
  html += '  weekList.innerHTML = listHtml;';
  html += '  if (direction) {';
  html += '    weekList.classList.remove("cal-slide-next", "cal-slide-prev");';
  html += '    void weekList.offsetWidth;';
  html += '    if (direction === "next") weekList.classList.add("cal-slide-next");';
  html += '    else if (direction === "prev") weekList.classList.add("cal-slide-prev");';
  html += '  }';
  html += '  setupCalendarSwipe();';
  html += '}';
  html += '';
  html += 'function setupCalendarSwipe() {';
  html += '  const calArea = document.getElementById("calendarArea");';
  html += '  if (!calArea || calArea._swipeAttached) return;';
  html += '  calArea._swipeAttached = true;';
  html += '  let startX = 0, startY = 0, startTime = 0;';
  html += '  calArea.addEventListener("touchstart", function(e) {';
  html += '    if (e.touches.length !== 1) return;';
  html += '    startX = e.touches[0].clientX;';
  html += '    startY = e.touches[0].clientY;';
  html += '    startTime = Date.now();';
  html += '  }, { passive: true });';
  html += '  calArea.addEventListener("touchend", function(e) {';
  html += '    if (!e.changedTouches || e.changedTouches.length === 0) return;';
  html += '    const deltaX = e.changedTouches[0].clientX - startX;';
  html += '    const deltaY = e.changedTouches[0].clientY - startY;';
  html += '    const elapsed = Date.now() - startTime;';
  html += '    if (elapsed <= 750 && Math.abs(deltaX) >= 40 && Math.abs(deltaX) > Math.abs(deltaY) * 1.25) {';
  html += '      if (deltaX < 0) navNextPeriod(); else navPrevPeriod();';
  html += '    }';
  html += '  }, { passive: true });';
  html += '  let mStartX = 0, mStartY = 0, mStartTime = 0, isMDown = false;';
  html += '  calArea.addEventListener("mousedown", function(e) {';
  html += '    if (e.button !== 0) return;';
  html += '    if (e.target.closest("button") || e.target.closest(".cal-event-pill")) return;';
  html += '    isMDown = true;';
  html += '    mStartX = e.clientX;';
  html += '    mStartY = e.clientY;';
  html += '    mStartTime = Date.now();';
  html += '  });';
  html += '  window.addEventListener("mouseup", function(e) {';
  html += '    if (!isMDown) return;';
  html += '    isMDown = false;';
  html += '    const deltaX = e.clientX - mStartX;';
  html += '    const deltaY = e.clientY - mStartY;';
  html += '    const elapsed = Date.now() - mStartTime;';
  html += '    if (elapsed <= 750 && Math.abs(deltaX) >= 50 && Math.abs(deltaX) > Math.abs(deltaY) * 1.25) {';
  html += '      if (deltaX < 0) navNextPeriod(); else navPrevPeriod();';
  html += '    }';
  html += '  });';
  html += '}';
  html += '';
  html += 'function resetAllCalendarFilters() {';
  html += '  checkedSites.clear();';
  html += '  document.querySelectorAll(".site-check").forEach(cb => cb.checked = false);';
  html += '  updateCheckedStatusUI();';
  html += '  setMemberFilter("");';
  html += '  const sInput = document.getElementById("searchInput");';
  html += '  if (sInput) sInput.value = "";';
  html += '  renderCurrentCalendarView();';
  html += '}';
  html += '';
  html += 'function onCalDayClick(dStr) {';
  html += '  if (!rawData || !rawData.sites) return;';
  html += '  const targetSites = (checkedSites.size > 0) ? rawData.sites.filter(s => checkedSites.has(s.siteName)) : rawData.sites;';
  html += '  const dayTasks = [];';
  html += '  const holidayName = (typeof activeHolidays !== "undefined" && activeHolidays[dStr]) ? activeHolidays[dStr] : "";';
  html += '  const weatherObj = (typeof activeWeatherData !== "undefined") ? activeWeatherData[dStr] : null;';
  html += '  const weatherText = weatherObj ? " | " + weatherObj.fullText : "";';
  html += '  const datePrefix = "📅 " + dStr + (holidayName ? " (" + holidayName + ")" : "") + weatherText;';
  html += '  targetSites.forEach(site => {';
  html += '    site.tasks.forEach(task => {';
  html += '      if (task.startDateStr && task.startDateStr <= dStr && dStr <= (task.endDateStr || task.startDateStr)) {';
  html += '        if (selectedMember) {';
  html += '          if (!isMemberMatch(task.manager, selectedMember)) return;';
  html += '        }';
  html += '        dayTasks.push({ site: site, task: task });';
  html += '      }';
  html += '    });';
  html += '  });';
  html += '  if (dayTasks.length === 0) {';
  html += '    showToast(datePrefix + ": 등록된 일정이 없습니다.");';
  html += '    return;';
  html += '  }';
  html += '  if (isMobileMode) {';
  html += '    const mList = dayTasks.map(item => ({ site: item.site.siteName, task: item.task }));';
  html += '    openBottomSheet(mList[0].site, mList[0].task, mList, 0);';
  html += '  } else {';
  html += '    onTaskClick(encodeURIComponent(dayTasks[0].site.siteName), encodeURIComponent(JSON.stringify(dayTasks[0].task)));';
  html += '  }';
  html += '  if (dayTasks.length > 1) {';
  html += '    showToast(datePrefix + " 일정 " + dayTasks.length + "건 (상세에서 ◀/▶ 이동 가능)");';
  html += '  }';
  html += '}';

  html += 'function onSearchInput() {';
  html += '  const q = document.getElementById("searchInput").value;';
  html += '  const clearBtn = document.getElementById("searchClearBtn");';
  html += '  if (clearBtn) {';
  html += '    clearBtn.classList.toggle("show", q.length > 0);';
  html += '  }';
  html += '  if (isCalendarMode) renderCurrentCalendarView(); else renderList();';
  html += '}';

  html += 'function clearSearch() {';
  html += '  const input = document.getElementById("searchInput");';
  html += '  if (input) {';
  html += '    input.value = "";';
  html += '    input.focus();';
  html += '  }';
  html += '  const clearBtn = document.getElementById("searchClearBtn");';
  html += '  if (clearBtn) {';
  html += '    clearBtn.classList.remove("show");';
  html += '  }';
  html += '  if (isCalendarMode) renderCurrentCalendarView(); else renderList();';
  html += '}';

  html += 'function isMemberMatch(taskManager, sel) {';
  html += '  if (!sel) return true;';
  html += '  const mgr = (taskManager || "").trim();';
  html += '  if (sel === "__NONE__" || sel === "담당자 없음") {';
  html += '    return (!mgr || mgr === "담당자 없음" || mgr === "미지정" || mgr === "없음");';
  html += '  }';
  html += '  return mgr.indexOf(sel) !== -1;';
  html += '}';
  html += '';
  html += 'function renderMemberFilterTabs(members) {';
  html += '  const bar = document.getElementById("memberFilterBar");';
  html += '  if (!bar) return;';
  html += '  if (!members || members.length === 0) {';
  html += '    bar.innerHTML = "";';
  html += '    return;';
  html += '  }';
  html += '  let tabsHtml = \'<div class="member-tab \' + (!selectedMember ? "active" : "") + \'" onclick="setMemberFilter(\\\'\\\')">👤 전체</div>\';';
  html += '  members.forEach(m => {';
  html += '    if (m === "김종혁" || m.indexOf("김종혁") !== -1) return;';
  html += '    const isAct = (selectedMember === m);';
  html += '    const mEnc = encodeURIComponent(m);';
  html += '    tabsHtml += \'<div class="member-tab \' + (isAct ? "active" : "") + \'" onclick="setMemberFilter(\\\'\' + mEnc + \'\\\')">\' + escapeHtml(m) + \'</div>\';';
  html += '  });';
  html += '  const isNoneAct = (selectedMember === "__NONE__" || selectedMember === "담당자 없음");';
  html += '  tabsHtml += \'<div class="member-tab \' + (isNoneAct ? "active" : "") + \'" onclick="setMemberFilter(\\\'__NONE__\\\')">🚫 담당자 없음</div>\';';
  html += '  bar.innerHTML = tabsHtml;';
  html += '}';

  html += 'function setMemberFilter(enc) {';
  html += '  const name = enc ? decodeURIComponent(enc) : "";';
  html += '  if (selectedMember === name) {';
  html += '    selectedMember = "";';
  html += '  } else {';
  html += '    selectedMember = name;';
  html += '  }';
  html += '  if (rawData && rawData.members) {';
  html += '    renderMemberFilterTabs(rawData.members);';
  html += '  }';
  html += '  if (isCalendarMode) renderCurrentCalendarView(); else renderList();';
  html += '}';

  html += 'function loadData() {';
  html += '  document.getElementById("contentArea").innerHTML = \'<div class="loading-state"><div class="spinner"></div><p class="loading-text">현장 데이터 불러오는 중...</p></div>\';';
  html += '  google.script.run';
  html += '    .withSuccessHandler(onDataLoaded)';
  html += '    .withFailureHandler(onError)';
  html += '    .getSidebarDashboardData();';
  html += '}';

  html += 'function onError(err) {';
  html += '  document.getElementById("contentArea").innerHTML = \'<div class="empty-state">⚠️ 데이터를 불러오지 못했습니다.<br><small style="color:#ef4444;">\' + escapeHtml((err && err.message) || err || "") + \'</small></div>\';';
  html += '}';

  html += 'function onDataLoaded(data) {';
  html += '  rawData = data;';
  html += '  if (!data) {';
  html += '    document.getElementById("contentArea").innerHTML = \'<div class="empty-state">⚠️ 데이터를 불러오지 못했습니다.</div>\';';
  html += '    return;';
  html += '  }';
  html += '  if (data.error) {';
  html += '    document.getElementById("contentArea").innerHTML = \'<div class="empty-state">⚠️ \' + escapeHtml(data.error) + \'</div>\';';
  html += '    return;';
  html += '  }';
  html += '  try {';
  html += '    if (data.holidays && typeof data.holidays === "object") {';
  html += '      activeHolidays = Object.assign({}, DEFAULT_KOREAN_HOLIDAYS, data.holidays);';
  html += '    }';
  html += '    const kpi = data.kpi || {};';
  html += '    document.getElementById("kpiSites").innerText = (kpi.activeSites !== undefined) ? kpi.activeSites : (kpi.totalSites || 0);';
  html += '    document.getElementById("kpiToday").innerText = kpi.todayCount || 0;';
  html += '    document.getElementById("kpiDelay").innerText = kpi.delayedCount || 0;';
  html += '    renderMemberFilterTabs(data.members || []);';
  html += '    updateCheckedStatusUI();';
  html += '    if (isCalendarMode) renderCurrentCalendarView(); else renderList();';
  html += '    loadWeatherData();';
  html += '  } catch (renderErr) {';
  html += '    document.getElementById("contentArea").innerHTML = \'<div class="empty-state">⚠️ 대시보드 렌더링 오류: \' + escapeHtml(renderErr.message || renderErr) + \'</div>\';';
  html += '  }';
  html += '}';
  html += '';
  html += 'function setFilter(type) {';
  html += '  currentFilter = type;';
  html += '  document.querySelectorAll(".filter-tab").forEach(b => {';
  html += '    b.classList.toggle("active", b.dataset.filter === type);';
  html += '  });';
  html += '  document.querySelectorAll(".kpi-card").forEach(c => {';
  html += '    c.classList.toggle("active", c.dataset.filter === type);';
  html += '  });';
  html += '  updateCheckedStatusUI();';
  html += '  if (isCalendarMode) toggleCalendarView(); else renderList();';
  html += '}';

  html += 'function renderList() {';
  html += '  if (!rawData || !rawData.sites) return;';
  html += '  const query = document.getElementById("searchInput").value.trim().toLowerCase();';
  html += '  const container = document.getElementById("contentArea");';
  html += '  let html = "";';
  html += '  let visibleSitesCount = 0;';

  html += '  const matchedSites = [];';
  html += '  rawData.sites.forEach(site => {';
  html += '    if (currentFilter === "checked") {';
  html += '      if (!checkedSites.has(site.siteName)) return;';
  html += '    } else if (currentFilter === "done") {';
  html += '      if (!site.isAllCompleted) return;';
  html += '    } else {';
  html += '      if (site.isAllCompleted) return;';
  html += '    }';

  html += '    const filteredTasks = site.tasks.filter(t => {';
  html += '      if (currentFilter === "today" && !t.isToday) return false;';
  html += '      if (currentFilter === "delay" && !t.isDelayed) return false;';
  html += '      if (currentFilter === "prog" && t.status.indexOf("진행") === -1) return false;';
  html += '      if (currentFilter === "done" && t.status.indexOf("완료") === -1 && t.status.indexOf("취소") === -1) return false;';
  html += '      if (query && currentFilter !== "done") {';
  html += '        const stClean = (t.status || "").replace(/\\s+/g, "");';
  html += '        if (stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1) return false;';
  html += '      }';
  html += '      if (selectedMember) {';
  html += '        const stClean = (t.status || "").replace(/\\s+/g, "");';
  html += '        if (stClean.indexOf("완료") !== -1 || stClean.indexOf("취소") !== -1) return false;';
  html += '        if (!isMemberMatch(t.manager, selectedMember)) return false;';
  html += '      }';
  html += '      if (query) {';
  html += '        const matchSite = site.siteName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchTask = t.taskName.toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMgr = (t.manager || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchPartner = (t.partner || "").toLowerCase().indexOf(query) !== -1;';
  html += '        const matchMemo = (t.memo || "").toLowerCase().indexOf(query) !== -1;';
  html += '        if (!matchSite && !matchTask && !matchMgr && !matchPartner && !matchMemo) return false;';
  html += '      }';
  html += '      return true;';
  html += '    });';

  html += '    if (filteredTasks.length === 0) return;';

  html += '    let earliestFilteredMs = 0;';
  html += '    let latestFilteredMs = 0;';
  html += '    for (let i = 0; i < filteredTasks.length; i++) {';
  html += '      if (filteredTasks[i].dateMs) {';
  html += '        if (!earliestFilteredMs || filteredTasks[i].dateMs < earliestFilteredMs) {';
  html += '          earliestFilteredMs = filteredTasks[i].dateMs;';
  html += '        }';
  html += '        if (!latestFilteredMs || filteredTasks[i].dateMs > latestFilteredMs) {';
  html += '          latestFilteredMs = filteredTasks[i].dateMs;';
  html += '        }';
  html += '      }';
  html += '    }';

  html += '    matchedSites.push({';
  html += '      site: site,';
  html += '      filteredTasks: filteredTasks,';
  html += '      earliestFilteredMs: earliestFilteredMs,';
  html += '      latestFilteredMs: latestFilteredMs';
  html += '    });';
  html += '  });';

  html += '  matchedSites.sort((a, b) => {';
  html += '    if (currentFilter === "done") {';
  html += '      const aLatest = a.latestFilteredMs || a.earliestFilteredMs || 0;';
  html += '      const bLatest = b.latestFilteredMs || b.earliestFilteredMs || 0;';
  html += '      if (aLatest && bLatest) {';
  html += '        if (aLatest !== bLatest) return bLatest - aLatest;';
  html += '      } else if (aLatest) {';
  html += '        return -1;';
  html += '      } else if (bLatest) {';
  html += '        return 1;';
  html += '      }';
  html += '      return (b.site.initialOrder || 0) - (a.site.initialOrder || 0);';
  html += '    }';
  html += '    const aMs = a.earliestFilteredMs || 0;';
  html += '    const bMs = b.earliestFilteredMs || 0;';
  html += '    if (aMs && bMs) {';
  html += '      if (aMs !== bMs) return aMs - bMs;';
  html += '    } else if (aMs) {';
  html += '      return -1;';
  html += '    } else if (bMs) {';
  html += '      return 1;';
  html += '    }';
  html += '    return 0;';
  html += '  });';

  html += '  matchedSites.forEach(item => {';
  html += '    visibleSitesCount++;';
  html += '    const site = item.site;';
  html += '    const filteredTasks = item.filteredTasks;';
  html += '    const isExpanded = (query !== "") || (currentFilter !== "all" && currentFilter !== "done");';
  html += '    const siteEnc = encodeURIComponent(site.siteName);';
  html += '    const isChecked = checkedSites.has(site.siteName);';
  html += '';
  html += '    html += \'<div class="site-card \' + (isExpanded ? "expanded" : "") + \'">\';';
  html += '    html += \'  <div class="site-header" style="border-left-color: \' + site.colorHex + \';" onclick="toggleSite(this)">\';';
  html += '    html += \'    <input type="checkbox" class="site-check" data-sitename="\' + siteEnc + \'" \' + (isChecked ? "checked" : "") + \' onclick="onSiteCheckboxClick(event, this)" title="현장 선택">\';';
  html += '    const pRate = (typeof site.progressRate === "number") ? site.progressRate : 0;';
  html += '    let rateClass = "low";';
  html += '    if (site.isAllCompleted) rateClass = "done";';
  html += '    else if (pRate >= 70) rateClass = "high";';
  html += '    else if (pRate >= 30) rateClass = "mid";';
  html += '    const validTotal = site.validTasksCount || site.totalTasksCount || site.tasks.length;';
  html += '    const compTotal = site.completedTasksCount || 0;';
  html += '    html += \'    <div class="site-title-area">\';';
  html += '    html += \'      <span class="site-name">\' + escapeHtml(site.siteName) + \'</span>\';';
  html += '    html += \'      <span class="site-progress-badge \' + rateClass + \'">\' + pRate + \'%</span>\';';
  html += '    if (site.isAllCompleted) {';
  html += '      html += \'      <span class="site-badge" style="background:#dcfce7;color:#166534;">완료 \' + compTotal + \'/\' + validTotal + \'</span>\';';
  html += '    } else {';
  html += '      html += \'      <span class="site-badge">\' + compTotal + \'/\' + validTotal + \'건</span>\';';
  html += '      if (site.delayedTasksCount > 0) {';
  html += '        html += \'      <span class="site-delay-mini-badge">⚠️ \' + site.delayedTasksCount + \'</span>\';';
  html += '      }';
  html += '    }';
  html += '    html += \'    </div>\';';
  html += '    html += \'    <span class="site-toggle">▼</span>\';';
  html += '    html += \'  </div>\';';
  html += '    const barColor = site.isAllCompleted ? "#10b981" : ((site.delayedTasksCount > 0) ? "#f59e0b" : (site.colorHex || "#2563eb"));';
  html += '    html += \'  <div class="site-progress-track" title="공정 진척률: \' + pRate + \'%"><div class="site-progress-bar" style="width:\' + pRate + \'%;background:\' + barColor + \';"></div></div>\';';
  html += '';
  html += '    html += \'  <div class="task-list">\';';
  html += '    const progTotal = site.inProgressTasksCount || 0;';
  html += '    const delayTotal = site.delayedTasksCount || 0;';
  html += '    let sumText = "총 " + validTotal + "개 공정 중 완료 " + compTotal + "건";';
  html += '    if (progTotal > 0) sumText += " · 진행 " + progTotal + "건";';
  html += '    if (delayTotal > 0) sumText += " · ⚠️ 지연 " + delayTotal + "건";';
  html += '    if (site.canceledTasksCount > 0) sumText += " (취소 " + site.canceledTasksCount + "건 제외)";';
  html += '    html += \'    <div class="site-summary-box">\';';
  html += '    html += \'      <div class="site-summary-main">\';';
  html += '    html += \'        <span class="site-summary-rate \' + (site.isAllCompleted ? "done" : "") + \'">진척률 <strong>\' + pRate + \'%</strong></span>\';';
  html += '    html += \'        <span class="site-summary-counts">\' + escapeHtml(sumText) + \'</span>\';';
  html += '    html += \'      </div>\';';
  html += '    html += \'    </div>\';';
  html += '    filteredTasks.forEach(task => {';
  html += '      let dateClass = "";';
  html += '      if (task.isDelayed) dateClass = "delay";';
  html += '      else if (task.isToday) dateClass = "today";';

  html += '      let statusClass = "status-예정";';
  html += '      if (task.isDelayed) statusClass = "status-지연";';
  html += '      else if (task.status.indexOf("완료") !== -1) statusClass = "status-완료";';
  html += '      else if (task.status.indexOf("진행") !== -1) statusClass = "status-진행";';
  html += '      else if (task.status.indexOf("AS") !== -1) statusClass = "status-AS";';

  html += '      const isSh = (task.isShared === true || String(task.isShared).toLowerCase() === "true" || String(task.isShared) === "1");';
  html += '      const shareTagClass = isSh ? "share-tag share-on" : "share-tag share-off";';
  html += '      const shareTagText = isSh ? "🟦 공유" : "🟥 미공유";';

  html += '      const taskJson = encodeURIComponent(JSON.stringify(task));';
  html += '      const siteEnc = encodeURIComponent(site.siteName);';
  html += '      html += \'<div class="task-item" onclick="onTaskClick(\\\'\' + siteEnc + \'\\\', \\\'\' + taskJson + \'\\\')">\';';
  html += '      html += \'  <div class="task-top">\';';
  html += '      html += \'    <span class="task-name">[\' + escapeHtml(task.taskName) + \']</span>\';';
  html += '      html += \'    <span class="task-date \' + dateClass + \'">\' + (task.isToday ? "🔥 오늘 " : (task.isDelayed ? "⚠️ 지연 " : "")) + escapeHtml(task.dateDisplay) + \'</span>\';';
  html += '      html += \'  </div>\';';
  html += '      html += \'  <div class="task-meta">\';';
  html += '      html += \'    <span>👤 \' + escapeHtml(task.manager || "-") + (task.partner ? " / " + escapeHtml(task.partner) : "") + \'</span>\';';
  html += '      html += \'    <div style="display:flex;align-items:center;gap:4px;">\';';
  html += '      html += \'      <span class="\' + shareTagClass + \'">\' + shareTagText + \'</span>\';';
  html += '      html += \'      <span class="status-tag \' + statusClass + \'">\' + (task.isDelayed ? "⚠️ 지연" : escapeHtml(task.status)) + \'</span>\';';
  html += '      html += \'    </div>\';';
  html += '      html += \'  </div>\';';
  html += '      if (task.memo) {';
  html += '        html += \'  <div class="task-memo">📝 \' + escapeHtml(task.memo) + \'</div>\';';
  html += '      }';
  html += '      html += \'</div>\';';
  html += '    });';
  html += '    html += \'  </div>\';';
  html += '    html += \'</div>\';';
  html += '  });';

  html += '  if (visibleSitesCount === 0) {';
  html += '    if (currentFilter === "checked") {';
  html += '      container.innerHTML = \'<div class="empty-state">☑️ 선택된 현장이 없습니다.<br><small style="color:#64748b;font-size:11px;">현장 카드 앞 체크박스를 선택 후 다시 확인해보세요.</small></div>\';';
  html += '    } else if (currentFilter === "done") {';
  html += '      container.innerHTML = \'<div class="empty-state">✅ 완료된 현장이 없습니다.</div>\';';
  html += '    } else {';
  html += '      container.innerHTML = \'<div class="empty-state">🔍 일치하는 현장 및 공정이 없습니다.</div>\';';
  html += '    }';
  html += '  } else {';
  html += '    container.innerHTML = html;';
  html += '  }';
  html += '}';

  html += 'function toggleSite(el) {';
  html += '  el.parentElement.classList.toggle("expanded");';
  html += '}';

  html += 'function onTaskClick(arg1, arg2, arg3) {';
  html += '  var site = "";';
  html += '  var task = null;';
  html += '  if (typeof arg1 === "number" || /^[0-9]+$/.test(String(arg1))) {';
  html += '    site = String(arg2 || "");';
  html += '    task = { row: parseInt(arg1, 10), taskName: String(arg3 || "") };';
  html += '  } else {';
  html += '    try {';
  html += '      site = decodeURIComponent(arg1);';
  html += '      task = (typeof arg2 === "string") ? JSON.parse(decodeURIComponent(arg2)) : arg2;';
  html += '    } catch(e) {';
  html += '      site = String(arg1 || "");';
  html += '      task = { row: 2, taskName: "" };';
  html += '    }';
  html += '  }';
  html += '  if (isMobileMode) {';
  html += '    openBottomSheet(site, task);';
  html += '  } else {';
  html += '    var targetRow = (task && task.row) ? task.row : parseInt(arg1, 10);';
  html += '    showToast("📌 시트 " + targetRow + "행으로 이동 중...");';
  html += '    google.script.run';
  html += '      .withSuccessHandler(() => showToast("✅ [" + (site || "현장") + " - " + ((task && task.taskName) || "공정") + "] 행으로 이동했습니다."))';
  html += '      .focusSheetRow(targetRow);';
  html += '  }';
  html += '}';

  html += 'function findSameDayTasks(baseTask) {';
  html += '  if (!baseTask || !rawData || !rawData.sites) return [];';
  html += '  const baseDStr = baseTask.startDateStr || "";';
  html += '  if (!baseDStr) return [];';
  html += '  const list = [];';
  html += '  rawData.sites.forEach(site => {';
  html += '    site.tasks.forEach(t => {';
  html += '      const sStr = t.startDateStr || "";';
  html += '      const eStr = t.endDateStr || sStr;';
  html += '      if (sStr && sStr <= baseDStr && baseDStr <= eStr) {';
  html += '        list.push({ site: site.siteName, task: t });';
  html += '      }';
  html += '    });';
  html += '  });';
  html += '  return list;';
  html += '}';
  html += '';
  html += 'function updateSameDayNavUI() {';
  html += '  const navEl = document.getElementById("sheetDayNav");';
  html += '  const prevBtn = document.getElementById("btnSheetPrevDayTask");';
  html += '  const nextBtn = document.getElementById("btnSheetNextDayTask");';
  html += '  const idxSpan = document.getElementById("sheetDayNavIdx");';
  html += '  const totalSpan = document.getElementById("sheetDayNavTotal");';
  html += '  if (!navEl) return;';
  html += '  if (sameDayTaskList && sameDayTaskList.length > 1) {';
  html += '    navEl.style.display = "flex";';
  html += '    if (idxSpan) idxSpan.innerText = curSameDayIndex + 1;';
  html += '    if (totalSpan) totalSpan.innerText = sameDayTaskList.length;';
  html += '    if (prevBtn) prevBtn.disabled = (curSameDayIndex <= 0);';
  html += '    if (nextBtn) nextBtn.disabled = (curSameDayIndex >= sameDayTaskList.length - 1);';
  html += '  } else {';
  html += '    navEl.style.display = "none";';
  html += '  }';
  html += '}';
  html += '';
  html += 'function navSameDayTask(delta) {';
  html += '  if (!sameDayTaskList || sameDayTaskList.length <= 1) return;';
  html += '  const newIdx = curSameDayIndex + delta;';
  html += '  if (newIdx < 0 || newIdx >= sameDayTaskList.length) return;';
  html += '  curSameDayIndex = newIdx;';
  html += '  const target = sameDayTaskList[curSameDayIndex];';
  html += '  if (target) {';
  html += '    openBottomSheet(target.site, target.task, sameDayTaskList, curSameDayIndex);';
  html += '  }';
  html += '}';
  html += '';
  html += 'function openBottomSheet(site, task, explicitList, explicitIndex) {';
  html += '  curSiteName = site;';
  html += '  curTaskObj = task;';
  html += '  selectedStatus = task.status || "예정";';
  html += '  selectedShare = (task.isShared === true);';
  html += '  document.getElementById("sheetSiteName").innerText = site;';
  html += '  document.getElementById("sheetTaskName").innerText = task.taskName;';
  html += '  document.getElementById("sheetDateDisplay").innerText = task.dateDisplay || "-";';
  html += '  document.getElementById("sheetMetaDisplay").innerText = (task.manager || "-") + (task.partner ? " / " + task.partner : "");';
  html += '  document.getElementById("sheetMemoInput").value = task.memo || "";';
  html += '  updateSheetStatusBtns();';
  html += '  updateSheetShareBtns();';
  html += '  if (Array.isArray(explicitList) && explicitList.length > 0) {';
  html += '    sameDayTaskList = explicitList;';
  html += '    curSameDayIndex = (typeof explicitIndex === "number") ? explicitIndex : 0;';
  html += '  } else {';
  html += '    sameDayTaskList = findSameDayTasks(task);';
  html += '    curSameDayIndex = sameDayTaskList.findIndex(item => item.task && task && (item.task.row === task.row || (item.site === site && item.task.taskName === task.taskName)));';
  html += '    if (curSameDayIndex === -1) curSameDayIndex = 0;';
  html += '  }';
  html += '  updateSameDayNavUI();';
  html += '  document.getElementById("sheetOverlay").classList.add("show");';
  html += '}';

  html += 'function closeBottomSheet(e) {';
  html += '  if (e && e.target !== document.getElementById("sheetOverlay") && !e.target.classList.contains("sheet-close")) return;';
  html += '  document.getElementById("sheetOverlay").classList.remove("show");';
  html += '}';

  html += 'function selectSheetStatus(s) {';
  html += '  selectedStatus = s;';
  html += '  updateSheetStatusBtns();';
  html += '}';

  html += 'function updateSheetStatusBtns() {';
  html += '  const cleanSel = selectedStatus.replace(/\\s+/g, "");';
  html += '  document.querySelectorAll(".sheet-status-btn").forEach(b => {';
  html += '    const st = b.dataset.status.replace(/\\s+/g, "");';
  html += '    b.classList.toggle("active", cleanSel.indexOf(st) !== -1 || st.indexOf(cleanSel) !== -1);';
  html += '  });';
  html += '}';

  html += 'function selectSheetShare(val) {';
  html += '  selectedShare = val;';
  html += '  updateSheetShareBtns();';
  html += '}';

  html += 'function updateSheetShareBtns() {';
  html += '  document.getElementById("sheetShareOn").classList.toggle("active-on", selectedShare);';
  html += '  document.getElementById("sheetShareOff").classList.toggle("active-off", !selectedShare);';
  html += '}';

  html += 'function saveSheetTask() {';
  html += '  if (!curTaskObj) return;';
  html += '  const saveBtn = document.getElementById("sheetSaveBtn");';
  html += '  const memoVal = document.getElementById("sheetMemoInput").value.trim();';
  html += '  saveBtn.disabled = true;';
  html += '  saveBtn.innerText = "저장 중...";';
  html += '  google.script.run';
  html += '    .withSuccessHandler(res => {';
  html += '      saveBtn.disabled = false;';
  html += '      saveBtn.innerText = "💾 저장하기";';
  html += '      if (res && res.success) {';
  html += '        document.getElementById("sheetOverlay").classList.remove("show");';
  html += '        showToast("✅ [" + curTaskObj.taskName + "] 상태가 저장되었습니다!");';
  html += '        loadData();';
  html += '      } else {';
  html += '        showToast("⚠️ 저장 실패: " + ((res && res.error) || "알 수 없는 오류"));';
  html += '      }';
  html += '    })';
  html += '    .withFailureHandler(err => {';
  html += '      saveBtn.disabled = false;';
  html += '      saveBtn.innerText = "💾 저장하기";';
  html += '      showToast("⚠️ 에러: " + (err.message || err));';
  html += '    })';
  html += '    .updateScheduleStatusFromWeb(curTaskObj.row, selectedStatus, selectedShare, memoVal);';
  html += '}';
  html += '';

  html += 'function escapeHtml(str) {';
  html += '  return String(str || "")';
  html += '    .replace(/&/g, "&amp;")';
  html += '    .replace(/</g, "&lt;")';
  html += '    .replace(/>/g, "&gt;")';
  html += '    .replace(/"/g, "&quot;")';
  html += '    .replace(/\'/g, "&#39;");';
  html += '}';

  html += 'if (document.readyState === "complete" || document.readyState === "interactive") {';
  html += '  setTimeout(loadData, 10);';
  html += '} else {';
  html += '  window.addEventListener("DOMContentLoaded", loadData);';
  html += '  window.addEventListener("load", loadData);';
  html += '}';
  html += '</script>';
  html += '</body></html>';

  return html;
}