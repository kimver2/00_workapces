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
// 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 메뉴 자동 등록
// ========================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 현장 관리")
    .addItem("📅 구글 캘린더 전체 동기화 (시트 ➔ 캘린더 일괄 반영)", "syncAllSiteSchedulesToGoogleCalendar")
    .addItem("⚡ 실시간 캘린더 연동 켜기 (시트 수정 시 실시간 캘린더 반영)", "enableRealtimeCalendarSync")
    .addItem("❌ 실시간 캘린더 연동 끄기", "disableRealtimeCalendarSync")
    .addSeparator()
    .addItem("📥 캘린더 수정사항 가져오기 (수동 즉시 반영)", "syncFromCalendarToSheet")
    .addItem("⚡ 캘린더 ➔ 시트 1분 주기 자동 동기화 켜기 (초고속)", "enableCalendarToSheet1MinSync")
    .addItem("⏰ 캘린더 ➔ 시트 5분 주기 자동 동기화 켜기 (표준)", "enableCalendarToSheet5MinSync")
    .addItem("⏹️ 캘린더 ➔ 시트 자동 동기화 끄기", "disableCalendarToSheetAutoSync")
    .addSeparator()
    .addItem("⚙️ 연동 캘린더 변경 / 확인", "configureTargetCalendar")
    .addSeparator()
    .addItem("🔄 대시보드 달력 새로고침 (선택된 월 동기화)", "syncAllSiteSchedulesToDashboard")
    .addSeparator()
    .addItem("➕ 신규 현장 추가 (상단 4개 행 + 아코디언)", "addNewSiteAtTop")
    .addItem("📖 모든 현장 펼치기 (전체 공정 보기)", "expandAllSites")
    .addItem("📁 모든 현장 접기 (현장명만 요약 보기)", "collapseAllSites")
    .addSeparator()
    .addItem("🗑️ 150행 하단 행 일괄 삭제", "deleteRowsBelow")
    .addItem("🎨 서식 및 아코디언 완전 재설정", "applySiteScheduleFormatting")
    .addToUi();
}


// ========================================================
// 2. [실시간 스프레드시트 UI 감지 (단순 트리거)]
// - 대시보드에서 '월' 또는 '년도' 변경 시 달력 즉시 새로고침
// - 현장데이터 시트에서 A, B, C열 수정 시 대시보드 달력 즉시 갱신
// ========================================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  // ① [대시보드 시트] 상단 컨트롤(1~3행)의 년도 또는 월 변경 시
  if (isDashboardSheet(sheet)) {
    if (row <= 3) {
      refreshDashboardCalendar(ss);
      return;
    }
  }

  // ② [현장데이터 시트] 데이터 행(2행 이상)의 A, B, C열(현장, 날짜, 공종) 수정 시 대시보드 갱신
  if (isSiteScheduleSheet(sheet)) {
    if (row >= 2 && (col === 1 || col === 2 || col === 3)) {
      refreshDashboardCalendar(ss);
      return;
    }
  }
}


// ========================================================
// 3. [구글 캘린더 실시간 연동 (설치형 트리거: 시트 ➔ 캘린더)]
// ========================================================

/**
 * 실시간 캘린더 자동 연동 켜기 (설치형 onEdit 트리거 자동 등록)
 */
function enableRealtimeCalendarSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var functionName = "handleCalendarEditTrigger";

  // 이미 설치되어 있는지 확인
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      SpreadsheetApp.getUi().alert(
        "⚡ 실시간 연동 활성화 상태",
        "이미 실시간 구글 캘린더 연동이 켜져 있습니다!\n\n'현장데이터' 시트에서 날짜나 공종을 입력하면 자동으로 구글 캘린더에 즉시 반영됩니다.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
  }

  // 신규 설치형 트리거 생성
  ScriptApp.newTrigger(functionName)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ 실시간 캘린더 연동 활성화 완료",
    "실시간 구글 캘린더 연동이 성공적으로 켜졌습니다!\n\n" +
    "앞으로 '현장데이터' 시트에서 날짜나 공종, 담당자를 입력하거나 수정하면 실시간으로 구글 캘린더에 자동 추가/수정/삭제됩니다.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 실시간 캘린더 자동 연동 끄기
 */
function disableRealtimeCalendarSync() {
  var functionName = "handleCalendarEditTrigger";
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  if (removed > 0) {
    SpreadsheetApp.getUi().alert("실시간 구글 캘린더 연동이 정상적으로 해제되었습니다.");
  } else {
    SpreadsheetApp.getUi().alert("현재 실시간 연동이 켜져 있지 않습니다.");
  }
}

/**
 * 설치형 onEdit 트리거 핸들러 (실시간 캘린더 수정/등록/삭제 및 작성자/수정자 추적)
 */
function handleCalendarEditTrigger(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (!isSiteScheduleSheet(sheet)) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row < 2) return;

  // A~H열(1~8열: 현장, 날짜, 공종, 협력업체, 체크박스, 비고, 담당자, 진행상태) 수정 시
  if (col >= 1 && col <= 8) {
    try {
      ensureCalendarSyncColumns(sheet);

      var now = new Date();
      var nowStr = formatDateTime(now);
      var currentUser = getCurrentUserIdentifier();

      // 최초작성일/작성자(K, L열)가 비어있으면 최초 작성 정보 기록
      var cellCreatedDate = sheet.getRange(row, 11);
      if (!cellCreatedDate.getValue()) {
        cellCreatedDate.setValue(nowStr);
        sheet.getRange(row, 12).setValue(currentUser);
      }

      // 최종수정일/수정자(M, N열) 갱신
      sheet.getRange(row, 13).setValue(nowStr);
      sheet.getRange(row, 14).setValue(currentUser);

      // 시트의 최종 수정 시각 밀리초 타임스탬프 기록 (J열: 10열)
      sheet.getRange(row, 10).setValue(now.getTime());

      // 캘린더 동기화 실행 (Last-Modified-Wins 검사 포함)
      syncSingleRowToCalendar(sheet, row);
    } catch (err) {
      Logger.log("캘린더 실시간 동기화 에러 (행 " + row + "): " + err.toString());
    }
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

  for (var i = 0; i < values.length; i++) {
    var row = i + 2;
    var rowVals = values[i];
    var existingId = String(rowVals[8] || "").trim(); // I열: 캘린더ID
    var sheetUpdated = parseInt(rowVals[9], 10) || 0; // J열: 시트수정시각
    var currentSheetDate = rowVals[1];               // B열: 날짜

    if (!existingId) continue;

    // 캘린더에서 이벤트 조회
    var event = null;
    try {
      event = cal.getEventById(existingId);
    } catch (e) {
      event = null;
    }

    if (!event) continue;

    var calUpdatedDate = event.getLastUpdated();
    var calUpdated = calUpdatedDate.getTime();

    // 캘린더의 날짜 추출
    var isAllDay = event.isAllDayEvent();
    var calStart = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
    var calEnd = isAllDay ? event.getAllDayEndDate() : event.getEndTime();
    if (!calStart) continue;

    // 종일 일정의 종료일 처리 (구글 캘린더 종일 종료일은 exclusive)
    var realCalEnd = null;
    if (isAllDay && calEnd) {
      var diffDays = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
      if (diffDays > 1) {
        realCalEnd = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
      }
    } else if (!isAllDay && calEnd) {
      if (!isDateEqual(calStart, calEnd)) {
        realCalEnd = calEnd;
      }
    }

    // 시트의 현재 날짜 파싱
    var currentRange = parseScheduleDateRange(currentSheetDate);
    var currentStart = currentRange ? currentRange.startDate : null;
    var currentEnd = (currentRange && currentRange.isRange) ? currentRange.displayEndDate : null;

    // 날짜 일치 여부 비교
    var isSame = isDateEqual(calStart, currentStart) && isDateEqual(realCalEnd, currentEnd);

    if (!isSame) {
      // 날짜가 서로 다름! [마지막 수정 시간 기반 비교 (Last-Modified-Wins)]
      if (calUpdated > sheetUpdated) {
        // [캘린더가 더 최신] ➔ 캘린더의 날짜를 시트에 반영!
        var newDateStr = realCalEnd ? (formatDateYYMMDD(calStart) + " ~ " + formatDateYYMMDD(realCalEnd)) : formatDateYYMMDD(calStart);
        sheet.getRange(row, 2).setValue(newDateStr);
        sheet.getRange(row, 10).setValue(calUpdated);

        // 캘린더에서 수정한 것이므로 최종수정일/수정자 기록!
        sheet.getRange(row, 13).setValue(formatDateTime(calUpdatedDate));
        sheet.getRange(row, 14).setValue("구글 캘린더");

        updatedCount++;
      } else {
        // [시트가 더 최신] ➔ 시트의 날짜를 캘린더에 덮어쓰기!
        if (currentRange) {
          if (currentRange.isRange) {
            event.setTime(currentRange.startDate, currentRange.endDate);
          } else {
            event.setAllDayDate(currentRange.startDate);
          }
          sheet.getRange(row, 10).setValue(Date.now());
          updatedCount++;
        }
      }
    }
  }

  if (updatedCount > 0) {
    SpreadsheetApp.flush();
    refreshDashboardCalendar(ss);
    ss.toast("최종 수정 시간 기준으로 " + updatedCount + "건의 일정이 최신화되었습니다.", "동기화 완료", 3);
  } else {
    ss.toast("캘린더와 시트의 일정이 이미 동일합니다. (변경사항 없음)", "동기화 확인", 2);
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
function buildEventDetails(rowValues, rowNumber) {
  var siteName = String(rowValues[0] || "").trim(); // A열: 현장
  var dateVal = rowValues[1];                      // B열: 날짜
  var taskName = String(rowValues[2] || "").trim(); // C열: 공종
  var partner = String(rowValues[3] || "").trim();  // D열: 협력업체 (인덱스 3)
  var memo = String(rowValues[5] || "").trim();     // F열: 비고 (인덱스 5)
  var manager = String(rowValues[6] || "").trim();  // G열: 담당자 (인덱스 6)
  var status = String(rowValues[7] || "").trim();   // H열: 진행상태 (인덱스 7)

  // 이력 정보 (K, L, M, N열)
  var createdAt = String(rowValues[10] || "").trim(); // K열: 최초작성일
  var createdBy = String(rowValues[11] || "").trim(); // L열: 최초작성자
  var updatedAt = String(rowValues[12] || "").trim(); // M열: 최종수정일
  var updatedBy = String(rowValues[13] || "").trim(); // N열: 최종수정자

  if (!siteName && !taskName) return null;

  var parsedDate = parseScheduleDateRange(dateVal);
  if (!parsedDate) return null;

  // 제목 포맷: [현장명] 공종 (비고가 있으면 비고 포함하여 중복 방지 및 구분)
  var title = "";
  var taskDisplay = taskName + (memo ? " (" + memo + ")" : "");
  if (siteName && taskName) {
    title = "[" + siteName + "] " + taskDisplay;
  } else if (siteName) {
    title = "[" + siteName + "]" + (memo ? " (" + memo + ")" : "");
  } else {
    title = taskDisplay;
  }

  // 설명(Description) 포맷
  var descLines = [
    "📌 현장명: " + (siteName || "-"),
    "🔨 공종: " + (taskName || "-")
  ];
  if (manager) descLines.push("👤 담당자: " + manager);
  if (status) descLines.push("🚦 진행상태: " + status);
  if (partner) descLines.push("🏢 협력업체: " + partner);
  if (memo) descLines.push("📝 비고: " + memo);
  descLines.push("--------------------------------");
  if (createdAt) descLines.push("⏱️ 최초작성: " + createdAt + (createdBy ? " (" + createdBy + ")" : ""));
  if (updatedAt) descLines.push("🔄 최종수정: " + updatedAt + (updatedBy ? " (" + updatedBy + ")" : ""));
  descLines.push("📊 스프레드시트 연동 (Row: " + rowNumber + ")");

  return {
    title: title,
    siteName: siteName,
    parsedDate: parsedDate,
    description: descLines.join("\n")
  };
}

/**
 * 단일 행을 구글 캘린더에 동기화 (Last-Modified-Wins 적용)
 */
function syncSingleRowToCalendar(sheet, row) {
  if (row < 2) return;
  var rowValues = sheet.getRange(row, 1, 1, 14).getValues()[0];
  var existingEventId = String(rowValues[8] || "").trim(); // I열: 캘린더ID
  var sheetUpdated = parseInt(rowValues[9], 10) || 0;     // J열: 시트최종수정시각

  var details = buildEventDetails(rowValues, row);
  var cal = getTargetCalendar();
  if (!cal) return;

  // 1) 날짜나 필수 정보가 없는 경우: 기존 캘린더 이벤트가 있다면 삭제
  if (!details) {
    if (existingEventId) {
      try {
        var ev = cal.getEventById(existingEventId);
        if (ev) ev.deleteEvent();
      } catch (err) {}
      sheet.getRange(row, 9).clearContent();
      sheet.getRange(row, 10).clearContent();
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

  if (event) {
    var calUpdatedDate = event.getLastUpdated();
    var calUpdated = calUpdatedDate.getTime();

    // 만약 캘린더가 시트보다 3초 이상 더 최근에 수정되었다면:
    // (캘린더의 수정을 덮어쓰지 않고 오히려 캘린더 날짜를 시트로 가져옴!)
    if (calUpdated > sheetUpdated + 3000) {
      var isAllDay = event.isAllDayEvent();
      var calStart = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
      var calEnd = isAllDay ? event.getAllDayEndDate() : event.getEndTime();
      var realCalEnd = null;
      if (isAllDay && calEnd) {
        var diffDays = Math.round((calEnd.getTime() - calStart.getTime()) / (24 * 60 * 60 * 1000));
        if (diffDays > 1) {
          realCalEnd = new Date(calEnd.getTime() - 24 * 60 * 60 * 1000);
        }
      } else if (!isAllDay && calEnd) {
        if (!isDateEqual(calStart, calEnd)) realCalEnd = calEnd;
      }
      var newDateStr = realCalEnd ? (formatDateYYMMDD(calStart) + " ~ " + formatDateYYMMDD(realCalEnd)) : formatDateYYMMDD(calStart);
      sheet.getRange(row, 2).setValue(newDateStr);
      sheet.getRange(row, 10).setValue(calUpdated);
      sheet.getRange(row, 13).setValue(formatDateTime(calUpdatedDate));
      sheet.getRange(row, 14).setValue("구글 캘린더");

      var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();
      refreshDashboardCalendar(ss);
      return;
    }

    // 시트가 더 최근에 수정된 경우 -> 캘린더 이벤트 속성 업데이트
    event.setTitle(details.title);
    event.setDescription(details.description);
    event.setLocation(details.siteName);
    if (details.parsedDate.isRange) {
      event.setTime(details.parsedDate.startDate, details.parsedDate.endDate);
    } else {
      event.setAllDayDate(details.parsedDate.startDate);
    }
    sheet.getRange(row, 10).setValue(Date.now());
  } else {
    // 3) 기존 ID가 없거나 유효하지 않은 경우 -> 중복 방지 검색 후 신규 생성
    var searchStart = new Date(details.parsedDate.startDate.getTime() - 24 * 60 * 60 * 1000);
    var searchEnd = new Date(details.parsedDate.startDate.getTime() + 48 * 60 * 60 * 1000);
    var candidateEvents = cal.getEvents(searchStart, searchEnd);
    var foundExisting = null;

    for (var i = 0; i < candidateEvents.length; i++) {
      if (candidateEvents[i].getTitle() === details.title) {
        foundExisting = candidateEvents[i];
        break;
      }
    }

    if (foundExisting) {
      event = foundExisting;
      event.setDescription(details.description);
      event.setLocation(details.siteName);
      if (details.parsedDate.isRange) {
        event.setTime(details.parsedDate.startDate, details.parsedDate.endDate);
      } else {
        event.setAllDayDate(details.parsedDate.startDate);
      }
    } else {
      var opt = {
        description: details.description,
        location: details.siteName
      };
      if (details.parsedDate.isRange) {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, details.parsedDate.endDate, opt);
      } else {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, opt);
      }
    }

    if (event) {
      sheet.getRange(row, 9).setValue(event.getId());
      sheet.getRange(row, 10).setValue(Date.now());
    }
  }

  ensureCalendarSyncColumns(sheet);
}

/**
 * 현장데이터 시트의 전체 일정을 구글 캘린더에 일괄 동기화
 */
function syncAllSiteSchedulesToGoogleCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

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
  var now = new Date();
  var nowMs = now.getTime();
  var nowStr = formatDateTime(now);
  var currentUser = getCurrentUserIdentifier();

  for (var i = 0; i < values.length; i++) {
    var row = i + 2;
    var rowVals = values[i];
    var existingId = String(rowVals[8] || "").trim();

    // 작성일이 비어있으면 보완
    if (!rowVals[10]) {
      sheet.getRange(row, 11).setValue(nowStr);
      sheet.getRange(row, 12).setValue(currentUser);
      rowVals[10] = nowStr;
      rowVals[11] = currentUser;
    }
    if (!rowVals[12]) {
      sheet.getRange(row, 13).setValue(nowStr);
      sheet.getRange(row, 14).setValue(currentUser);
      rowVals[12] = nowStr;
      rowVals[13] = currentUser;
    }

    var details = buildEventDetails(rowVals, row);

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

    if (event) {
      event.setTitle(details.title);
      event.setDescription(details.description);
      event.setLocation(details.siteName);
      if (details.parsedDate.isRange) {
        event.setTime(details.parsedDate.startDate, details.parsedDate.endDate);
      } else {
        event.setAllDayDate(details.parsedDate.startDate);
      }
      sheet.getRange(row, 10).setValue(nowMs);
      updatedCount++;
    } else {
      var opt = { description: details.description, location: details.siteName };
      if (details.parsedDate.isRange) {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, details.parsedDate.endDate, opt);
      } else {
        event = cal.createAllDayEvent(details.title, details.parsedDate.startDate, opt);
      }
      if (event) {
        sheet.getRange(row, 9).setValue(event.getId());
        sheet.getRange(row, 10).setValue(nowMs);
        syncedCount++;
      }
    }
  }

  SpreadsheetApp.flush();
  var calName = cal.getName();
  SpreadsheetApp.getUi().alert(
    "📅 구글 캘린더 동기화 완료",
    "대상 캘린더: [" + calName + "]\n\n" +
    "- 신규 등록: " + syncedCount + "건\n" +
    "- 업데이트: " + updatedCount + "건\n\n" +
    "구글 캘린더에서 등록된 현장 일정을 확인해보세요!",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
  }

  // 9, 10열(시스템 열) 숨김
  try {
    sheet.hideColumns(9, 2);
  } catch(e) {}
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

  // 5. 현장데이터 시트에서 전체 데이터 읽기 (1~6열: 현장, 날짜, 공종, 협력업체, 공유, 비고)
  var lastRow = siteSheet.getLastRow();
  if (lastRow < 2) return;

  var siteData = siteSheet.getRange(2, 1, lastRow - 1, 6).getValues();

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

    if (!siteName || !dateVal || !taskName) continue;

    var rangeInfo = parseScheduleDateRange(dateVal);
    if (!rangeInfo) continue;

    // 연도 비교 (연도가 지정된 경우)
    if (rangeInfo.year && targetYear && rangeInfo.year !== targetYear) continue;

    // 월 비교 (현재 대시보드에 선택된 월과 일치해야 함)
    if (rangeInfo.month !== targetMonth) continue;

    var day = rangeInfo.day;
    if (!schedulesByDay[day]) schedulesByDay[day] = [];

    // 동일 현장/날짜에 같은 공종이 여러 개일 때 비고가 있으면 비고를 붙여서 구분
    var itemText = "[" + siteName + "] " + taskName + (memo ? " (" + memo + ")" : "");

    // 비고까지 완전히 똑같은 중복이 아니라면 모두 추가
    schedulesByDay[day].push(itemText);
    matchedCount++;
  }

  // 8. 해당 월의 일정들을 달력의 해당 날짜 밑 칸에 쓰기
  for (var day in schedulesByDay) {
    var pos = cellPositions[day];
    if (pos) {
      dashSheet.getRange(pos.row, pos.col).setValue(schedulesByDay[day].join("\n"));
    }
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
  var defYear = 2026;
  var defMonth = 10;
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
      if (n >= 2020 && n <= 2035 && !foundYear) {
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
 * 반환값: { startDate, endDate, displayEndDate, isRange, year, month, day }
 */
function parseScheduleDateRange(val) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    var d = new Date(val.getFullYear(), val.getMonth(), val.getDate());
    return {
      startDate: d,
      endDate: null,
      displayEndDate: null,
      isRange: false,
      year: d.getFullYear(),
      month: d.getMonth() + 1,
      day: d.getDate()
    };
  }

  var str = String(val).trim();
  if (!str) return null;

  // 기간 날짜 구분자 감지: "Oct 21 → Oct 22", "10/21 ~ 10/22", "26.10.21 - 26.10.22"
  var separatorRegex = /\s*(?:→|~|\s-\s|\s~|\.\.\.)\s*/;
  if (separatorRegex.test(str)) {
    var parts = str.split(separatorRegex);
    if (parts.length >= 2) {
      var p1 = parseDateComponents(parts[0].trim());
      var p2 = parseDateComponents(parts[1].trim());
      if (p1 && p2) {
        var startYear = p1.year || p2.year || 2026;
        var endYear = p2.year || p1.year || startYear;
        var sDate = new Date(startYear, p1.month - 1, p1.day);
        // 구글 캘린더 종일 일정의 종료일은 exclusive이므로 +1일
        var eDate = new Date(endYear, p2.month - 1, p2.day + 1);
        var dispEnd = new Date(endYear, p2.month - 1, p2.day);
        return {
          startDate: sDate,
          endDate: eDate,
          displayEndDate: dispEnd,
          isRange: true,
          year: startYear,
          month: p1.month,
          day: p1.day
        };
      }
    }
  }

  // 단일 날짜 파싱
  var comp = parseDateComponents(str);
  if (comp) {
    var y = comp.year || 2026;
    var sDate = new Date(y, comp.month - 1, comp.day);
    return {
      startDate: sDate,
      endDate: null,
      displayEndDate: null,
      isRange: false,
      year: y,
      month: comp.month,
      day: comp.day
    };
  }

  return null;
}

/**
 * 단일 날짜 문자열 파싱 유틸리티 ({ year, month, day })
 */
function parseDateComponents(val) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    return { year: val.getFullYear(), month: val.getMonth() + 1, day: val.getDate() };
  }

  var str = String(val).trim();
  if (!str) return null;

  // YY.MM.DD 또는 YYYY.MM.DD (예: 26.10.09, 2026-10-09)
  var mYmd = str.match(/^(\d{2,4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (mYmd) {
    var y = parseInt(mYmd[1], 10);
    if (y < 100) y += 2000;
    return { year: y, month: parseInt(mYmd[2], 10), day: parseInt(mYmd[3], 10) };
  }

  // MM/DD 또는 MM.DD (예: 10/9)
  var mMd = str.match(/^(\d{1,2})[-./](\d{1,2})$/);
  if (mMd) {
    return { year: 2026, month: parseInt(mMd[1], 10), day: parseInt(mMd[2], 10) };
  }

  // 10월 9일
  var mKor = str.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (mKor) {
    return { year: 2026, month: parseInt(mKor[1], 10), day: parseInt(mKor[2], 10) };
  }

  // Oct 9, Nov 14 등 영문 월
  var months = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  var mEng = str.match(/^([A-Za-z]{3})\s+(\d{1,2})/);
  if (mEng) {
    var mon = months[mEng[1].toLowerCase()];
    if (mon) return { year: 2026, month: mon, day: parseInt(mEng[2], 10) };
  }

  return null;
}


// ========================================================
// 8. [시트 판별 헬퍼]
// ========================================================
function isDashboardSheet(sheet) {
  if (!sheet) return false;
  var name = sheet.getName().toLowerCase();
  return name.indexOf("dashboard") !== -1 || name.indexOf("대시보드") !== -1 || sheet.getSheetId() === 244837509;
}

function isSiteScheduleSheet(sheet) {
  if (!sheet) return false;
  var name = sheet.getName();
  if (name.indexOf("현장") !== -1) return true;
  if (sheet.getLastRow() >= 1 && sheet.getRange(1, 1).getValue() === "현장") return true;
  return false;
}

function findDashboardSheet(ss) {
  var s = ss.getSheetByName("dashboard") || ss.getSheetByName("Dashboard") || ss.getSheetByName("대시보드");
  if (s) return s;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (isDashboardSheet(sheets[i])) return sheets[i];
  }
  return null;
}

function findSiteScheduleSheet(ss) {
  var s = ss.getSheetByName("현장데이터") || ss.getSheetByName("현장 데이터") || ss.getSheetByName("01_현장데이터");
  if (s) return s;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (isSiteScheduleSheet(sheets[i])) return sheets[i];
  }
  return sheets[0];
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
  sheet.getRange(2, 5, 4, 1).insertCheckboxes();

  [1, 2, 3, 5, 7, 8, 11, 12, 13, 14].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, 4, 1).setHorizontalAlignment("left");
  });

  try {
    var sourceRow = sheet.getLastRow();
    sheet.getRange(sourceRow, 7).copyTo(sheet.getRange(2, 7, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sheet.getRange(sourceRow, 3).copyTo(sheet.getRange(2, 3, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sheet.getRange(sourceRow, 8).copyTo(sheet.getRange(2, 8, 4, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
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
 * 지정 행(기본 150행) 하단의 불필요한 모든 빈 행 삭제
 */
function deleteRowsBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "🗑️ 하단 행 일괄 삭제",
    "몇 번째 행 아래의 행들을 모두 삭제하시겠습니까?\n(그냥 확인을 누르면 150행 하단 전체가 삭제됩니다)",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var targetRow = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(targetRow) || targetRow <= 1) targetRow = 150;

  var maxRows = sheet.getMaxRows();
  sheet.showRows(1, maxRows);
  sheet.expandAllRowGroups();
  SpreadsheetApp.flush();

  maxRows = sheet.getMaxRows();
  if (maxRows > targetRow) {
    var rowsToDelete = maxRows - targetRow;
    sheet.deleteRows(targetRow + 1, rowsToDelete);
    ss.toast(targetRow + "행 아래의 " + rowsToDelete + "개 행이 영구 삭제되었습니다.", "삭제 완료", 3);
  } else {
    ss.toast("현재 시트의 총 행 수가 " + targetRow + "행 이하입니다.", "삭제할 행 없음", 3);
  }
}

/**
 * 현장데이터 시트 서식 및 아코디언 완전 재설정
 */
function applySiteScheduleFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();

  var maxRows = sheet.getMaxRows();
  var lastRow = Math.max(sheet.getLastRow(), 2);

  sheet.showRows(1, maxRows);

  try {
    sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch(e) {}

  ensureCalendarSyncColumns(sheet);

  sheet.getRange(2, 5, maxRows - 1, 1).insertCheckboxes();

  var rangeC = sheet.getRange(2, 3, maxRows - 1, 1);
  var rangeG = sheet.getRange(2, 7, maxRows - 1, 1);
  var rangeH = sheet.getRange(2, 8, maxRows - 1, 1);
  var rules = [];

  // 상태별 색상
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("예정").setBackground("#fef3c7").setFontColor("#b45309").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("진행 중").setBackground("#dbeafe").setFontColor("#1d4ed8").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("완료").setBackground("#dcfce7").setFontColor("#15803d").setBold(true).setRanges([rangeH]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("AS").setBackground("#fee2e2").setFontColor("#b91c1c").setBold(true).setRanges([rangeH]).build());

  // 담당자별 색상
  var managerColorMap = [
    { key: "조항만", bg: "#dbeafe", font: "#1e40af" },
    { key: "송지은", bg: "#fce7f3", font: "#be185d" },
    { key: "전예린", bg: "#f3e8ff", font: "#7e22ce" },
    { key: "영하",   bg: "#ccfbf1", font: "#0f766e" }
  ];
  managerColorMap.forEach(function(item) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(item.key).setBackground(item.bg).setFontColor(item.font).setBold(true).setRanges([rangeG]).build());
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
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(k).setBackground(item.bg).setFontColor(item.font).setBold(true).setRanges([rangeC]).build());
    });
  });

  sheet.setConditionalFormatRules(rules);

  // 헤더 스타일링 (1~14열)
  var lastCol = Math.max(sheet.getLastColumn(), 14);
  sheet.getRange(1, 1, 1, 8)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // 정렬
  [1, 2, 3, 5, 7, 8, 11, 12, 13, 14].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("left");
  });

  // 열 너비 설정
  var colWidths = [180, 140, 110, 150, 70, 220, 100, 95];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // 아코디언 그룹 재생성
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

  sheet.showRows(1, maxRows);
  sheet.expandAllRowGroups();
  try {
    sheet.hideColumns(9, 2);
  } catch(e) {}
  SpreadsheetApp.flush();

  ss.toast("모든 서식과 아코디언이 정상적으로 복구되었습니다.", "새로고침 완료", 3);
}
