const fs = require("fs");
const vm = require("vm");
const content = fs.readFileSync("현장서식.js", "utf-8");

const sandbox = {
  console: console,
  window: {},
  document: {},
  formatDateYYMMDD: (d) => "2026-10-23",
  escapeHtml: (s) => s,
  isMobileMode: false,
};
vm.createContext(sandbox);
vm.runInContext(content, sandbox);

console.log("1. 현장서식.js VM 컴파일 성공");

if (typeof sandbox.buildSidebarDashboardHtml !== "function") {
  throw new Error("buildSidebarDashboardHtml is not defined");
}

const html = sandbox.buildSidebarDashboardHtml("todo");
console.log("2. buildSidebarDashboardHtml(\"todo\") HTML 생성 성공, 길이:", html.length);

const scriptMatches = html.match(/<script[^>]*>([\s\S]*?)<\/script>/gi);
if (!scriptMatches || scriptMatches.length < 2) {
  throw new Error("Script tags missing in generated HTML");
}

let script2 = scriptMatches[1].replace(/<script[^>]*>/i, "").replace(/<\/script>/i, "");

// 정적 구문 분석
new vm.Script(script2);
console.log("3. 클라이언트 Script 구문 분석 (Syntax Check) 100% 통과!");

// 브라우저 런타임 시뮬레이션
let domHtml = "";
const browserSandbox = {
  console: console,
  window: {
    innerWidth: 400,
    addEventListener: () => {},
    removeEventListener: () => {},
    location: { href: "https://example.com" }
  },
  document: {
    readyState: "complete",
    addEventListener: () => {},
    getElementById: (id) => ({
      style: {},
      classList: { add: () => {}, remove: () => {}, contains: () => false },
      set innerHTML(val) { domHtml = val; },
      get innerHTML() { return domHtml; }
    })
  },
  fetch: (url) => Promise.resolve({
    json: () => Promise.resolve({ success: true })
  }),
  isMobileMode: true,
  setTimeout: () => {},
  clearTimeout: () => {},
  showToast: (msg) => console.log("   [Toast]:", msg)
};

vm.createContext(browserSandbox);
vm.runInContext(script2, browserSandbox);

// Mock Data 주입
vm.runInContext(`
todoMembersList = ["홍길동", "김철수", "이영희"];
todoDataList = [
  { id: "TD1", title: "배관 청소", manager: "홍길동", status: "미완료", dueDate: "2026-10-22", isImportant: false },
  { id: "TD2", title: "현장 안전 점검", manager: "김철수", status: "중요", dueDate: "2026-10-23", isImportant: true },
  { id: "TD3", title: "자재 입고 확인", manager: "이영희", status: "완료", dueDate: "2026-10-20", isImportant: false, completedAt: "2026-10-21 10:00" },
  { id: "TD4", title: "폐기물 처리", manager: "", status: "미완료", dueDate: "", isImportant: false },
];
currentTodoMemberFilter = "all";
`, browserSandbox);

// Test 1: [방안 1] 담당자별 세로 아코디언 카드 뷰
vm.runInContext(`setMobileTodoViewType("manager");`, browserSandbox);
if (!domHtml.includes("홍길동") || !domHtml.includes("배관 청소") || !domHtml.includes("mobile-todo-switcher")) {
  throw new Error("Manager view failed to render");
}
console.log("4. [방안 1] 모바일 담당자별 아코디언 카드 뷰 렌더링 검증 통과 (HTML 길이: " + domHtml.length + ")");

// Test 2: [방안 2] 상태별 3대 탭 칸반 뷰
vm.runInContext(`setMobileTodoViewType("status");`, browserSandbox);
vm.runInContext(`setMobileTodoStatusTab("important");`, browserSandbox);
if (!domHtml.includes("★ 중요") || !domHtml.includes("현장 안전 점검")) {
  throw new Error("Status view (important) failed");
}
vm.runInContext(`setMobileTodoStatusTab("uncompleted");`, browserSandbox);
if (!domHtml.includes("배관 청소") || !domHtml.includes("폐기물 처리")) {
  throw new Error("Status view (uncompleted) failed");
}
vm.runInContext(`setMobileTodoStatusTab("completed");`, browserSandbox);
if (!domHtml.includes("자재 입고 확인")) {
  throw new Error("Status view (completed) failed");
}
console.log("5. [방안 2] 모바일 상태별 3대 탭 전환 칸반 뷰 검증 통과");

// Test 3: [방안 3] 우선순위 타임라인 피드 뷰
vm.runInContext(`setMobileTodoViewType("priority");`, browserSandbox);
if (!domHtml.includes("기한 지난 업무") || !domHtml.includes("중요 업무") || !domHtml.includes("완료된 업무")) {
  throw new Error("Priority view failed to render sections");
}
console.log("6. [방안 3] 모바일 우선순위 타임라인 피드 뷰 검증 통과");

// Test 4: 모바일 원터치 빠른 상태 변경 바 (changeTodoStatusQuick) & 중요 여부 보존 검증
// 4-1: 미완료 -> 중요 변경
vm.runInContext(`changeTodoStatusQuick("TD1", "중요");`, browserSandbox);
let td1 = vm.runInContext(`todoDataList.find(it => it.id === "TD1");`, browserSandbox);
if (td1.status !== "중요" || td1.isImportant !== true) {
  throw new Error("TD1 중요 설정 실패");
}

// 4-2: 중요 -> 완료 변경 시 중요 여부(isImportant: true) 100% 보존 검증
vm.runInContext(`changeTodoStatusQuick("TD1", "완료");`, browserSandbox);
td1 = vm.runInContext(`todoDataList.find(it => it.id === "TD1");`, browserSandbox);
if (td1.status !== "완료" || td1.isImportant !== true || !td1.completedAt) {
  throw new Error("TD1 완료 변경 시 중요 여부(isImportant) 보존 실패");
}
console.log("7-1. 중요 상태에서 완료로 변경 시 isImportant === true 유지 검증 통과!");

// 4-3: 체크박스로 완료 해제 시 원래 중요 상태("중요")로 자동 복원 검증
vm.runInContext(`toggleTodoComplete("TD1", false);`, browserSandbox);
td1 = vm.runInContext(`todoDataList.find(it => it.id === "TD1");`, browserSandbox);
if (td1.status !== "중요" || td1.isImportant !== true || td1.completedAt !== "") {
  throw new Error("TD1 완료 해제 시 중요 상태 자동 복원 실패: status=" + td1.status);
}
console.log("7-2. 완료 해제 시 중요(중요) 상태 자동 복원 검증 통과!");

// 4-4: 미완료로 명시적 변경 시 중요 해제 검증
vm.runInContext(`changeTodoStatusQuick("TD1", "미완료");`, browserSandbox);
td1 = vm.runInContext(`todoDataList.find(it => it.id === "TD1");`, browserSandbox);
if (td1.status !== "미완료" || td1.isImportant !== false) {
  throw new Error("TD1 미완료 변경 시 중요 해제 실패");
}
console.log("7-3. 모바일 원터치 상태 변경 바 & 중요 상태 보존 및 복원 검증 통과");

// Test 5: PC 모드 2차원 매트릭스 테이블 뷰 보존 검증
const pcSandbox = {
  console: console,
  window: {
    innerWidth: 1200,
    addEventListener: () => {},
    removeEventListener: () => {},
    location: { href: "https://example.com" }
  },
  document: {
    readyState: "complete",
    addEventListener: () => {},
    getElementById: (id) => ({
      style: {},
      classList: { add: () => {}, remove: () => {}, contains: () => false },
      set innerHTML(val) { domHtml = val; },
      get innerHTML() { return domHtml; }
    })
  },
  formatDateYYMMDD: (d) => "2026-10-23",
  escapeHtml: (s) => String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"),
  isMobileMode: false,
  setTimeout: () => {},
  clearTimeout: () => {},
  showToast: () => {}
};
vm.createContext(pcSandbox);
vm.runInContext(script2, pcSandbox);
vm.runInContext(`
todoMembersList = ["홍길동", "김철수", "이영희"];
todoDataList = [
  { id: "TD1", title: "배관 청소", manager: "홍길동", status: "미완료", dueDate: "2026-10-22", isImportant: false },
];
currentTodoMemberFilter = "all";
renderTodoDashboard();
`, pcSandbox);

if (!domHtml.includes("todo-matrix-table")) {
  throw new Error("PC Matrix table header missing");
}
console.log("8. PC 2차원 매트릭스 테이블 보존 및 렌더링 검증 통과 (HTML 길이: " + domHtml.length + ")");

console.log("\n=======================================================");
console.log("🎉 ALL TESTS PASSED: 모바일 3대 뷰 & PC 매트릭스 테이블 100% 정상 작동!");
console.log("=======================================================");
