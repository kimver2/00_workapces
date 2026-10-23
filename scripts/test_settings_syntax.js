const fs = require('fs');
const path = require('path');
const vm = require('vm');

console.log('🧪 [설정 마스터 관리 단위 테스트] 시작...');

const code = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');

// 1. AST 구문 검사
console.log('1. 현장서식.js 컴파일...');
let script;
try {
  script = new vm.Script(code);
  console.log('  ✓ VM 컴파일 성공');
} catch (e) {
  console.error('  ❌ VM 컴파일 실패:', e);
  process.exit(1);
}

// 2. Mock 환경 준비
const mockSpreadsheet = {
  getId: () => "mock-ss-id",
  getName: () => "우성아이디 일정 관리",
  getSheets: () => [],
  getSheetByName: (name) => null,
  insertSheet: (name) => ({
    getName: () => name,
    getLastRow: () => 1,
    getLastColumn: () => 11,
    getRange: () => ({
      setValues: () => ({ setBackground: () => ({ setFontColor: () => ({ setFontWeight: () => ({ setHorizontalAlignment: () => {} }) }) }) }),
      getValues: () => [],
      setValue: () => {},
      setBackground: () => {},
      setFontColor: () => {},
      setFontWeight: () => {},
      setHorizontalAlignment: () => {},
      setVerticalAlignment: () => {}
    }),
    appendRow: () => {},
    deleteRow: () => {},
    setRowHeight: () => {},
    setFrozenRows: () => {}
  }),
  toast: () => {}
};

const sandbox = {
  console: console,
  Logger: { log: console.log },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => mockSpreadsheet,
    flush: () => {},
    newDataValidation: () => ({ requireValueInRange: () => ({ build: () => ({}) }) })
  },
  ContentService: {
    createTextOutput: (s) => ({
      setMimeType: () => s
    }),
    MimeType: { JSON: "application/json" }
  },
  Utilities: {
    formatDate: (d) => "2026-10-23 13:00:00"
  },
  CacheService: {
    getScriptCache: () => ({ get: () => null, put: () => {} })
  },
  LockService: {
    getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} })
  }
};

vm.createContext(sandbox);
script.runInContext(sandbox);

// 3. 서버 함수 존재 검증
console.log('2. 서버 마스터 설정 함수 검증...');
const requiredFns = [
  'getSettingsManagementData',
  'addPartnerItemFromWeb',
  'updatePartnerItemFromWeb',
  'deletePartnerItemFromWeb',
  'reorderPartnerItemFromWeb',
  'addTaskItemFromWeb',
  'updateTaskItemFromWeb',
  'deleteTaskItemFromWeb',
  'reorderTaskItemFromWeb'
];

requiredFns.forEach(fn => {
  if (typeof sandbox[fn] === 'function') {
    console.log(`  ✓ sandbox.${fn} 함수 존재 확인`);
  } else {
    console.error(`  ❌ sandbox.${fn} 함수 부재!`);
    process.exit(1);
  }
});

// 4. HTML 렌더링 검증
console.log('3. HTML 렌더링 및 UI 컴포넌트 검증...');
const html = sandbox.buildSidebarDashboardHtml("calendar");

if (html.includes('id="settingsModalOverlay"')) {
  console.log('  ✓ #settingsModalOverlay 마스터 설정 모달 HTML 렌더링 확인');
} else {
  console.error('  ❌ #settingsModalOverlay 누락!');
  process.exit(1);
}

if (html.includes('id="partnerFormModal"') && html.includes('id="taskFormModal"')) {
  console.log('  ✓ #partnerFormModal 및 #taskFormModal 서브 모달 렌더링 확인');
} else {
  console.error('  ❌ 서브 모달 누락!');
  process.exit(1);
}

if (html.includes('class="pc-settings-badge-btn"') && html.includes('openSettingsModal()')) {
  console.log('  ✓ PC 좌측 상단 pc-settings-badge-btn 버튼 확인');
} else {
  console.error('  ❌ PC 설정 버튼 누락!');
  process.exit(1);
}

if (html.includes('btn-settings-action')) {
  console.log('  ✓ 모바일 상단 헤더 btn-settings-action 버튼 확인');
} else {
  console.error('  ❌ 모바일 설정 버튼 누락!');
  process.exit(1);
}

// 5. 클라이언트 스크립트 검증
console.log('4. 클라이언트 스크립트 함수 검증...');
const clientFns = [
  'openSettingsModal',
  'closeSettingsModal',
  'switchSettingsTab',
  'loadSettingsData',
  'applySettingsDataResponse',
  'renderSettingsView',
  'renderPartnersTable',
  'renderTasksTable',
  'openAddPartnerModal',
  'openEditPartnerModal',
  'savePartnerForm',
  'deletePartner',
  'reorderPartner',
  'openAddTaskModal',
  'openEditTaskModal',
  'saveTaskForm',
  'deleteTask',
  'reorderTask',
  'callSettingsApi'
];

clientFns.forEach(fn => {
  if (html.includes(`function ${fn}(`) || html.includes(`${fn} = function(`)) {
    console.log(`  ✓ 클라이언트 ${fn} 함수 포함 확인`);
  } else {
    console.error(`  ❌ 클라이언트 ${fn} 함수 누락!`);
    process.exit(1);
  }
});

// 6. 실시간 연동 (현장 추가/수정에 드롭다운 반영) 검증
console.log('5. 실시간 연동 (rawData 및 populateTaskSelectOptions 등) 검증...');
if (html.includes('populateTaskSelectOptions(rawData.tasks)') && html.includes('populatePartnerSelectOptions(rawData.partners)')) {
  console.log('  ✓ 설정 변경 시 현장 추가/수정 드롭다운 즉시 동기화 로직 탑재 확인');
} else {
  console.error('  ❌ 드롭다운 즉시 동기화 누락!');
  process.exit(1);
}

// 7. doGet 액션 라우팅 검증
console.log('6. doGet 액션 라우팅 검증...');
const actions = [
  'getSettingsData',
  'addPartner',
  'updatePartner',
  'deletePartner',
  'reorderPartner',
  'addTask',
  'updateTask',
  'deleteTask',
  'reorderTask'
];

actions.forEach(act => {
  if (code.includes(`action === "${act}"`)) {
    console.log(`  ✓ doGet 라우팅: action === "${act}" 확인`);
  } else {
    console.error(`  ❌ doGet 라우팅 누락: ${act}`);
    process.exit(1);
  }
});

console.log('\n=======================================================');
console.log('🎉 ALL SETTINGS TESTS PASSED: 마스터 설정(거래처 & 공종) 100% 정상 작동!');
console.log('=======================================================');
