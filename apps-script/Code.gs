// 수업 허브 서버 (Google Apps Script)
// 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 그대로 붙여 넣고
// 배포 > 배포 관리 > (연필) > 버전: 새 버전 > 배포 로 다시 배포하세요.
// ('새 배포'를 만들면 주소가 바뀌어 index.html, teacher.html의 API 주소도 바꿔야 합니다)

const KEY = "daejeong1121";            // 수업 코드: 원하는 값으로 바꾸세요 (index.html, teacher.html과 같아야 함)
const TABS = ["progress", "help", "challenge", "share"];
const DEFAULT_TABS = ["progress", "help", "challenge"];   // tabs를 지정하지 않으면 돌려줄 시트
const MAX_CELL = 45000;                // 시트 한 칸은 최대 5만 자

// =, +, -, @ 로 시작하는 글이 수식으로 바뀌지 않게 글자로 저장
function safe(v) {
  const s = String(v == null ? "" : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.key !== KEY || TABS.indexOf(d.type) < 0) return out({ ok: false });
  const extra = JSON.stringify(d.extra || {});
  if (extra.length > MAX_CELL) return out({ ok: false, error: "too_long" });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(d.type) || ss.insertSheet(d.type);
    sh.appendRow([new Date(), safe(d.student), safe(d.mission), safe(d.value), extra]);
  } finally {
    lock.releaseLock();
  }
  return out({ ok: true });
}

// ?key=...&tabs=share 처럼 필요한 시트만 골라 받을 수 있습니다
function doGet(e) {
  if (e.parameter.key !== KEY) return out({ ok: false });
  const want = e.parameter.tabs
    ? String(e.parameter.tabs).split(",").filter(function (t) { return TABS.indexOf(t) >= 0; })
    : DEFAULT_TABS;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {};
  want.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    data[name] = sh ? sh.getDataRange().getValues() : [];
  });
  return out({ ok: true, data: data });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
