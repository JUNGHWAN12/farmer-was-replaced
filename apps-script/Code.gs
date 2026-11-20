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

// doPost / doGet 은 학생·선생님 화면이 요청할 때 Google이 실행하는 함수입니다.
// 편집기에서 직접 실행하지 말고, 점검할 때는 위쪽 함수 목록에서 testServer 를 골라 실행하세요.
function doPost(e) {
  if (!e || !e.postData) return noRequest("doPost");
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
  if (!e || !e.parameter) return noRequest("doGet");
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

function noRequest(name) {
  Logger.log(name + " 은(는) 편집기에서 직접 실행하는 함수가 아닙니다. testServer 를 실행해 점검하세요.");
  return out({ ok: false, error: "no_request" });
}

// 편집기 점검용: 함수 목록에서 testServer 를 고르고 ▶ 실행 → 아래 '실행 로그'에서 결과 확인
// (기록을 쓰지 않고 읽기만 합니다. 처음 실행하면 권한 승인 창이 뜹니다)
function testServer() {
  const res = JSON.parse(doGet({ parameter: { key: KEY, tabs: TABS.join(",") } }).getContent());
  if (!res.ok) { Logger.log("실패: 응답이 ok가 아닙니다"); return; }
  TABS.forEach(function (name) {
    Logger.log(name + " 시트: " + res.data[name].length + "줄");
  });
  Logger.log("정상입니다. share 시트는 학생이 처음 공유할 때 자동으로 만들어집니다.");
}
