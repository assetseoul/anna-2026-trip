/**
 * Scouting Trip 2026 — Apps Script 중간 서버
 * 배포: 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹 앱
 *       (실행 계정: 나 / 액세스: 모든 사용자)  → 발급된 URL을 앱의 CONFIG.API_URL에 입력
 * 일정 탭은 읽기만 하고(assignIds 수동 실행 시에만 일정ID 칸을 채움), 미팅기록·영수증 탭과 드라이브 폴더에만 쓴다.
 */
// 접속 코드는 코드에 쓰지 않고 스크립트 속성(프로젝트 설정 > 스크립트 속성 > ACCESS_CODE)에 저장한다.
function accessCode_() { return PropertiesService.getScriptProperties().getProperty("ACCESS_CODE") || ""; }

var SCHEDULE_TAB = "TRIP SCHEDULE";
var LODGING_TAB  = "HOTEL";
var FLIGHT_TAB   = "FLIGHT & TRAIN";
var AGENCY_TAB   = "AGENCY INFO(필요한 경우 입력,,)";
var NOTES_TAB    = "미팅기록";      // 앱이 자동 생성
var RECEIPT_TAB  = "영수증";        // 앱이 자동 생성
var DRIVE_FOLDER_NAME = "ScoutingTrip2026_영수증";

// 앱 필드 → 일정 탭 헤더 (실제 시트 기준)
var HEADER_MAP = {
  id:"일정ID", city:"도시", date:"날짜", time:"시작시간", duration:"미팅소요시간(분)", type:"일정유형",
  agency:"에이전시명", agent:"미팅 담당자", placeType:"장소유형", addr:"미팅 주소", models:"미팅예정모델 인원",
  modelList:"미팅예정 모델명", gift:"에이전시 선물", upcoming:"업커밍 모델", past:"과거 인타운 모델",
  status:"일정상태", note:"비고"
};
var CITY_ALIAS = {"항저우":"hangzhou","상하이":"shanghai","베이징":"beijing","울란바토르":"ulaanbaatar"};

// 도시 값이 "항저우 -> 상하이"(이동 행)이면 도착 도시를 쓴다.
function cityKey_(v) {
  var parts = String(v || "").split(/->|→|-\>/).map(function (x) { return x.trim(); });
  var last = parts[parts.length - 1];
  return CITY_ALIAS[last] || CITY_ALIAS[String(last).toLowerCase()] || last.toLowerCase();
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var code = accessCode_();
    if (code && req.code !== code) return out({error: "forbidden"});
    if (req.action === "pull") return out(pull());
    if (req.action === "push") return out(push(req.ops || [], req.user));
    return out({error: "unknown action"});
  } catch (err) { return out({error: String(err)}); }
}
function doGet() { return out({ok: true}); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet(name, headers) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh && headers) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function rows(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getDisplayValues(), h = v[0];
  return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}

function pull() {
  var lodging = rows(sheet(LODGING_TAB)).filter(function (r) { return r["도시"]; }).map(function (r) {
    return {city: cityKey_(r["도시"]), from: normDate(r["체크인 날짜"]), to: normDate(r["체크아웃 날짜"]),
            name: r["호텔명"], addr: r["호텔 주소"], addrLocal: r["호텔 주소"], booking: ""};
  });
  var agencies = {};
  rows(sheet(AGENCY_TAB)).forEach(function (r) { if (r["에이전시명"]) agencies[String(r["에이전시명"]).trim().toLowerCase()] = r; });

  var sched = rows(sheet(SCHEDULE_TAB)).filter(function (r) { return r[HEADER_MAP.date]; }).map(function (r) {
    var m = {};
    for (var k in HEADER_MAP) m[k] = r[HEADER_MAP[k]] || "";
    m.city = cityKey_(m.city);
    m.date = normDate(m.date);
    // 장소유형이 호텔인데 주소가 비어 있으면 그날 묵는 호텔 정보를 사용
    if (!m.addr && m.placeType === "호텔") {
      var h = lodging.filter(function (l) { return l.city === m.city && l.from <= m.date && m.date < l.to; })[0];
      if (h) { m.addr = h.addr; m.hotel = h.name; }
    }
    m.addrLocal = m.addr;
    // AGENCY INFO 탭에 있으면 연락처·위챗·오피스 주소 보강
    var a = agencies[String(m.agency).trim().toLowerCase()];
    if (a) {
      var c = String(a["담당자 연락처 / WeChat"] || "").trim();
      if (/^\+?[\d\s\-()]{7,}$/.test(c)) m.phone = c; else m.wechat = c;
      if (!m.addr && a["오피스 주소 (현지어)"]) { m.addr = a["오피스 주소 (현지어)"]; m.addrLocal = m.addr; }
    }
    return m;
  });
  var flights = rows(sheet(FLIGHT_TAB)).filter(function (r) { return r["탑승일"]; }).map(function (r) {
    return {from: r["탑승역"], to: r["도착역"], no: r["편명"], date: normDate(r["탑승일"]), dep: r["탑승시간"], arr: r["도착시간"]};
  });
  var notes = {};
  rows(sheet(NOTES_TAB, ["미팅 ID","미팅 내용","논의 모델","실물 미팅 모델","피드백","팔로업","작성자","수정 시각"])).forEach(function (r) {
    notes[r["미팅 ID"]] = {content: r["미팅 내용"], discussed: j(r["논의 모델"]), met: j(r["실물 미팅 모델"]), followups: j(r["팔로업"]), by: r["작성자"], at: Number(r["수정 시각"]) || null};
  });
  var receipts = rows(sheet(RECEIPT_TAB, ["ID","날짜","분류","금액","통화","메모","연결 미팅 ID","사진 링크","작성자","등록 시각"])).map(function (r) {
    var links = (r["사진 링크"] || "").split("\n").filter(String);
    return {id: r["ID"], date: r["날짜"], cat: r["분류"], amount: Number(r["금액"]), cur: r["통화"], memo: r["메모"], meetingId: r["연결 미팅 ID"], photos: links, by: r["작성자"], at: Number(r["등록 시각"]) || 0};
  });
  return {schedule: sched, lodging: lodging, flights: flights, notes: notes, receipts: receipts};
}
function j(s) { try { return JSON.parse(s || "[]"); } catch (e) { return []; } }
function normDate(s) {
  s = String(s || "").trim();
  var m = s.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  return m ? m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2) : s;
}

function push(ops, user) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var ok = [];
  try {
    ops.forEach(function (op) {
      try {
        if (op.type === "note") saveNote(op.data);
        else if (op.type === "receipt") saveReceipt(op.data);
        else if (op.type === "receiptDelete") deleteReceipt(op.data.id);
        ok.push(op.opId);
      } catch (e) { /* 실패한 건은 앱에 남아 다음 동기화 때 재시도 */ }
    });
  } finally { lock.releaseLock(); }
  return {okIds: ok};
}
function findRow(sh, col, val) {
  var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 2;
  return -1;
}
function saveNote(d) {
  var sh = sheet(NOTES_TAB, ["미팅 ID","미팅 내용","논의 모델","실물 미팅 모델","피드백","팔로업","작성자","수정 시각"]);
  var feedback = (d.met || []).map(function (r) { return r.name + ": " + (r.feedback || ""); }).join("\n");
  var row = [d.meetingId, d.content || "", JSON.stringify(d.discussed || []), JSON.stringify(d.met || []), feedback, JSON.stringify(d.followups || []), d.by || "", d.at || Date.now()];
  var r = findRow(sh, 1, d.meetingId);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function folderFor(dateStr) {
  var roots = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var root = roots.hasNext() ? roots.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  var f = root.getFoldersByName(dateStr);
  return f.hasNext() ? f.next() : root.createFolder(dateStr);
}
function saveReceipt(d) {
  var sh = sheet(RECEIPT_TAB, ["ID","날짜","분류","금액","통화","메모","연결 미팅 ID","사진 링크","작성자","등록 시각"]);
  if (findRow(sh, 1, d.id) > 0) return; // 중복 전송 방지
  var links = [];
  (d.photos || []).forEach(function (p, i) {
    if (String(p).indexOf("data:") !== 0) { links.push(p); return; }
    var b64 = p.split(",")[1];
    var name = [d.date, d.cat, d.amount, d.by].join("_") + (i ? "_" + (i + 1) : "") + ".jpg";
    var file = folderFor(d.date).createFile(Utilities.newBlob(Utilities.base64Decode(b64), "image/jpeg", name));
    links.push(file.getUrl());
  });
  sh.appendRow([d.id, d.date, d.cat, d.amount, d.cur, d.memo || "", d.meetingId || "", links.join("\n"), d.by || "", d.at || Date.now()]);
}
function deleteReceipt(id) {
  var sh = sheet(RECEIPT_TAB); if (!sh) return;
  var r = findRow(sh, 1, id); if (r > 0) sh.deleteRow(r);
}

/**
 * (선택) 일정 탭의 비어 있는 "일정ID" 칸에 고유 번호(M001…)를 채운다. 스크립트 편집기에서 한 번만 직접 실행.
 * 시간·에이전시명이 바뀌어도 미팅 기록과의 연결이 끊기지 않게 하는 용도. 이 함수만 일정 탭에 값을 쓴다.
 * 행을 삭제·정렬한 뒤에는 이미 채운 ID가 그대로 따라가므로 안전하다(ID 열 값은 덮어쓰지 않음).
 */
function assignIds() {
  var sh = sheet(SCHEDULE_TAB); if (!sh) return;
  var v = sh.getDataRange().getValues(), h = v[0];
  var ci = h.indexOf(HEADER_MAP.id), di = h.indexOf(HEADER_MAP.date); if (ci < 0 || di < 0) return;
  var max = 0; v.slice(1).forEach(function (r) { var m = /^M(\d+)$/.exec(String(r[ci])); if (m) max = Math.max(max, +m[1]); });
  for (var i = 1; i < v.length; i++) {
    if (v[i][di] && !v[i][ci]) { max++; sh.getRange(i + 1, ci + 1).setValue("M" + ("00" + max).slice(-3)); }
  }
}
