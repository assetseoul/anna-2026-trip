/**
 * Scouting Trip 2026 — Apps Script 중간 서버
 * 배포: 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹 앱
 *       (실행 계정: 나 / 액세스: 모든 사용자)  → 발급된 URL을 앱의 CONFIG.API_URL에 입력
 * 일정 탭은 읽기만 하고, 미팅기록·영수증 탭과 드라이브 폴더에만 쓴다.
 */
var ACCESS_CODE = "";                 // 앱 CONFIG.ACCESS_CODE와 동일하게
var SCHEDULE_TAB = "일정";
var LODGING_TAB  = "숙소";
var NOTES_TAB    = "미팅기록";
var RECEIPT_TAB  = "영수증";
var DRIVE_FOLDER_NAME = "ScoutingTrip2026_영수증";

// 앱 필드 → 시트 헤더(실제 열 이름에 맞게 수정)
var HEADER_MAP = {
  id:"미팅 ID", date:"날짜", city:"도시", time:"시간", agency:"에이전시명", agent:"담당 에이전트",
  models:"예정 모델 인원", gift:"선물", transport:"이동수단", travelMin:"이동시간(분)", status:"상태",
  addr:"주소", addrLocal:"중국어/현지어 주소", lat:"위도", lng:"경도", phone:"연락처", wechat:"위챗/메신저 ID",
  modelList:"모델 명단", giftReady:"선물 준비", note:"비고"
};
var CITY_ALIAS = {"항저우":"hangzhou","hangzhou":"hangzhou","상하이":"shanghai","shanghai":"shanghai","베이징":"beijing","beijing":"beijing","울란바토르":"ulaanbaatar","ulaanbaatar":"ulaanbaatar"};

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (ACCESS_CODE && req.code !== ACCESS_CODE) return out({error: "forbidden"});
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
  var sched = rows(sheet(SCHEDULE_TAB)).filter(function (r) { return r[HEADER_MAP.id]; }).map(function (r) {
    var m = {};
    for (var k in HEADER_MAP) m[k] = r[HEADER_MAP[k]] || "";
    m.city = CITY_ALIAS[String(m.city).toLowerCase().trim()] || CITY_ALIAS[m.city] || m.city;
    m.date = normDate(m.date);
    return m;
  });
  var lodging = rows(sheet(LODGING_TAB)).map(function (r) {
    return {city: CITY_ALIAS[String(r["도시"]).toLowerCase()] || r["도시"], from: normDate(r["체크인"]), to: normDate(r["체크아웃"]),
            name: r["숙소명"], addr: r["주소"], addrLocal: r["현지어 주소"], booking: r["예약 정보"]};
  });
  var notes = {};
  rows(sheet(NOTES_TAB, ["미팅 ID","미팅 내용","논의 모델","실물 미팅 모델","피드백","팔로업","작성자","수정 시각"])).forEach(function (r) {
    notes[r["미팅 ID"]] = {content: r["미팅 내용"], discussed: j(r["논의 모델"]), met: j(r["실물 미팅 모델"]), followups: j(r["팔로업"]), by: r["작성자"], at: Number(r["수정 시각"]) || null};
  });
  var receipts = rows(sheet(RECEIPT_TAB, ["ID","날짜","분류","금액","통화","메모","연결 미팅 ID","사진 링크","작성자","등록 시각"])).map(function (r) {
    var links = (r["사진 링크"] || "").split("\n").filter(String);
    return {id: r["ID"], date: r["날짜"], cat: r["분류"], amount: Number(r["금액"]), cur: r["통화"], memo: r["메모"], meetingId: r["연결 미팅 ID"], photos: links, by: r["작성자"], at: Number(r["등록 시각"]) || 0};
  });
  return {schedule: sched, lodging: lodging, notes: notes, receipts: receipts};
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
