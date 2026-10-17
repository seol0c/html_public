/**
 * seolcount → 구글 스프레드시트
 *
 * 1. 구글 스프레드시트를 새로 만든다
 * 2. 확장 프로그램 → Apps Script
 * 3. 이 코드를 전부 붙여넣고 저장
 * 4. 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 * 5. 나온 주소를 앱 설정의 "전송주소" 칸에 붙여넣는다
 *
 * 시트 두 개가 자동으로 생깁니다.
 *   기록 : 보낸 기록이 한 줄씩 쌓입니다. 평소엔 이것만 보시면 됩니다.
 *   현황 : 폰 여러 대를 쓸 때 합계를 내려고 쓰는 표입니다. 한 대만 쓰면 안 보셔도 됩니다.
 */

var S_REC  = '기록';
var S_LIVE = '현황';

// 이 시간(분) 넘게 소식 없는 곳은 합계에서 뺀다
var STALE_MIN = 30;

var H_REC  = ['장소','기록','기록자','시작','종료','경과',
              '총입장','총퇴장','남은인원','가장많을때','가장많을때시각',
              '수용인원','수용률','키'];
var REC_KEY_COL = 14;   // 맨 끝 '키' 열 (앱의 기록에서도 같은 값을 볼 수 있습니다)

var H_LIVE = ['장소','상태','갱신시각','현재인원','수용인원','사용률','총입장','총퇴장','기록자'];


function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return out('BUSY'); }
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.type === 'live') live(d); else record(d);
    return out('OK');
  } catch (err) {
    return out('ERR: ' + err);
  } finally {
    lock.releaseLock();
  }
}


/** 폰이 여러 대일 때, 모든 곳의 현재 인원을 합쳐 돌려준다 */
function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'total') return out('seolcount 대기중');
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LIVE);
    if (!sh || sh.getLastRow() < 2) return json({inside:0, gates:0});

    var rows = sh.getRange(2, 1, sh.getLastRow()-1, H_LIVE.length).getValues();
    var now = new Date().getTime();
    var total = 0, gates = 0, cap = 0;

    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][1]) !== '기록중') continue;                             // 정지한 곳은 뺀다
      var t = rows[i][2];
      if (t instanceof Date && (now - t.getTime()) > STALE_MIN*60000) continue;  // 끊긴 곳도 뺀다
      total += num(rows[i][3]);
      cap = Math.max(cap, num(rows[i][4]));
      gates++;
    }
    return json({inside: total, gates: gates, maxCap: cap});
  } catch (err) {
    return json({error: String(err)});
  }
}


/** 기록 한 건 = 한 줄. 같은 기록을 다시 보내면 새 줄이 아니라 그 줄을 고쳐 쓴다. */
function record(d) {
  var sh  = sheet(S_REC, H_REC);
  var cap = num(d.maxCap);
  var row = [str(d.place), str(d.recName), str(d.writer),
             str(d.start), str(d.end), str(d.elapsed),
             num(d.inCnt), num(d.outCnt), num(d.inside),
             num(d.peak), str(d.peakAt),
             cap, cap > 0 ? num(d.peak)/cap : '',      // 수용률 = 가장많을때 ÷ 수용인원
             str(d.key)];
  var at = find(sh, REC_KEY_COL, str(d.key));
  if (at === 0) at = sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setValues([row]);
  sh.getRange(at, 13).setNumberFormat('0.0%');
}


/** 장소마다 한 줄을 계속 덮어쓴다 (폰 여러 대 합계용) */
function live(d) {
  var sh  = sheet(S_LIVE, H_LIVE);
  var cap = num(d.maxCap);
  var row = [str(d.place), str(d.state), new Date(), num(d.inside), cap,
             cap > 0 ? num(d.inside)/cap : '',
             num(d.inCnt), num(d.outCnt), str(d.writer)];
  var at = find(sh, 1, str(d.place));
  if (at === 0) at = sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setValues([row]);
  sh.getRange(at, 3).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(at, 6).setNumberFormat('0.0%');
}


function find(sh, col, key) {
  var last = sh.getLastRow();
  if (last < 2 || !key) return 0;
  var v = sh.getRange(2, col, last-1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === key) return i + 2;
  return 0;
}

function sheet(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function num(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function str(v) { return (v === undefined || v === null) ? '' : String(v); }
function out(m)  { return ContentService.createTextOutput(m).setMimeType(ContentService.MimeType.TEXT); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
