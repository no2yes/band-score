/**
 * 합주 전자악보 - 구글 시트 저장용 Code.gs
 * 시트에 'songs' 탭이 자동으로 만들어지고, 곡 한 줄이 한 행으로 저장됩니다.
 * 접근 코드는 코드에 적지 않고 '스크립트 속성'의 CODE 값으로 관리합니다.
 */
const SHEET_NAME = 'songs';
const HEADER = ['id', 'title', 'artist', 'key', 'bpm', 'time', 'body', 'ts', 'deleted'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function props_() {
  return PropertiesService.getScriptProperties();
}

function authed_(code) {
  const c = props_().getProperty('CODE');
  return !!c && String(code) === c;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 시트가 수식으로 오해하지 않도록 '=' 로 시작하는 글자는 앞에 ' 를 붙입니다.
function safe_(x) {
  const s = String(x == null ? '' : x);
  return s.charAt(0) === '=' ? "'" + s : s;
}

function readAll_() {
  const v = getSheet_().getDataRange().getValues();
  const songs = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    if (!row[0]) continue;
    songs.push({
      id: String(row[0]),
      title: String(row[1]),
      artist: String(row[2]),
      key: String(row[3]),
      bpm: Number(row[4]) || 100,
      time: String(row[5]) || '4/4',
      body: String(row[6]),
      ts: Number(row[7]) || 0,
      deleted: row[8] === true || row[8] === 'TRUE'
    });
  }
  return songs;
}

function readOrder_() {
  try {
    return JSON.parse(props_().getProperty('ORDER') || 'null') || { ids: [], ts: 0 };
  } catch (e) {
    return { ids: [], ts: 0 };
  }
}

// 브라우저로 웹앱 주소를 열었을 때 동작 확인용
function doGet() {
  return json_({ ok: true, message: '합주 전자악보 서버가 동작 중이에요' });
}

// 앱이 보내는 요청: 바뀐 곡을 저장하고, 전체 곡 목록과 순서를 돌려줍니다.
function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'json' });
  }
  if (!authed_(d.code)) return json_({ ok: false, error: 'code' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const v = sh.getDataRange().getValues();
    const rowOf = {};
    for (let r = 1; r < v.length; r++) if (v[r][0]) rowOf[String(v[r][0])] = r + 1;

    (d.songs || []).forEach(function (s) {
      if (!s || !s.id) return;
      const id = String(s.id);
      const ts = Number(s.ts) || 0;
      const r = rowOf[id];
      if (r && (Number(v[r - 1][7]) || 0) > ts) return; // 시트에 더 최신 내용이 있으면 건너뜀

      if (s.deleted) {
        if (r) sh.getRange(r, 8, 1, 2).setValues([[ts, true]]);
        else sh.appendRow([id, '', '', '', 100, '4/4', '', ts, true]);
        return;
      }
      const line = [
        id, safe_(s.title), safe_(s.artist), safe_(s.key),
        Number(s.bpm) || 100, safe_(s.time || '4/4'), safe_(s.body), ts, false
      ];
      if (r) {
        sh.getRange(r, 1, 1, line.length).setValues([line]);
      } else {
        sh.appendRow(line);
        rowOf[id] = sh.getLastRow();
      }
    });

    if (d.order && d.order.ids) {
      const cur = readOrder_();
      const ots = Number(d.order.ts) || 0;
      if (ots >= cur.ts) {
        props_().setProperty('ORDER', JSON.stringify({ ids: d.order.ids.slice(0, 500), ts: ots }));
      }
    }
    return json_({ ok: true, songs: readAll_(), order: readOrder_() });
  } finally {
    lock.releaseLock();
  }
}
