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
  if (d.action === 'ping') return json_({ ok: true });
  if (d.action === 'ask') return ask_(d);

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

/* ===== AI 설명서 검색 (Gemini) =====
 * 위에서부터 차례로 쓰고, 모델별 하루 한도(LIMIT)를 다 쓰면 다음 모델로 내려갑니다.
 * 스크립트 속성 MODELS(쉼표로 구분)로 순서를, LIMIT로 한도를 바꿀 수 있어요.
 * 필요한 스크립트 속성: GEMINI_API_KEY
 */
const AI_MODELS = ['gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-2.5-flash-lite'];
const AI_LIMIT = 500;

function aiModels_() {
  const p = props_().getProperty('MODELS');
  const list = p ? p.split(',').map(function (s) { return s.trim(); }).filter(String) : AI_MODELS;
  return list.length ? list : AI_MODELS;
}

function aiLimit_() {
  return Number(props_().getProperty('LIMIT')) || AI_LIMIT;
}

// 구글의 일일 한도는 미국 태평양 시간 자정에 초기화되므로 그 기준으로 날짜를 셉니다.
function aiUsage_() {
  const day = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyyMMdd');
  let u = null;
  try { u = JSON.parse(props_().getProperty('USAGE') || 'null'); } catch (e) {}
  if (!u || u.day !== day) u = { day: day, n: {}, off: {} };
  return u;
}

function geminiPost_(model, key, body) {
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  return { code: res.getResponseCode(), text: res.getContentText() };
}

// 생각 과정을 최소로 줄여 빠르게 답하게 합니다. 모델이 이 설정을 받지 않으면 빼고 다시 보냅니다.
function callGemini_(model, key, body) {
  const cache = CacheService.getScriptCache();
  const skipThink = cache.get('NT_' + model);
  const b = JSON.parse(JSON.stringify(body));
  if (!skipThink) {
    b.generationConfig.thinkingConfig = /gemini-2\./.test(model) ? { thinkingBudget: 0 } : { thinkingLevel: 'minimal' };
  }
  let r = geminiPost_(model, key, b);
  if (r.code === 400 && !skipThink && /think/i.test(r.text)) {
    cache.put('NT_' + model, '1', 21600);
    delete b.generationConfig.thinkingConfig;
    r = geminiPost_(model, key, b);
  }
  if (r.code !== 200) return { ok: false, code: r.code, msg: r.text };
  try {
    const parts = JSON.parse(r.text).candidates[0].content.parts;
    const text = parts.map(function (p) { return p.text || ''; }).join('').trim();
    if (text) return { ok: true, text: text };
  } catch (e) {}
  return { ok: false, code: 0, msg: 'empty' };
}

function ask_(d) {
  const key = props_().getProperty('GEMINI_API_KEY');
  if (!key) return json_({ ok: false, error: 'nokey' });
  const q = String(d.q || '').trim().slice(0, 200);
  if (!q) return json_({ ok: false, error: 'empty' });
  const ctx = String(d.ctx || '').slice(0, 4000);
  const limit = aiLimit_();
  const u = aiUsage_();
  const body = {
    systemInstruction: {
      parts: [{
        text: '너는 밴드 합주용 전자악보 앱의 도움말 안내원이다. ' +
          '[설명서 발췌]에 있는 내용만 근거로, 음악을 잘 모르는 초보자도 알아듣게 쉬운 말로 2~4문장 안에 답한다. ' +
          '화면의 버튼 이름은 「 」로 감싼다. 순서가 필요하면 1. 2. 3. 으로 짧게 적는다. ' +
          '발췌에 답이 없으면 "설명서에 없는 내용이에요."라고 말하고 가장 가까운 항목을 알려 준다. ' +
          '인사말과 군더더기는 쓰지 않는다.'
      }]
    },
    contents: [{ role: 'user', parts: [{ text: '[설명서 발췌]\n' + ctx + '\n\n[질문]\n' + q }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 400 }
  };

  let exhausted = false, failed = false;
  const models = aiModels_();
  for (let i = 0; i < models.length; i++) {
    const m = models[i];
    if (u.off[m]) continue;
    if ((u.n[m] || 0) >= limit) { exhausted = true; continue; }
    const r = callGemini_(m, key, body);
    if (r.ok) {
      u.n[m] = (u.n[m] || 0) + 1;
      props_().setProperty('USAGE', JSON.stringify(u));
      return json_({ ok: true, answer: r.text, model: m, used: u.n[m], limit: limit });
    }
    if (r.code === 400 && /API key/i.test(r.msg)) return json_({ ok: false, error: 'badkey' });
    if (r.code === 429) {
      // 하루 한도면 오늘은 이 모델을 건너뛰고, 분당 제한이면 이번 요청만 다음 모델로 넘깁니다.
      if (/PerDay/i.test(r.msg)) { u.n[m] = limit; exhausted = true; } else { failed = true; }
    } else if (r.code === 404 || r.code === 403 || r.code === 400) {
      u.off[m] = 1; // 이 키로 쓸 수 없는 모델은 오늘 건너뜁니다.
    } else {
      failed = true;
    }
  }
  props_().setProperty('USAGE', JSON.stringify(u));
  return json_({ ok: false, error: exhausted ? 'quota' : (failed ? 'busy' : 'nomodel') });
}

// 처음 한 번 편집기에서 실행해 권한을 승인하고, 동작을 확인하세요.
function testAsk() {
  const out = ask_({ q: '노래가 너무 높아요', ctx: '[키 바꾸기]\n1. 노래가 너무 높으면 「키 낮추기」를 누릅니다.' });
  Logger.log(out.getContent());
}

// 내 키로 쓸 수 있는 flash-lite 모델 이름을 실행 로그에 보여 줍니다.
function checkModels() {
  const key = props_().getProperty('GEMINI_API_KEY');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true
  });
  const names = (JSON.parse(res.getContentText()).models || [])
    .map(function (m) { return m.name.replace('models/', ''); })
    .filter(function (n) { return /flash-lite/.test(n); });
  Logger.log(names.join('\n') || res.getContentText());
}
