/**
 * 손글씨 단어노트 - Google Apps Script 백엔드 (v7: 등록일(일자별 등록) · 폴더/서브폴더 + 단어 가져오기 · 간격 반복 복습 + 어근 팁 + 난이도 + 인식·저장 한 번에 + 상세 정보(사용빈도·예문·변형·파생어))
 *
 * [설정]
 * 1) 구글시트를 새로 만들고, 확장 프로그램 > Apps Script 에 이 코드를 붙여넣기
 * 2) 프로젝트 설정 > 스크립트 속성에 추가
 *      ANTHROPIC_API_KEY = (Anthropic API 키)
 *    (선택) MODEL = claude-haiku-5-5   (기본값)
 * 3) 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) > URL 복사
 *    -> index_voca.html 의 DEFAULT_GAS_URL 에 붙여넣기 (또는 앱 ⚙️ 설정에 입력)
 *
 * [이미 배포한 적이 있다면]
 *   코드를 이 파일로 교체한 뒤 배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포
 *   (기존 '단어' 시트에는 단계·다음복습·어근팁·난이도 열이 자동으로 추가됩니다. 데이터는 그대로 유지돼요.)
 */

const SHEET_WORDS = '단어';
const SHEET_LOG = '학습기록';
const HEADERS_WORDS = ['id', '단어', '품사', '뜻', '예문', '예문뜻', '등록일', '정답수', '오답수', '마지막학습', '단계', '다음복습', '어근팁', '난이도', '추가정보', '폴더'];
const HEADERS_LOG = ['시각', 'id', '단어', '유형', '결과'];

// 간격 반복: 단계(0~5)별로 "다음 복습까지 며칠"
//  정답이면 단계 +1(최대 5), 오답이면 단계 0(내일 다시)
//  단계 1 → 1일 뒤, 2 → 3일 뒤, 3 → 7일 뒤, 4 → 14일 뒤, 5(완전 암기) → 30일 뒤
const DAYS = [1, 1, 3, 7, 14, 30];
const TZ = 'Asia/Seoul';

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    // 예전 버전으로 만든 시트: 부족한 열 머리글만 채운다 (기존 데이터는 건드리지 않음)
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

// 'yyyy-MM-dd' → 그날 정오(한국 시간). 형식이 아니면 null
function parseYmd_(s) {
  s = String(s || '');
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(s + 'T12:00:00+09:00') : null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function addDays_(ymd, n) {
  const p = String(ymd).split('-');
  const d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// 모든 요청은 text/plain POST 로 받는다 (CORS preflight 회피)
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'recognize': {
        const r = recognize_(req);
        // autosave: 인식 결과가 확실하면 같은 요청 안에서 바로 저장 (왕복 1회 절약)
        if (r.ok && req.autosave) {
          const x = r.result || {};
          const w = String(x.word || '').trim();
          const conf = typeof x.confidence === 'number' ? x.confidence : 1;
          if (w && conf >= 0.7) r.saved = addWord_({ word: w, pos: x.pos, meaning: x.meaning, example: x.example, example_ko: x.example_ko, tip: x.tip, level: x.level, date: req.date });
        }
        return json_(r);
      }
      case 'add':       return json_(addWord_(req));
      case 'list':      return json_(listWords_());
      case 'delete':    return json_(deleteWord_(req));
      case 'update':    return json_(updateWord_(req));
      case 'record':    return json_(record_(req));
      case 'rate':      return json_(rate_(req));
      case 'detail':    return json_(detail_(req));
      case 'import':    return json_(import_(req));
      default:          return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, msg: 'wordnote api alive', version: 7 });
}

/* ---------- Claude: 손글씨 인식 + 뜻/예문/어근 팁 ---------- */
function recognize_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';

  const prompt =
    '이미지는 영어 단어 하나를 손으로 쓴 것입니다. 글자를 정확히 읽어서 JSON 으로만 답하세요.\n' +
    '{"word":"읽은 영단어(소문자, 표제어 원형)","pos":"품사(명/동/형/부 등 한글)",' +
    '"meaning":"한국어 뜻(핵심 1~3개, 쉼표 구분)","example":"해당 단어를 쓴 자연스러운 영어 예문 1개(고등학생 수준)",' +
    '"example_ko":"예문의 한국어 번역",' +
    '"tip":"외우기 팁 한 줄(한국어). 어근·접두사·접미사 구성, 같은 어근의 단어 1~2개, 또는 연상 암기법. 마땅한 게 없으면 빈 문자열",' +
    '"level":한국 고등학생(수능) 기준 난이도 정수 1~5 (1=중학 기초, 2=고1 기본, 3=고2·3 필수·수능 빈출, 4=수능 고난도, 5=상위권·대학 수준),' +
    '"confidence":0~1 숫자}\n' +
    '글자를 알아볼 수 없으면 word 를 빈 문자열로 두세요. JSON 외의 설명은 쓰지 마세요.';

  const payload = {
    model: model,
    max_tokens: 400,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/png', data: req.image } },
        { type: 'text', text: prompt }
      ]
    }]
  };
  return callClaude_(key, payload);
}

function callClaude_(key, payload) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) return { ok: false, error: 'Claude API ' + code + ': ' + body.slice(0, 300) };
  const data = JSON.parse(body);
  const text = (data.content || []).map(function (c) { return c.text || ''; }).join('');
  const m = text.match(/\{[\s\S]*\}/);
  if (!m) return { ok: false, error: '응답 해석 실패: ' + text.slice(0, 200) };
  try {
    return { ok: true, result: JSON.parse(m[0]) };
  } catch (e) {
    return { ok: false, error: 'JSON 파싱 실패: ' + m[0].slice(0, 200) };
  }
}

/* ---------- 단어 CRUD ---------- */
function addWord_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const w = String(req.word || '').trim().toLowerCase();
  if (!w) return { ok: false, error: '단어가 비어 있습니다.' };

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === w) return { ok: false, error: '이미 저장된 단어입니다: ' + w, duplicate: true };
  }
  const id = Utilities.getUuid().slice(0, 8);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  sh.appendRow([id, w, req.pos || '', req.meaning || '', req.example || '', req.example_ko || '', parseYmd_(req.date) || new Date(), 0, 0, '', 0, today, req.tip || '', Number(req.level) || '', '', req.folder || '']);
  return { ok: true, id: id };
}

function listWords_() {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), word: r[1], pos: r[2], meaning: r[3], example: r[4], example_ko: r[5],
      added: r[6] instanceof Date ? r[6].toISOString() : String(r[6]),
      right: Number(r[7]) || 0, wrong: Number(r[8]) || 0,
      last: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
      box: Number(r[10]) || 0, due: dstr_(r[11]), tip: String(r[12] || ''), level: Number(r[13]) || 0, extra: String(r[14] || ''), folder: String(r[15] || '')
    });
  }
  return { ok: true, words: out };
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function deleteWord_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const row = findRow_(sh, req.id);
  if (row < 0) return { ok: false, error: '단어를 찾을 수 없습니다.' };
  sh.deleteRow(row);
  return { ok: true };
}

function updateWord_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const row = findRow_(sh, req.id);
  if (row < 0) return { ok: false, error: '단어를 찾을 수 없습니다.' };
  sh.getRange(row, 2, 1, 5).setValues([[
    String(req.word || '').trim().toLowerCase(), req.pos || '', req.meaning || '', req.example || '', req.example_ko || ''
  ]]);
  sh.getRange(row, 13, 1, 2).setValues([[req.tip || '', Number(req.level) || '']]);
  if (req.folder !== undefined) sh.getRange(row, 16).setValue(String(req.folder || ''));
  const dd = parseYmd_(req.date);
  if (dd) sh.getRange(row, 7).setValue(dd);   // 등록일 변경
  return { ok: true };
}

/* ---------- 학습 기록 + 복습 일정 계산 ---------- */
function record_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const log = getSheet_(SHEET_LOG, HEADERS_LOG);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const updates = [];
  try {
    const data = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < data.length; i++) rowOf[String(data[i][0])] = i;

    const now = new Date();
    const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const logRows = [];

    (req.items || []).forEach(function (it) { // [{id, word, type, correct}]
      const i = rowOf[String(it.id)];
      if (i !== undefined) {
        const r = data[i];
        let right = Number(r[7]) || 0, wrong = Number(r[8]) || 0, box = Number(r[10]) || 0;
        if (it.correct) { right++; box = Math.min(box + 1, 5); } else { wrong++; box = 0; }
        const due = addDays_(today, DAYS[box]);
        r[7] = right; r[8] = wrong; r[9] = now; r[10] = box; r[11] = due; // 같은 단어가 또 오면 이어서 계산
        sh.getRange(i + 1, 8, 1, 5).setValues([[right, wrong, now, box, due]]);
        updates.push({ id: String(it.id), box: box, due: due });
      }
      logRows.push([now, it.id, it.word, it.type, it.correct ? '정답' : '오답']);
    });
    if (logRows.length) log.getRange(log.getLastRow() + 1, 1, logRows.length, 5).setValues(logRows);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, updates: updates };
}

/* ---------- 이미 저장된 단어의 난이도 일괄 매기기 (한 번에 최대 30개) ---------- */
function rate_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';

  const items = (req.items || []).slice(0, 30);
  if (!items.length) return { ok: true, levels: {} };
  const list = items.map(function (it) { return it.id + '\t' + it.word + '\t' + (it.meaning || ''); }).join('\n');
  const prompt =
    '다음은 영어 단어 목록입니다(id, 단어, 뜻을 탭으로 구분). 한국 고등학생(수능) 기준 난이도를 1~5 정수로 매기세요.\n' +
    '1=중학 기초, 2=고1 기본, 3=고2·3 필수·수능 빈출, 4=수능 고난도, 5=상위권·대학 수준.\n' +
    'JSON 으로만 답하세요: {"levels":{"<id>":숫자, ...}}\n\n' + list;

  const r = callClaude_(key, { model: model, max_tokens: 1500, messages: [{ role: 'user', content: prompt }] });
  if (!r.ok) return r;
  const levels = (r.result && r.result.levels) || {};

  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const data = sh.getDataRange().getValues();
  const rowOf = {};
  for (let i = 1; i < data.length; i++) rowOf[String(data[i][0])] = i + 1;

  const out = {};
  Object.keys(levels).forEach(function (id) {
    const v = Math.round(Number(levels[id]));
    if (v >= 1 && v <= 5 && rowOf[id]) {
      sh.getRange(rowOf[id], 14).setValue(v);
      out[id] = v;
    }
  });
  return { ok: true, levels: out };
}

/* ---------- 카드 상세 정보: 사용빈도·예문 3개·형태 변화·파생어 (필요할 때 한 단어씩) ---------- */
function detail_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';
  const word = String(req.word || '').trim();
  if (!word) return { ok: false, error: '단어가 비어 있습니다.' };

  const prompt =
    '영어 단어 "' + word + '"' + (req.meaning ? ' (뜻: ' + req.meaning + ')' : '') + ' 에 대해 JSON 으로만 답하세요.\n' +
    '{"freq":일상·시험에서 쓰이는 빈도 1~5 정수(5=매우 자주),' +
    '"examples":[{"en":"자연스러운 영어 예문(고등학생 수준)","ko":"한국어 번역"}, 서로 다른 상황으로 3개],' +
    '"forms":"형태 변화를 한국어로 간결하게. 동사는 3인칭 단수 현재·과거·과거분사·현재분사, 명사는 복수형, 형용사·부사는 비교급·최상급. 변하지 않으면 빈 문자열",' +
    '"derivs":["파생어 한 개당 한 문자열: 단어 (품사) 뜻  — 최대 4개, 없으면 빈 배열"]}\n' +
    'JSON 외의 설명은 쓰지 마세요.';

  const r = callClaude_(key, { model: model, max_tokens: 700, messages: [{ role: 'user', content: prompt }] });
  if (!r.ok) return r;
  const x = r.result || {};
  const extra = {
    freq: Math.max(1, Math.min(5, Math.round(Number(x.freq)) || 3)),
    examples: (x.examples || []).slice(0, 3).map(function (e) { return { en: String(e.en || ''), ko: String(e.ko || '') }; }),
    forms: String(x.forms || ''),
    derivs: (x.derivs || []).slice(0, 4).map(String)
  };
  if (req.id) {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(req.id)) { sh.getRange(i + 1, 15).setValue(JSON.stringify(extra)); break; }
    }
  }
  return { ok: true, extra: extra };
}

/* ---------- 단어 묶음 가져오기 (폴더/서브폴더 지정) ----------
   req.items = [{ word, pos, meaning, example, example_ko, folder }]  (한 번에 최대 200개)
   - 이미 있는 단어: 새로 만들지 않고, 폴더가 비어 있으면 그 폴더로만 넣어 준다 */
function import_(req) {
  const items = (req.items || []).slice(0, 200);
  if (!items.length) return { ok: true, added: 0, moved: 0, skipped: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < data.length; i++) if (data[i][0]) rowOf[String(data[i][1]).toLowerCase()] = i;
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const now = new Date();
    const rows = [];
    let moved = 0, skipped = 0;
    items.forEach(function (it, n) {
      const w = String(it.word || '').trim().toLowerCase();
      if (!w) return;
      const folder = String(it.folder || '').trim();
      if (w in rowOf) {
        const i = rowOf[w];
        if (!String(data[i][15] || '') && folder) { sh.getRange(i + 1, 16).setValue(folder); moved++; } else skipped++;
        return;
      }
      rowOf[w] = -1;
      rows.push([Utilities.getUuid().slice(0, 8), w, it.pos || '', it.meaning || '', it.example || '', it.example_ko || '', now, 0, 0, '', 0, today, it.tip || '', Number(it.level) || '', '', folder]);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, added: rows.length, moved: moved, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}
