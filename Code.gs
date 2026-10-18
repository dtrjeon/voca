/**
 * handvoca - Google Apps Script 백엔드 (v9: 백업 복원 · v8: 연어(Collocation) 자동 생성 · v7: 등록일(일자별 등록) · 폴더/서브폴더 + 단어 가져오기 · 간격 반복 복습 + 어근 팁 + 난이도 + 인식·저장 한 번에 + 상세 정보(사용빈도·예문·변형·파생어))
 *
 * [설정]
 * 1) 구글시트를 새로 만들고, 확장 프로그램 > Apps Script 에 이 코드를 붙여넣기
 * 2) 프로젝트 설정 > 스크립트 속성에 추가
 *      ANTHROPIC_API_KEY = (Anthropic API 키)
 *    (선택) MODEL = claude-haiku-5-5   (기본값)
 *    (선택) TEACHER_CODE = (돕는이 코드)  ← 넣으면 AI 기능(손글씨 인식·뜻/예문·상세·연어·영영)이 코드를 아는 앱에서만 동작해요. 코드는 이 파일에 쓰지 마세요.
 *    (선택) 러너 로그인: 'TEACHER_CODE' 를 정하면 잠금이 켜져요. 돕는이는 이 코드로, 러너는 '학생' 시트의 이름+코드로 들어와요.
 *      '학생' 시트(자동 생성)의 A열 이름, B열 코드(서식: 일반 텍스트)를 선생님이 적거나 앱의 '학생 현황'에서 추가하세요.
 * 3) 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) > URL 복사
 *    -> index_voca.html 의 DEFAULT_GAS_URL 에 붙여넣기 (또는 앱 ⚙️ 설정에 입력)
 *
 * [이미 배포한 적이 있다면]
 *   코드를 이 파일로 교체한 뒤 배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포
 *   (기존 '단어' 시트에는 단계·다음복습·어근팁·난이도 열이 자동으로 추가됩니다. 데이터는 그대로 유지돼요.)
 */

const SHEET_WORDS = '단어';
const SHEET_LOG = '학습기록';
const HEADERS_WORDS = ['id', '단어', '품사', '뜻', '예문', '예문뜻', '등록일', '정답수', '오답수', '마지막학습', '단계', '다음복습', '어근팁', '난이도', '추가정보', '폴더', '구분', '소유자'];
const HEADERS_LOG = ['시각', 'id', '단어', '유형', '결과', '이름'];
const SHEET_STU = '학생', HEADERS_STU = ['이름', '코드', '목표대학', '시작일', '가입일', '마지막접속', '앱버전'];
const SHEET_STATE = '진행', HEADERS_STATE = ['이름', '순번', '내용'];   // 토마토 표시·회독 날짜 등을 JSON 조각으로 저장
const SHEET_PROG = '진도', HEADERS_PROG = ['이름', 'id', '정답수', '오답수', '마지막학습', '단계', '다음복습'];   // 학생별 암기 진도(공유 단어)
let CUR_AU_ = { role: 'teacher' };

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
    // 권한: 스크립트 속성 TEACHER_CODE 가 있으면 돕는이 코드 또는 학생(이름+코드)만 쓸 수 있다. 없으면 예전처럼 모두 열려 있다.
    const au = auth_(req);
    if (au.blocked) return json_({ ok: false, auth: true, error: '잠시 후에 다시 시도해 주세요 (코드를 여러 번 틀렸어요)' });
    if (req.action === 'login') return json_(login_(req, au));
    if (au.role === 'none') return json_({ ok: false, auth: true, error: '로그인이 필요해요 (이름과 코드를 입력하세요)' });
    if (au.role === 'student' && !STUDENT_OK_[req.action]) return json_({ ok: false, forbidden: true, error: '돕는이만 할 수 있어요' });
    CUR_AU_ = au;
    switch (req.action) {
      case 'recognize': {
        const t0 = Date.now();
        const r = recognize_(req);
        const t1 = Date.now();
        // autosave: 인식 결과가 확실하면 같은 요청 안에서 바로 저장 (왕복 1회 절약)
        if (r.ok && req.autosave) {
          const x = r.result || {};
          const w = String(x.word || '').trim();
          const conf = typeof x.confidence === 'number' ? x.confidence : 1;
          if (w && conf >= 0.7) r.saved = addWord_({ word: w, pos: x.pos, meaning: x.meaning, example: x.example, example_ko: x.example_ko, tip: x.tip, level: x.level, date: req.date });
        }
        r.t = { ai: t1 - t0, save: Date.now() - t1 };   // 걸린 시간(ms): AI 인식 / 시트 저장
        return json_(r);
      }
      case 'enrich':    return json_(enrich_(req));
      case 'addmany':   return json_(addMany_(req));
      case 'add':       return json_(addWord_(req));
      case 'list':      return json_(listWords_());
      case 'delete':    return json_(deleteWord_(req));
      case 'update':    return json_(updateWord_(req));
      case 'record':    return json_(record_(req));
      case 'rate':      return json_(rate_(req));
      case 'detail':    return json_(detail_(req));
      case 'import':    return json_(import_(req));
      case 'colloc':    return json_(colloc_(req));
      case 'defen':     return json_(defen_(req));
      case 'renamefolder': return json_(renameFolder_(req));
      case 'undoimport': return json_(undoImport_(req));
      case 'restore':   return json_(restore_(req));
      case 'profile':   return json_(profile_(req));
      case 'statesave': return json_(stateSave_(req));
      case 'students':  return json_(students_());
      case 'studentadd': return json_(studentAdd_(req));
      default:          return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 학생이 쓸 수 있는 기능 (나머지는 선생님만)
const STUDENT_OK_ = { list: 1, add: 1, addmany: 1, recognize: 1, enrich: 1, detail: 1, record: 1, update: 1, delete: 1, profile: 1, statesave: 1 };

function doGet() {
  return json_({ ok: true, msg: 'wordnote api alive', version: 22, locked: !!PropertiesService.getScriptProperties().getProperty('TEACHER_CODE') });
}

/* ---------- Claude: 손글씨 인식 + 뜻/예문/어근 팁 ---------- */
function recognize_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';

  // 빠른 인식(quick): 철자·품사·뜻·난이도만 먼저 돌려준다. 예문·팁은 enrich 가 따로 채운다.
  const quick = !!req.quick;
  const prompt = quick ?
    '이미지는 영어 단어 하나를 손으로 쓴 것입니다. 글자를 정확히 읽어서 JSON 으로만 답하세요.\n' +
    '{"word":"쓴 글씨 그대로 읽은 영단어(소문자). 철자만 바로잡고, 원형으로 바꾸거나 -ing/-ed/복수형을 임의로 고치지 마세요",' +
    '"confidence":0~1 숫자,"pos":"품사(명/동/형/부 등 한글)","meaning":"한국어 뜻(핵심 1~3개, 쉼표 구분)",' +
    '"level":한국 고등학생(수능) 기준 난이도 정수 1~5 (1=중학 기초, 2=고1 기본, 3=고2·3 필수·수능 빈출, 4=수능 고난도, 5=상위권·대학 수준)}\n' +
    '글자를 알아볼 수 없으면 word 를 빈 문자열로 두세요. JSON 외의 설명은 쓰지 마세요.' :
    '이미지는 영어 단어 하나를 손으로 쓴 것입니다. 글자를 정확히 읽어서 JSON 으로만 답하세요. 아래 순서를 지키세요.\n' +
    '{"word":"쓴 글씨 그대로 읽은 영단어(소문자). 철자만 바로잡고, 원형으로 바꾸거나 -ing/-ed/복수형을 임의로 고치지 마세요(예: booming 은 booming)",' +
    '"confidence":0~1 숫자,' +
    '"pos":"품사(명/동/형/부 등 한글)",' +
    '"meaning":"한국어 뜻(핵심 1~3개, 쉼표 구분)",' +
    '"level":한국 고등학생(수능) 기준 난이도 정수 1~5 (1=중학 기초, 2=고1 기본, 3=고2·3 필수·수능 빈출, 4=수능 고난도, 5=상위권·대학 수준),' +
    '"example":"해당 단어를 쓴 자연스러운 영어 예문 1개(고등학생 수준, 짧게)",' +
    '"example_ko":"예문의 한국어 번역",' +
    '"tip":"외우기 팁 한 줄(한국어). 어근·접두사·접미사 구성, 같은 어근의 단어 1~2개, 또는 연상 암기법. 마땅한 게 없으면 빈 문자열"}\n' +
    '글자를 알아볼 수 없으면 word 를 빈 문자열로 두세요. JSON 외의 설명은 쓰지 마세요.';

  const payload = {
    model: model,
    max_tokens: quick ? 250 : 900,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/png', data: req.image } },
        { type: 'text', text: prompt }
      ]
    }]
  };
  const r = callClaudeText_(key, payload);
  if (!r.ok) return r;
  const parsed = parseRecog_(r.text);
  if (!parsed) return { ok: false, error: '응답 해석 실패: ' + String(r.text).slice(0, 200) };
  return { ok: true, result: parsed };
}

// 인식 응답 해석: 정상 JSON 이면 그대로, 답이 중간에 잘렸으면 읽은 항목만 건져 쓴다
function parseRecog_(text) {
  text = String(text || '');
  const m = text.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch (e) { /* 아래에서 건진다 */ } }
  const s = text.indexOf('{'); if (s < 0) return null;
  const body = text.slice(s);
  const str = function (k) {
    const r = body.match(new RegExp('"' + k + '"\\s*:\\s*"((?:[^"\\\\]|\\\\.)*)"'));
    return r ? r[1].replace(/\\"/g, '"').replace(/\\n/g, ' ') : '';
  };
  const num = function (k) {
    const r = body.match(new RegExp('"' + k + '"\\s*:\\s*([0-9.]+)'));
    return r ? Number(r[1]) : undefined;
  };
  const word = str('word'); if (!word) return null;
  return { word: word, pos: str('pos'), meaning: str('meaning'), example: str('example'), example_ko: str('example_ko'), tip: str('tip'), level: num('level'), confidence: num('confidence') };
}

// 예문·번역·암기 팁만 글자(텍스트)로 채운다 — 이미지가 없어서 빠르다
function enrich_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';
  const word = String(req.word || '').trim();
  if (!word) return { ok: false, error: '단어가 비어 있습니다.' };
  const prompt =
    '영어 단어 "' + word + '"' + (req.meaning ? ' (뜻: ' + req.meaning + ')' : '') + ' 에 대해 JSON 으로만 답하세요.\n' +
    '{"example":"해당 단어를 쓴 자연스러운 영어 예문 1개(고등학생 수준, 짧게)",' +
    '"example_ko":"예문의 한국어 번역",' +
    '"tip":"외우기 팁 한 줄(한국어). 어근·접두사·접미사 구성, 같은 어근의 단어 1~2개, 또는 연상 암기법. 마땅한 게 없으면 빈 문자열"}\n' +
    'JSON 외의 설명은 쓰지 마세요.';
  const t0 = Date.now();
  const r = callClaudeText_(key, { model: model, max_tokens: 400, messages: [{ role: 'user', content: prompt }] });
  if (!r.ok) return r;
  const m = String(r.text).match(/\{[\s\S]*\}/);
  let x = null;
  if (m) { try { x = JSON.parse(m[0]); } catch (e) { x = null; } }
  if (!x) {
    const body = String(r.text);
    const str = function (k) { const q = body.match(new RegExp('"' + k + '"\\s*:\\s*"((?:[^"\\\\]|\\\\.)*)"')); return q ? q[1].replace(/\\"/g, '"') : ''; };
    x = { example: str('example'), example_ko: str('example_ko'), tip: str('tip') };
  }
  return { ok: true, example: String(x.example || ''), example_ko: String(x.example_ko || ''), tip: String(x.tip || ''), t: { ai: Date.now() - t0 } };
}

function callClaudeText_(key, payload) {
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
  return { ok: true, text: (data.content || []).map(function (c) { return c.text || ''; }).join('') };
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

  // 중복 검사는 '단어' 열(2번째)과 '소유자' 열만 읽는다. 공유 단어와 내(학생) 단어만 비교한다.
  const owner = CUR_AU_.role === 'student' ? CUR_AU_.name : '';
  const last = sh.getLastRow();
  if (last > 1) {
    const col = sh.getRange(2, 2, last - 1, 1).getValues(), ow = sh.getRange(2, 18, last - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      const o = String(ow[i][0] || '');
      if ((!o || o === owner) && String(col[i][0]).toLowerCase() === w) return { ok: false, error: '이미 저장된 단어입니다: ' + w, duplicate: true };
    }
  }
  const id = Utilities.getUuid().slice(0, 8);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  sh.appendRow([id, w, req.pos || '', req.meaning || '', req.example || '', req.example_ko || '', parseYmd_(req.date) || new Date(), 0, 0, '', 0, today, req.tip || '', Number(req.level) || '', '', req.folder || '', req.src === 'import' ? 'import' : 'hand', owner]);
  return { ok: true, id: id };
}

// 여러 단어를 한 번에 저장한다 (시트를 한 번만 읽고 한 번만 쓴다)
function addMany_(req) {
  const items = (req.items || []).slice(0, 100);
  if (!items.length) return { ok: true, results: [] };
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const results = [], rows = [];
  try {
    const have = {};
    const owner = CUR_AU_.role === 'student' ? CUR_AU_.name : '';
    const last = sh.getLastRow();
    if (last > 1) {
      const ws = sh.getRange(2, 2, last - 1, 1).getValues(), ow = sh.getRange(2, 18, last - 1, 1).getValues();
      ws.forEach(function (r, i) { const o = String(ow[i][0] || ''); if (!o || o === owner) have[String(r[0]).toLowerCase()] = true; });
    }
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    items.forEach(function (it) {
      const w = String(it.word || '').trim().toLowerCase();
      if (!w) { results.push({ lid: it.lid, ok: false, error: '단어가 비어 있습니다.' }); return; }
      if (have[w]) { results.push({ lid: it.lid, ok: false, duplicate: true }); return; }
      have[w] = true;
      const id = Utilities.getUuid().slice(0, 8);
      rows.push([id, w, it.pos || '', it.meaning || '', it.example || '', it.example_ko || '', parseYmd_(it.date) || new Date(), 0, 0, '', 0, today, it.tip || '', Number(it.level) || '', '', it.folder || '', it.src === 'import' ? 'import' : 'hand', owner]);
      results.push({ lid: it.lid, ok: true, id: id });
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, results: results };
}

function listWords_() {
  const au = CUR_AU_, stu = au.role === 'student';
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const rows = sh.getDataRange().getValues();
  const prog = stu ? progOf_(au.name) : null;
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const owner = String(r[17] || '');
    if (stu && owner && owner !== au.name) continue;          // 다른 학생이 쓴 단어는 보이지 않는다
    const id = String(r[0]), p = stu ? prog[id] : null;
    out.push({
      id: id, word: r[1], pos: r[2], meaning: r[3], example: r[4], example_ko: r[5],
      added: r[6] instanceof Date ? r[6].toISOString() : String(r[6]),
      right: stu ? (p ? p.right : 0) : (Number(r[7]) || 0), wrong: stu ? (p ? p.wrong : 0) : (Number(r[8]) || 0),
      last: stu ? (p ? p.last : '') : (r[9] instanceof Date ? r[9].toISOString() : String(r[9] || '')),
      box: stu ? (p ? p.box : 0) : (Number(r[10]) || 0), due: stu ? (p ? p.due : dstr_(r[11])) : dstr_(r[11]),
      tip: String(r[12] || ''), level: Number(r[13]) || 0, extra: String(r[14] || ''), folder: String(r[15] || ''), src: String(r[16] || ''), owner: owner
    });
  }
  return { ok: true, words: out };
}

// 학생 한 명의 암기 진도(공유 단어용): id -> {right, wrong, last, box, due}
function progOf_(name) {
  const sh = getSheet_(SHEET_PROG, HEADERS_PROG), m = {};
  const last = sh.getLastRow();
  if (last < 2) return m;
  sh.getRange(2, 1, last - 1, 7).getValues().forEach(function (r) {
    if (String(r[0]) !== name) return;
    m[String(r[1])] = { right: Number(r[2]) || 0, wrong: Number(r[3]) || 0, last: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''), box: Number(r[5]) || 0, due: dstr_(r[6]) };
  });
  return m;
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

// 학생은 자기 이름으로 쓴 손글씨 단어만 고치거나 지울 수 있다 (선생님은 모두)
function mayEdit_(sh, row) {
  if (CUR_AU_.role !== 'student') return true;
  const v = sh.getRange(row, 17, 1, 2).getValues()[0];
  return String(v[0]) === 'hand' && String(v[1] || '') === CUR_AU_.name;
}

function deleteWord_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const row = findRow_(sh, req.id);
  if (row < 0) return { ok: false, error: '단어를 찾을 수 없습니다.' };
  if (!mayEdit_(sh, row)) return { ok: false, forbidden: true, error: '내가 쓴 단어만 지울 수 있어요' };
  sh.deleteRow(row);
  return { ok: true };
}

function updateWord_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const row = findRow_(sh, req.id);
  if (row < 0) return { ok: false, error: '단어를 찾을 수 없습니다.' };
  if (!mayEdit_(sh, row)) return { ok: false, forbidden: true, error: '내가 쓴 단어만 고칠 수 있어요' };
  const stu = CUR_AU_.role === 'student';
  sh.getRange(row, 2, 1, 5).setValues([[
    String(req.word || '').trim().toLowerCase(), req.pos || '', req.meaning || '', req.example || '', req.example_ko || ''
  ]]);
  sh.getRange(row, 13, 1, 2).setValues([[req.tip || '', Number(req.level) || '']]);
  if (!stu && req.folder !== undefined) sh.getRange(row, 16).setValue(String(req.folder || ''));
  if (!stu && (req.src === 'hand' || req.src === 'import')) sh.getRange(row, 17).setValue(req.src);   // 구분(손쓰기/업로드)
  const dd = parseYmd_(req.date);
  if (dd) sh.getRange(row, 7).setValue(dd);   // 등록일 변경
  return { ok: true };
}

/* ---------- 학습 기록 + 복습 일정 계산 ---------- */
// ───────── 학생 / 선생님 로그인 ─────────
function normCode_(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ''); }

// 요청이 누구에게서 왔는지: {role:'teacher'|'student'|'none', name, blocked}
function auth_(req) {
  const tc = PropertiesService.getScriptProperties().getProperty('TEACHER_CODE');
  if (!tc) return { role: 'teacher' };                       // 돕는이 코드를 안 정했으면 예전처럼 모두 열림
  if (req.code && normCode_(req.code) === normCode_(tc)) return { role: 'teacher' };
  const name = String(req.user || '').trim();
  if (!name || !req.ucode) return { role: 'none' };
  const cache = CacheService.getScriptCache(), key = 'fail_' + name;
  if (Number(cache.get(key) || 0) >= 10) return { role: 'none', blocked: true };
  const f = findStudent_(name);
  const inp = normCode_(req.ucode);
  if (f && normCode_(f.rec[1]) !== '' && (normCode_(f.rec[1]) === inp || (typeof f.rec[1] === 'number' && String(f.rec[1]) === inp.replace(/^0+/, '')))) {
    return { role: 'student', name: String(f.rec[0]).trim(), row: f.row };
  }
  cache.put(key, String(Number(cache.get(key) || 0) + 1), 600);   // 10분 안에 10번 틀리면 잠시 막기
  return { role: 'none' };
}

function findStudent_(name) {
  const sh = getSheet_(SHEET_STU, HEADERS_STU);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const d = sh.getRange(2, 1, last - 1, 7).getValues();
  for (let i = 0; i < d.length; i++) if (String(d[i][0]).trim() === name) return { row: i + 2, rec: d[i] };
  return null;
}

function profileOf_(rec) {
  return { name: String(rec[0]).trim(), univ: String(rec[2] || ''), start: dstr_(rec[3]) };
}

function login_(req, au) {
  if (au.role === 'teacher') return { ok: true, role: 'teacher', locked: !!PropertiesService.getScriptProperties().getProperty('TEACHER_CODE') };
  if (au.role !== 'student') return { ok: false, auth: true, error: '이름이나 코드가 맞지 않아요' };
  const sh = getSheet_(SHEET_STU, HEADERS_STU);
  sh.getRange(au.row, 6, 1, 2).setValues([[new Date(), String(req.ver || '')]]);
  const f = findStudent_(au.name);
  if (f && !f.rec[4]) sh.getRange(au.row, 5).setValue(new Date());
  return { ok: true, role: 'student', profile: profileOf_(f.rec), state: loadState_(au.name) };
}

// 학생 본인 정보 저장: 목표대학, 시작일
function profile_(req) {
  if (CUR_AU_.role !== 'student') return { ok: false, error: '러너만 쓸 수 있어요' };
  const sh = getSheet_(SHEET_STU, HEADERS_STU);
  const row = CUR_AU_.row;
  if (req.univ !== undefined) sh.getRange(row, 3).setValue(String(req.univ || '').slice(0, 40));
  if (req.start !== undefined) {
    const c = sh.getRange(row, 4); c.setNumberFormat('@'); c.setValue(String(req.start || '').slice(0, 10));
  }
  return { ok: true, profile: profileOf_(sh.getRange(row, 1, 1, 7).getValues()[0]) };
}

// 토마토 표시·회독 날짜 등 학생별 학습 상태(JSON 문자열)를 '진행' 시트에 조각으로 저장
function loadState_(name) {
  const sh = getSheet_(SHEET_STATE, HEADERS_STATE), last = sh.getLastRow();
  if (last < 2) return '';
  const parts = [];
  sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) { if (String(r[0]) === name) parts.push([Number(r[1]) || 0, String(r[2])]); });
  parts.sort(function (a, b) { return a[0] - b[0]; });
  return parts.map(function (x) { return x[1]; }).join('');
}

function stateSave_(req) {
  if (CUR_AU_.role !== 'student') return { ok: false, error: '러너만 쓸 수 있어요' };
  const name = CUR_AU_.name;
  const txt = typeof req.state === 'string' ? req.state : JSON.stringify(req.state || {});
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(SHEET_STATE, HEADERS_STATE);
    sh.getRange('C:C').setNumberFormat('@');
    const last = sh.getLastRow();
    if (last > 1) {
      const names = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = names.length - 1; i >= 0; i--) if (String(names[i][0]) === name) sh.deleteRow(i + 2);
    }
    const rows = [];
    for (let i = 0, n = 0; i < txt.length || n === 0; i += 40000, n++) rows.push([name, n, txt.slice(i, i + 40000)]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// 선생님: 학생 현황
function students_() {
  const sh = getSheet_(SHEET_STU, HEADERS_STU), last = sh.getLastRow();
  if (last < 2) return { ok: true, students: [] };
  const wsh = getSheet_(SHEET_WORDS, HEADERS_WORDS), wl = wsh.getLastRow();
  const own = {};
  if (wl > 1) wsh.getRange(2, 18, wl - 1, 1).getValues().forEach(function (r) { const o = String(r[0] || ''); if (o) own[o] = (own[o] || 0) + 1; });
  const out = sh.getRange(2, 1, last - 1, 7).getValues().filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    const name = String(r[0]).trim();
    let marks = 0;
    try { const st = JSON.parse(loadState_(name) || '{}'); const m = st.vn_tm_marks ? (typeof st.vn_tm_marks === 'string' ? JSON.parse(st.vn_tm_marks) : st.vn_tm_marks) : {}; marks = Object.keys(m || {}).length; } catch (e) { marks = 0; }
    return { name: name, code: normCode_(r[1]), univ: String(r[2] || ''), start: dstr_(r[3]), joined: dstr_(r[4]), lastSeen: r[5] instanceof Date ? r[5].toISOString() : String(r[5] || ''), ver: String(r[6] || ''), marks: marks, own: own[name] || 0 };
  });
  return { ok: true, students: out };
}

// 선생님: 학생 추가 (이름 + 개인 코드)
function studentAdd_(req) {
  const name = String(req.name || '').trim();
  if (!name) return { ok: false, error: '이름을 입력하세요' };
  if (findStudent_(name)) return { ok: false, error: '이미 있는 이름이에요: ' + name };
  const sh = getSheet_(SHEET_STU, HEADERS_STU), r = sh.getLastRow() + 1;
  // 코드는 자동으로 정한다: 다른 러너와 겹치지 않는 4자리 숫자 (직접 넣고 싶으면 studentCode 로 보내도 됨)
  const used = {};
  if (r > 2) sh.getRange(2, 2, r - 2, 1).getValues().forEach(function (x) { used[normCode_(x[0])] = true; });
  let code = normCode_(req.studentCode);
  if (!code) { do { code = String(1000 + Math.floor(Math.random() * 9000)); } while (used[code]); }
  sh.getRange(r, 2).setNumberFormat('@');
  sh.getRange(r, 1, 1, 3).setValues([[name, code, String(req.univ || '')]]);
  return { ok: true, name: name, code: code };
}

// 학생의 암기 기록: 공유 단어는 '진도' 시트에 학생별로 저장하고, 학습기록에는 이름을 남긴다
function recordStudent_(req) {
  const name = CUR_AU_.name;
  const sh = getSheet_(SHEET_PROG, HEADERS_PROG);
  const log = getSheet_(SHEET_LOG, HEADERS_LOG);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const updates = [];
  try {
    const last = sh.getLastRow();
    const data = last > 1 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
    const rowOf = {};
    data.forEach(function (r, i) { if (String(r[0]) === name) rowOf[String(r[1])] = i; });
    const now = new Date(), today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const logRows = [], adds = [];
    (req.items || []).forEach(function (it) {
      const id = String(it.id);
      let i = rowOf[id], right = 0, wrong = 0, box = 0;
      if (i !== undefined) { right = Number(data[i][2]) || 0; wrong = Number(data[i][3]) || 0; box = Number(data[i][5]) || 0; }
      if (it.correct) { right++; box = Math.min(box + 1, 5); } else { wrong++; box = 0; }
      const due = addDays_(today, DAYS[box]);
      const rec = [name, id, right, wrong, now, box, due];
      if (i !== undefined) { data[i] = rec; sh.getRange(i + 2, 1, 1, 7).setValues([rec]); }
      else { rowOf[id] = data.length; data.push(rec); adds.push(rec); }
      updates.push({ id: id, box: box, due: due });
      logRows.push([now, id, it.word, it.type, it.correct ? '정답' : '오답', name]);
    });
    if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, 7).setValues(adds);
    if (logRows.length) log.getRange(log.getLastRow() + 1, 1, logRows.length, 6).setValues(logRows);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, updates: updates };
}

function record_(req) {
  if (CUR_AU_.role === 'student') return recordStudent_(req);
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
      logRows.push([now, it.id, it.word, it.type, it.correct ? '정답' : '오답', '']);
    });
    if (logRows.length) log.getRange(log.getLastRow() + 1, 1, logRows.length, 6).setValues(logRows);
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
  let old = {};
  if (req.id) {
    const sh0 = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const d0 = sh0.getDataRange().getValues();
    for (let i = 1; i < d0.length; i++) if (String(d0[i][0]) === String(req.id)) { try { old = JSON.parse(d0[i][14] || '{}') || {}; } catch (e) { old = {}; } break; }
  }
  const extra = Object.assign({}, old, {
    freq: Math.max(1, Math.min(5, Math.round(Number(x.freq)) || 3)),
    examples: (x.examples || []).slice(0, 3).map(function (e) { return { en: String(e.en || ''), ko: String(e.ko || '') }; }),
    forms: String(x.forms || ''),
    derivs: (x.derivs || []).slice(0, 4).map(String)
  });
  if (req.id) {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(req.id)) { sh.getRange(i + 1, 15).setValue(JSON.stringify(extra)); break; }
    }
  }
  return { ok: true, extra: extra };
}

/* ---------- 연어(Collocation) 자동 생성 ----------
   req.items = [{ id, word, meaning }]  (한 번에 최대 25개)
   결과를 '추가정보(JSON)' 열에 colloc / colloc_ko 로 합쳐서 저장한다 (기존 상세 정보는 그대로 유지) */
function colloc_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';
  const items = (req.items || []).slice(0, 25);
  if (!items.length) return { ok: true, items: [] };
  const list = items.map(function (it, n) { return n + '. ' + String(it.word || '') + (it.meaning ? ' (' + String(it.meaning).slice(0, 40) + ')' : ''); }).join('\n');
  const prompt =
    '아래 영어 단어(또는 숙어) 각각에 대해, 가장 대표적이고 자주 쓰이는 연어(collocation) 하나와 그 한국어 뜻을 만드세요.\n' +
    '- 연어는 2~4단어의 짧은 구입니다. 예: authority → legal authority / 법적 권위, shift → shift focus / 초점을 전환하다, substitute → substitute A for B / B 대신 A를 쓰다, appreciate → appreciate art / 예술을 감상하다\n' +
    '- 해당 단어(활용형 가능)가 반드시 들어가야 합니다. 동사는 목적어와, 형용사는 명사와, 명사는 수식어나 동사와 함께 씁니다.\n' +
    '- 한국어 뜻은 짧게(10자 안팎).\n\n' + list + '\n\n' +
    '반드시 한 줄에 하나씩, 아래 형식의 텍스트로만 답하세요(JSON·따옴표·설명 금지).\n' +
    '번호|영어 연어|한국어 뜻\n예) 0|legal authority|법적 권위';
  const r = callClaudeText_(key, { model: model, max_tokens: 2000, messages: [{ role: 'user', content: prompt }] });
  if (!r.ok) return r;
  const got = {};
  String(r.text || '').split(/\r?\n/).forEach(function (ln) {
    const p = ln.replace(/^[\s\-*]+/, '').split('|');
    if (p.length < 3) return;
    const n = parseInt(p[0], 10);
    if (isNaN(n)) return;
    got[n] = { c: p[1].replace(/^["'\s]+|["'\s]+$/g, ''), ck: p.slice(2).join('|').replace(/^["'\s]+|["'\s]+$/g, '') };
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < data.length; i++) if (data[i][0]) rowOf[String(data[i][0])] = i;
    const out = [];
    items.forEach(function (it, n) {
      const g = got[n]; if (!g || !g.c) return;
      const i = rowOf[String(it.id)]; if (i === undefined) return;
      let ex = {}; try { ex = JSON.parse(data[i][14] || '{}') || {}; } catch (e) { ex = {}; }
      ex.colloc = g.c; ex.colloc_ko = g.ck;
      sh.getRange(i + 1, 15).setValue(JSON.stringify(ex));
      out.push({ id: String(it.id), colloc: g.c, colloc_ko: g.ck });
    });
    return { ok: true, items: out };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 영영 버전: 단어 뜻·연어 뜻을 영어로 (한 번에 최대 25개) ----------
   req.items = [{ id, word, meaning, colloc, defs?:["verb: …", …], onlyColloc?:true }]
   - defs 가 있으면(앱이 영영사전에서 가져온 풀이) 그중 한국어 뜻에 맞는 것을 골라 줄여서 쓴다 → meaning_en_src = 'dict'
   - 없으면 AI 가 직접 쓴다 → 'ai'.   연어 뜻은 사전에 없어서 항상 AI.
   결과는 extra.meaning_en / colloc_en / meaning_en_src 에 저장 */
function defen_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';
  const items = (req.items || []).slice(0, 25);
  if (!items.length) return { ok: true, items: [] };
  const list = items.map(function (it, n) {
    let t = n + '. ' + String(it.word || '') + (it.meaning ? ' (' + String(it.meaning).slice(0, 40) + ')' : '') + (it.colloc ? ' [연어: ' + String(it.colloc).slice(0, 60) + ']' : '');
    if (it.onlyColloc) t += ' [뜻은 이미 있음: 연어 뜻만]';
    else if (it.defs && it.defs.length) t += '\n   [사전] ' + it.defs.slice(0, 6).map(function (d, k) { return (k + 1) + ') ' + String(d).slice(0, 220); }).join(' ');
    return t;
  }).join('\n');
  const prompt =
    '아래 영어 단어(또는 숙어)마다 영영 풀이를 만드세요. 한국어 뜻은 어떤 의미인지 정하는 참고용입니다.\n' +
    '- 영어 뜻: [사전] 목록이 있으면 그 중에서 한국어 뜻과 가장 잘 맞는 항목 하나를 골라, 사전 표현을 최대한 그대로 살려 10단어 안팎으로 줄이세요(없는 내용을 지어내지 마세요). 이 경우 4번째 칸에 D.\n' +
    '  [사전] 이 없거나 맞는 항목이 없으면 학습자 사전처럼 쉬운 영어로 8단어 이내로 직접 쓰고(동사는 to 로 시작, 명사는 a/an/the 로 시작) 4번째 칸에 A.\n' +
    '- 연어 뜻: [연어: …] 가 있으면 그 구의 뜻을 짧은 영어로(8단어 이내). 연어가 없으면 비워 두세요.\n' +
    '- [뜻은 이미 있음] 표시가 있으면 영어 뜻 칸은 비우고 연어 뜻만 쓰세요(4번째 칸 A).\n\n' + list + '\n\n' +
    '반드시 한 줄에 하나씩, 아래 형식의 텍스트로만 답하세요(JSON·따옴표·설명 금지).\n' +
    '번호|영어 뜻|연어의 영어 뜻|D 또는 A\n예) 0|to leave completely|to stop having something|D';
  const r = callClaudeText_(key, { model: model, max_tokens: 2500, messages: [{ role: 'user', content: prompt }] });
  if (!r.ok) return r;
  const got = {};
  String(r.text || '').split(/\r?\n/).forEach(function (ln) {
    const p = ln.replace(/^[\s\-*]+/, '').split('|');
    if (p.length < 2) return;
    const n = parseInt(p[0], 10);
    if (isNaN(n)) return;
    const clean = function (x) { return String(x || '').replace(/^["'\s]+|["'\s]+$/g, ''); };
    const flag = p.length >= 4 ? clean(p[p.length - 1]).toUpperCase() : '';
    got[n] = { m: clean(p[1]), c: p.length >= 3 ? clean(p[2]) : '', f: flag === 'D' ? 'D' : 'A' };
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < data.length; i++) if (data[i][0]) rowOf[String(data[i][0])] = i;
    const out = [];
    items.forEach(function (it, n) {
      const g = got[n]; if (!g) return;
      if (!it.onlyColloc && !g.m) return;
      const i = rowOf[String(it.id)]; if (i === undefined) return;
      let ex = {}; try { ex = JSON.parse(data[i][14] || '{}') || {}; } catch (e) { ex = {}; }
      const o = { id: String(it.id) };
      if (!it.onlyColloc) {
        ex.meaning_en = g.m; ex.meaning_en_src = (g.f === 'D' && it.defs && it.defs.length) ? 'dict' : 'ai';
        o.meaning_en = ex.meaning_en; o.meaning_en_src = ex.meaning_en_src;
      }
      if (g.c) { ex.colloc_en = g.c; o.colloc_en = g.c; }
      sh.getRange(i + 1, 15).setValue(JSON.stringify(ex));
      out.push(o);
    });
    return { ok: true, items: out };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 백업 복원 ----------
   req.items = [{ word, pos, meaning, example, example_ko, added, right, wrong, last, box, due, tip, level, extra, folder }]  (한 번에 최대 150개)
   - 시트에 이미 있는 단어는 건드리지 않고, 없는 단어만 학습 기록까지 그대로 채워 넣는다 */
function restore_(req) {
  const items = (req.items || []).slice(0, 150);
  if (!items.length) return { ok: true, added: 0, skipped: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const have = {};
    for (let i = 1; i < data.length; i++) if (data[i][0]) have[String(data[i][1]).toLowerCase()] = true;
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const rows = []; let skipped = 0;
    items.forEach(function (it) {
      const w = String(it.word || '').trim().toLowerCase();
      if (!w || have[w]) { skipped++; return; }
      have[w] = true;
      const d = it.added ? new Date(it.added) : new Date();
      let ex = '';
      if (it.extra && typeof it.extra === 'object') ex = JSON.stringify(it.extra); else if (typeof it.extra === 'string') ex = it.extra;
      rows.push([Utilities.getUuid().slice(0, 8), w, it.pos || '', it.meaning || '', it.example || '', it.example_ko || '', isNaN(d) ? new Date() : d,
        Number(it.right) || 0, Number(it.wrong) || 0, it.last || '', Number(it.box) || 0, it.due || today, it.tip || '', Number(it.level) || '', ex, it.folder || '', (it.src === 'hand' || it.src === 'import') ? it.src : '']);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, added: rows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 단어 묶음 가져오기 (폴더/서브폴더 지정) ----------
   req.items = [{ word, pos, meaning, example, example_ko, folder }]  (한 번에 최대 200개)
   - 이미 있는 단어: 새로 만들지 않고, 폴더가 비어 있으면 그 폴더로만 넣어 준다 */
/* ---------- 폴더 이름 바꾸기 ----------
   req = { from:'폴더' | '폴더/서브', to:'새폴더' | '새폴더/새서브', kind:'root' | 'sub' }
   root: from 과 같거나 from + '/' 로 시작하는 모든 단어의 폴더를 바꾼다.  sub: from 과 정확히 같은 단어만. */
function renameFolder_(req) {
  const from = String(req.from || '').trim(), to = String(req.to || '').trim(), kind = req.kind === 'sub' ? 'sub' : 'root';
  if (!from || !to) return { ok: false, error: '이름이 비어 있어요' };
  if (from === to) return { ok: true, changed: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const n = sh.getLastRow() - 1;
    if (n < 1) return { ok: true, changed: 0 };
    const rng = sh.getRange(2, 16, n, 1), col = rng.getValues();
    let changed = 0;
    for (let i = 0; i < n; i++) {
      const f = String(col[i][0] || '');
      if (!f) continue;
      if (f === from || (kind === 'root' && f.indexOf(from + '/') === 0)) { col[i][0] = to + f.slice(from.length); changed++; }
    }
    if (changed) rng.setValues(col);
    return { ok: true, changed: changed };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 단어 가져오기 ----------
   req = { items:[…], mode:'skip'|'move' }
   - 새 단어: 추가(출처 import)
   - 이미 있는 단어: 폴더가 비어 있으면 그 폴더로. 폴더가 다르면 mode='move' 일 때만 옮기고, 아니면 건너뜀
   돌려주는 값: added, moved, skipped, ids(새로 만든 id), movedList([{id,from}]), skippedList([{word,folder}]) */
function import_(req) {
  const items = (req.items || []).slice(0, 200);
  const move = req.mode === 'move';
  const dup = req.mode === 'dup';   // 같은 단어가 다른 폴더에 있어도 이 폴더에 따로 등록 (같은 폴더 안에서만 중복 제외)
  if (!items.length) return { ok: true, added: 0, moved: 0, skipped: 0, ids: [], movedList: [], skippedList: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < data.length; i++) if (data[i][0]) {
      const k = String(data[i][1]).toLowerCase();
      if (dup) rowOf[k + '\u0001' + String(data[i][15] || '')] = i; else rowOf[k] = i;
    }
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const now = new Date();
    const rows = [], ids = [], movedList = [], skippedList = [];
    let moved = 0, skipped = 0, filled = 0;
    const xcol = {};   // 연어 채우기용: 행 번호 -> 추가정보(JSON)
    items.forEach(function (it) {
      const w = String(it.word || '').trim().toLowerCase();
      if (!w) return;
      const folder = String(it.folder || '').trim();
      const rk = dup ? w + '\u0001' + folder : w;
      if (rk in rowOf) {
        const i = rowOf[rk];
        if (i < 0) { skipped++; return; }                 // 같은 묶음 안의 중복
        const cur = String(data[i][15] || '');
        if (it.colloc && (dup || cur === folder)) {              // 이미 있는 단어에 연어가 비어 있으면 채운다
          let ex = {}; try { ex = data[i][14] ? JSON.parse(data[i][14]) : {}; } catch (e) { ex = {}; }
          let ch = false;
          if (!ex.colloc) { ex.colloc = String(it.colloc); ex.colloc_ko = String(it.colloc_ko || ''); ch = true; }
          if (it.colloc_en && !ex.colloc_en) { ex.colloc_en = String(it.colloc_en); ch = true; }
          if (ch) { xcol[i] = ex; filled++; }
        }
        if (folder && (!cur || (move && cur !== folder))) { sh.getRange(i + 1, 16).setValue(folder); moved++; movedList.push({ id: String(data[i][0]), from: cur }); }
        else { skipped++; if (skippedList.length < 100) skippedList.push({ word: w, folder: cur }); }
        return;
      }
      rowOf[rk] = -1;
      const id = Utilities.getUuid().slice(0, 8);
      ids.push(id);
      rows.push([id, w, it.pos || '', it.meaning || '', it.example || '', it.example_ko || '', now, 0, 0, '', 0, today, it.tip || '', Number(it.level) || '', it.colloc ? JSON.stringify({ colloc: String(it.colloc), colloc_ko: String(it.colloc_ko || ''), colloc_en: String(it.colloc_en || '') }) : '', folder, 'import']);
    });
    Object.keys(xcol).forEach(function (i) { sh.getRange(Number(i) + 1, 15).setValue(JSON.stringify(xcol[i])); });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, added: rows.length, moved: moved, skipped: skipped, filled: filled, ids: ids, movedList: movedList, skippedList: skippedList };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 가져오기 되돌리기: 새로 만든 단어는 삭제(출처 import 만), 옮긴 단어는 원래 폴더로 ---------- */
function undoImport_(req) {
  const ids = {}; (req.ids || []).forEach(function (x) { ids[String(x)] = true; });
  const moves = {}; (req.moves || []).forEach(function (m) { moves[String(m.id)] = String(m.from || ''); });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
    const data = sh.getDataRange().getValues();
    const del = []; let restored = 0;
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0]);
      if (!id) continue;
      if (ids[id] && String(data[i][16] || '') === 'import') del.push(i + 1);
      else if (id in moves) { sh.getRange(i + 1, 16).setValue(moves[id]); restored++; }
    }
    // 아래쪽부터, 이어진 줄은 한 번에 지운다
    let k = del.length - 1;
    while (k >= 0) {
      let start = del[k], n = 1;
      while (k - 1 >= 0 && del[k - 1] === start - 1) { start = del[k - 1]; n++; k--; }
      sh.deleteRows(start, n); k--;
    }
    return { ok: true, removed: del.length, restored: restored };
  } finally {
    lock.releaseLock();
  }
}
