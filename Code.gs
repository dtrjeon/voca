/**
 * 손글씨 단어노트 - Google Apps Script 백엔드
 *
 * [설정]
 * 1) 구글시트를 새로 만들고, 확장 프로그램 > Apps Script 에 이 코드를 붙여넣기
 * 2) 프로젝트 설정 > 스크립트 속성에 추가
 *      ANTHROPIC_API_KEY = (Anthropic API 키)
 *    (선택) MODEL = claude-haiku-5-5   (기본값)
 * 3) 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) > URL 복사
 *    -> index.html 의 GAS_URL 에 붙여넣기
 */

const SHEET_WORDS = '단어';
const SHEET_LOG = '학습기록';
const HEADERS_WORDS = ['id', '단어', '품사', '뜻', '예문', '예문뜻', '등록일', '정답수', '오답수', '마지막학습'];
const HEADERS_LOG = ['시각', 'id', '단어', '유형', '결과'];

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 모든 요청은 text/plain POST 로 받는다 (CORS preflight 회피)
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'recognize': return json_(recognize_(req));
      case 'add':       return json_(addWord_(req));
      case 'list':      return json_(listWords_());
      case 'delete':    return json_(deleteWord_(req));
      case 'update':    return json_(updateWord_(req));
      case 'record':    return json_(record_(req));
      default:          return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, msg: 'wordnote api alive' });
}

/* ---------- Claude: 손글씨 인식 + 뜻/예문 ---------- */
function recognize_(req) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_API_KEY 가 스크립트 속성에 없습니다.' };
  const model = props.getProperty('MODEL') || 'claude-haiku-5-5';

  const prompt =
    '이미지는 영어 단어 하나를 손으로 쓴 것입니다. 글자를 정확히 읽어서 JSON 으로만 답하세요.\n' +
    '{"word":"읽은 영단어(소문자, 표제어 원형)","pos":"품사(명/동/형/부 등 한글)",' +
    '"meaning":"한국어 뜻(핵심 1~3개, 쉼표 구분)","example":"해당 단어를 쓴 자연스러운 영어 예문 1개",' +
    '"example_ko":"예문의 한국어 번역","confidence":0~1 숫자}\n' +
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
  sh.appendRow([id, w, req.pos || '', req.meaning || '', req.example || '', req.example_ko || '', new Date(), 0, 0, '']);
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
      last: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || '')
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
  return { ok: true };
}

/* ---------- 학습 기록 ---------- */
function record_(req) {
  const sh = getSheet_(SHEET_WORDS, HEADERS_WORDS);
  const log = getSheet_(SHEET_LOG, HEADERS_LOG);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const items = req.items || []; // [{id, word, type, correct}]
    const now = new Date();
    items.forEach(function (it) {
      const row = findRow_(sh, it.id);
      if (row > 0) {
        const col = it.correct ? 8 : 9;
        sh.getRange(row, col).setValue((Number(sh.getRange(row, col).getValue()) || 0) + 1);
        sh.getRange(row, 10).setValue(now);
      }
      log.appendRow([now, it.id, it.word, it.type, it.correct ? '정답' : '오답']);
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}
