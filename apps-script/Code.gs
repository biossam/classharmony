/**
 * 클래스 하모니 · 학생 욕구 설문 저장소 (Google Apps Script, 무료)
 * 학생이 자기 기기에서 제출한 설문 결과를 선생님 구글 시트에 모읍니다.
 * 선생님 화면은 '선생님 코드'를 알아야만 결과를 볼 수 있어요.
 *
 * 설치 방법은 apps-script/설치방법.md 를 보세요.
 * 선생님 코드(TEACHER_KEY)와 학생 입장 코드(JOIN_CODE)는 이 파일에 적지 말고
 * '프로젝트 설정 → 스크립트 속성'에 넣습니다. (저장소가 공개라서)
 */

const SHEET_NAME = '설문';
const HEAD = ['제출 시각', '이름', '생존', '사랑/소속', '힘/성취', '자유', '즐거움', '대표 욕구'];
const NEEDS = ['survival', 'love', 'power', 'freedom', 'fun'];
const NEED_KOR = { survival: '생존', love: '사랑/소속', power: '힘/성취', freedom: '자유', fun: '즐거움' };
const MAX_POSTS_PER_10MIN = 200;

function props_() { return PropertiesService.getScriptProperties(); }
function teacherKey_() { return props_().getProperty('TEACHER_KEY') || ''; }
function joinCode_() { return props_().getProperty('JOIN_CODE') || '1234'; }

function book_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = props_().getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만듦 */ } }
  const ss = SpreadsheetApp.create('클래스 하모니 설문');
  props_().setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 시트에 글을 넣을 때 = + - @ 로 시작하면 수식으로 실행되지 않게 앞에 ' 를 붙임
function safe_(v) { const t = String(v == null ? '' : v); return /^[=+\-@\t\r]/.test(t) ? "'" + t : t; }
function name_(v) { return String(v || '').replace(/\s+/g, ' ').trim().slice(0, 20); }

function tooMany_() {
  const c = CacheService.getScriptCache();
  const k = 'posts_' + Math.floor(Date.now() / 600000);
  const n = Number(c.get(k) || 0) + 1;
  c.put(k, String(n), 700);
  return n > MAX_POSTS_PER_10MIN;
}

function isTeacher_(d) {
  const key = teacherKey_();
  return key && String(d.key || '') === key;
}

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, 8).getValues();
}

// 학생 설문 제출: 같은 이름이면 새 결과로 바꿈
function submit_(d) {
  if (String(d.code || '').trim().toLowerCase() !== joinCode_().toLowerCase()) return json_({ ok: false, error: '입장 코드가 맞지 않아요. 선생님께 다시 물어보세요.' });
  const name = name_(d.name);
  if (!name) return json_({ ok: false, error: '이름을 입력해 주세요.' });
  const s = d.scores || {};
  const vals = NEEDS.map((k) => Number(s[k]));
  if (vals.some((v) => !(v >= 3 && v <= 25) || Math.round(v) !== v)) return json_({ ok: false, error: '설문 점수가 올바르지 않아요.' });
  let dom = 0;
  vals.forEach((v, i) => { if (v > vals[dom]) dom = i; });
  const sh = sheet_();
  const rows = rows_();
  const row = [new Date(), safe_(name)].concat(vals, [NEED_KOR[NEEDS[dom]]]);
  const at = rows.findIndex((r) => String(r[1]) === name);
  if (at >= 0) sh.getRange(at + 2, 1, 1, 8).setValues([row]);
  else sh.appendRow(row);
  return json_({ ok: true });
}

function list_() {
  const students = rows_().filter((r) => r[1] !== '').map((r) => {
    const scoresSum = {};
    NEEDS.forEach((k, i) => { scoresSum[k] = Number(r[2 + i]) || 0; });
    return { name: String(r[1]), scoresSum: scoresSum, time: r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'M/d HH:mm') : '' };
  });
  return json_({ ok: true, students: students, joinCode: joinCode_() });
}

function remove_(d) {
  const name = name_(d.name);
  const rows = rows_();
  for (let i = rows.length - 1; i >= 0; i--) if (String(rows[i][1]) === name) sheet_().deleteRow(i + 2);
  return list_();
}

// 초기화: 지금 결과를 '보관_날짜' 시트로 옮겨 두고 새로 시작 (지우지 않음)
function reset_() {
  const sh = sheet_();
  if (sh.getLastRow() >= 2) {
    sh.setName('보관_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm'));
    sheet_();
  }
  return list_();
}

function setCode_(d) {
  const code = String(d.code || '').trim().slice(0, 20);
  if (!code) return json_({ ok: false, error: '입장 코드를 1글자 이상 입력해 주세요.' });
  props_().setProperty('JOIN_CODE', code);
  return list_();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'check') {
      if (tooMany_()) return json_({ ok: false, error: '잠시 뒤 다시 해 주세요.' });
      return String(d.code || '').trim().toLowerCase() === joinCode_().toLowerCase() ? json_({ ok: true }) : json_({ ok: false, error: '입장 코드가 맞지 않아요. 선생님께 다시 물어보세요.' });
    }
    if (d.action === 'submit') {
      if (tooMany_()) return json_({ ok: false, error: '지금 제출하는 사람이 많아요. 잠시 뒤 다시 해 주세요.' });
      return submit_(d);
    }
    if (!teacherKey_()) return json_({ ok: false, needSetup: true, error: '스크립트 속성에 TEACHER_KEY(선생님 코드)를 먼저 넣어 주세요.' });
    if (!isTeacher_(d)) return json_({ ok: false, needKey: true, error: '선생님 코드가 맞지 않아요.' });
    if (d.action === 'list') return list_();
    if (d.action === 'delete') return remove_(d);
    if (d.action === 'reset') return reset_();
    if (d.action === 'setCode') return setCode_(d);
    return json_({ ok: false, error: '알 수 없는 요청이에요.' });
  } catch (err) {
    return json_({ ok: false, error: '처리하지 못했어요. 잠시 뒤 다시 해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}

// 주소를 브라우저로 열었을 때 연결 확인용 (학생 결과는 보여 주지 않음)
function doGet() {
  return json_({ ok: true, app: 'classharmony' });
}
