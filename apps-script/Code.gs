/**
 * 우리 땅 임대 장부 — 구글 시트 연결 스크립트
 *
 * 이 스크립트는 장부 데이터가 들어 있는 구글 시트에 붙여서 씁니다.
 * (시트 메뉴: 확장 프로그램 → Apps Script)
 *
 * 처음 한 번 할 일
 * 1. 왼쪽 톱니바퀴(프로젝트 설정) → 맨 아래 '스크립트 속성' → '스크립트 속성 추가'
 *    속성: FAMILY_PASSWORD   값: 가족끼리 쓸 비밀번호 (예: 숫자 6자리 이상)
 * 2. 오른쪽 위 '배포' → '새 배포' → 유형 '웹 앱'
 *    다음 사용자 인증 정보로 실행: 나
 *    액세스 권한이 있는 사용자: 모든 사용자
 *    → '배포' → 권한 승인 → 나오는 '웹 앱 URL'을 복사
 *
 * 비밀번호가 맞아야만 시트를 읽고 쓸 수 있습니다.
 * 비밀번호는 이 스크립트 속성에만 저장되고, GitHub 코드에는 들어가지 않습니다.
 */

const LEASE_COLS = ['id', 'region', 'parcels', 'kind', 'start', 'end', 'tenant', 'phone', 'rent', 'area',
  'facility', 'noticeDone', 'reportedAt', 'customDeadline', 'customDeadlineLabel', 'note',
  'addresses', 'parcelCodes', 'issues', 'updatedAt'];
const CONTACT_COLS = ['id', 'name', 'phone', 'memo', 'order'];
const JSON_COLS = { addresses: [], parcelCodes: {}, issues: [] };
const BOOL_COLS = ['facility', 'noticeDone'];
const KINDS = ['wituk', 'private', 'forest', 'self', 'idle'];

function doGet() {
  return out({ ok: true, message: '우리 땅 임대 장부 연결 스크립트가 동작 중이에요.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return out({ ok: false, error: 'bad_request' });
  }
  const pw = PropertiesService.getScriptProperties().getProperty('FAMILY_PASSWORD');
  if (!pw) return out({ ok: false, error: 'no_password_set' });
  if (String(req.key || '') !== pw) {
    Utilities.sleep(800); // 비밀번호 무작위 대입을 느리게
    return out({ ok: false, error: 'wrong_password' });
  }

  const lock = LockService.getScriptLock();
  try {
    if (req.action === 'list') {
      return out({ ok: true, leases: readAll('leases', LEASE_COLS), contacts: readAll('contacts', CONTACT_COLS) });
    }
    lock.waitLock(15000);
    if (req.action === 'saveLease') {
      const d = req.data || {};
      if (!String(d.region || '').trim()) return out({ ok: false, error: 'region_required' });
      if (KINDS.indexOf(d.kind) < 0) d.kind = 'idle';
      d.updatedAt = new Date().toISOString();
      return out({ ok: true, item: upsert('leases', LEASE_COLS, d) });
    }
    if (req.action === 'deleteLease') {
      remove('leases', req.id);
      return out({ ok: true });
    }
    if (req.action === 'saveContact') {
      const d = req.data || {};
      if (!String(d.name || '').trim() || !String(d.phone || '').trim()) return out({ ok: false, error: 'contact_required' });
      return out({ ok: true, item: upsert('contacts', CONTACT_COLS, d) });
    }
    if (req.action === 'deleteContact') {
      remove('contacts', req.id);
      return out({ ok: true });
    }
    return out({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out({ ok: false, error: 'server_error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function sheetFor(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setNumberFormat('@').setValues([cols]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(name, cols) {
  const sh = sheetFor(name, cols);
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const header = values[0];
  return values.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    header.forEach(function (h, i) {
      if (!h) return;
      let v = r[i];
      if (Object.prototype.hasOwnProperty.call(JSON_COLS, h)) {
        try { v = v ? JSON.parse(v) : JSON_COLS[h]; } catch (err) { v = JSON_COLS[h]; }
      } else if (BOOL_COLS.indexOf(h) >= 0) {
        v = String(v).toUpperCase() === 'TRUE';
      }
      o[h] = v;
    });
    return o;
  });
}

function toRow(cols, d) {
  return cols.map(function (c) {
    const v = d[c];
    if (Object.prototype.hasOwnProperty.call(JSON_COLS, c)) return JSON.stringify(v == null ? JSON_COLS[c] : v);
    if (BOOL_COLS.indexOf(c) >= 0) return v ? 'TRUE' : 'FALSE';
    return v == null ? '' : String(v);
  });
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function upsert(name, cols, d) {
  const sh = sheetFor(name, cols);
  if (!d.id || !/^[A-Za-z0-9_-]{1,80}$/.test(String(d.id))) d.id = Utilities.getUuid();
  const row = toRow(cols, d);
  let r = findRow(sh, d.id);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
  return d;
}

function remove(name, id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) return;
  const r = findRow(sh, String(id || ''));
  if (r > 0) sh.deleteRow(r);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
