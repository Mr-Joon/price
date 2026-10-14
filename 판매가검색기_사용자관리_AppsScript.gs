/**
 * 판매가 검색기 - 사용자 관리 (구글 Apps Script)
 *
 * 시트 구성 (첫 번째 시트):  1행 제목
 *   A: 기기번호 | B: 이름 | C: 허용 | D: 변경 시간
 *   허용 칸 값:  O = 허용 / X = 차단 / 대기(또는 빈칸) = 승인 요청 / 관리자 = 관리자 기기
 *
 * 버전: 앱 v1.0.15용 (사용자 추가 'add' 기능 추가)
 *   ※ 코드를 바꾼 뒤에는 배포 → 배포 관리 → ✏ 수정 → 버전: 새 버전 → 배포
 *     (웹 앱 URL은 그대로 유지됩니다)
 *
 * 설치 (처음 한 번):
 *   1) 사용자 명단 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 저장
 *   2) 왼쪽 ⚙ 프로젝트 설정 → 스크립트 속성 → 속성 추가: ADMIN_PIN = 6자리 숫자
 *   3) 오른쪽 위 배포 → 새 배포 → 유형: 웹 앱
 *        실행 사용자: 나 / 액세스 권한이 있는 사용자: 모든 사용자 → 배포 → 권한 허용
 *   4) 나온 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 전달
 */
const SHEET_ID = '1_OMhj2bGozo0Q8UGMpFAVLvf5-F3WHrTRM1xGJOYdJA';
const MAX_ROWS = 300;          // 명단 최대 줄 수 (승인 요청 남용 방지)
const MAX_FAILS = 5;           // 관리자 비밀번호 연속 실패 허용 횟수
const LOCK_SECONDS = 600;      // 실패 초과 시 잠금 시간 (10분)

function doGet() {
  return out_({ ok: true, msg: '판매가 검색기 사용자 관리' });
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return out_(handle_(p));
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() { return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; }

// 기기번호 정리: 대소문자·공백·하이픈 무시 → "ABCD-1234"
function normId_(v) {
  const s = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return s.length === 8 ? s.slice(0, 4) + '-' + s.slice(4) : '';
}

// 이름 정리: 20자, 줄바꿈 제거, 수식으로 해석되지 않게 처리
function cleanName_(v) {
  let s = String(v || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, 20);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }

function rows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(function (r, i) {
    return {
      row: i + 2,
      id: normId_(r[0]),
      name: String(r[1] || ''),
      status: String(r[2] || '').trim(),
      time: r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(r[3] || '')
    };
  });
}

function find_(list, id) {
  for (let i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  return null;
}

function isAdmin_(s) { return String(s).trim() === '관리자'; }

function handle_(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = sheet_();
    if (sh.getLastRow() < 1) sh.appendRow(['기기번호', '이름', '허용', '변경 시간']);
    const me = normId_(p.deviceId);
    if (!me) return { ok: false, error: '기기번호가 올바르지 않습니다' };
    const list = rows_(sh);

    // ---------- 직원: 승인 요청 (비밀번호 필요 없음) ----------
    if (p.action === 'request') {
      const name = cleanName_(p.name);
      if (!name) return { ok: false, error: '이름을 입력하세요' };
      const r = find_(list, me);
      if (r) {
        if (r.status === '대기' || r.status === '') {
          sh.getRange(r.row, 2, 1, 3).setValues([[name, '대기', now_()]]);
          return { ok: true, status: '대기' };
        }
        return { ok: true, status: r.status };   // 이미 허용/차단된 기기는 바꾸지 않음
      }
      if (list.length >= MAX_ROWS) return { ok: false, error: '명단이 가득 찼습니다. 관리자에게 문의하세요' };
      sh.appendRow([me, name, '대기', now_()]);
      return { ok: true, status: '대기' };
    }

    // ---------- 관리자 전용 (관리자 기기 + 비밀번호) ----------
    const cache = CacheService.getScriptCache();
    const failKey = 'fail_' + me;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= MAX_FAILS) return { ok: false, error: '비밀번호를 여러 번 틀려 10분간 잠겼습니다' };

    const admin = find_(list, me);
    if (!admin || !isAdmin_(admin.status)) return { ok: false, error: '관리자 기기가 아닙니다' };

    const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
    if (!pin) return { ok: false, error: '관리자 비밀번호가 설정되지 않았습니다 (스크립트 속성 ADMIN_PIN)' };
    if (String(p.pin || '') !== String(pin).trim()) {
      cache.put(failKey, String(fails + 1), LOCK_SECONDS);
      return { ok: false, error: '비밀번호가 틀렸습니다 (' + (fails + 1) + '/' + MAX_FAILS + ')', badPin: true };
    }
    cache.remove(failKey);

    if (p.action === 'list') {
      return {
        ok: true,
        users: list.filter(function (x) { return x.id; })
                   .map(function (x) { return { id: x.id, name: x.name, status: x.status, time: x.time }; })
      };
    }

    const target = normId_(p.id);
    if (!target) return { ok: false, error: '기기번호가 올바르지 않습니다 (예: ABCD-1234)' };
    const t = find_(list, target);

    if (p.action === 'add') {          // 새 기기 추가 (허용). 이미 있으면 확인을 받은 뒤(overwrite)에만 변경
      if (t && isAdmin_(t.status)) return { ok: false, error: '관리자 기기는 시트에서 직접 변경하세요' };
      if (t && !p.overwrite) {
        return { ok: false, exists: true, error: '이미 등록된 기기입니다',
                 user: { id: t.id, name: t.name, status: t.status, time: t.time } };
      }
      const name = cleanName_(p.name) || (t ? t.name : '');   // 이름을 비우면 기존 이름 유지
      if (t) {
        sh.getRange(t.row, 2, 1, 3).setValues([[name, 'O', now_()]]);
      } else {
        if (list.length >= MAX_ROWS) return { ok: false, error: '명단이 가득 찼습니다' };
        sh.appendRow([target, name, 'O', now_()]);
      }
      return { ok: true };
    }

    if (p.action === 'set') {          // 허용(O) / 차단(X)
      const st = { O: 'O', X: 'X' }[String(p.status)];
      if (!st) return { ok: false, error: '상태 값이 올바르지 않습니다' };
      if (t && isAdmin_(t.status)) return { ok: false, error: '관리자 기기는 시트에서 직접 변경하세요' };
      const name = cleanName_(p.name) || (t ? t.name : '');   // 이름을 비우면 기존 이름 유지
      if (t) {
        sh.getRange(t.row, 2, 1, 3).setValues([[name, st, now_()]]);
      } else {
        if (list.length >= MAX_ROWS) return { ok: false, error: '명단이 가득 찼습니다' };
        sh.appendRow([target, name, st, now_()]);
      }
      return { ok: true };
    }

    if (p.action === 'delete') {
      if (!t) return { ok: true };
      if (isAdmin_(t.status)) return { ok: false, error: '관리자 기기는 시트에서 직접 삭제하세요' };
      sh.deleteRow(t.row);
      return { ok: true };
    }

    return { ok: false, error: '알 수 없는 요청' };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
