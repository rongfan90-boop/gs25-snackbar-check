/**
 * GS25 스낵바 입고검수 - 회차별 기록 저장소
 * 구글시트 > 확장 프로그램 > Apps Script 에 전체 붙여넣기 → 배포 > 새 배포 > 웹 앱
 *  - 다음 사용자 인증정보로 실행: 나
 *  - 액세스 권한: 모든 사용자
 */
const SHEET_NAME = '입고기록';
const HEAD = ['id','저장시각','고객사','납품일','회차','품목수','납품수량','입고수량','입고율(%)','미입고·부족 품목수','카톡문구','상세(JSON)'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f1fc');
    s.setFrozenRows(1);
    s.getRange('A:L').setNumberFormat('@');
    s.setColumnWidth(11, 260);
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function rows_() {
  const s = sheet_();
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, HEAD.length).getDisplayValues();
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || 'list';
    if (action === 'list') {
      const data = rows_().filter(r => r[0] && (!p.store || r[2] === p.store)).map(r => ({
        id: r[0], savedAt: r[1], store: r[2], date: r[3], round: r[4],
        itemCount: Number(r[5]), plan: Number(r[6]), recv: Number(r[7]), pct: Number(r[8]), left: Number(r[9])
      }));
      return out_({ ok: true, data: data });
    }
    if (action === 'get') {
      const r = rows_().find(r => r[0] === p.id);
      return out_(r ? { ok: true, data: JSON.parse(r[11]) } : { ok: false, error: '기록 없음' });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const s = sheet_();
    const ids = rows_().map(r => r[0]);
    if (d.action === 'save') {
      const r = d.record;
      const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
      r.savedAt = now;
      const row = [r.id, now, r.store, r.date, r.round, r.itemCount, r.plan, r.recv, r.pct, r.left, r.msg || '', JSON.stringify(r)]
        .map(v => String(v == null ? '' : v));
      const i = ids.indexOf(r.id);
      const at = i >= 0 ? i + 2 : s.getLastRow() + 1;
      s.getRange(at, 1, 1, HEAD.length).setNumberFormat('@').setValues([row]);
      return out_({ ok: true, data: { id: r.id, savedAt: now } });
    }
    if (d.action === 'delete') {
      const i = ids.indexOf(d.id);
      if (i >= 0) s.deleteRow(i + 2);
      return out_({ ok: true, data: null });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
