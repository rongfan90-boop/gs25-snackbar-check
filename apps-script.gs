/**
 * GS25 스낵바 입고검수 - 저장소 v2 (점포 로그인)
 * 1) 아래 ADMIN_PIN 을 원하는 관리자 비밀번호(숫자)로 바꾸기
 * 2) 저장 → 배포 > 배포 관리 > ✏️ 편집 > 버전: "새 버전" > 배포   (주소는 그대로 유지됨)
 */
const ADMIN_PIN = '0000';          // ★ 관리자 비밀번호 - 꼭 바꾸세요
const TZ = 'Asia/Seoul';
const SHOP_SHEET = '점포';
const REC_SHEET = '입고기록v2';
const SHOP_HEAD = ['점포명', '점포키', '비밀번호', '등록일시', '최근사용'];
const REC_HEAD = ['id', '저장시각', '점포', '고객사', '납품일', '주차/회차', '품목수', '납품수량', '입고수량', '입고율(%)', '미입고·부족 품목수', '카톡문구', '상세(JSON)'];

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function key_(name) { return String(name || '').replace(/\s+/g, '').toUpperCase(); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f1fc');
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), head.length).setNumberFormat('@');
  }
  return s;
}
function rows_(s, n) { const last = s.getLastRow(); return last < 2 ? [] : s.getRange(2, 1, last - 1, n).getDisplayValues(); }
function writeRow_(s, rowIdx, vals) { s.getRange(rowIdx, 1, 1, vals.length).setNumberFormat('@').setValues([vals.map(v => String(v == null ? '' : v))]); }

/* 로그인 확인 (5회 틀리면 10분 잠금) */
function auth_(a) {
  a = a || {};
  const cache = CacheService.getScriptCache();
  const k = a.admin ? 'ADMIN' : key_(a.store);
  const fk = 'fail_' + k;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('비밀번호를 여러 번 틀려서 10분간 잠겼어요');
  const bad = () => { cache.put(fk, String(fails + 1), 600); throw new Error('점포명 또는 비밀번호가 맞지 않아요'); };
  if (a.admin) { if (String(a.pin) !== String(ADMIN_PIN)) bad(); cache.remove(fk); return { admin: true }; }
  const s = sheet_(SHOP_SHEET, SHOP_HEAD);
  const rows = rows_(s, SHOP_HEAD.length);
  const i = rows.findIndex(r => r[1] === k);
  if (i < 0 || rows[i][2] !== String(a.pin)) bad();
  cache.remove(fk);
  s.getRange(i + 2, 5).setValue(now_());
  return { admin: false, store: rows[i][0], key: k, row: i + 2 };
}

function doGet() { return out_({ ok: true, data: 'GS25 snackbar api v2' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const act = d.action;

    if (act === 'ping') return out_({ ok: true, data: 'v2' });

    if (act === 'register') {
      const name = String(d.store || '').trim(), pin = String(d.pin || '');
      if (!name) throw new Error('점포명을 입력해주세요');
      if (!/^\d{4,6}$/.test(pin)) throw new Error('비밀번호는 숫자 4~6자리로 정해주세요');
      const s = sheet_(SHOP_SHEET, SHOP_HEAD);
      if (rows_(s, SHOP_HEAD.length).some(r => r[1] === key_(name))) throw new Error('이미 등록된 점포예요. 로그인해주세요');
      writeRow_(s, s.getLastRow() + 1, [name, key_(name), pin, now_(), now_()]);
      return out_({ ok: true, data: { store: name } });
    }

    const who = auth_(d.auth);
    const rs = sheet_(REC_SHEET, REC_HEAD);

    if (act === 'login') return out_({ ok: true, data: who.admin ? { admin: true } : { store: who.store } });

    if (act === 'changePin') {
      if (who.admin) throw new Error('관리자 비밀번호는 스크립트에서 바꿔주세요');
      if (!/^\d{4,6}$/.test(String(d.newPin))) throw new Error('비밀번호는 숫자 4~6자리로 정해주세요');
      sheet_(SHOP_SHEET, SHOP_HEAD).getRange(who.row, 3).setNumberFormat('@').setValue(String(d.newPin));
      return out_({ ok: true, data: null });
    }

    const recs = rows_(rs, REC_HEAD.length);
    const mine = r => who.admin || key_(r[2]) === who.key;

    if (act === 'list') {
      const data = recs.filter(r => r[0] && mine(r) && (!d.shop || key_(r[2]) === key_(d.shop))).map(r => ({
        id: r[0], savedAt: r[1], shop: r[2], cust: r[3], date: r[4], round: r[5],
        itemCount: Number(r[6]), plan: Number(r[7]), recv: Number(r[8]), pct: Number(r[9]), left: Number(r[10])
      }));
      return out_({ ok: true, data: data });
    }
    if (act === 'get') {
      const r = recs.find(r => r[0] === d.id && mine(r));
      if (!r) throw new Error('기록이 없어요');
      return out_({ ok: true, data: JSON.parse(r[12]) });
    }
    if (act === 'save') {
      if (who.admin) throw new Error('관리자 계정으로는 저장하지 않아요');
      const r = d.record;
      r.shop = who.store;
      r.id = who.key + '|' + String(r.id || '');
      r.savedAt = now_();
      const i = recs.findIndex(x => x[0] === r.id);
      writeRow_(rs, i >= 0 ? i + 2 : rs.getLastRow() + 1,
        [r.id, r.savedAt, r.shop, r.cust || '', r.date, r.round, r.itemCount, r.plan, r.recv, r.pct, r.left, r.msg || '', JSON.stringify(r)]);
      return out_({ ok: true, data: { id: r.id, savedAt: r.savedAt } });
    }
    if (act === 'delete') {
      const i = recs.findIndex(x => x[0] === d.id && mine(x));
      if (i >= 0) rs.deleteRow(i + 2);
      return out_({ ok: true, data: null });
    }
    if (act === 'shops') {
      if (!who.admin) throw new Error('관리자만 볼 수 있어요');
      const cnt = {};
      recs.forEach(r => { const k = key_(r[2]); cnt[k] = (cnt[k] || 0) + 1; });
      const data = rows_(sheet_(SHOP_SHEET, SHOP_HEAD), SHOP_HEAD.length).filter(r => r[0]).map(r => ({
        store: r[0], pin: r[2], created: r[3], lastUsed: r[4], records: cnt[r[1]] || 0
      }));
      return out_({ ok: true, data: data });
    }
    throw new Error('unknown action');
  } catch (err) {
    return out_({ ok: false, error: err.message || String(err) });
  } finally {
    lock.releaseLock();
  }
}
