// おにぎりEXPRESS 注文受付スクリプト(Googleスプレッドシート用)

// ★ここを自分で決めたパスワードに変えてください(注文管理ページを開くときに使います)
const KEY = 'ここを自分のパスワードに変える';

const SHEET = 'orders';
const HEAD = ['注文番号', '受付日時', '日にち', '受け取り', 'お名前', 'セット数', '合計', 'おにぎり', '受け渡し済み'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    // 「2-7」や「10/10」「12:30」が日付や時刻に勝手に変わらないよう、文字として扱う
    ['A:A', 'C:D', 'H:H'].forEach(r => sh.getRange(r).setNumberFormat('@'));
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// お客さんの注文を受け取る / 管理ページの操作(受け渡し済み・削除)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = sheet_();

    if (d.action) {
      if (d.key !== KEY) return out_({ ok: false, error: 'auth' });
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(d.code) && String(rows[i][2]) === String(d.day)) {
          if (d.action === 'done') sh.getRange(i + 1, 9).setValue(d.value ? '済' : '');
          if (d.action === 'delete') sh.deleteRow(i + 1);
          return out_({ ok: true });
        }
      }
      return out_({ ok: false, error: 'notfound' });
    }

    if (!d.name || !d.day || !d.time || !(d.qty >= 1 && d.qty <= 10) || !Array.isArray(d.picks) || d.picks.length !== d.qty * 2) {
      return out_({ ok: false, error: 'bad' });
    }
    // 同じ日・同じ受け取り時間の中で、注文順に 1, 2, 3... と番号を付ける
    const rows = sh.getDataRange().getValues().slice(1);
    let max = 0;
    rows.forEach(r => {
      if (String(r[2]) === String(d.day) && String(r[3]) === String(d.time)) {
        max = Math.max(max, Number(String(r[0]).split('-')[1]) || 0);
      }
    });
    const code = (Number(d.slotIndex) || 0) + '-' + (max + 1);
    sh.appendRow([code, new Date(), d.day, d.time, String(d.name).slice(0, 30), d.qty, d.total, d.picks.join('・'), '']);
    return out_({ ok: true, code: code });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 管理ページが注文一覧を読むとき(パスワードが合わないと返しません)
function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out_({ ok: false, error: 'auth' });
  const rows = sheet_().getDataRange().getValues().slice(1);
  const orders = rows.map(r => ({
    code: String(r[0]), at: r[1], day: String(r[2]), time: String(r[3]), name: String(r[4]),
    qty: Number(r[5]), total: Number(r[6]), picks: String(r[7]), done: r[8] === '済'
  }));
  return out_({ ok: true, orders: orders });
}
