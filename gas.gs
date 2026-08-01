/**
 * 勤怠管理システム スプレッドシート連携(Google Apps Script)
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成(名前は任意。例:「勤怠データ」)
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. このファイルの内容をすべて貼り付けて保存
 * 4. 下の TOKEN を任意の文字列に変更する(例:'mise2026abc')
 * 5. 「デプロイ」→「新しいデプロイ」→ 種類:「ウェブアプリ」
 *    - 実行するユーザー:自分
 *    - アクセスできるユーザー:全員
 *    でデプロイし、表示された「ウェブアプリのURL」をコピー
 * 6. 勤怠管理の管理画面 →「従業員・設定」→「スプレッドシート連携」に
 *    URLと合言葉(TOKENと同じ文字列)を入力し、「連携を有効にする」にチェックして保存
 * 7. 「テスト送信」を押し、このスプレッドシートの「打刻ログ」シートに行が追加されれば完了
 *
 * 【シート構成(自動作成)】
 * - 勤怠記録:記録IDごとに常に最新の状態(修正・削除も反映)
 * - 打刻ログ:すべての操作の履歴(追記のみ。監査ログとして残ります)
 *
 * 【注意】コードを修正した場合は「デプロイ」→「デプロイを管理」→ 編集 →
 *        バージョン「新バージョン」で再デプロイしないと反映されません。
 */

const TOKEN = 'kintai'; // ←必ず変更してください(管理画面の「合言葉」と一致させる)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return out({ ok: false, error: 'token' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = getSheet(ss, '打刻ログ',
      ['受信日時', '操作', '記録ID', '日付', '氏名', '区分', '出勤', '退勤', '休憩(分)', '実働(h)', '残業(h)', '深夜(h)', 'メモ']);
    const master = getSheet(ss, '勤怠記録',
      ['記録ID', '日付', '氏名', '区分', '出勤', '退勤', '休憩(分)', '実働(h)', '残業(h)', '深夜(h)', 'メモ', '最終更新']);

    const items = data.action === 'bulk'
      ? (data.records || []).map(r => ({ action: 'upsert', record: r }))
      : [{ action: data.action, record: data.record }];

    items.forEach(it => apply(log, master, it.action, it.record || {}));
    return out({ ok: true, count: items.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function apply(log, master, action, r) {
  // 履歴ログ(常に追記)
  log.appendRow([new Date(), action, r.id || '', r.date || '', r.empName || '', r.empType || '',
    r.clockIn || '', r.clockOut || '', r.breakMin || 0, r.workH || '', r.otH || '', r.nightH || '', r.note || '']);

  if (action !== 'upsert' && action !== 'delete') return; // testなどはログのみ

  // 勤怠記録シート(記録IDでupsert)
  const last = master.getLastRow();
  let idx = -1;
  if (last >= 2) {
    const ids = master.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    idx = ids.indexOf(String(r.id));
  }
  if (action === 'delete') {
    if (idx > -1) master.deleteRow(idx + 2);
    return;
  }
  const row = [r.id, r.date, r.empName, r.empType, r.clockIn, r.clockOut,
    r.breakMin, r.workH, r.otH, r.nightH, r.note, new Date()];
  if (idx > -1) master.getRange(idx + 2, 1, 1, row.length).setValues([row]);
  else master.appendRow(row);
}

function getSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#173A52').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
