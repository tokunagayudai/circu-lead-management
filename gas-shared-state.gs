/**
 * ===== チーム共有ストレージ（アカプラ編集内容の共有） =====
 *
 * リード管理ポータルのアカプラ編集内容（組織図の移動・合体・攻略メモ・
 * 手動追加した人物など）をスプレッドシートに保存し、チーム全員のブラウザで
 * 同じ状態を共有するためのGoogle Apps Script追加コードです。
 *
 * 【セットアップ手順】
 * 1. ポータルがデータ取得に使っているGASプロジェクト
 *    （GAS_URLのWebアプリを提供しているApps Script）を開く
 * 2. このファイルの内容を新しいスクリプトファイルとして貼り付ける
 * 3. 既存の doGet(e) 関数の先頭に、次の1行を追加する:
 *
 *      if(e && e.parameter && e.parameter.action === 'state_get') return stateGet_();
 *
 * 4. 既に doPost(e) が存在する場合は、このファイルの doPost の中身を
 *    既存の doPost にマージする（存在しなければこのファイルのdoPostがそのまま使われる）
 * 5. 「デプロイ」→「デプロイを管理」→ 既存のWebアプリデプロイを「編集」して
 *    新しいバージョンとして再デプロイする
 *    （アクセス権限は既存と同じ「全員」のままでOK）
 *
 * スプレッドシートに「shared_state」というシートが自動作成され、
 * key / value / updated_at の3列で編集内容が保存されます。
 */

var STATE_SHEET_NAME = 'shared_state';

/** 共有ステートの全件取得（doGetから action=state_get で呼ばれる） */
function stateGet_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(STATE_SHEET_NAME);
  var out = {};
  if (sh && sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    vals.forEach(function(row) {
      if (row[0]) out[String(row[0])] = String(row[1] == null ? '' : row[1]);
    });
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 共有ステートの保存（ポータルからPOSTで呼ばれる） */
function doPost(e) {
  var res = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.action === 'state_set' && body.items) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var ss = SpreadsheetApp.getActive();
        var sh = ss.getSheetByName(STATE_SHEET_NAME) || ss.insertSheet(STATE_SHEET_NAME);
        if (sh.getLastRow() === 0) sh.appendRow(['key', 'value', 'updated_at']);
        var lastRow = sh.getLastRow();
        var rowByKey = {};
        if (lastRow > 1) {
          var keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
          keys.forEach(function(r, i) { if (r[0]) rowByKey[String(r[0])] = i + 2; });
        }
        var now = new Date();
        Object.keys(body.items).forEach(function(k) {
          var v = String(body.items[k] == null ? '__DEL__' : body.items[k]);
          if (v.length > 49000) return; // Sheetsのセル上限(5万文字)対策: 巨大値はスキップ
          var row = rowByKey[k];
          if (row) {
            sh.getRange(row, 2, 1, 2).setValues([[v, now]]);
          } else {
            sh.appendRow([k, v, now]);
            rowByKey[k] = sh.getLastRow();
          }
        });
        res.ok = true;
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    res.error = String(err);
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}
