// ===== TOTONOI Xmasプラン予約フォーム受信用 GAS =====
// 1. 新しいGoogleスプレッドシート → 拡張機能 → Apps Script に貼り付け
// 2. CHATWORK_TOKEN / ROOM_ID を設定（不要なら空のまま）
// 3. デプロイ → 新しいデプロイ → ウェブアプリ（実行：自分／アクセス：全員）
// 4. 発行されたURLを xmas.html の GAS_URL に貼り付け

const CHATWORK_TOKEN = '';
const CHATWORK_ROOM_ID = '';

const COLS = [
  ['受付日時','_ts'],['ステータス','_st'],['Xmasプラン','xmas'],['チェックイン','checkin'],['泊数','nights'],
  ['大人','adults'],['子ども','children'],['ワンちゃん','dogs'],['希望ドーム','dome'],['夕食','dinner'],
  ['サウナ','sauna'],['体験オプション','exp'],['氏名','name'],['ふりがな','kana'],['メール','email'],
  ['性別','gender'],['生年月日','birth'],['郵便番号','zip'],['住所','addr'],['建物名','bldg'],
  ['電話番号','tel'],['交通手段','access'],['Instagram','insta'],['備考','memo']
];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  d._ts = new Date(); d._st = '未対応';
  const sh = getSheet_();
  sh.appendRow(COLS.map(c => d[c[1]] || ''));
  notify_(d);
  return ContentService.createTextOutput('ok');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Xmas予約');
  if (!sh) {
    sh = ss.insertSheet('Xmas予約');
    sh.appendRow(COLS.map(c => c[0]));
    sh.getRange(1, 1, 1, COLS.length).setBackground('#b88931').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['未対応','連絡済','予約確定','お断り']).build();
    sh.getRange('B2:B1000').setDataValidation(rule);
  }
  return sh;
}

function notify_(d) {
  if (!CHATWORK_TOKEN || !CHATWORK_ROOM_ID) return;
  const body = '[info][title]🎄Xmasプラン DM予約リクエスト[/title]'
    + `Xmasプラン：${d.xmas}\n日程：${d.checkin}（${d.nights}）\n人数：大人${d.adults}／子ども${d.children}／犬${d.dogs}\n`
    + `ドーム：${d.dome}\n夕食：${d.dinner}\nサウナ：${d.sauna}\n体験：${d.exp}\n`
    + `氏名：${d.name}（${d.kana}）\nTEL：${d.tel}\nMail：${d.email}\nInsta：${d.insta || '-'}\n交通：${d.access}\n備考：${d.memo || '-'}[/info]`;
  UrlFetchApp.fetch(`https://api.chatwork.com/v2/rooms/${CHATWORK_ROOM_ID}/messages`, {
    method: 'post', headers: { 'X-ChatWorkToken': CHATWORK_TOKEN }, payload: { body }
  });
}
