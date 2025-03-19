/************ Google Apps Script (GAS) チートシート ************/

// ===== スプレッドシートの基本操作 =====

// アクティブなスプレッドシートを取得
const ss = SpreadsheetApp.getActiveSpreadsheet();

// IDからスプレッドシートを取得
const ss = SpreadsheetApp.openById('スプレッドシートID');

// URLからスプレッドシートを取得
const ss = SpreadsheetApp.openByUrl('スプレッドシートURL');

// 新規スプレッドシート作成
const ss = SpreadsheetApp.create('新規スプレッドシート名');

// アクティブなシートを取得
const sheet = SpreadsheetApp.getActiveSheet();

// シート名からシートを取得
const sheet = ss.getSheetByName('シート名');

// インデックスからシートを取得（0から始まる）
const sheet = ss.getSheets()[0];

// 新規シート作成
const newSheet = ss.insertSheet('新規シート名');

// シートのコピー
const copiedSheet = sheet.copyTo(ss).setName('コピーシート名');

// ===== セルと範囲の操作 =====

// 単一セルの値を取得
const value = sheet.getRange('A1').getValue();
// または
const value = sheet.getRange(1, 1).getValue();

// 範囲の値を取得（2次元配列で返される）
const values = sheet.getRange('A1:C3').getValues();
// または
const values = sheet.getRange(1, 1, 3, 3).getValues(); // 開始行, 開始列, 行数, 列数

// 列全体を取得
const columnValues = sheet.getRange('A:A').getValues();

// 行全体を取得
const rowValues = sheet.getRange('1:1').getValues();

// データが存在する最終行を取得
const lastRow = sheet.getLastRow();

// データが存在する最終列を取得
const lastColumn = sheet.getLastColumn();

// 単一セルに値を設定
sheet.getRange('A1').setValue('テキスト');

// 範囲に値を設定（2次元配列）
const data = [
  ['名前', '年齢', '都市'],
  ['田中', 25, '東京'],
  ['鈴木', 30, '大阪']
];
sheet.getRange('A1:C3').setValues(data);

// 数式を設定
sheet.getRange('D1').setFormula('=SUM(A1:C1)');

// 書式を設定
sheet.getRange('A1').setFontWeight('bold').setBackground('#f3f3f3');

// セルの値をクリア
sheet.getRange('A1').clearContent();

// セルの書式をクリア
sheet.getRange('A1').clearFormat();

// セルのすべてをクリア（値、書式、検証など）
sheet.getRange('A1').clear();

// 行を削除
sheet.deleteRow(1);

// 複数行を削除
sheet.deleteRows(1, 3); // 開始行, 削除する行数

// 列を削除
sheet.deleteColumn(1);

// 複数列を削除
sheet.deleteColumns(1, 3); // 開始列, 削除する列数

// ===== データの処理 =====

// フィルターを作成
const filter = sheet.getRange('A1:D10').createFilter();

// フィルター条件を設定（例：B列が「東京」のみ表示）
const filterCriteria = SpreadsheetApp.newFilterCriteria()
  .whenTextEqualTo('東京')
  .build();
filter.setColumnFilterCriteria(2, filterCriteria); // 2は列のインデックス（B列）

// フィルターを削除
filter.remove();

// 範囲をソート（A列で昇順）
sheet.getRange('A2:D10').sort(1); // 1はA列を表す

// 複数条件でソート（A列で昇順、B列で降順）
sheet.getRange('A2:D10').sort([
  {column: 1, ascending: true},
  {column: 2, ascending: false}
]);

// ===== 制御構文 =====

// 基本的なif文
if (条件) {
  // 条件が真の場合の処理
} else if (別の条件) {
  // 別の条件が真の場合の処理
} else {
  // すべての条件が偽の場合の処理
}

// 例
const value = sheet.getRange('A1').getValue();
if (value > 100) {
  console.log('100より大きい');
} else if (value > 50) {
  console.log('50より大きく100以下');
} else {
  console.log('50以下');
}

// 三項演算子
const result = (条件) ? '真の場合の値' : '偽の場合の値';

// 基本的なfor文
for (let i = 0; i < 10; i++) {
  console.log(i);
}

// 配列のforEachループ
const array = [1, 2, 3, 4, 5];
array.forEach(function(item, index) {
  console.log(item, index);
});

// for...ofループ（ES6）
for (const item of array) {
  console.log(item);
}

// for...inループ（オブジェクトのプロパティに対して）
const obj = {a: 1, b: 2, c: 3};
for (const key in obj) {
  console.log(key, obj[key]);
}

// 2次元配列のループ（スプレッドシートデータの処理に便利）
const values = sheet.getRange('A1:C3').getValues();
for (let i = 0; i < values.length; i++) {
  for (let j = 0; j < values[i].length; j++) {
    console.log(values[i][j]);
  }
}

// while文
let i = 0;
while (i < 10) {
  console.log(i);
  i++;
}

// do...while文
let j = 0;
do {
  console.log(j);
  j++;
} while (j < 10);

// ===== 関数と変数 =====

// 関数の定義
function myFunction(param1, param2) {
  // 処理
  return 結果;
}

// アロー関数（ES6）
const myArrowFunction = (param1, param2) => {
  // 処理
  return 結果;
};

// 変数の定義
let variable1 = '変更可能な変数';
const constant1 = '変更不可能な定数';
var oldVariable = '古い変数宣言方法（避けるべき）';

// ===== 日付と時間 =====

// 現在の日時を取得
const now = new Date();

// 特定の日時を作成
const date = new Date(2023, 0, 1); // 2023年1月1日（月は0から始まる）

// Utilities.formatDateを使って日付をフォーマット
const formattedDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

// セルに日付を設定
sheet.getRange('A1').setValue(new Date());

// ===== UI操作とユーザー対話 =====

// アラートを表示
SpreadsheetApp.getUi().alert('メッセージ');

// 確認ダイアログを表示
const response = SpreadsheetApp.getUi().alert('確認', '続行しますか？', SpreadsheetApp.getUi().ButtonSet.YES_NO);
if (response == SpreadsheetApp.getUi().Button.YES) {
  // YESが選択された場合の処理
}

// プロンプトでユーザー入力を受け取る
const input = SpreadsheetApp.getUi().prompt('入力', '名前を入力してください', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
if (input.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
  const text = input.getResponseText();
  // 入力されたテキストで処理
}

// カスタムダイアログを表示（HTMLファイルが必要）
const html = HtmlService.createHtmlOutputFromFile('Page')
  .setWidth(400)
  .setHeight(300);
SpreadsheetApp.getUi().showModalDialog(html, 'カスタムダイアログ');

// ===== トリガーの設定 =====

// 時間ベースのトリガーを設定
function createTimeDrivenTriggers() {
  // 毎日午前9時に実行
  ScriptApp.newTrigger('myFunction')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  
  // 1時間ごとに実行
  ScriptApp.newTrigger('myFunction')
    .timeBased()
    .everyHours(1)
    .create();
}

// スプレッドシートの変更時に実行するトリガー
function createOnEditTrigger() {
  ScriptApp.newTrigger('myFunction')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// ===== エラー処理 =====

// try-catch文でエラーをキャッチ
try {
  // エラーが発生する可能性のあるコード
} catch (e) {
  // エラー発生時の処理
  console.error('エラーが発生しました: ' + e.message);
  SpreadsheetApp.getUi().alert('エラー: ' + e.message);
} finally {
  // エラーの有無にかかわらず実行される処理
}

// ===== 外部サービスとの連携 =====

// HTTPリクエストを送信
const response = UrlFetchApp.fetch('https://api.example.com/data');
const data = JSON.parse(response.getContentText());

// POSTリクエストを送信
const options = {
  'method': 'post',
  'contentType': 'application/json',
  'payload': JSON.stringify({key: 'value'})
};
const response = UrlFetchApp.fetch('https://api.example.com/data', options);

// Gmail操作
const threads = GmailApp.search('label:inbox');
const messages = GmailApp.getMessagesForThreads(threads);

// Gmailで新規メール送信
GmailApp.sendEmail('recipient@example.com', '件名', '本文');
