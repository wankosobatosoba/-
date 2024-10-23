// === 基本的なスプレッドシート/シートの取得 ===

// アクティブなスプレッドシートを取得
function getActiveSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss;
}

// IDからスプレッドシートを取得
function getSpreadsheetById() {
  const ss = SpreadsheetApp.openById('スプレッドシートのID');
  return ss;
}

// URLからスプレッドシートを取得
function getSpreadsheetByUrl() {
  const ss = SpreadsheetApp.openByUrl('スプレッドシートのURL');
  return ss;
}

// シート名からシートを取得
function getSheetByName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1');
  return sheet;
}

// === 範囲の取得と設定 ===

// A1表記で範囲を取得
function getRangeByA1() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1:D10');
  return range;
}

// 行と列の番号で範囲を取得
function getRangeByPosition() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(1, 1, 10, 4); // 開始行, 開始列, 行数, 列数
  return range;
}

// 最終行までの範囲を取得
function getRangeToLastRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(`A1:D${lastRow}`);
  return range;
}

// 最終列までの範囲を取得
function getRangeToLastColumn() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, 1, lastColumn);
  return range;
}

// データが存在する最終行を取得
function getLastRowWithData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getDataRange().getLastRow();
  return lastRow;
}

// === データの取得 ===

// 範囲の値を2次元配列で取得
function getValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getRange('A1:D10').getValues();
  return values;
}

// 範囲の数式を取得
function getFormulas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const formulas = sheet.getRange('A1:D10').getFormulas();
  return formulas;
}

// 範囲の表示値を取得
function getDisplayValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const displayValues = sheet.getRange('A1:D10').getDisplayValues();
  return displayValues;
}

// 特定の列のデータを取得
function getColumnData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const columnData = sheet.getRange(`A1:A${lastRow}`).getValues();
  return columnData;
}

// === データの設定 ===

// 範囲に値を設定
function setValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = [['データ1', 'データ2'], ['データ3', 'データ4']];
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

// 範囲に数式を設定
function setFormulas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const formulas = [['=SUM(B1:B10)', '=AVERAGE(C1:C10)']];
  sheet.getRange(1, 1, formulas.length, formulas[0].length).setFormulas(formulas);
}

// 最終行の次の行にデータを追加
function appendRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['データ1', 'データ2', 'データ3']);
}

// === フォーマット設定 ===

// 背景色を設定
function setBackground() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1:D1').setBackground('#f3f3f3');
}

// フォントの色を設定
function setFontColor() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1:D1').setFontColor('#ff0000');
}

// セルの表示形式を設定
function setNumberFormat() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1:A10').setNumberFormat('@'); // テキスト
  sheet.getRange('B1:B10').setNumberFormat('#,##0'); // 数値
  sheet.getRange('C1:C10').setNumberFormat('yyyy/mm/dd'); // 日付
}

// セルの整列を設定
function setAlignment() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1:D10');
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
}

// === シートの操作 ===

// 新しいシートを追加
function addNewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSheet = ss.insertSheet('新しいシート');
  return newSheet;
}

// シートの並び順を変更
function moveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.activate();
  sheet.moveToEnd(); // または .moveToStart()
}

// シートをコピー
function copySheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet.copyTo(ss).setName('コピーしたシート');
}

// === 行と列の操作 ===

// 行を挿入
function insertRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRows(1, 5); // 1行目に5行挿入
}

// 列を挿入
function insertColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumns(1, 3); // A列に3列挿入
}

// 行を削除
function deleteRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.deleteRows(1, 5); // 1行目から5行削除
}

// === 保護の設定 ===

// 範囲を保護
function protectRange() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1:D10');
  const protection = range.protect();
  
  // 特定のユーザーに編集権限を付与
  protection.addEditor('user@example.com');
  
  // 保護の説明を設定
  protection.setDescription('保護された範囲');
}

// === フィルタの操作 ===

// フィルタを設定
function createFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1:D10');
  const filter = range.createFilter();
  return filter;
}

// === 条件付き書式 ===

// 条件付き書式を設定
function setConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1:D10');
  
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1000)
    .setBackground('#FF0000')
    .setRanges([range])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// === 検索と置換 ===

// テキストを検索
function findText() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const textFinder = sheet.createTextFinder('検索テキスト');
  const ranges = textFinder.findAll();
  return ranges;
}

// テキストを置換
function replaceText() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const textFinder = sheet.createTextFinder('置換前テキスト');
  textFinder.replaceAllWith('置換後テキスト');
}

// === ソートと並べ替え ===

// 範囲をソート
function sortRange() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A2:D10');
  range.sort([{
    column: 1, // A列でソート
    ascending: true
  }, {
    column: 2, // B列で次にソート
    ascending: false
  }]);
}

// === ピボットテーブル ===

// ピボットテーブルを作成
function createPivotTable() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sourceData = sheet.getRange('A1:D100');
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  
  const pivotTable = targetSheet.getRange('A1').createPivotTable(sourceData);
  
  pivotTable.addRowGroup(1); // 1列目を行としてグループ化
  pivotTable.addColumnGroup(2); // 2列目を列としてグループ化
  pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // 4列目の合計を計算
}
