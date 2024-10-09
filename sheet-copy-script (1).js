function copySheetsBetweenSpreadsheets() {
  // スプレッドシートのIDを指定（実際のIDに置き換えてください）
  var spreadsheetA_ID = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
  var spreadsheetB_ID = "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY";
  var spreadsheetC_ID = "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ";

  // スプレッドシートを開く
  var ssA = SpreadsheetApp.openById(spreadsheetA_ID);
  var ssB = SpreadsheetApp.openById(spreadsheetB_ID);
  var ssC = SpreadsheetApp.openById(spreadsheetC_ID);

  // AとBの最初のシートを取得
  var sheetA = ssA.getSheets()[0];
  var sheetB = ssB.getSheets()[0];

  // Cに新しいシートを作成
  var newSheetA = ssC.insertSheet("From_A");
  var newSheetB = ssC.insertSheet("From_B");

  // シートをコピー
  copySheet(sheetA, newSheetA);
  copySheet(sheetB, newSheetB);

  Logger.log("コピーが完了しました。");
}

function copySheet(sourceSheet, targetSheet) {
  var range = sourceSheet.getDataRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // 値をコピー
  var values = range.getValues();
  targetSheet.getRange(1, 1, numRows, numCols).setValues(values);

  // 書式をコピー
  var formats = range.getTextStyles();
  targetSheet.getRange(1, 1, numRows, numCols).setTextStyles(formats);

  var backgrounds = range.getBackgrounds();
  targetSheet.getRange(1, 1, numRows, numCols).setBackgrounds(backgrounds);

  var fontColors = range.getFontColors();
  targetSheet.getRange(1, 1, numRows, numCols).setFontColors(fontColors);

  // 列幅をコピー
  for (var i = 1; i <= numCols; i++) {
    targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
  }

  // 行の高さをコピー
  for (var i = 1; i <= numRows; i++) {
    targetSheet.setRowHeight(i, sourceSheet.getRowHeight(i));
  }
}
