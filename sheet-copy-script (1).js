function copySheetWithFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("A");
  var sheetB = ss.getSheetByName("B");
  
  if (!sheetA || !sheetB) {
    Logger.log("シートAまたはシートBが見つかりません。");
    return;
  }
  
  var newSheetA = ss.insertSheet("C_A");
  var newSheetB = ss.insertSheet("C_B");
  
  copySheet(sheetA, newSheetA);
  copySheet(sheetB, newSheetB);
  
  Logger.log("コピーが完了しました。");
}

function copySheet(sourceSheet, targetSheet) {
  var range = sourceSheet.getDataRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  
  // 値と書式をコピー
  range.copyTo(targetSheet.getRange(1, 1, numRows, numCols), {contentsOnly: false, formatOnly: false});
  
  // 列幅をコピー
  for (var i = 1; i <= numCols; i++) {
    targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
  }
  
  // 行の高さをコピー
  for (var i = 1; i <= numRows; i++) {
    targetSheet.setRowHeight(i, sourceSheet.getRowHeight(i));
  }
}
