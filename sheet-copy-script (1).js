function filterAndMoveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("A");
  var sheetB = ss.getSheetByName("B");
  
  // ヘッダー行を取得（Aシートは6行目、Bシートも6行目）
  var headersA = sheetA.getRange(6, 1, 1, sheetA.getLastColumn()).getValues()[0];
  var headersB = sheetB.getRange(6, 1, 1, sheetB.getLastColumn()).getValues()[0];

  // 必要なカラムのインデックスを取得
  var colIndices = {
    status: headersA.indexOf("状態"),
    bsColumn: headersB.indexOf("BS"),
    sortColumn: headersB.indexOf("ソート基準")
  };

  // Aシートのデータ範囲を取得（7行目から最終行まで）
  var dataRange = sheetA.getRange(7, 1, sheetA.getLastRow() - 6, sheetA.getLastColumn());
  var data = dataRange.getValues();
  
  // フィルター条件に合うデータを抽出
  var filteredData = data.filter(function(row) {
    return row[colIndices.status] === "PSAX撤去" && row[colIndices.bsColumn] === "o";
  });

  if (filteredData.length > 0) {
    // Bシートの既存データをクリア（7行目以降）
    var lastRow = Math.max(sheetB.getLastRow(), 7);
    sheetB.getRange(7, 1, lastRow - 6, sheetB.getLastColumn()).clear();

    // フィルターしたデータをBシートの7行目以降に貼り付け
    sheetB.getRange(7, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    
    // ソート基準カラムで昇順にソート（7行目以降）
    var sortRange = sheetB.getRange(7, 1, filteredData.length, sheetB.getLastColumn());
    sortRange.sort({column: colIndices.sortColumn + 1, ascending: true});
  }
}
