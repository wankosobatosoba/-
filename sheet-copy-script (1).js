function filterAndMoveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("A");
  var sheetB = ss.getSheetByName("B");
  
  // ヘッダー行を取得（6行目）
  var headersA = sheetA.getRange(6, 1, 1, sheetA.getLastColumn()).getValues()[0];
  var headersB = sheetB.getRange(6, 1, 1, sheetB.getLastColumn()).getValues()[0];

  // 必要なカラムのインデックスを取得
  var colIndices = {
    status: headersA.indexOf("状態"),
    bsColumn: headersB.indexOf("BS"),
    sortColumn: headersB.indexOf("ソート基準")
  };

  // データ行を取得（7行目）
  var dataRange = sheetA.getRange(7, 1, 1, sheetA.getLastColumn()).getValues()[0];
  
  // 状態が "PSAX撤去" かどうかをチェック
  if (dataRange[colIndices.status] === "PSAX撤去") {
    // Bシートに貼り付け
    var targetRange = sheetB.getRange(7, 1, 1, dataRange.length);
    targetRange.setValues([dataRange]);
    
    // BSカラムが "o" のものだけにフィルター
    var bsCell = targetRange.offset(0, colIndices.bsColumn);
    if (bsCell.getValue() === "o") {
      // ソート基準カラムで昇順にソート
      var sortRange = sheetB.getRange(7, 1, sheetB.getLastRow() - 6, sheetB.getLastColumn());
      sortRange.sort({column: colIndices.sortColumn + 1, ascending: true});
    }
  }
}
