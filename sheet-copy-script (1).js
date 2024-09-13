function filterAndMoveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("A");
  var sheetB = ss.getSheetByName("B");
  
  // AシートからデータをコピーAColumnName
  var dataRange = sheetA.getRange("A7:BR7").getValues()[0];
  
  // E列（4番目）の値がBSAX撮去かどうかをチェック
  if (dataRange[4] === "PSAX撤去") {
    // Bシートに貼り付け
    sheetB.getRange(7, 1, 1, dataRange.length).setValues([dataRange]);
    
    // BS列（71番目）が "o" のものだけにフィルター
    if (dataRange[70] === "o") {
      // G列（6番目）で昇順にソート
      sheetB.getRange("A7:BR").sort({column: 7, ascending: true});
    }
  }
}
