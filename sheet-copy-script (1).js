function manipulateSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = spreadsheet.getSheetByName('A');
  var sheetB = spreadsheet.getSheetByName('B');
  
  // Aシートからデータを取得
  var dataA = sheetA.getRange(7, 1, sheetA.getLastRow() - 6, 70).getValues(); // A列からBR列まで
  
  // F列が"PSAX撤去"のデータのみをフィルター
  var filteredData = dataA.filter(function(row) {
    return row[5] === 'PSAX撤去'; // F列は0から数えて5番目
  });
  
  // Bシートのデータをクリア
  sheetB.getRange(7, 1, sheetB.getLastRow() - 6, 70).clearContent();
  
  // フィルターしたデータをBシートに貼り付け
  if (filteredData.length > 0) {
    sheetB.getRange(7, 1, filteredData.length, 70).setValues(filteredData);
  }
  
  // BS列でフィルター
  var range = sheetB.getRange(7, 1, sheetB.getLastRow() - 6, 70);
  var filter = range.getFilter() || range.createFilter();
  filter.setColumnFilterCriteria(71, SpreadsheetApp.newFilterCriteria().whenTextEqualTo('o')); // BS列は71番目
  
  // G列で昇順に並べ替え
  range.sort({column: 7, ascending: true}); // G列は7番目
}
