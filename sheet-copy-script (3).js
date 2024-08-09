function copyFilteredDataBetweenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("masterSheet");
  const psSheet = ss.getSheetByName("psSheet");
  
  const headerRange = masterSheet.getRange(6, 1, 1, masterSheet.getLastColumn());
  const header = headerRange.getValues()[0];
  
  const projectNameIndex = header.indexOf("案件名2");
  const woIdIndex = header.indexOf("WO-ID");
  const dateIndex = header.indexOf("日付");
  const duplicateCheckIndex = header.indexOf("重複チェック");
  
  if (projectNameIndex === -1 || woIdIndex === -1 || dateIndex === -1 || duplicateCheckIndex === -1) {
    throw new Error("必要な列が見つかりません。");
  }
  
  // masterSheetのフィルター処理
  const filterMaster = masterSheet.getFilter();
  if (filterMaster) {
    filterMaster.remove();
  }
  const rangeMaster = masterSheet.getRange(6, 1, masterSheet.getLastRow() - 5, masterSheet.getLastColumn());
  rangeMaster.createFilter();
  
  // データと数式を取得
  const dataRange = masterSheet.getRange(7, 1, masterSheet.getLastRow() - 6, masterSheet.getLastColumn());
  const data = dataRange.getValues();
  const formulas = dataRange.getFormulas();
  
  // PSAXのみをフィルター
  const filteredData = data.map((row, index) => {
    if (row[projectNameIndex] === "PSAX") {
      return formulas[index].map((formula, colIndex) => formula || row[colIndex]);
    }
    return null;
  }).filter(row => row !== null);
  
  // psSheetの処理
  const lastRow = psSheet.getLastRow();
  const lastColumn = psSheet.getLastColumn();
  if (lastRow > 6 && lastColumn > 0) {
    psSheet.getRange(7, 1, lastRow - 6, lastColumn).clear({contentsOnly: true, skipFilteredRows: true});
  }
  
  const targetHeaderRange = psSheet.getRange(6, 1, 1, header.length);
  headerRange.copyTo(targetHeaderRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
  
  if (filteredData.length > 0) {
    psSheet.getRange(7, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
  
  // psSheetのフィルター処理
  const filterPs = psSheet.getFilter();
  if (filterPs) {
    filterPs.remove();
  }
  const rangePs = psSheet.getRange(6, 1, Math.max(2, psSheet.getLastRow() - 5), psSheet.getLastColumn());
  const newFilter = rangePs.createFilter();
  
  // 日付列のフィルター設定（空白を非表示）
  newFilter.setColumnFilterCriteria(dateIndex + 1, 
    SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build()
  );
  
  // 重複チェック列のフィルター設定（○のみ表示）
  newFilter.setColumnFilterCriteria(duplicateCheckIndex + 1, 
    SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo("○")
    .build()
  );
  
  // WO-ID列を基準に昇順で並び替え
  rangePs.sort({column: woIdIndex + 1, ascending: true});
}
