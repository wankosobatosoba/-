function copyFilteredDataBetweenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("masterSheet");
  const psSheet = ss.getSheetByName("psSheet");
  
  const headerRange = masterSheet.getRange(6, 1, 1, masterSheet.getLastColumn());
  const header = headerRange.getValues()[0];
  
  const projectNameIndex = header.indexOf("案件名2");
  const woIdIndex = header.indexOf("WO-ID");
  const dateIndex = header.indexOf("日付"); // 日付列のインデックスを取得
  
  if (projectNameIndex === -1 || woIdIndex === -1 || dateIndex === -1) {
    throw new Error("必要な列が見つかりません。");
  }
  
  // masterSheetのフィルター処理
  const filterMaster = masterSheet.getFilter();
  if (filterMaster) {
    filterMaster.remove();
  }
  const rangeMaster = masterSheet.getRange(6, 1, masterSheet.getLastRow() - 5, masterSheet.getLastColumn());
  rangeMaster.createFilter();
  
  const data = masterSheet.getRange(7, 1, masterSheet.getLastRow() - 6, masterSheet.getLastColumn()).getValues();
  
  const filteredData = data.filter(row => row[projectNameIndex].toString().includes("PSAX"));
  
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
  rangePs.createFilter();
  
  // WO-ID列を基準に昇順で並び替え
  const range = psSheet.getRange(7, 1, psSheet.getLastRow() - 6, psSheet.getLastColumn());
  range.sort({column: woIdIndex + 1, ascending: true});
  
  // 重複チェック列を追加
  const duplicateCheckIndex = header.length;
  psSheet.getRange(6, duplicateCheckIndex + 1).setValue("重複チェック");
  
  // 重複チェックと日付列の空白チェック
  const psData = psSheet.getRange(7, 1, psSheet.getLastRow() - 6, psSheet.getLastColumn()).getValues();
  const duplicateChecks = [];
  const hiddenRows = [];
  
  for (let i = 0; i < psData.length; i++) {
    // 重複チェック
    if (i === 0 || psData[i][woIdIndex] !== psData[i-1][woIdIndex]) {
      duplicateChecks.push(["○"]);
    } else {
      duplicateChecks.push(["✕"]);
    }
    
    // 日付列の空白チェック
    if (psData[i][dateIndex] === "") {
      hiddenRows.push(i + 7); // 7を加えて実際の行番号に調整
    }
  }
  
  // 重複チェック結果を書き込み
  psSheet.getRange(7, duplicateCheckIndex + 1, duplicateChecks.length, 1).setValues(duplicateChecks);
  
  // 日付が空白の行を非表示に
  psSheet.hideRows(hiddenRows);
}
