function copyFilteredDataBetweenSheets() {
  // スプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // masterSheetとpsSheetを取得
  const masterSheet = ss.getSheetByName("masterSheet");
  const psSheet = ss.getSheetByName("psSheet");
  
  // ヘッダー行（6行目）を取得
  const headerRange = masterSheet.getRange(6, 1, 1, masterSheet.getLastColumn());
  const header = headerRange.getValues()[0];
  
  // "案件名2"と"WO-ID"の列のインデックスを見つける
  const projectNameIndex = header.indexOf("案件名2");
  const woIdIndex = header.indexOf("WO-ID");
  
  if (projectNameIndex === -1) {
    throw new Error("'案件名2'の列が見つかりません。");
  }
  if (woIdIndex === -1) {
    throw new Error("'WO-ID'の列が見つかりません。");
  }
  
  // masterSheetのフィルターを更新または作成
  const filterMaster = masterSheet.getFilter();
  if (filterMaster) {
    filterMaster.remove();
  }
  const rangeMaster = masterSheet.getRange(6, 1, masterSheet.getLastRow() - 5, masterSheet.getLastColumn());
  rangeMaster.createFilter();
  
  // masterSheetからデータを取得（7行目以降）
  const data = masterSheet.getRange(7, 1, masterSheet.getLastRow() - 6, masterSheet.getLastColumn()).getValues();
  
  // "PSAX"を含む行だけをフィルター
  const filteredData = data.filter(function(row) {
    return row[projectNameIndex].toString().includes("PSAX");
  });
  
  // psSheetをクリア（ヘッダー行より下、フィルターを除く）
  const lastRow = psSheet.getLastRow();
  const lastColumn = psSheet.getLastColumn();
  if (lastRow > 6 && lastColumn > 0) {
    psSheet.getRange(7, 1, lastRow - 6, lastColumn).clear({contentsOnly: true, skipFilteredRows: true});
  }
  
  // psSheetのヘッダー行（6行目）の書式をmasterSheetのヘッダー行に合わせる
  const targetHeaderRange = psSheet.getRange(6, 1, 1, header.length);
  headerRange.copyTo(targetHeaderRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
  
  // フィルターされたデータがある場合、psSheetに貼り付け
  if (filteredData.length > 0) {
    psSheet.getRange(7, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
  
  // psSheetのフィルターを更新または作成
  const filterPs = psSheet.getFilter();
  if (filterPs) {
    filterPs.remove();
  }
  const rangePs = psSheet.getRange(6, 1, Math.max(2, psSheet.getLastRow() - 5), psSheet.getLastColumn());
  rangePs.createFilter();
  
  // WO-ID列を基準に昇順で並び替え
  const range = psSheet.getRange(7, 1, psSheet.getLastRow() - 6, psSheet.getLastColumn());
  range.sort({column: woIdIndex + 1, ascending: true});
}
