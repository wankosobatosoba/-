function countTotalCells() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let totalCells = 0;
  let result = [];
  
  // シート毎の全セル数を計算
  for (const sheet of sheets) {
    const numRows = sheet.getMaxRows();
    const numCols = sheet.getMaxColumns();
    const sheetCells = numRows * numCols;
    
    totalCells += sheetCells;
    result.push({
      sheetName: sheet.getName(),
      rows: numRows,
      cols: numCols,
      cells: sheetCells
    });
  }
  
  // スプレッドシートの制限値
  const CELL_LIMIT = 10000000; // 1000万セル
  const remainingCells = CELL_LIMIT - totalCells;
  
  // 結果をログに出力
  Logger.log('=== セル数（空白セル含む） ===');
  Logger.log(`総セル数: ${totalCells.toLocaleString()} セル`);
  Logger.log(`残りセル数: ${remainingCells.toLocaleString()} セル`);
  Logger.log('\nシート別セル数:');
  result.forEach(item => {
    Logger.log(`${item.sheetName}: ${item.cells.toLocaleString()} セル (${item.rows} 行 × ${item.cols} 列)`);
  });
  
  return {
    totalCells: totalCells,
    remainingCells: remainingCells,
    sheetDetails: result
  };
}

// UI付きで実行するための関数
function showTotalCellCount() {
  const result = countTotalCells();
  const ui = SpreadsheetApp.getUi();
  
  let message = `総セル数: ${result.totalCells.toLocaleString()} セル\n`;
  message += `残りセル数: ${result.remainingCells.toLocaleString()} セル\n\n`;
  message += 'シート別セル数:\n';
  result.sheetDetails.forEach(item => {
    message += `${item.sheetName}: ${item.cells.toLocaleString()} セル`;
    message += ` (${item.rows} 行 × ${item.cols} 列)\n`;
  });
  
  ui.alert('セル数カウント結果', message, ui.ButtonSet.OK);
}

// メニューに追加するための関数
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムツール')
    .addItem('全セル数をカウント', 'showTotalCellCount')
    .addToUi();
}
