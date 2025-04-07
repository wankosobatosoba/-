/**
 * 大規模データ処理のためのGASスクリプト
 * ソーススプレッドシートから特定の列(A, BA-BD)だけを抽出し、
 * A列でソートした後、ターゲットシートに貼り付ける
 */
function processLargeDataset() {
  // スプレッドシートの取得 (URLは適宜変更してください)
  const sourceSpreadsheet = SpreadsheetApp.openById('ソーススプレッドシートのID');
  const targetSpreadsheet = SpreadsheetApp.openById('ターゲットスプレッドシートのID');
  
  // シートの取得
  const sourceSheet = sourceSpreadsheet.getSheetByName('シート名');
  const targetSheet = targetSpreadsheet.getSheetByName('シート名');
  
  // 処理開始のログ
  console.log('データ処理を開始します');
  console.time('処理時間');
  
  // データの範囲を取得 (3行目から30000行まで、全列)
  const dataRange = sourceSheet.getRange(3, 1, sourceSheet.getLastRow() - 2, sourceSheet.getLastColumn());
  const values = dataRange.getValues();
  
  // 必要な列のインデックスを定義
  // A列は0、BA列は52、BB列は53、BC列は54、BD列は55 (0ベースのインデックス)
  const colIndices = [0, 52, 53, 54, 55];
  
  // 必要な列だけを抽出
  const extractedData = [];
  values.forEach(row => {
    const newRow = [];
    colIndices.forEach(idx => {
      newRow.push(row[idx]);
    });
    extractedData.push(newRow);
  });
  
  // A列（抽出後のデータでは0番目）でソート
  extractedData.sort((a, b) => {
    // 数値の場合
    if (typeof a[0] === 'number' && typeof b[0] === 'number') {
      return a[0] - b[0];
    }
    // 文字列の場合
    return String(a[0]).localeCompare(String(b[0]));
  });
  
  // A列を除去（インデックス1以降を取得）
  const finalData = extractedData.map(row => row.slice(1));
  
  // ターゲットシートのデータが既に存在する場合はクリア
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow > 0) {
    targetSheet.getRange(1, 53, targetLastRow, 4).clearContent();
  }
  
  // データを分割して書き込み (一度に書き込むデータが多すぎるとエラーになるため)
  const CHUNK_SIZE = 1000;
  for (let i = 0; i < finalData.length; i += CHUNK_SIZE) {
    const chunk = finalData.slice(i, i + CHUNK_SIZE);
    targetSheet.getRange(i + 1, 53, chunk.length, 4).setValues(chunk);
    
    // 進捗ログ
    console.log(`${i + chunk.length}/${finalData.length} 行処理完了`);
    
    // 処理が重くならないように少し待機
    Utilities.sleep(100);
  }
  
  // 処理終了のログ
  console.timeEnd('処理時間');
  console.log('データ処理が完了しました');
}

/**
 * 軽量化バージョン: スプレッドシートAPIを使用
 * より高速な処理が必要な場合に使用
 */
function processLargeDatasetOptimized() {
  // スプレッドシートIDの設定
  const sourceId = 'ソーススプレッドシートのID';
  const targetId = 'ターゲットスプレッドシートのID';
  const sourceSheetName = 'シート名';
  const targetSheetName = 'シート名';
  
  // 処理開始のログ
  console.log('最適化されたデータ処理を開始します');
  console.time('処理時間');
  
  // Sheets APIを使用して必要な列だけを一度に取得
  // A列を取得
  const aColumnRange = `${sourceSheetName}!A3:A`;
  const aColumnData = Sheets.Spreadsheets.Values.get(sourceId, aColumnRange).values || [];
  
  // BA-BD列を取得
  const baToFdRange = `${sourceSheetName}!BA3:BD`;
  const baToFdData = Sheets.Spreadsheets.Values.get(sourceId, baToFdRange).values || [];
  
  // 二つのデータを結合
  const combinedData = [];
  for (let i = 0; i < aColumnData.length; i++) {
    if (i < baToFdData.length) {
      combinedData.push([
        aColumnData[i][0], 
        ...(baToFdData[i] || [])
      ]);
    }
  }
  
  // A列（インデックス0）でソート
  combinedData.sort((a, b) => {
    if (typeof a[0] === 'number' && typeof b[0] === 'number') {
      return a[0] - b[0];
    }
    return String(a[0]).localeCompare(String(b[0]));
  });
  
  // A列を除去
  const finalData = combinedData.map(row => row.slice(1));
  
  // データを一括で書き込み (Sheets APIを使用)
  Sheets.Spreadsheets.Values.update(
    {
      values: finalData
    },
    targetId,
    `${targetSheetName}!BA1:BD${finalData.length}`,
    {
      valueInputOption: 'USER_ENTERED'
    }
  );
  
  // 処理終了のログ
  console.timeEnd('処理時間');
  console.log('最適化されたデータ処理が完了しました');
}

/**
 * もっと軽量化したバージョン: バッチ処理とキャッシュを使用
 */
function processLargeDatasetUltraOptimized() {
  // スプレッドシートの取得
  const sourceSpreadsheet = SpreadsheetApp.openById('ソーススプレッドシートのID');
  const targetSpreadsheet = SpreadsheetApp.openById('ターゲットスプレッドシートのID');
  
  const sourceSheet = sourceSpreadsheet.getSheetByName('シート名');
  const targetSheet = targetSpreadsheet.getSheetByName('シート名');
  
  console.log('超最適化されたデータ処理を開始します');
  console.time('処理時間');
  
  // まずA列だけを取得してソート用のインデックスを作成
  const aColumnRange = sourceSheet.getRange(3, 1, sourceSheet.getLastRow() - 2, 1);
  const aColumnValues = aColumnRange.getValues();
  
  // ソート用のインデックス配列を作成 (値と元の行番号のペア)
  const sortIndices = aColumnValues.map((value, index) => ({ value: value[0], originalIndex: index }));
  
  // インデックス配列をA列の値でソート
  sortIndices.sort((a, b) => {
    if (typeof a.value === 'number' && typeof b.value === 'number') {
      return a.value - b.value;
    }
    return String(a.value).localeCompare(String(b.value));
  });
  
  // BA-BD列を取得
  const baToFdRange = sourceSheet.getRange(3, 53, sourceSheet.getLastRow() - 2, 4);
  const baToFdValues = baToFdRange.getValues();
  
  // ソートされたインデックスに基づいて新しいデータセットを作成
  const sortedData = sortIndices.map(item => baToFdValues[item.originalIndex]);
  
  // ターゲットシートに書き込む
  const CHUNK_SIZE = 1000;
  for (let i = 0; i < sortedData.length; i += CHUNK_SIZE) {
    const chunk = sortedData.slice(i, i + CHUNK_SIZE);
    targetSheet.getRange(i + 1, 53, chunk.length, 4).setValues(chunk);
    
    // 進捗ログと軽いスリープ
    console.log(`${i + chunk.length}/${sortedData.length} 行処理完了`);
    Utilities.sleep(50);
  }
  
  console.timeEnd('処理時間');
  console.log('超最適化されたデータ処理が完了しました');
}
