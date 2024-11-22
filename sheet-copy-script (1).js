function transferDataFromSlidesToSheets() {
  // スライドとスプレッドシートのIDを設定
  const PRESENTATION_ID = ''; // スライドのIDを入力
  const SPREADSHEET_ID = ''; // スプレッドシートのIDを入力
  const SHEET_NAME = ''; // シート名を入力
  const TARGET_TITLE = '3GSS GC設備撤去 進捗状況'; // 検索するスライドのタイトルに含まれる文字列
  
  try {
    // スライドとスプレッドシートを取得
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // スライドを検索
    const slides = presentation.getSlides();
    let targetSlide = null;
    
    for (const slide of slides) {
      // スライド内のすべてのシェイプを取得
      const shapes = slide.getShapes();
      for (const shape of shapes) {
        // シェイプにテキストがある場合、タイトルを検索
        if (shape.getText) {
          const text = shape.getText().asString();
          if (text.includes(TARGET_TITLE)) {
            targetSlide = slide;
            break;
          }
        }
      }
      if (targetSlide) break;
    }
    
    if (!targetSlide) {
      throw new Error('対象のスライドが見つかりませんでした: ' + TARGET_TITLE);
    }
    
    // テーブルを取得（最初のテーブルを想定）
    const tables = targetSlide.getTables();
    if (tables.length === 0) {
      throw new Error('テーブルが見つかりませんでした');
    }
    const table = tables[0];
    
    // データを格納する2次元配列
    const data = [];
    
    // テーブルの各行をループ
    for (let i = 0; i < table.getNumRows(); i++) {
      const row = [];
      const tableRow = table.getRow(i);
      
      for (let j = 0; j < table.getNumColumns(); j++) {
        const cell = tableRow.getCell(j);
        const value = cell.getText().asString().trim();
        row.push(value);
      }
      data.push(row);
    }
    
    // メインデータ部分の更新（1行目から8行目まで）
    const mainDataRange = sheet.getRange(1, 1, 8, data[0].length);
    mainDataRange.setValues(data.slice(0, 8));
    
    Logger.log('データ転記が完了しました');
    
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error);
    throw error; // エラーを再スロー
  }
}

// トリガーを設定する関数
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('transferDataFromSlidesToSheets')
    .timeBased()
    .everyHours(1)
    .create();
}

// 手動実行用の関数
function manualRun() {
  transferDataFromSlidesToSheets();
}
