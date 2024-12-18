function transferAndUpdateSheetData() {
  // アクティブなスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1つ目と2つ目のシートを取得
  const sheet1 = ss.getSheets()[0];  // 1つ目のシート
  const sheet2 = ss.getSheets()[1];  // 2つ目のシート
  const sheet4 = ss.getSheets()[3];  // 4つ目のシート
  
  // 現在時刻のタイムスタンプを取得してC1セルに記録
  const timestamp = new Date();
  sheet1.getRange("C1")
    .setValue(timestamp)
    .setNumberFormat('"日付:"yyyy/mm/dd hh:mm');
  
  // 当月と翌月の日付を生成
  const today = new Date();
  const currentMonth = new Date(today.getFullYear(), today.getMonth());
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1);
  
  // yyyy/mm/dd形式の文字列に変換
  const formatDate = (date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
  };
  
  const currentMonthStr = formatDate(currentMonth);
  const nextMonthStr = formatDate(nextMonth);
  
  // シート4の10行目のデータを取得
  const row10 = sheet4.getRange("1:1").getValues()[0];
  
  // 日付が一致する列を探し、背景色を設定
  row10.forEach((cell, index) => {
    if (cell instanceof Date) {
      const cellDateStr = formatDate(cell);
      if (cellDateStr === currentMonthStr || cellDateStr === nextMonthStr) {
        // 対象列の10行目から30行目までの範囲を取得
        const column = sheet4.getRange(10, index + 1, 21, 1);
        // 背景色を黄色に設定
        column.setBackground("yellow");
      }
    }
  });
  
  // D4からAE48までの範囲を指定
  const range = sheet1.getRange("D4:AE48");
  
  // 1つ目のシートから値を取得
  const values = range.getValues();
  
  // 2つ目のシートの同じ範囲に値を貼り付け
  sheet2.getRange("D4:AE48").setValues(values);
  
  // 別のスプレッドシートを開く（スプレッドシートIDを指定する必要があります）
  const otherSpreadsheetId = "YOUR_SPREADSHEET_ID_HERE"; // ここに別のスプレッドシートのIDを入力
  const otherSS = SpreadsheetApp.openById(otherSpreadsheetId);
  
  // 3つ目のシートを取得
  const sheet3 = otherSS.getSheets()[2];  // 3つ目のシート
  
  // 3つ目のシートから新しい値を取得
  const newValues = sheet3.getRange("D4:AE48").getValues();
  
  // 1つ目のシートの値を更新
  sheet1.getRange("D4:AE48").setValues(newValues);
}
