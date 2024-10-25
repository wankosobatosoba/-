// 新しい処理：管理番号ごとの作成者とメール送付回数を取得
let controlNumbers = [];
let creators = [];
let mailCounts = [];
let monthCounts = new Map(); // 管理番号ごとの月別カウントを保持
let dailyChecks = new Map(); // 管理番号ごとの日次チェックを保持

for (let i = 1; i < data_check.length; i++) {
  if (data_check[i][CHECK_CHK] !== "") {
    let controlNumber = data_check[i][CHECK_NNM];
    let creator = data_check[i][CHECK_MAK];
    let checkDate = data_check[i][CHECK_DAY];
    
    // 日付が正しく取得できた場合のみ処理
    if (checkDate instanceof Date && !isNaN(checkDate)) {
      // 年月日の文字列を作成（例: "2024-01-15"）
      let fullDate = Utilities.formatDate(checkDate, 'Asia/Tokyo', 'yyyy-MM-dd');
      let yearMonth = Utilities.formatDate(checkDate, 'Asia/Tokyo', 'yyyy-MM');
      
      // その日の最初のチェックかどうかを確認
      let dailyKey = `${controlNumber}_${fullDate}`;
      if (!dailyChecks.has(dailyKey)) {
        dailyChecks.set(dailyKey, true);
        
        let index = controlNumbers.indexOf(controlNumber);
        if (index === -1) {
          // 新しい管理番号の場合
          controlNumbers.push(controlNumber);
          creators.push(creator);
          
          // 月別カウントの初期化
          monthCounts.set(controlNumber, new Set([yearMonth]));
          mailCounts.push(1);
        } else {
          // 既存の管理番号の場合、その月のカウントが未登録なら追加
          let months = monthCounts.get(controlNumber);
          if (!months.has(yearMonth)) {
            months.add(yearMonth);
            monthCounts.set(controlNumber, months);
            mailCounts[index] = months.size;
          }
        }
      }
    }
  }
}

// デバッグ用のログ出力
console.log("日次チェック状況:");
for (let [key, value] of dailyChecks) {
  console.log(`${key}: ${value}`);
}
