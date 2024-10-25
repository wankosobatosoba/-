// 処理済みの管理番号を追跡するためのSetを作成
const processedNumbers = new Set();

for(i = 0; i < data_check.length; i++) {
  // 文字列型で取得できるのに、再変換しないと比較ができない
  data_check[i][CHECK_DAY] = Utilities.formatDate(new Date(data_check[i][CHECK_DAY]), 'JST', 'yyyy/MM/dd');
  
  if(data_check[i][CHECK_CHK] != "" && data_check[i][CHECK_DAY] == data_check[1][CHECK_DAY]) {
    // 現在の管理番号を取得
    const currentNumber = data_check[i][CHECK_NNM];
    
    // この管理番号がまだ処理されていない場合のみ処理を実行
    if (!processedNumbers.has(currentNumber)) {
      for(j = 0; j < data_targetList.length; j++) {
        // NW構築課の振切りはGC施工管理課に送付
        let targetDivusion = data_check[i][CHECK_DIV];
        if(targetDivusion == "NW構築課") {
          targetDivusion = "GC施工管理課";
        }
        
        if(data_targetList[j][0] == targetDivusion) {
          // 管理番号挿入
          data_targetList[j][2] += data_check[i][CHECK_NNM] + "\n";
          // 処理済みの管理番号を記録
          processedNumbers.add(currentNumber);
        }
      }
    }
  }
}
