function inputpivot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("3GSS_マスターファイル_V1.0");
  const targetSheet = ss.getSheetByName("週次報告集計 のコピー 2");

  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const headers = sourceData[5]; // 6行目がヘッダー行
  const colIndexes = {
    hColumn: headers.indexOf("局舎名"),
    baColumn: headers.indexOf("撤去_WOID"),
    bzColumn: headers.indexOf("週次ピボット"),
    bwColumn: headers.indexOf("撤去予定月"),
    ankenName2Column: headers.indexOf("案件名２"),
    acColumn: headers.indexOf("実地局回線完了日")
  };

  const removalTypes = ['PSAX撤去', 'MURS撤去', 'SHDSL撤去'];
  const results = {};

  removalTypes.forEach(removalType => {
    results[removalType] = processRemovalType(removalType, sourceData, targetData, colIndexes);
  });

  // 結果をシートに書き込み
  writeResultsToSheet(targetData, results, targetSheet);

  // ログ出力
  logResults(results);
}

function processRemovalType(removalType, sourceData, targetData, colIndexes) {
  const monthCounts = {
    completed: {}, planned: {}, tentative: {}, selfApplied: {}, possibleRemoval: {}
  };
  const uniqueCombinations = {
    completed: new Set(), planned: new Set(), tentative: new Set(),
    selfApplied: new Set(), possibleRemoval: new Set()
  };
  const logMessages = [];

  // ターゲットシートの日付行（71行目）を取得し、日付列のインデックスを特定
  const dateRow = targetData[70];
  const dateColumns = {};
  for (let i = 6; i <= 17; i++) {
    if (dateRow[i]) {
      const formattedDateString = formatDate(dateRow[i]);
      dateColumns[formattedDateString] = i;
      Object.keys(monthCounts).forEach(key => {
        monthCounts[key][formattedDateString] = 0;
      });
    }
  }

  // 撤去タイプでフィルターをかける処理
  const removalData = sourceData.filter((row, index) => 
    index > 5 && row[colIndexes.ankenName2Column] === removalType
  );

  // データ処理
  removalData.forEach(row => {
    const bzColumnValue = row[colIndexes.bzColumn];
    const category = getCategoryFromBzValue(bzColumnValue);
    if (!category) return;

    const combinationKey = `${row[colIndexes.hColumn]}_${row[colIndexes.baColumn]}`;
    
    // ユニークな組み合わせの場合のみ処理を続行
    if (!uniqueCombinations[category].has(combinationKey)) {
      uniqueCombinations[category].add(combinationKey);

      const rawDate = category === 'possibleRemoval' ? row[colIndexes.acColumn] : row[colIndexes.bwColumn];
      processUniqueDate(rawDate, category, monthCounts, logMessages);
    }
  });

  return { monthCounts, uniqueCombinations, logMessages, dateColumns, totalProcessed: removalData.length };
}

function formatDate(dateValue) {
  if (dateValue instanceof Date) {
    if (isNaN(dateValue.getTime())) return null;
    return `${dateValue.getFullYear().toString().slice(-2)}/${(dateValue.getMonth() + 1).toString().padStart(2, '0')}`;
  } else if (typeof dateValue === 'string') {
    const parts = dateValue.split('/');
    if (parts.length >= 2) {
      return `${parts[0].slice(-2)}/${parts[1].padStart(2, '0')}`;
    }
  }
  return null;
}

function getCategoryFromBzValue(bzValue) {
  const categoryMap = {
    "N検完了": 'completed',
    "N検予定": 'planned',
    "仮定": 'tentative',
    "撤去自前申請": 'selfApplied',
    "撤去可能ビル": 'possibleRemoval'
  };
  return categoryMap[bzValue];
}

function processUniqueDate(rawDate, category, monthCounts, logMessages) {
  if (rawDate instanceof Date || typeof rawDate === 'string') {
    const formattedDate = formatDate(rawDate);
    if (formattedDate && formattedDate.match(/^\d{2}\/\d{2}$/)) {
      if (formattedDate in monthCounts[category]) {
        monthCounts[category][formattedDate]++;
      } else {
        logMessages.push(`警告: ${formattedDate} は有効な月のリストにございません。rawDate: ${rawDate}`);
      }
    } else {
      logMessages.push(`警告: 無効な日付形式 "${rawDate}" が見つかりました。フォーマット後: ${formattedDate}`);
    }
  } else if (rawDate === undefined) {
    logMessages.push(`警告: 日付が未定義です。カテゴリー: ${category}`);
  } else {
    logMessages.push(`警告: 無効な日付値 "${rawDate}" が見つかりました。タイプ: ${typeof rawDate}`);
  }
}

function writeResultsToSheet(targetData, results, targetSheet) {
  const patterns = [
    { suffix: '単月撤去可能ビル', key: 'possibleRemoval' },
    { suffix: '単月撤去自前申請', key: 'selfApplied' },
    { suffix: '単月N検完了', key: 'completed' },
    { suffix: '単月N検予定', key: 'planned' },
    { suffix: '単月仮予定', key: 'tentative' }
  ];

  Object.entries(results).forEach(([removalType, data]) => {
    patterns.forEach(pattern => {
      let rowPrefix;
      if (removalType === 'MURS撤去') {
        rowPrefix = 'MU-RS';
      } else {
        rowPrefix = removalType.split('撤去')[0];
      }
      const rowName = `${rowPrefix}${pattern.suffix}`;
      const targetRow = targetData.findIndex(row => row[4] === rowName) + 1;
      if (targetRow === 0) {
        console.log(`エラー: targetSheetに列名"${rowName}"の行が見つかりませんでした。`);
      } else {
        Object.entries(data.monthCounts[pattern.key]).forEach(([dateString, count]) => {
          const col = data.dateColumns[dateString] + 1;
          targetSheet.getRange(targetRow, col).setValue(count);
        });
        console.log(`${rowName}のカウント結果をtargetSheetに入力しました。`);
      }
    });
  });
}

function logResults(results) {
  console.log("\n=== 処理結果のサマリー ===\n");

  const categoryNames = {
    possibleRemoval: '撤去可能ビル',
    selfApplied: '撤去自前申請',
    completed: 'N検完了',
    planned: 'N検予定',
    tentative: '仮予定'
  };

  const categoryOrder = ['possibleRemoval', 'selfApplied', 'completed', 'planned', 'tentative'];

  Object.entries(results).forEach(([removalType, data]) => {
    const displayType = removalType === 'MURS撤去' ? 'MU-RS撤去' : removalType;
    console.log(`\n${displayType}の結果:`);

    console.log("\n1. 月別カウント結果:");
    categoryOrder.forEach(category => {
      const counts = data.monthCounts[category];
      console.log(`\n  ${categoryNames[category]}:`);
      Object.entries(counts)
        .sort(([a], [b]) => a.localeCompare(b))
        .forEach(([month, count]) => {
          console.log(`    ${month}: ${count}件`);
        });
    });

    console.log("\n2. 全体の処理行数:", data.totalProcessed);

    console.log("\n3. カテゴリー別ユニークな組み合わせ数:");
    categoryOrder.forEach(category => {
      console.log(`  ${categoryNames[category]}: ${data.uniqueCombinations[category].size}件`);
    });

    if (data.logMessages.length > 0) {
      console.log("\n4. 警告・情報メッセージ:");
      const undefinedCount = data.logMessages.filter(msg => msg.includes('日付が未定義です')).length;
      console.log(`  未定義の日付: ${undefinedCount}件`);
      
      data.logMessages.forEach((msg, index) => {
        if (!msg.includes('日付が未定義です')) {
          console.log(`  ${index + 1}. ${msg}`);
        }
      });
    }
  });

  console.log("\n=== 処理完了 ===");
}
