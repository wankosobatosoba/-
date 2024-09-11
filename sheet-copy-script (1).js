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
    
    // ユニークな組み合わせの場合のみ処理を行う
    if (!uniqueCombinations[category].has(combinationKey)) {
      uniqueCombinations[category].add(combinationKey);

      const rawDate = category === 'possibleRemoval' ? row[colIndexes.acColumn] : row[colIndexes.bwColumn];
      processDate(rawDate, category, monthCounts, logMessages);
    }
  });

  return { monthCounts, uniqueCombinations, logMessages, dateColumns, totalProcessed: removalData.length };
}

function processDate(rawDate, category, monthCounts, logMessages) {
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
