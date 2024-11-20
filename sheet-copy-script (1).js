// エラーログを格納する配列
let errorLogs = [];

// 抹去ID未合致の場合
if (flag == false) {
  // WOIDの値を取得（列番号は適宜調整してください）
  const woid = data_3GSS[i][WOID_COLUMN];
  
  if (!woid) {
    // WOIDが空白の場合、3GSSマスターから取得
    BA.push([data_3GSS[i][CLM_BA]]);
    BB.push([data_3GSS[i][CLM_BB]]);
    BF.push([data_3GSS[i][CLM_BF]]);
    BG.push([data_3GSS[i][CLM_BG]]);
    BH.push([data_3GSS[i][CLM_BH]]);
    BJ.push([data_3GSS[i][CLM_BJ]]);
    BK.push([data_3GSS[i][CLM_BK]]);
    BL.push([data_3GSS[i][CLM_BL]]);
    BM.push([data_3GSS[i][CLM_BM]]);
  } else if (woid.toString().startsWith('DSB')) {
    // WOIDがDSBから始まる場合
    let jaburoFound = false;
    for (let j = 0; j < data_jb.length; j++) {
      if (data_3GSS[i][51] === data_jb[j][0]) {
        // ジャブローから取得
        AT.push([data_jb[j][1]]);
        AU.push([data_jb[j][2]]);
        jaburoFound = true;
        break;
      }
    }
    
    if (!jaburoFound) {
      // ログにエラーを追加
      errorLogs.push(`撤去ID: ${data_3GSS[i][51]} がジャブローデータに存在しません。（WOID: ${woid}）`);
    }
  } else if (woid.toString().startsWith('KSB')) {
    // WOIDがKSBから始まる場合
    // ※オデッサのデータ取得部分は環境に合わせて調整が必要です
    let odessaFound = false;
    for (let j = 0; j < data_odessa.length; j++) {
      if (data_3GSS[i][51] === data_odessa[j][0]) {
        // オデッサから取得
        AT.push([data_odessa[j][1]]);
        AU.push([data_odessa[j][2]]);
        odessaFound = true;
        break;
      }
    }
    
    if (!odessaFound) {
      // ログにエラーを追加
      errorLogs.push(`撤去ID: ${data_3GSS[i][51]} がオデッサデータに存在しません。（WOID: ${woid}）`);
    }
  }
}

// 最後にログを表示
if (errorLogs.length > 0) {
  const logMessage = "以下のエラーが発生しました：\n\n" + errorLogs.join("\n");
  Browser.msgBox("処理完了", logMessage, Browser.Buttons.OK);
} else {
  Browser.msgBox("処理完了", "すべてのデータが正常に処理されました。", Browser.Buttons.OK);
}

// データ更新処理（既存のコードをそのまま使用）
sheet_3GSS.getRange(6, CLM_AT + 1, AT.length, 1).setValues(AT);
sheet_3GSS.getRange(6, CLM_AU + 1, AU.length, 1).setValues(AU);
sheet_3GSS.getRange(6, CLM_BA + 1, BA.length, 1).setValues(BA);
// ... 他のデータ更新処理も同様
