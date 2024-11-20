// 抹去ID未合致の場合
if (flag == false) {
  // WOIDの値を取得（列番号は適宜調整してください）
  const woid = data_3GSS[i][WOID_COLUMN];
  
  if (!woid) {
    // WOIDが空白の場合
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
    const jaburoExists = checkJaburoExists(data_3GSS[i][51]); // ジャブローデータ存在確認
    
    if (jaburoExists) {
      // ジャブローから取得
      AT.push([data_jb[j][1]]);
      AU.push([data_jb[j][2]]);
    } else {
      // ログを記録
      logNotFound('ジャブロー', data_3GSS[i][51]);
    }
  } else if (woid.toString().startsWith('KSB')) {
    // WOIDがKSBから始まる場合
    const odessaExists = checkOdessaExists(data_3GSS[i][51]); // オデッサデータ存在確認
    
    if (odessaExists) {
      // オデッサから取得（列番号は適宜調整してください）
      AT.push([data_odessa[j][1]]);
      AU.push([data_odessa[j][2]]);
    } else {
      // ログを記録
      logNotFound('オデッサ', data_3GSS[i][51]);
    }
  }
}

// ヘルパー関数
function checkJaburoExists(id) {
  return data_jb.some(row => row[0] === id);
}

function checkOdessaExists(id) {
  return data_odessa.some(row => row[0] === id);
}

function logNotFound(system, id) {
  Logger.log(`${system}データに撤去ID: ${id} が見つかりませんでした`);
}

// 最後にポップアップを表示せずに終了
return;
