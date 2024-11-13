for (let i = 0; i < values1.length; i++) {
    // values2[i]が存在する場合のみ処理を行う
    if (values2[i]) {
        for (let j = 0; j < values1[i].length; j++) {
            // values2[i][j]が存在する場合のみ比較を行う
            if (values2[i][j] !== undefined && formulas2[i][j] !== undefined) {
                if (values1[i][j] != values2[i][j] || formulas1[i][j] != formulas2[i][j]) {
                    range = dataSheet.getRange(i + 1, j + 1);
                    range.setBackground('yellow');
                    result++;
                }
            }
        }
    }
}
