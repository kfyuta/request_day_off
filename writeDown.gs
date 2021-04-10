/**
 * 条件付き書式をシートに設定する
 * return {void}
 */
function setConditinalRule() {
  const ss = SpreadsheetApp.getActive();
  const origin = ss.getSheetByName("例");
  const rules = origin.getConditionalFormatRules();
  
  //31行分条件付き書式を作成する
  for (let row = 10; row < 41; row++) {
    const range = origin.getRange(row, 2, 1, 15);
    // 土日祝日の場合
    const formula = `=OR($C${row}="土", $C${row}="日", $D${row}="★")`;
    const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(formula)
            .setBackground("#ffff99") // 薄い黄色
            .setRanges([range])
            .build();
    rules.push(rule);
  }
  origin.setConditionalFormatRules(rules);
}