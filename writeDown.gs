function setConditinalRule() {
  const ss = SpreadsheetApp.getActive();
  const origin = ss.getSheetByName("例");
  const rules = origin.getConditionalFormatRules();
  for (let row = 10; row < 41; row++) {
    const range = origin.getRange(row, 2, 1, 15);
    const formula = `=OR($C${row}="土", $C${row}="日", $D${row}="★")`;
    const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(formula)
            .setBackground("#ffff99")
            .setRanges([range])
            .build();
    rules.push(rule);
  }
  origin.setConditionalFormatRules(rules);
}