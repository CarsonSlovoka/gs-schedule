function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('🧙我的自定義選單')
    .addItem('生成班表', 'showPrompt')
    .addToUi()
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi()
  var yearResult = ui.prompt('請輸入年份 (例如: 2024)')
  if (yearResult.getSelectedButton() === ui.Button.OK) {
    var year = yearResult.getResponseText()
    var monthResult = ui.prompt('請輸入月份 (1-12)')
    if (monthResult.getSelectedButton() === ui.Button.OK) {
      var month = monthResult.getResponseText()
      generateSchedule(parseInt(year), parseInt(month))
    }
  }
}
