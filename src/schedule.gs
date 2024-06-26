function generateSchedule(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  // const lastRow = sheet.getLastRow()
  // const cell = sheet.getRange("A1")
  // const cell = sheet.getRange(row, column)
  const cell = sheet.getActiveCell()

  const ruleCheckbox = SpreadsheetApp.newDataValidation().requireCheckbox().build()
  const holidayMap = GetHolidayMap()

  let row = cell.getRow()
  const beginRow = row
  let col = cell.getColumn()
  const beginCol = col

  // 標題
  const titleRange = sheet.getRange(row, col, 1, 8)
  titleRange.merge().setValue(year + "/" + (month < 10 ? "0" + month : month))
  titleRange.setHorizontalAlignment("center")
  titleRange.setVerticalAlignment("middle")
  titleRange.setFontSize(36)
  row++
  var weekdays = ["一", "二", "三", "四", "五", "六", "日"]
  const endCol = col + 7 + 1 // 有一列是空白分隔
  for (var i = 0; i < 7; i++) {
    sheet.getRange(row, col + i + 1).setValue(weekdays[i]);
  }
  row++

  const date = new Date(year, month - 1, 1) // month = 0為1月
  const endData = new Date(year, month, 1)
  if (date.getDay() !== 1) { // 0 sunDay, (0 to 6)
    date.setDate(date.getDate() - (date.getDay() - 1)) // 如果1日是在星期五，那麼要退回到5-1等於4天前
  }
  const itemHeader = ['日期', '值班人', '日巡查', '周巡查', '二周', '夜間', '備註', '日期備註']
  let endRow = row
  while (date < endData) {
    const beginCol = col
    // 縱向
    const values = itemHeader.map(e => [e]) // [['日期], []'值班人'], ...]
    sheet.getRange(row, col, itemHeader.length, 1).setValues(values)
    sheet.getRange(row, col).setBackground("#46BDC6") // 日期
    sheet.getRange(row, col).setFontColor("white")
    col++

    // 橫向: 日期1,2,...31
    [...Array(7)].map((_, i) => {
      if (date < endData) { // 避免因為+1之後已經不同月份了
        const curRange = sheet.getRange(row, col)
        curRange.setValue(date.getDate())
        const holidayObj = holidayMap[date]
        if (holidayObj !== undefined) {
          if (holidayObj.isHoliday) {
            sheet.getRange(row, col).setBackground("#FF99CC")
          }
          if (holidayObj.desc !== "") {
            sheet.getRange(row + 7, col).setValue(holidayObj.desc) // 我們知道這裡往下7列會到日期備註
          }
        }
        /*
        if (i === 5 || i=== 6) { // 星期六,日
          sheet.getRange(row, col).setBackground("#FF99CC")
        }
        */
        date.setDate(date.getDate() + 1)
        col++
      }
    })

    // 橫向: 日巡查
    // 從日巡查開始 的 1列 7欄 改成checkbox
    const rangeDaily = sheet.getRange(row + 2, // 日巡查
      beginCol + 1, 1, 7)
    rangeDaily.setValues([
      [...Array(7)].map((_, i) => "False")
    ]) // [[True, True, ...]]
    rangeDaily.setDataValidation(ruleCheckbox)
    rangeDaily.setFontColor("red")

    row += itemHeader.length
    endRow = row
    col = beginCol
  }

  // 全域格式設定
  const entireRange = sheet.getRange(beginRow, beginCol,
    endRow - beginRow,
    endCol - beginCol,
  )
  entireRange.setHorizontalAlignment("center")
  entireRange.setVerticalAlignment("middle")
  entireRange.setBorder(
    true, true, true, true, true, true,
    "black", // 邊框顏色
    SpreadsheetApp.BorderStyle.SOLID // SpreadsheetApp.BorderStyle.SOLID_MEDIUM // 邊框樣式 https://developers.google.com/apps-script/reference/spreadsheet/border-style
  )

  sheet.getRange(endRow + 1, beginCol + 3).setValue("上班日")
}

function TestGenerateSchedule() {
  generateSchedule(2024, 1)
}
