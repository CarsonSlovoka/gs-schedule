// https://jsdoc.app/tags-typedef

/**
 * 休假日
 * @typedef {Object} Holiday
 * @property {boolean} isHoliday
 * @property {string} desc
 */

function GetHolidayData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getSheetByName('放假清單')
  const range = sheet.getRange('A2:D') // 跳過標題
  const values = range.getValues().filter(e => e[0] !== "") // 空列就跳過
  return values.map(e => { // 將日期轉成Date物件
    const dateStr = e[0] // 記得要調整此欄成文字，不是數字
    const year = dateStr.substring(0, 4)
    const month = dateStr.substring(4, 6)
    const day = dateStr.substring(6, 8)
    const date = new Date(year, month - 1, day)
    return [date, e[1], e[2], e[3]]
  })
}

class HolidayDB {
  constructor() {
    const m = {}
    GetHolidayData().forEach(e => {
      m[e[0]] = {
        isHoliday: e[2] === 2,
        desc: e[3],
      }
    })
    this.m = m
  }

  /**
   * @return Holiday
   **/
  get(date) {
    return this.m[date]
  }
}
