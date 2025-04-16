/** 將多列資料整合在一起，並且轉成一欄方式呈現
 * @param {Array} ranges 範圍, 可以給多個範圍
 * @return {Array} [[item1], [item2], ...]
 * @customfunction
 */
function RowToColumnSet(...ranges) {
  const all = new Set()
  ranges.forEach(row=>{
    row.forEach(cols=>{
      cols.forEach(cell=>{
        if (cell !== "") {
          all.add(cell)
        }
      })
    })
  })
  return  [...all.values()].map(e=>[e])
}

/** 轉成鎖定的範圍
 * @param {Array} ranges
 * @return {String} "A1:B10"
 * @customfunction
 */
function GetLockedA1Notation(range) {
  // 取得範圍的 A1 表示法
  let a1Notation = range.getA1Notation();

  // 如果是單一儲存格
  if (!a1Notation.includes(':')) {
    let col = a1Notation.match(/[A-Z]+/)[0];
    let row = a1Notation.match(/\d+/)[0];
    return `$${col}$${row}`;
  }

  // 如果是範圍（包含 ":"）
  let [start, end] = a1Notation.split(':');
  let startCol = start.match(/[A-Z]+/)[0];
  let startRow = start.match(/\d+/)[0];
  let endCol = end.match(/[A-Z]+/)[0];
  let endRow = end.match(/\d+/)[0];

  return `$${startCol}$${startRow}:$${endCol}$${endRow}`;
}


/** 取得台灣指定年份的行事曆
 * 範例: https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/2025.json
 * @param {String} year
 * @customfunction
 */
function GetCalendar(year) {
  const url = `https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/${year}.json`
  const response = UrlFetchApp.fetch(url)
  const json = JSON.parse(response.getContentText())

  const holidays = []
  json.forEach(item => {
    const dateStr = item.date
    const year = dateStr.substring(0, 4)
    const month = dateStr.substring(4, 6)
    const day = dateStr.substring(6, 8)
    const date = new Date(year, month - 1, day)
    holidays.push([date, item.week, item.isHoliday, item.description])
  })
  return holidays
}


function TestRowToColumnSet() {
  // const sheet = ss.getActiveSheet()
  // const a = RowToColumnSet(sheet.getRange("B225:H225").getValues())
  const b = RowToColumnSet(
    [["c1", "c2"]],
    [["c3", "", "c4"]],
  )
  console.log(b)
  // Output:
  // [[c1], [c2], [c3], [c4]]
}


function TestGetLockedA1Notation() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet.getRange('A1:B10'); // 範圍範例
  let lockedRange = getLockedA1Notation(range);
  Logger.log(lockedRange); // 輸出：$A$1:$B$10
}
