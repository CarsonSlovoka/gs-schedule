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
