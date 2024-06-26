const ruleCheckbox = SpreadsheetApp.newDataValidation().requireCheckbox().build()
const ruleEmployee = SpreadsheetApp.newDataValidation()
  .requireValueInList(
    GetEmployeeData()
      .filter(e=>e[2] === true) // 設定在職員工為下拉式選單的內容
      .map(e=>e[1]), // 挑選出姓名欄位
    true) // showDropdown
  .setAllowInvalid(false)
  .build()
