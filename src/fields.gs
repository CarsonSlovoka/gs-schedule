const FieldDate = '日期'
const FieldOnDutyPerson = '值班人'
const FieldOnDutyInspection = '日巡查'
const FieldWeeklyInspection = '周巡查'
const FieldBiWeekly = '二周'
const FieldNight = '夜間'
const FieldRemarks = '備註'
const FieldDateRemarks = '日期備註'

const itemHeader = ['日期', '值班人', '日巡查', '周巡查', '二周', '夜間', '備註', '日期備註']

function getHeaderIndex(name) {
  return itemHeader.findIndex(e=> e === name)
}
