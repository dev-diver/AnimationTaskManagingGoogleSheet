function updateWorkerDropdown(partCol) {

  //드롭다운 정보 범위 선택
  const partRange = getRangeByName('파트시작');
  const partNameRow = partRange.getRow()
  const dropdownInfoRange = getColumnRange(getSheetByName("설정"), partNameRow+1, partCol);
  
  //적용 범위 선택
  const partName = getSheetByName("설정").getRange(partNameRow, partCol).getValue()
  const applyFieldRange = getRangeByName("작업자필드")
  const applyRange = makeApplyRange(partName+" 파트", applyFieldRange, getCutCount())

  applyDropdown(dropdownInfoRange,applyRange)
}

function updateProgressDropdown(sheetName) {

  const progressRange = getRangeByName('진행상태');
  const startRow = progressRange.getRow();
  const dataColumn = progressRange.getColumn() + 1;
  const dropdownInfoRange = getColumnRange(getSheetByName("설정"), startRow, dataColumn);

  const applyFieldRange = getRangeByName("진행현황필드")
  const applyRange = makeApplyRange(sheetName,applyFieldRange, getCutCount())

  applyDropdown(dropdownInfoRange,applyRange)
}

function makeApplyRange(sheetName, applyFieldRange, cutCount){
  const dataRow = applyFieldRange.getRow() + 1;
  const dataColumn = applyFieldRange.getColumn();
  const applyRange = getSheetByName(sheetName).getRange(dataRow, dataColumn, cutCount);
  return applyRange
}

function applyDropdown(infoRange, applyRange){
  if (isSameDropdown(infoRange, applyRange)) {
    console.log("same")
    return;
  }
  clearDropdown(applyRange)
  applyDropdownText(infoRange,applyRange)
  applyDropdownColor(infoRange,applyRange)
}

function applyDropdownText(infoRange, applyRange){
  // 드롭다운 목록을 만들기 위한 데이터 유효성 객체 생성
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infoRange.getValues().flat()) // 값 범위를 배열로 변환하여 지정
    .setAllowInvalid(false)
    .build();
  applyRange.setDataValidation(rule);
}

function applyDropdownColor(infoRange,applyRange){
  const sheet = applyRange.getSheet()
  const colors = infoRange.getBackgrounds().flat()
  const values = infoRange.getValues().flat()
  const rules = sheet.getConditionalFormatRules(); //기존룰

  for(let i=0;i<values.length;i++){
    let rule = makeConditionalFormattingRule_(values[i],colors[i],applyRange) //필터 없을 때 규칙
    rules.push(rule)
  }
  sheet.setConditionalFormatRules(rules);
}

//조건부 서식 규칙 생성
function makeConditionalFormattingRule_(text,color,rng) {
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(color)
    .setRanges([rng])
    .build();
  return rule
}

function clearDropdown(rng){
  clearDropdownText(rng)
  clearDropdownColor(rng)
}

function clearDropdownColor(rng){
  const sheet = rng.getSheet()
  let rules = sheet.getConditionalFormatRules();
  rules = rules.filter(rule=>{ //모든 룰 중에서
    const ruleRange = rule.getRanges()
    return !ruleRange.some(range=>
    {
      // console.log(range.getA1Notation(), rng.getA1Notation())
      let bool = RangeIntersect_(range,rng)  //원하는 range만 고름
      return bool
    })
  })
  // console.log(rules.length)
  sheet.setConditionalFormatRules(rules);
}

function clearDropdownText(rng) {
  // 드롭다운 목록을 제거하기 위해 데이터 유효성 규칙을 제거합니다.
  rng.setDataValidation(null);
}

function getCutCount(){
  const cutCountRange = getRangeByName("컷수")
  return cutCountRange.offset(0,1).getValue()
}

function isSameDropdown(infoRange, applyRange) {
  const currentRule = applyRange.getDataValidation();
  const newRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infoRange.getValues().flat()) // 값 범위를 배열로 변환하여 지정
    .setAllowInvalid(false)
    .build();

  console.log(currentRule,newRule)
  const currentRuleValues = currentRule.getCriteriaValues()
  const newRuleValues = newRule.getCriteriaValues()
  
  console.log(currentRuleValues, newRuleValues)
  return JSON.stringify(currentRuleValues) === JSON.stringify(newRuleValues);
}

function RangeIntersect_(R1, R2) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}