function updateWorkerDropdown(sheetName : string) : void{

  const progressRange = getRangeByName('진행상태');
  const startRow = progressRange.getRow();
  const dataColumn = progressRange.getColumn();
  const dropdownInfoRange = getColumnRange(getSheetByName("설정"), startRow, dataColumn);
  
  //적용 범위 선택
  const applyFieldRange = getRangeByName("작업자필드")
  const applyRange = makeApplyRange(sheetName, applyFieldRange, getCutCount())

  applyDropdown(dropdownInfoRange,applyRange)
}

function updateProgressDropdown(sheetName : string) : void {

  const progressRange = getRangeByName('진행상태');
  const startRow = progressRange.getRow();
  const dataColumn = progressRange.getColumn();
  const dropdownInfoRange = getColumnRange(getSheetByName("설정"), startRow, dataColumn);

  //적용 범위 선택
  const applyFieldRange = getRangeByName("진행현황필드")
  const applyRange = makeApplyRange(sheetName ,applyFieldRange, getCutCount())

  applyDropdown(dropdownInfoRange, applyRange)
}

function makeApplyRange(sheetName : string, applyFieldRange : Range, cutCount : number) : Range {
  const dataRow = applyFieldRange.getRow() + 1;
  const dataColumn = applyFieldRange.getColumn();
  const applyRange = getSheetByName(sheetName).getRange(dataRow, dataColumn, cutCount);
  return applyRange
}

function applyDropdown(infoRange : Range, applyRange : Range) : void {
  // const beforeRowCount = applyRange.getSheet().getLastRow()-applyRange.getRow()+1;
  // const beforeColumnCount = applyRange.getNumColumns();
  // console.log(beforeRowCount,beforeColumnCount)
  // if (getCutCount() == beforeRowCount || isSameDropdown(infoRange, applyRange)) {
  //   console.log("same")
  //   return;
  // }
  const clearRange = applyRange
  clearDropdown(clearRange)
  applyDropdownText(infoRange,applyRange)
  applyDropdownColor(infoRange,applyRange)
}

function applyDropdownText(infoRange : Range, applyRange : Range) : void {
  // 드롭다운 목록을 만들기 위한 데이터 유효성 객체 생성
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infoRange.getValues().flat()) // 값 범위를 배열로 변환하여 지정
    .setAllowInvalid(false)
    .build();
  applyRange.setDataValidation(rule);
}

function applyDropdownColor(infoRange : Range,applyRange : Range) : void {
  const sheet = applyRange.getSheet()
  const colors = infoRange.getBackgrounds().flat()
  const values = infoRange.getValues().flat()
  const rules = sheet.getConditionalFormatRules(); //기존룰

  for(let i=0;i<values.length;i++){
    let rule = makeConditionalFormattingRule(values[i],colors[i],applyRange) //필터 없을 때 규칙
    rules.push(rule)
  }
  sheet.setConditionalFormatRules(rules);
}

//조건부 서식 규칙 생성
function makeConditionalFormattingRule(text : string ,color: string, rng : Range) : ConditionalFormatRule{
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(color)
    .setRanges([rng])
    .build();
  return rule
}

function clearDropdown(rng : Range) : void {
  clearDropdownText(rng)
  clearDropdownColor(rng)
}

function clearDropdownColor(rng : Range) : void {
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

function clearDropdownText(rng : Range) : void {
  rng.setDataValidation(null);
}

function isSameDropdown(infoRange : Range, applyRange : Range) : boolean {
  const currentRule = applyRange.getDataValidation();
  const newRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infoRange.getValues().flat()) // 값 범위를 배열로 변환하여 지정
    .setAllowInvalid(false)
    .build();

  console.log(currentRule,newRule)
  const currentRuleValues = currentRule?.getCriteriaValues()
  const newRuleValues = newRule.getCriteriaValues()
  
  console.log(currentRuleValues, newRuleValues)
  return JSON.stringify(currentRuleValues) === JSON.stringify(newRuleValues);
}
