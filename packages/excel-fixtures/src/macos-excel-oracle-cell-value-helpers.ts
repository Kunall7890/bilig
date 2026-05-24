export function cellValueAppleScriptHelpers(): string {
  return `on formulaText(cellFormula)
  if cellFormula is missing value then
    return ""
  end if
  return cellFormula as string
end formulaText

on typedCellValue(cellValue, renderedValue)
  if cellValue is missing value then
    if my isExcelErrorDisplayText(renderedValue) then
      return "error" & (ASCII character 9) & renderedValue
    end if
    return "blank" & (ASCII character 9)
  end if
  set valueClass to class of cellValue
  if valueClass is boolean then
    if cellValue then
      return "boolean" & (ASCII character 9) & "true"
    end if
    return "boolean" & (ASCII character 9) & "false"
  end if
  if valueClass is integer or valueClass is real then
    return "number" & (ASCII character 9) & (cellValue as string)
  end if
  return "string" & (ASCII character 9) & (cellValue as string)
end typedCellValue

on isExcelErrorDisplayText(displayText)
  if displayText is "#DIV/0!" then return true
  if displayText is "#REF!" then return true
  if displayText is "#VALUE!" then return true
  if displayText is "#NAME?" then return true
  if displayText is "#N/A" then return true
  if displayText is "#SPILL!" then return true
  if displayText is "#BLOCKED!" then return true
  if displayText is "#NUM!" then return true
  if displayText is "#NULL!" then return true
  if displayText is "#CALC!" then return true
  if displayText is "#FIELD!" then return true
  if displayText is "#UNKNOWN!" then return true
  if displayText is "#GETTING_DATA" then return true
  return false
end isExcelErrorDisplayText
`
}
