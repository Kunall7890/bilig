import { strFromU8, strToU8, zipSync, type Unzipped } from 'fflate'

const xlsxWorksheetXmlPathPattern = /^xl\/worksheets\/[^/]+\.xml$/u
const xmlAttributePattern = /\s([A-Za-z_:][\w:.-]*)=("[^"]*"|'[^']*')/gu

function forEachXmlAttribute(tag: string, visit: (name: string, value: string) => boolean | void): void {
  xmlAttributePattern.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = xmlAttributePattern.exec(tag)) !== null) {
    const name = match[1]
    const quotedValue = match[2]
    if (!name || !quotedValue) {
      continue
    }
    if (visit(name, quotedValue.slice(1, -1)) === false) {
      return
    }
  }
}

function cellAttributeValues(tag: string): Map<string, string> {
  const attributes = new Map<string, string>()
  forEachXmlAttribute(tag, (name, value) => {
    attributes.set(name, value)
  })
  return attributes
}

function readSheetDefaultRowHeight(sheetXml: string): string | undefined {
  const sheetFormatTag = /<sheetFormatPr\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(sheetXml)?.[0]
  if (!sheetFormatTag) {
    return undefined
  }
  return cellAttributeValues(sheetFormatTag).get('defaultRowHeight')
}

function isStyleOnlyBlankCellXml(cellXml: string): boolean {
  const openingTag = /^<c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(cellXml)?.[0]
  if (!openingTag) {
    return false
  }
  if (!styleOnlyBlankCellAttributesMatch(openingTag)) {
    return false
  }
  if (openingTag.endsWith('/>')) {
    return true
  }
  return cellXml.slice(openingTag.length, -'</c>'.length).trim().length === 0
}

function styleOnlyBlankCellAttributesMatch(openingTag: string): boolean {
  let sawAddress = false
  let sawStyle = false
  let sawZeroType = false
  let attributeCount = 0
  let unsupported = false
  forEachXmlAttribute(openingTag, (name, value) => {
    attributeCount += 1
    if (name === 'r') {
      sawAddress = true
      return
    }
    if (name === 's') {
      sawStyle = true
      return
    }
    if (name === 't' && value === 'z') {
      sawZeroType = true
      return
    }
    unsupported = true
    return false
  })
  return !unsupported && sawAddress && sawStyle && (attributeCount === 2 || (attributeCount === 3 && sawZeroType))
}

function isNoOpEmptyRowXml(rowXml: string, defaultRowHeight: string | undefined): boolean {
  const openingTag = /^<row\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(rowXml)?.[0]
  if (!openingTag) {
    return false
  }
  if (/<c\b/u.test(rowXml)) {
    return false
  }
  const attributes = cellAttributeValues(openingTag)
  for (const name of attributes.keys()) {
    if (name === 'r' || name === 'spans') {
      continue
    }
    if (name === 'x14ac:dyDescent') {
      continue
    }
    if (name === 'customHeight' && attributes.get(name) === '1' && attributes.get('ht') === defaultRowHeight) {
      continue
    }
    if (name === 'ht' && attributes.get(name) === defaultRowHeight) {
      continue
    }
    return false
  }
  return true
}

function stripNoOpEmptyRows(sheetXml: string): string {
  const defaultRowHeight = readSheetDefaultRowHeight(sheetXml)
  return sheetXml.replace(/<row\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<row\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/row>/gu, (rowXml) =>
    isNoOpEmptyRowXml(rowXml, defaultRowHeight) ? '' : rowXml,
  )
}

function stripStyleOnlyBlankCells(sheetXml: string): string {
  const withoutBlankCells = sheetXml.replace(/<c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<c\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/c>/gu, (cellXml) =>
    isStyleOnlyBlankCellXml(cellXml) ? '' : cellXml,
  )
  return stripNoOpEmptyRows(withoutBlankCells)
}

function stripWorkbookWorksheetXml(data: Uint8Array, zip: Unzipped, stripWorksheet: (sheetXml: string) => string): Uint8Array {
  let changed = false
  for (const path of Object.keys(zip)) {
    if (!xlsxWorksheetXmlPathPattern.test(path)) {
      continue
    }
    const worksheetBytes = zip[path]
    if (!worksheetBytes) {
      continue
    }
    const worksheetXml = strFromU8(worksheetBytes)
    const strippedWorksheetXml = stripWorksheet(worksheetXml)
    if (strippedWorksheetXml === worksheetXml) {
      continue
    }
    zip[path] = strToU8(strippedWorksheetXml)
    changed = true
  }
  return changed ? zipSync(zip) : data
}

export function stripNoOpEmptyRowsFromXlsx(data: Uint8Array, zip: Unzipped): Uint8Array {
  return stripWorkbookWorksheetXml(data, zip, stripNoOpEmptyRows)
}

export function stripStyleOnlyBlankCellsForSheetJs(data: Uint8Array, zip: Unzipped): Uint8Array {
  return stripWorkbookWorksheetXml(data, zip, stripStyleOnlyBlankCells)
}
