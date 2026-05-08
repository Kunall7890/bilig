import { strFromU8, strToU8, zipSync, type Unzipped } from 'fflate'

const xlsxWorksheetXmlPathPattern = /^xl\/worksheets\/[^/]+\.xml$/u

function isStyleOnlyBlankCellTag(tag: string): boolean {
  const attributes = [...tag.matchAll(/\s([A-Za-z_:][\w:.-]*)=(?:"[^"]*"|'[^']*')/gu)].map((match) => match[1])
  const attributeNames = new Set(attributes)
  return attributeNames.size === 2 && attributeNames.has('r') && attributeNames.has('s')
}

function stripStyleOnlyBlankCells(sheetXml: string): string {
  return sheetXml.replace(/<c\b[^>]*\/>/gu, (tag) => (isStyleOnlyBlankCellTag(tag) ? '' : tag))
}

export function stripStyleOnlyBlankCellsForSheetJs(data: Uint8Array, zip: Unzipped): Uint8Array {
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
    const strippedWorksheetXml = stripStyleOnlyBlankCells(worksheetXml)
    if (strippedWorksheetXml === worksheetXml) {
      continue
    }
    zip[path] = strToU8(strippedWorksheetXml)
    changed = true
  }
  return changed ? zipSync(zip) : data
}
