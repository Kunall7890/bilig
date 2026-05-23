import * as XLSX from 'xlsx'

import { escapeXmlAttribute } from './xlsx-export-xml.js'

interface MissingCellXml {
  readonly address: string
  readonly xml: string
}

const worksheetRowElementPattern = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/gu
const worksheetRowOpeningTagPattern = /^<row\b[^>]*(?:\/>|>)/u
const worksheetCellElementPattern = /<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gu
const worksheetCellOpeningTagPattern = /^<c\b[^>]*(?:\/>|>)/u

function rowNumberForAddress(address: string): number {
  return XLSX.utils.decode_cell(address).r + 1
}

function columnIndexForAddress(address: string): number {
  return XLSX.utils.decode_cell(address).c
}

function updateRangeWithAddress(range: XLSX.Range | null, address: string): XLSX.Range | null {
  try {
    const decoded = XLSX.utils.decode_cell(address)
    return range
      ? {
          s: {
            r: Math.min(range.s.r, decoded.r),
            c: Math.min(range.s.c, decoded.c),
          },
          e: {
            r: Math.max(range.e.r, decoded.r),
            c: Math.max(range.e.c, decoded.c),
          },
        }
      : {
          s: { r: decoded.r, c: decoded.c },
          e: { r: decoded.r, c: decoded.c },
        }
  } catch {
    return range
  }
}

function existingWorksheetDimensionRange(sheetXml: string): XLSX.Range | null {
  const dimensionXml = findXmlElement(sheetXml, 'dimension')?.xml
  const ref = dimensionXml ? /\bref="([^"]+)"/u.exec(dimensionXml)?.[1] : undefined
  if (!ref) {
    return null
  }
  try {
    return XLSX.utils.decode_range(ref)
  } catch {
    return null
  }
}

function setWorksheetDimensionRef(sheetXml: string, range: XLSX.Range): string {
  const ref = escapeXmlAttribute(XLSX.utils.encode_range(range))
  const dimensionElement = findXmlElement(sheetXml, 'dimension')
  if (dimensionElement) {
    const dimensionXml = /\bref="[^"]*"/u.test(dimensionElement.xml)
      ? dimensionElement.xml.replace(/\bref="[^"]*"/u, `ref="${ref}"`)
      : `<dimension ref="${ref}"/>`
    return `${sheetXml.slice(0, dimensionElement.start)}${dimensionXml}${sheetXml.slice(dimensionElement.end)}`
  }
  return sheetXml.replace(/<worksheet\b[^>]*>/u, (openingTag) => `${openingTag}<dimension ref="${ref}"/>`)
}

function expandWorksheetDimensionForMissingCells(sheetXml: string, cells: readonly MissingCellXml[]): string {
  let range = existingWorksheetDimensionRange(sheetXml)
  const originalRange = range
  for (const cell of cells) {
    range = updateRangeWithAddress(range, cell.address)
  }
  if (!range || (originalRange && XLSX.utils.encode_range(range) === XLSX.utils.encode_range(originalRange))) {
    return sheetXml
  }
  return setWorksheetDimensionRef(sheetXml, range)
}

function rowCellXml(rowCells: readonly MissingCellXml[]): string[] {
  return rowCells
    .toSorted((left, right) => columnIndexForAddress(left.address) - columnIndexForAddress(right.address))
    .map((cell) => cell.xml)
}

function buildRowsXml(entries: readonly [number, MissingCellXml[]][]): string {
  return entries.map(([rowNumber, rowCells]) => `<row r="${String(rowNumber)}">${rowCellXml(rowCells).join('')}</row>`).join('')
}

function rowNumberForRowXml(rowXml: string): number | null {
  const rowTag = worksheetRowOpeningTagPattern.exec(rowXml)?.[0]
  const rowNumber = rowTag ? Number(/\br="([0-9]+)"/u.exec(rowTag)?.[1] ?? NaN) : NaN
  return Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber : null
}

function appendCellsToRowXml(rowXml: string, cells: readonly MissingCellXml[]): string {
  if (rowXml.endsWith('/>')) {
    return rowXml.replace(/\/>$/u, `>${rowCellXml(cells).join('')}</row>`)
  }

  const openingTag = worksheetRowOpeningTagPattern.exec(rowXml)?.[0]
  if (!openingTag || !rowXml.endsWith('</row>')) {
    return rowXml.replace('</row>', `${rowCellXml(cells).join('')}</row>`)
  }

  const body = rowXml.slice(openingTag.length, -'</row>'.length)
  const existingCells: MissingCellXml[] = []
  let nonCellXml = ''
  let lastIndex = 0
  for (const match of body.matchAll(worksheetCellElementPattern)) {
    const cellXml = match[0]
    nonCellXml += body.slice(lastIndex, match.index)
    lastIndex = match.index + cellXml.length
    const cellOpeningTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = cellOpeningTag ? /\br="([^"]+)"/u.exec(cellOpeningTag)?.[1] : undefined
    if (!address) {
      nonCellXml += cellXml
      continue
    }
    existingCells.push({ address, xml: cellXml })
  }
  nonCellXml += body.slice(lastIndex)

  return `${openingTag}${rowCellXml([...existingCells, ...cells]).join('')}${nonCellXml}</row>`
}

function findXmlElement(xml: string, localName: string): { readonly start: number; readonly end: number; readonly xml: string } | null {
  const start = xml.indexOf(`<${localName}`)
  if (start === -1) {
    return null
  }
  const tagEnd = xml.indexOf('>', start + localName.length + 1)
  if (tagEnd === -1) {
    return null
  }
  if (xml[tagEnd - 1] === '/') {
    return { start, end: tagEnd + 1, xml: xml.slice(start, tagEnd + 1) }
  }
  const closeTag = `</${localName}>`
  const closeStart = xml.indexOf(closeTag, tagEnd + 1)
  return closeStart === -1 ? null : { start, end: closeStart + closeTag.length, xml: xml.slice(start, closeStart + closeTag.length) }
}

function addMissingCellsToSheetDataXml(sheetDataXml: string, rowEntries: readonly [number, MissingCellXml[]][]): string {
  const sheetDataOpeningTag = /^<sheetData\b[^>]*>/u.exec(sheetDataXml)?.[0]
  if (!sheetDataOpeningTag || !sheetDataXml.endsWith('</sheetData>')) {
    return sheetDataXml
  }

  const bodyStart = sheetDataOpeningTag.length
  const bodyEnd = sheetDataXml.length - '</sheetData>'.length
  const sheetDataBody = sheetDataXml.slice(bodyStart, bodyEnd)
  let outputBody = ''
  let lastIndex = 0
  let missingIndex = 0

  for (const match of sheetDataBody.matchAll(worksheetRowElementPattern)) {
    const rowXml = match[0]
    const rowNumber = rowNumberForRowXml(rowXml)
    outputBody += sheetDataBody.slice(lastIndex, match.index)
    if (rowNumber === null) {
      outputBody += rowXml
      lastIndex = match.index + rowXml.length
      continue
    }
    while (missingIndex < rowEntries.length && rowEntries[missingIndex]![0] < rowNumber) {
      outputBody += buildRowsXml([rowEntries[missingIndex]!])
      missingIndex += 1
    }
    if (missingIndex < rowEntries.length && rowEntries[missingIndex]![0] === rowNumber) {
      outputBody += appendCellsToRowXml(rowXml, rowEntries[missingIndex]![1])
      missingIndex += 1
    } else {
      outputBody += rowXml
    }
    lastIndex = match.index + rowXml.length
  }

  outputBody += sheetDataBody.slice(lastIndex)
  while (missingIndex < rowEntries.length) {
    outputBody += buildRowsXml([rowEntries[missingIndex]!])
    missingIndex += 1
  }

  return `${sheetDataOpeningTag}${outputBody}</sheetData>`
}

export function addMissingCellsToSheetXml(sheetXml: string, cells: readonly MissingCellXml[]): string {
  if (cells.length === 0) {
    return sheetXml
  }
  const byRow = new Map<number, MissingCellXml[]>()
  for (const cell of cells) {
    const rowNumber = rowNumberForAddress(cell.address)
    byRow.set(rowNumber, [...(byRow.get(rowNumber) ?? []), cell])
  }
  const rowEntries = [...byRow.entries()].toSorted(([left], [right]) => left - right)
  const sheetDataElement = findXmlElement(sheetXml, 'sheetData')
  const selfClosingSheetData = sheetDataElement ? /^<sheetData\b([^>]*)\/>$/u.exec(sheetDataElement.xml) : null
  if (sheetDataElement && selfClosingSheetData) {
    const updatedSheetDataXml = `<sheetData${selfClosingSheetData[1] ?? ''}>${buildRowsXml(rowEntries)}</sheetData>`
    const updatedSheetXml = `${sheetXml.slice(0, sheetDataElement.start)}${updatedSheetDataXml}${sheetXml.slice(sheetDataElement.end)}`
    return expandWorksheetDimensionForMissingCells(updatedSheetXml, cells)
  }

  if (sheetDataElement) {
    const updatedSheetXml = `${sheetXml.slice(0, sheetDataElement.start)}${addMissingCellsToSheetDataXml(
      sheetDataElement.xml,
      rowEntries,
    )}${sheetXml.slice(sheetDataElement.end)}`
    return expandWorksheetDimensionForMissingCells(updatedSheetXml, cells)
  }

  return expandWorksheetDimensionForMissingCells(
    sheetXml.replace('</worksheet>', `<sheetData>${buildRowsXml(rowEntries)}</sheetData></worksheet>`),
    cells,
  )
}

export function addMissingFormattedCells(
  sheetXml: string,
  cells: readonly { readonly address: string; readonly styleIndex: number }[],
): string {
  return addMissingCellsToSheetXml(
    sheetXml,
    cells.map((cell) => {
      const address = escapeXmlAttribute(cell.address)
      return {
        address: cell.address,
        xml: `<c r="${address}" s="${String(cell.styleIndex)}"/>`,
      }
    }),
  )
}

export function addMissingBlankCells(sheetXml: string, addresses: readonly string[]): string {
  return addMissingCellsToSheetXml(
    sheetXml,
    addresses.map((address) => ({
      address,
      xml: `<c r="${escapeXmlAttribute(address)}"/>`,
    })),
  )
}
