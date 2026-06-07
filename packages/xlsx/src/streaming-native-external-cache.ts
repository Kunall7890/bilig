import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { decodeCellAddress } from './address.js'
import type { NativeFormulaCell, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'
import { readXmlAttribute } from './xml.js'
import { getZipText, normalizeZipPath, type XlsxZipEntries } from './zip-reader.js'

interface ExternalReferenceScanTarget {
  readonly bookIndex: number
  readonly sheetName: string
  readonly rows: Set<number>
}

const externalReferencePattern =
  /'\[([1-9][0-9]*)\]((?:[^']|'')+)'!((?:\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(?:\$?[A-Za-z]{1,3}\$?[0-9]+))?)|\[([1-9][0-9]*)\]([A-Za-z_][A-Za-z0-9_.]*)!((?:\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(?:\$?[A-Za-z]{1,3}\$?[0-9]+))?)/gu
const relationshipPattern = /<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const emptyCellValue: CellValue = Object.freeze({ tag: ValueTag.Empty })

export function normalizeExternalWorkbookReferences(formula: string): string {
  externalReferencePattern.lastIndex = 0
  return formula.replace(
    externalReferencePattern,
    (
      _match,
      quotedBookIndex: string | undefined,
      quotedSheetName: string | undefined,
      quotedRef: string | undefined,
      unquotedBookIndex: string | undefined,
      unquotedSheetName: string | undefined,
      unquotedRef: string | undefined,
    ) => {
      const bookIndex = Number(quotedBookIndex ?? unquotedBookIndex)
      const sheetName = quotedSheetName ? quotedSheetName.replaceAll("''", "'") : (unquotedSheetName ?? '')
      const ref = quotedRef ?? unquotedRef ?? ''
      return `${quoteFormulaSheetName(externalReferenceAlias(bookIndex, sheetName))}!${ref}`
    },
  )
}

export function readStreamingNativeExternalCachedRowsByAlias(
  zip: XlsxZipEntries,
  sheetScans: ReadonlyMap<string, SheetScanState>,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>> {
  const targetsByAlias = collectExternalReferenceScanTargets(sheetScans, resolveFormulaSource)
  if (targetsByAlias.size === 0) {
    return new Map()
  }
  const externalLinkPathsByBookIndex = readExternalLinkPathsByBookIndex(zip)
  const targetsByBookIndex = new Map<number, ExternalReferenceScanTarget[]>()
  for (const target of targetsByAlias.values()) {
    const targets = targetsByBookIndex.get(target.bookIndex) ?? []
    targets.push(target)
    targetsByBookIndex.set(target.bookIndex, targets)
  }
  const output = new Map<string, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  for (const [bookIndex, targets] of targetsByBookIndex.entries()) {
    const externalLinkPath = externalLinkPathsByBookIndex.get(bookIndex)
    if (!externalLinkPath) {
      continue
    }
    const externalLinkXml = getZipText(zip, externalLinkPath)
    if (!externalLinkXml) {
      continue
    }
    for (const [alias, rows] of parseExternalLinkCachedRows(bookIndex, externalLinkXml, targets).entries()) {
      output.set(alias, rows)
    }
  }
  return output
}

export function isStreamingNativeExternalReferenceAlias(sheetName: string): boolean {
  return sheetName.startsWith('__bilig_external_')
}

function collectExternalReferenceScanTargets(
  sheetScans: ReadonlyMap<string, SheetScanState>,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): ReadonlyMap<string, ExternalReferenceScanTarget> {
  const targets = new Map<string, ExternalReferenceScanTarget>()
  for (const scan of sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      let formula
      try {
        formula = resolveFormulaSource(scan, cell)
      } catch {
        continue
      }
      for (const reference of extractExternalFormulaReferences(formula)) {
        const alias = externalReferenceAlias(reference.bookIndex, reference.sheetName)
        const target = targets.get(alias) ?? { bookIndex: reference.bookIndex, sheetName: reference.sheetName, rows: new Set<number>() }
        for (let row = reference.startRow; row <= reference.endRow; row += 1) {
          target.rows.add(row)
        }
        targets.set(alias, target)
      }
    }
  }
  return targets
}

function extractExternalFormulaReferences(formula: string): readonly {
  readonly bookIndex: number
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
}[] {
  externalReferencePattern.lastIndex = 0
  return [...formula.matchAll(externalReferencePattern)].flatMap((match) => {
    const bookIndex = Number(match[1] ?? match[4])
    const sheetName = match[2] ? match[2].replaceAll("''", "'") : (match[5] ?? '')
    const ref = match[3] ?? match[6] ?? ''
    if (!Number.isSafeInteger(bookIndex) || bookIndex < 1 || sheetName.length === 0 || ref.length === 0) {
      return []
    }
    const [startRef, endRef] = ref.split(':')
    if (!startRef) {
      return []
    }
    let start
    let end
    try {
      start = decodeCellAddress(startRef.replaceAll('$', ''))
      end = decodeCellAddress((endRef ?? startRef).replaceAll('$', ''))
    } catch {
      return []
    }
    return [
      {
        bookIndex,
        sheetName,
        startRow: Math.min(start.r, end.r),
        endRow: Math.max(start.r, end.r),
      },
    ]
  })
}

function readExternalLinkPathsByBookIndex(zip: XlsxZipEntries): ReadonlyMap<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
  if (!workbookXml || !workbookRelationshipsXml) {
    return new Map()
  }
  const externalLinkTargetsByRelationshipId = new Map<string, string>()
  for (const match of workbookRelationshipsXml.matchAll(relationshipPattern)) {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    if (id && target && type && (type === externalLinkRelationshipType || type.endsWith('/externalLink'))) {
      externalLinkTargetsByRelationshipId.set(id, resolveTargetPath('xl/workbook.xml', target))
    }
  }
  const paths = new Map<number, string>()
  let bookIndex = 1
  for (const match of workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?externalReference\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)) {
    const relationshipId = readXmlAttribute(match[0], 'r:id') ?? readXmlAttribute(match[0], 'id')
    const target = relationshipId ? externalLinkTargetsByRelationshipId.get(relationshipId) : undefined
    if (target) {
      paths.set(bookIndex, target)
    }
    bookIndex += 1
  }
  return paths
}

function parseExternalLinkCachedRows(
  bookIndex: number,
  externalLinkXml: string,
  targets: readonly ExternalReferenceScanTarget[],
): ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>> {
  const sheetNames = [...externalLinkXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheetName\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap(
    (match) => {
      const name = readXmlAttribute(match[0], 'val')
      return name ? [name] : []
    },
  )
  const targetsBySheetName = new Map(targets.map((target) => [target.sheetName, target]))
  const output = new Map<string, Map<number, Map<number, PendingCellValue>>>()
  for (const match of externalLinkXml.matchAll(
    /<((?:[A-Za-z_][\w.-]*:)?sheetData)\b((?:[^/>"']|"[^"]*"|'[^']*')*)(?:\/>|>([\s\S]*?)<\/\1>)/gu,
  )) {
    const sheetIdText = readXmlAttribute(match[2] ?? '', 'sheetId')
    const sheetId = sheetIdText === null ? Number.NaN : Number(sheetIdText)
    const sheetName = Number.isSafeInteger(sheetId) && sheetId >= 0 ? sheetNames[sheetId] : undefined
    const target = sheetName ? targetsBySheetName.get(sheetName) : undefined
    if (!target || sheetName === undefined) {
      continue
    }
    const rows = readExternalCachedRows(match[3] ?? '', target.rows)
    if (rows.size > 0) {
      output.set(externalReferenceAlias(bookIndex, sheetName), rows)
    }
  }
  return output
}

function readExternalCachedRows(xml: string, targetRows: ReadonlySet<number>): Map<number, Map<number, PendingCellValue>> {
  const output = new Map<number, Map<number, PendingCellValue>>()
  for (const rowMatch of xml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?row)\b((?:[^>"']|"[^"]*"|'[^']*')*)>([\s\S]*?)<\/\1>/gu)) {
    const rowText = readXmlAttribute(rowMatch[2] ?? '', 'r')
    const rowNumber = rowText === null ? Number.NaN : Number(rowText)
    if (!Number.isSafeInteger(rowNumber) || rowNumber < 1) {
      continue
    }
    const row = rowNumber - 1
    if (!targetRows.has(row)) {
      continue
    }
    const rowValues = new Map<number, PendingCellValue>()
    for (const cellMatch of (rowMatch[3] ?? '').matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?cell)\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?cell)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\2>/gu,
    )) {
      const cellXml = cellMatch[0]
      const openingTag = /<(?:[A-Za-z_][\w.-]*:)?cell\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(cellXml)?.[0]
      const addressText = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (!openingTag || !addressText) {
        continue
      }
      let address
      try {
        address = decodeCellAddress(addressText.replaceAll('$', ''))
      } catch {
        continue
      }
      rowValues.set(address.c, readExternalCachedCellValue(cellXml, openingTag))
    }
    output.set(row, rowValues)
  }
  return output
}

function readExternalCachedCellValue(cellXml: string, openingTag: string): PendingCellValue {
  const type = readXmlAttribute(openingTag, 't')
  const rawValue = readElementText(cellXml, 'v')
  if (rawValue === null) {
    return emptyCellValue
  }
  if (type === 'str' || type === 'inlineStr') {
    return stringCellValue(type === 'inlineStr' ? readTextRuns(cellXml) : decodeXmlText(rawValue))
  }
  if (type === 'b') {
    return { tag: ValueTag.Boolean, value: rawValue === '1' || rawValue.toLowerCase() === 'true' }
  }
  if (type === 'e') {
    return { tag: ValueTag.Error, code: errorCodeForText(decodeXmlText(rawValue)) }
  }
  const numeric = Number(rawValue)
  return Number.isFinite(numeric) ? { tag: ValueTag.Number, value: numeric } : stringCellValue(decodeXmlText(rawValue))
}

function resolveTargetPath(fromPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const baseParts = normalizeZipPath(fromPath).split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part.length === 0 || part === '.') {
      continue
    }
    if (part === '..') {
      baseParts.pop()
      continue
    }
    baseParts.push(part)
  }
  return normalizeZipPath(baseParts.join('/'))
}

function quoteFormulaSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function externalReferenceAlias(bookIndex: number, sheetName: string): string {
  let encodedSheetName = ''
  for (const character of sheetName) {
    encodedSheetName += /^[A-Za-z0-9_]$/u.test(character) ? character : `_u${character.codePointAt(0)!.toString(16).padStart(4, '0')}_`
  }
  return `__bilig_external_${String(bookIndex)}_${encodedSheetName}`
}

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function errorCodeForText(value: string): ErrorCode {
  switch (value.toUpperCase()) {
    case '#DIV/0!':
      return ErrorCode.Div0
    case '#REF!':
      return ErrorCode.Ref
    case '#VALUE!':
      return ErrorCode.Value
    case '#NAME?':
      return ErrorCode.Name
    case '#N/A':
      return ErrorCode.NA
    case '#NUM!':
      return ErrorCode.Num
    case '#FIELD!':
      return ErrorCode.Field
    case '#NULL!':
      return ErrorCode.Null
    default:
      return ErrorCode.Value
  }
}

function readTextRuns(xml: string): string {
  const runs = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)].map(
    (match) => decodeXmlText(match[1] ?? ''),
  )
  return runs.length > 0 ? runs.join('') : ''
}

function readElementText(xml: string, name: string): string | null {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  return (
    new RegExp(
      `<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b(?:[^>"']|"[^"]*"|'[^']*')*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escapedName}>`,
      'u',
    ).exec(xml)?.[1] ?? null
  )
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}
