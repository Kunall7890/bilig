import {
  decodeCellRange,
  encodeCellAddress,
  writeSimpleXlsxWorkbook,
  type SimpleXlsxCell,
  type SimpleXlsxDefinedName,
  type SimpleXlsxSheet,
  type SimpleXlsxStyle,
  type SimpleXlsxWorkbook,
} from '@bilig/xlsx'
import type { WorkbookRichTextCellSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { buildExportDefinedNames } from './xlsx-defined-names.js'
import { encodeFormulaForXlsx } from './xlsx-formula-translation.js'

const maxGeneratedRangeCells = 200_000

const supportedWorkbookMetadataKeys = new Set(['definedNames', 'styles', 'styleArtifacts', 'formats', 'tables'])
const supportedSheetMetadataKeys = new Set([
  'rows',
  'columns',
  'styleRanges',
  'styleArtifacts',
  'formatRanges',
  'merges',
  'filters',
  'hyperlinks',
  'richTextArtifacts',
])
const excelJsDefaultPageMarginsXml = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
const excelJsDefaultPageSetupXml =
  '<pageSetup orientation="portrait" horizontalDpi="4294967295" verticalDpi="4294967295" scale="100" fitToWidth="1" fitToHeight="1"/>'
const biligDefaultCorePropertiesXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Bilig</dc:creator><cp:lastModifiedBy>Bilig</cp:lastModifiedBy></cp:coreProperties>'
const biligDefaultAppPropertiesPattern =
  /^<\?xml version="1\.0" encoding="UTF-8" standalone="yes"\?><Properties xmlns="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/extended-properties" xmlns:vt="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/docPropsVTypes"><Application>Bilig<\/Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>[1-9][0-9]*<\/vt:i4><\/vt:variant><\/vt:vector><\/HeadingPairs><\/Properties>$/u

interface SimpleRichTextPatch {
  readonly sharedStringIndex?: number
  readonly inlineStringXml?: string
}

interface SimpleRichTextExportPlan {
  readonly cellsBySheet: ReadonlyMap<WorkbookSnapshot['sheets'][number], ReadonlyMap<string, SimpleRichTextPatch>>
  readonly sharedStringsXml?: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasDefaultExcelJsDocumentPropertyArtifacts(coreXml: string, appXml: string): boolean {
  return (
    coreXml.includes('<dc:creator>Unknown</dc:creator>') &&
    coreXml.includes('<dc:title></dc:title>') &&
    coreXml.includes('<dc:subject></dc:subject>') &&
    coreXml.includes('<dc:description></dc:description>') &&
    coreXml.includes('<cp:keywords></cp:keywords>') &&
    coreXml.includes('<cp:category></cp:category>') &&
    coreXml.includes('<cp:lastModifiedBy>Unknown</cp:lastModifiedBy>') &&
    appXml.includes('<Application>Microsoft Excel</Application>')
  )
}

function hasDefaultBiligDocumentPropertyArtifacts(coreXml: string, appXml: string): boolean {
  return coreXml === biligDefaultCorePropertiesXml && biligDefaultAppPropertiesPattern.test(appXml)
}

function isIgnorableDefaultDocumentPropertyArtifacts(value: unknown): boolean {
  if (!isRecord(value)) {
    return false
  }
  if (!Object.keys(value).every((key) => key === 'core' || key === 'app')) {
    return false
  }
  const core = value['core']
  const app = value['app']
  if (!isRecord(core) || !isRecord(app) || typeof core['xml'] !== 'string' || typeof app['xml'] !== 'string') {
    return false
  }
  return (
    hasDefaultExcelJsDocumentPropertyArtifacts(core['xml'], app['xml']) || hasDefaultBiligDocumentPropertyArtifacts(core['xml'], app['xml'])
  )
}

function isIgnorableDefaultSheetFormat(value: unknown): boolean {
  if (!isRecord(value)) {
    return false
  }
  return Object.entries(value).every(([key, entry]) => {
    if (key === 'defaultRowHeight') {
      return entry === 15
    }
    if (key === 'outlineLevelRow' || key === 'outlineLevelCol') {
      return entry === 0
    }
    return false
  })
}

function isIgnorableExcelJsDefaultPrintPageSetup(value: unknown): boolean {
  return (
    isRecord(value) &&
    Object.keys(value).every((key) => key === 'pageMarginsXml' || key === 'pageSetupXml') &&
    value['pageMarginsXml'] === excelJsDefaultPageMarginsXml &&
    value['pageSetupXml'] === excelJsDefaultPageSetupXml
  )
}

function isIgnorableWorkbookMetadata(key: string, value: unknown): boolean {
  if (key === 'formulaAudit') {
    return value !== undefined
  }
  if (key === 'volatileContext') {
    return value !== undefined
  }
  if (key === 'calculationSettings') {
    return isRecord(value)
  }
  if (key === 'documentPropertyArtifacts') {
    return isIgnorableDefaultDocumentPropertyArtifacts(value)
  }
  return false
}

function isIgnorableSheetMetadata(key: string, value: unknown): boolean {
  if (key === 'sheetFormatPr') {
    return isIgnorableDefaultSheetFormat(value)
  }
  if (key === 'printPageSetup') {
    return isIgnorableExcelJsDefaultPrintPageSetup(value)
  }
  return false
}

function hasUnsupportedMetadata(
  metadata: object | undefined,
  supportedKeys: ReadonlySet<string>,
  isIgnorableMetadata: (key: string, value: unknown) => boolean,
): boolean {
  if (!metadata) {
    return false
  }
  return Object.entries(metadata).some(([key, value]) => {
    if (supportedKeys.has(key) || isIgnorableMetadata(key, value)) {
      return false
    }
    if (value === undefined || value === null) {
      return false
    }
    if (Array.isArray(value)) {
      return value.length > 0
    }
    if (typeof value === 'object') {
      return Object.keys(value).length > 0
    }
    return true
  })
}

function simpleRangeCellCount(startAddress: string, endAddress: string): number | null {
  try {
    const range = decodeCellRange(`${startAddress}:${endAddress}`)
    return (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1)
  } catch {
    return null
  }
}

function canMaterializeRanges(sheet: WorkbookSnapshot['sheets'][number]): boolean {
  let generatedCells = 0
  for (const range of [...(sheet.metadata?.styleRanges ?? []), ...(sheet.metadata?.formatRanges ?? [])]) {
    if (range.range.sheetName !== sheet.name) {
      return false
    }
    const count = simpleRangeCellCount(range.range.startAddress, range.range.endAddress)
    if (count === null) {
      return false
    }
    generatedCells += count
    if (generatedCells > maxGeneratedRangeCells) {
      return false
    }
  }
  return true
}

function styleArtifactAddressCount(sheet: WorkbookSnapshot['sheets'][number]): number {
  return (sheet.metadata?.styleArtifacts?.cellStyleIndexes.length ?? 0) + (sheet.metadata?.styleArtifacts?.blankCellAddresses?.length ?? 0)
}

function canMaterializeStyleArtifacts(snapshot: WorkbookSnapshot): boolean {
  const workbookStyleArtifacts = snapshot.workbook.metadata?.styleArtifacts
  const hasSheetStyleArtifacts = snapshot.sheets.some((sheet) => styleArtifactAddressCount(sheet) > 0)
  if (!workbookStyleArtifacts && !hasSheetStyleArtifacts) {
    return true
  }
  if (!workbookStyleArtifacts?.stylesXml || workbookStyleArtifacts.theme) {
    return false
  }
  if ((snapshot.workbook.metadata?.styles?.length ?? 0) > 0 || (snapshot.workbook.metadata?.formats?.length ?? 0) > 0) {
    return false
  }
  let generatedCells = 0
  for (const sheet of snapshot.sheets) {
    if ((sheet.metadata?.styleRanges?.length ?? 0) > 0 || (sheet.metadata?.formatRanges?.length ?? 0) > 0) {
      return false
    }
    if (sheet.cells.some((cell) => cell.format?.trim())) {
      return false
    }
    for (const entry of sheet.metadata?.styleArtifacts?.cellStyleIndexes ?? []) {
      if (!Number.isSafeInteger(entry.styleIndex) || entry.styleIndex < 0 || !isValidCellAddress(entry.address)) {
        return false
      }
    }
    for (const address of sheet.metadata?.styleArtifacts?.blankCellAddresses ?? []) {
      if (!isValidCellAddress(address)) {
        return false
      }
    }
    generatedCells += styleArtifactAddressCount(sheet)
    if (generatedCells > maxGeneratedRangeCells) {
      return false
    }
  }
  return true
}

function richTextArtifactCells(sheet: WorkbookSnapshot['sheets'][number]): readonly WorkbookRichTextCellSnapshot[] {
  return sheet.metadata?.richTextArtifacts?.cells ?? []
}

function richTextArtifactXmlMatchesStorage(artifact: WorkbookRichTextCellSnapshot): boolean {
  if (artifact.storage === 'sharedString') {
    return /^\s*<((?:[A-Za-z_][\w.-]*:)?si)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)\s*$/u.test(artifact.xml)
  }
  return /^\s*<((?:[A-Za-z_][\w.-]*:)?is)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)\s*$/u.test(artifact.xml)
}

function richTextSourceStringCellsByAddress(
  sheet: WorkbookSnapshot['sheets'][number],
): ReadonlyMap<string, WorkbookSnapshot['sheets'][number]['cells'][number]> {
  return new Map(
    sheet.cells.flatMap((cell) => (typeof cell.value === 'string' && !cell.formula ? ([[cell.address, cell]] as const) : ([] as const))),
  )
}

function matchingRichTextArtifacts(sheet: WorkbookSnapshot['sheets'][number]): readonly WorkbookRichTextCellSnapshot[] {
  const cellsByAddress = richTextSourceStringCellsByAddress(sheet)
  return richTextArtifactCells(sheet).filter((artifact) => {
    const cell = cellsByAddress.get(artifact.address)
    return cell?.value === artifact.text
  })
}

function canMaterializeRichTextArtifacts(snapshot: WorkbookSnapshot): boolean {
  let generatedCells = 0
  for (const sheet of snapshot.sheets) {
    for (const artifact of richTextArtifactCells(sheet)) {
      if (!isValidCellAddress(artifact.address)) {
        return false
      }
    }
    const artifacts = matchingRichTextArtifacts(sheet)
    generatedCells += artifacts.length
    if (generatedCells > maxGeneratedRangeCells) {
      return false
    }
    for (const artifact of artifacts) {
      if (!richTextArtifactXmlMatchesStorage(artifact)) {
        return false
      }
    }
  }
  return true
}

function hasOnlySimpleFilters(sheet: WorkbookSnapshot['sheets'][number]): boolean {
  return (sheet.metadata?.filters ?? []).every((filter) => !filter.criteria || filter.criteria.length === 0)
}

function hasOnlySimpleHyperlinks(sheet: WorkbookSnapshot['sheets'][number]): boolean {
  return (sheet.metadata?.hyperlinks ?? []).every(
    (hyperlink) => hyperlink.sheetName === sheet.name && hyperlink.target.trim().length > 0 && isValidCellAddress(hyperlink.address),
  )
}

function canUseBiligSimpleWriter(snapshot: WorkbookSnapshot): boolean {
  if (snapshot.sheets.length === 0) {
    return false
  }
  if (hasUnsupportedMetadata(snapshot.workbook.metadata, supportedWorkbookMetadataKeys, isIgnorableWorkbookMetadata)) {
    return false
  }
  if (!canMaterializeStyleArtifacts(snapshot)) {
    return false
  }
  if (!canMaterializeRichTextArtifacts(snapshot)) {
    return false
  }
  return snapshot.sheets.every(
    (sheet) =>
      !hasUnsupportedMetadata(sheet.metadata, supportedSheetMetadataKeys, isIgnorableSheetMetadata) &&
      hasOnlySimpleFilters(sheet) &&
      hasOnlySimpleHyperlinks(sheet) &&
      canMaterializeRanges(sheet),
  )
}

const invalidExportSheetNameCharacters = ['[', ']', ':', '*', '?', '/', '\\'] as const

function normalizeExportSheetName(name: string, order: number, usedNames: Set<string>): string {
  let sanitized = name
  for (const character of invalidExportSheetNameCharacters) {
    sanitized = sanitized.split(character).join(' ')
  }
  const baseName = sanitized.length > 0 ? sanitized : `Sheet${order + 1}`
  let candidate = baseName.slice(0, 31)
  candidate = candidate.length > 0 ? candidate : `Sheet${order + 1}`
  let suffix = 1
  while (usedNames.has(candidate)) {
    const suffixText = ` ${String(suffix)}`
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`
    suffix += 1
  }
  usedNames.add(candidate)
  return candidate
}

function rangeAddresses(startAddress: string, endAddress: string): string[] {
  const range = decodeCellRange(`${startAddress}:${endAddress}`)
  const output: string[] = []
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      output.push(encodeCellAddress({ r: row, c: col }))
    }
  }
  return output
}

function readCellCoordinates(address: string): { readonly row: number; readonly col: number } {
  const decoded = decodeCellRange(address)
  return { row: decoded.s.r, col: decoded.s.c }
}

function isValidCellAddress(address: string): boolean {
  try {
    readCellCoordinates(address)
    return true
  } catch {
    return false
  }
}

function ensureCell(cellsByAddress: Map<string, SimpleXlsxCell>, address: string): SimpleXlsxCell {
  const existing = cellsByAddress.get(address)
  if (existing) {
    return existing
  }
  const coordinates = readCellCoordinates(address)
  const cell: SimpleXlsxCell = { address, row: coordinates.row, col: coordinates.col }
  cellsByAddress.set(address, cell)
  return cell
}

function replaceCell(cellsByAddress: Map<string, SimpleXlsxCell>, cell: SimpleXlsxCell, patch: Partial<SimpleXlsxCell>): void {
  cellsByAddress.set(cell.address, {
    ...cell,
    ...patch,
  })
}

function simpleSheetCells(
  sheet: WorkbookSnapshot['sheets'][number],
  formatCodesById: ReadonlyMap<string, string>,
  richTextPatches: ReadonlyMap<string, SimpleRichTextPatch> | undefined,
): readonly SimpleXlsxCell[] {
  const cellsByAddress = new Map<string, SimpleXlsxCell>()
  for (const cell of sheet.cells) {
    const decoded = readCellCoordinates(cell.address)
    const richTextPatch = richTextPatches?.get(cell.address)
    cellsByAddress.set(cell.address, {
      address: cell.address,
      row: decoded.row,
      col: decoded.col,
      ...(cell.value !== undefined ? { value: cell.value } : {}),
      ...(cell.formula ? { formula: encodeFormulaForXlsx(cell.formula.replace(/^=/u, '')) } : {}),
      ...(cell.format?.trim() ? { numberFormat: cell.format.trim() } : {}),
      ...richTextPatch,
    })
  }
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    for (const address of rangeAddresses(styleRange.range.startAddress, styleRange.range.endAddress)) {
      const cell = ensureCell(cellsByAddress, address)
      replaceCell(cellsByAddress, cell, { styleId: styleRange.styleId })
    }
  }
  for (const entry of sheet.metadata?.styleArtifacts?.cellStyleIndexes ?? []) {
    const cell = ensureCell(cellsByAddress, entry.address)
    replaceCell(cellsByAddress, cell, { styleIndex: entry.styleIndex })
  }
  for (const address of sheet.metadata?.styleArtifacts?.blankCellAddresses ?? []) {
    ensureCell(cellsByAddress, address)
  }
  for (const formatRange of sheet.metadata?.formatRanges ?? []) {
    const format = formatCodesById.get(formatRange.formatId)?.trim()
    if (!format || format === 'General') {
      continue
    }
    for (const address of rangeAddresses(formatRange.range.startAddress, formatRange.range.endAddress)) {
      const cell = ensureCell(cellsByAddress, address)
      replaceCell(cellsByAddress, cell, { numberFormat: format })
    }
  }
  return [...cellsByAddress.values()].toSorted((left, right) => left.row - right.row || left.col - right.col)
}

function simpleStyles(snapshot: WorkbookSnapshot): readonly SimpleXlsxStyle[] | undefined {
  if (snapshot.workbook.metadata?.styleArtifacts?.stylesXml) {
    return undefined
  }
  const styles = snapshot.workbook.metadata?.styles
  return styles && styles.length > 0 ? styles : undefined
}

function simpleDefinedNames(
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
  exportSheetIndexesByOriginalName: ReadonlyMap<string, number>,
): readonly SimpleXlsxDefinedName[] | undefined {
  const definedNames = buildExportDefinedNames(
    snapshot.workbook.metadata?.definedNames,
    exportSheetNamesByOriginalName,
    exportSheetIndexesByOriginalName,
  )
  if (!definedNames) {
    return undefined
  }
  const output: SimpleXlsxDefinedName[] = []
  for (const definedName of definedNames) {
    if (!definedName.Name || !definedName.Ref) {
      continue
    }
    output.push({
      name: definedName.Name,
      formula: definedName.Ref,
      ...(definedName.Sheet !== undefined ? { localSheetIndex: definedName.Sheet } : {}),
    })
  }
  return output.length > 0 ? output : undefined
}

function sharedStringsXml(items: readonly string[]): string | undefined {
  if (items.length === 0) {
    return undefined
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${String(items.length)}" uniqueCount="${String(
      items.length,
    )}">`,
    ...items,
    '</sst>',
  ].join('')
}

function buildRichTextExportPlan(snapshot: WorkbookSnapshot): SimpleRichTextExportPlan {
  const cellsBySheet = new Map<WorkbookSnapshot['sheets'][number], Map<string, SimpleRichTextPatch>>()
  const sharedStringItems: string[] = []
  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    const cellPatches = new Map<string, SimpleRichTextPatch>()
    for (const artifact of matchingRichTextArtifacts(sheet)) {
      if (artifact.storage === 'sharedString') {
        cellPatches.set(artifact.address, { sharedStringIndex: sharedStringItems.length })
        sharedStringItems.push(artifact.xml)
        continue
      }
      cellPatches.set(artifact.address, { inlineStringXml: artifact.xml })
    }
    if (cellPatches.size > 0) {
      cellsBySheet.set(sheet, cellPatches)
    }
  }
  const outputSharedStringsXml = sharedStringsXml(sharedStringItems)
  return {
    cellsBySheet,
    ...(outputSharedStringsXml !== undefined ? { sharedStringsXml: outputSharedStringsXml } : {}),
  }
}

function simpleSheet(
  sheet: WorkbookSnapshot['sheets'][number],
  exportName: string,
  formatCodesById: ReadonlyMap<string, string>,
  richTextPatches: ReadonlyMap<string, SimpleRichTextPatch> | undefined,
): SimpleXlsxSheet {
  const rows = sheet.metadata?.rows?.map((row) => ({
    index: row.index,
    ...(typeof row.size === 'number' ? { size: row.size } : {}),
    ...(row.hidden === true || row.filterHidden === true ? { hidden: true } : {}),
  }))
  const columns = sheet.metadata?.columns?.map((column) => ({
    index: column.index,
    ...(typeof column.size === 'number' ? { size: column.size } : {}),
    ...(column.hidden === true ? { hidden: true } : {}),
  }))
  const merges = sheet.metadata?.merges?.map((merge) => ({
    startAddress: merge.startAddress,
    endAddress: merge.endAddress,
  }))
  const autoFilters = sheet.metadata?.filters?.map((filter) => ({
    startAddress: filter.startAddress,
    endAddress: filter.endAddress,
  }))
  const hyperlinks = sheet.metadata?.hyperlinks
    ?.filter((hyperlink) => hyperlink.sheetName === sheet.name && hyperlink.target.trim().length > 0)
    .map((hyperlink) => {
      const output: { address: string; target: string; tooltip?: string; display?: string } = {
        address: hyperlink.address,
        target: hyperlink.target,
      }
      if (hyperlink.tooltip) {
        output.tooltip = hyperlink.tooltip
      }
      if (hyperlink.display) {
        output.display = hyperlink.display
      }
      return output
    })
  return {
    name: exportName,
    cells: simpleSheetCells(sheet, formatCodesById, richTextPatches),
    ...(rows && rows.length > 0 ? { rows } : {}),
    ...(columns && columns.length > 0 ? { columns } : {}),
    ...(merges && merges.length > 0 ? { merges } : {}),
    ...(autoFilters && autoFilters.length > 0 ? { autoFilters } : {}),
    ...(hyperlinks && hyperlinks.length > 0 ? { hyperlinks } : {}),
  }
}

function simpleWorkbook(snapshot: WorkbookSnapshot): SimpleXlsxWorkbook {
  const usedNames = new Set<string>()
  const exportSheetNamesByOriginalName = new Map<string, string>()
  const exportSheetIndexesByOriginalName = new Map<string, number>()
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  for (const sheet of orderedSheets) {
    const exportSheetName = normalizeExportSheetName(sheet.name, sheet.order, usedNames)
    exportSheetNamesByOriginalName.set(sheet.name, exportSheetName)
    exportSheetIndexesByOriginalName.set(sheet.name, exportSheetIndexesByOriginalName.size)
  }
  const formatCodesById = new Map((snapshot.workbook.metadata?.formats ?? []).map((format) => [format.id, format.code]))
  const richTextPlan = buildRichTextExportPlan(snapshot)
  const styles = simpleStyles(snapshot)
  const definedNames = simpleDefinedNames(snapshot, exportSheetNamesByOriginalName, exportSheetIndexesByOriginalName)
  const stylesXml = snapshot.workbook.metadata?.styleArtifacts?.stylesXml
  return {
    sheets: orderedSheets.map((sheet) =>
      simpleSheet(
        sheet,
        exportSheetNamesByOriginalName.get(sheet.name) ?? sheet.name,
        formatCodesById,
        richTextPlan.cellsBySheet.get(sheet),
      ),
    ),
    ...(styles ? { styles } : {}),
    ...(stylesXml ? { stylesXml } : {}),
    ...(richTextPlan.sharedStringsXml ? { sharedStringsXml: richTextPlan.sharedStringsXml } : {}),
    ...(definedNames ? { definedNames } : {}),
  }
}

export function tryExportBiligSimpleXlsx(snapshot: WorkbookSnapshot): Uint8Array | null {
  return canUseBiligSimpleWriter(snapshot) ? writeSimpleXlsxWorkbook(simpleWorkbook(snapshot)) : null
}
