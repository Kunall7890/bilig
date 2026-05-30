import { WorkPaper, type RawCellContent, type WorkPaperCellAddress, type WorkPaperChange, type WorkPaperConfig } from '@bilig/headless'
import type { ImportedWorkbookDiagnostics, XlsxExternalWorkbookInput } from '@bilig/headless/xlsx'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'
import { ErrorCode, formatErrorCode, ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

export { WorkPaper } from '@bilig/headless'
export { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSnapshot>
type WorkbookSnapshot = Parameters<typeof WorkPaper.buildFromSnapshot>[0]
type WorkbookCalculationSettings = NonNullable<NonNullable<WorkbookSnapshot['workbook']['metadata']>['calculationSettings']>
interface FormulaErrorCache {
  readonly sheetName: string
  readonly address: string
  readonly value: string
}

export type XlsxFormulaRecalcCellValue = ReturnType<WorkPaperInstance['getCellValue']>

export interface XlsxFormulaRecalcEdit {
  readonly target: string
  readonly value: RawCellContent
}

export interface XlsxFormulaRecalcOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly config?: WorkPaperConfig
}

export interface XlsxFormulaRecalcResult {
  readonly xlsx: Uint8Array
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly WorkPaperChange[]
  readonly diagnostics?: ImportedWorkbookDiagnostics
}

export function recalculateXlsx(input: Uint8Array | ArrayBuffer | Buffer, options: XlsxFormulaRecalcOptions = {}): XlsxFormulaRecalcResult {
  const importOptions = options.externalWorkbooks
    ? { externalWorkbooks: options.externalWorkbooks, externalLinkCacheArtifactMode: 'replace-refreshed' as const }
    : {}
  const imported = importXlsx(toUint8Array(input), options.fileName ?? 'workbook.xlsx', importOptions)
  const originalCalculationSettings = imported.snapshot.workbook.metadata?.calculationSettings
  const workbook = WorkPaper.buildFromSnapshot(snapshotForFreshFormulaRecalculation(imported.snapshot), {
    evaluationTimeoutMs: 30_000,
    useColumnIndex: true,
    ...options.config,
  })

  try {
    const changes: WorkPaperChange[] = []
    for (const edit of options.edits ?? []) {
      changes.push(...workbook.setCellContents(parseQualifiedCellTarget(workbook, edit.target), edit.value))
    }
    changes.push(...workbook.rebuildAndRecalculate())

    const reads: Record<string, XlsxFormulaRecalcCellValue> = {}
    for (const target of options.reads ?? []) {
      reads[target] = workbook.getCellValue(parseQualifiedCellTarget(workbook, target))
    }

    const outputFormulaCaches = snapshotWithFormulaCachedValues(workbook, workbook.exportSnapshot())
    const outputSnapshot = outputFormulaCaches.snapshot
    const exportedXlsx = toUint8Array(
      exportXlsx(
        restoreOutputCalculationSettings(
          outputSnapshot,
          options.config?.calculationSettings === undefined ? originalCalculationSettings : undefined,
        ),
      ),
    )
    return {
      xlsx: addFormulaErrorCachesToXlsxBytes(exportedXlsx, outputSnapshot, outputFormulaCaches.errorCaches),
      warnings: imported.warnings,
      sheetNames: imported.sheetNames,
      reads,
      changes,
      ...(imported.diagnostics ? { diagnostics: imported.diagnostics } : {}),
    }
  } finally {
    workbook.dispose()
  }
}

export const recalculateSheetjsWorkbook = recalculateXlsx

export const xlsxCacheDoctorSchemaVersion = 'xlsx-cache-doctor.v1'

export type XlsxCacheInspectionLimit = number | 'all'
export type XlsxCacheStatus = 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'

export interface XlsxCacheInspectionOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly inspectLimit?: XlsxCacheInspectionLimit
  readonly config?: WorkPaperConfig
}

export interface XlsxCacheStatusSummary {
  readonly inspected: number
  readonly stale: number
  readonly fresh: number
  readonly missingCache: number
  readonly unsupportedRecalculation: number
}

export interface XlsxCacheFormulaInspection {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: RawCellContent
  readonly recalculatedValue: XlsxFormulaRecalcCellValue | undefined
  readonly literalRecalculatedValue?: RawCellContent | string
  readonly cacheStatus: XlsxCacheStatus
  readonly staleCachedValue: boolean | null
}

export interface XlsxCacheInspectionResult {
  readonly schemaVersion: typeof xlsxCacheDoctorSchemaVersion
  readonly sheetNames: readonly string[]
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: XlsxCacheInspectionLimit
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: XlsxCacheStatusSummary
  readonly suggestedReads: readonly string[]
  readonly formulas: readonly XlsxCacheFormulaInspection[]
  readonly warnings: readonly string[]
  readonly diagnostics?: ImportedWorkbookDiagnostics
  readonly inspectionCompleted: true
  readonly recalculationCompleted: true
  readonly excelParity: 'not_proven'
}

interface XlsxCacheFormulaCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: RawCellContent
}

export function inspectXlsxCache(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxCacheInspectionOptions = {},
): XlsxCacheInspectionResult {
  const importOptions = options.externalWorkbooks
    ? { externalWorkbooks: options.externalWorkbooks, externalLinkCacheArtifactMode: 'replace-refreshed' as const }
    : {}
  const imported = importXlsx(toUint8Array(input), options.fileName ?? 'workbook.xlsx', importOptions)
  const formulaCells = collectXlsxCacheFormulaCells(imported.snapshot)
  const inspectionLimit = normalizeXlsxCacheInspectionLimit(options.inspectLimit ?? 'all')
  const inspectedFormulaCells = inspectionLimit === 'all' ? formulaCells : formulaCells.slice(0, inspectionLimit)
  const uninspectedFormulaCellCount = formulaCells.length - inspectedFormulaCells.length
  const suggestedReads = inspectedFormulaCells.map((cell) => cell.target)
  const recalculated = recalculateXlsx(input, {
    ...(options.externalWorkbooks ? { externalWorkbooks: options.externalWorkbooks } : {}),
    ...(options.fileName ? { fileName: options.fileName } : {}),
    ...(options.edits ? { edits: options.edits } : {}),
    reads: suggestedReads,
    ...(options.config ? { config: options.config } : {}),
  })
  const formulas = inspectedFormulaCells.map((cell) => {
    const recalculatedValue = recalculated.reads[cell.target]
    const literalRecalculatedValue = literalValueForXlsxCacheInspection(recalculatedValue)
    const cacheStatus = xlsxCacheStatusForInspection(cell.cachedValue, literalRecalculatedValue)
    return {
      target: cell.target,
      formula: cell.formula,
      ...(cell.cachedValue !== undefined ? { cachedValue: cell.cachedValue } : {}),
      recalculatedValue,
      ...(literalRecalculatedValue !== undefined ? { literalRecalculatedValue } : {}),
      cacheStatus,
      staleCachedValue: staleCachedValueForXlsxCacheInspection(cacheStatus),
    }
  })

  return {
    schemaVersion: xlsxCacheDoctorSchemaVersion,
    sheetNames: imported.sheetNames,
    formulaCellCount: formulaCells.length,
    inspectedFormulaCellCount: inspectedFormulaCells.length,
    uninspectedFormulaCellCount,
    inspectionLimit,
    staleCachedFormulaCount: formulas.filter((formula) => formula.staleCachedValue === true).length,
    cacheStatusSummary: buildXlsxCacheStatusSummary(formulas),
    suggestedReads,
    formulas,
    warnings: recalculated.warnings,
    ...(recalculated.diagnostics ? { diagnostics: recalculated.diagnostics } : {}),
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
  }
}

function collectXlsxCacheFormulaCells(snapshot: WorkbookSnapshot): XlsxCacheFormulaCell[] {
  const cells: XlsxCacheFormulaCell[] = []
  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    for (const cell of sheet.cells.toSorted((left, right) => compareA1Addresses(left.address, right.address))) {
      if (typeof cell.formula !== 'string' || cell.formula.trim().length === 0) {
        continue
      }
      cells.push({
        target: formatQualifiedTarget(sheet.name, cell.address),
        formula: cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`,
        ...(cell.value !== undefined ? { cachedValue: cell.value } : {}),
      })
    }
  }
  return cells
}

function normalizeXlsxCacheInspectionLimit(limit: XlsxCacheInspectionLimit): XlsxCacheInspectionLimit {
  if (limit === 'all') {
    return limit
  }
  if (Number.isInteger(limit) && limit > 0) {
    return limit
  }
  throw new Error(`Expected inspectLimit to be "all" or a positive integer, received: ${String(limit)}`)
}

function compareA1Addresses(left: string, right: string): number {
  const leftParts = parseA1AddressForSort(left)
  const rightParts = parseA1AddressForSort(right)
  return leftParts.row - rightParts.row || leftParts.col - rightParts.col || left.localeCompare(right)
}

function parseA1AddressForSort(address: string): { readonly row: number; readonly col: number } {
  const match = /^([A-Z]+)(\d+)$/iu.exec(address)
  if (!match) {
    return { row: Number.MAX_SAFE_INTEGER, col: Number.MAX_SAFE_INTEGER }
  }
  const [, letters = '', rowText = ''] = match
  let col = 0
  for (const letter of letters.toUpperCase()) {
    col = col * 26 + letter.charCodeAt(0) - 64
  }
  return { row: Number(rowText), col }
}

function formatQualifiedTarget(sheetName: string, address: string): string {
  return `${quoteSheetNameForTarget(sheetName)}!${address}`
}

function quoteSheetNameForTarget(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function literalValueForXlsxCacheInspection(value: XlsxFormulaRecalcCellValue | undefined): RawCellContent | string | undefined {
  if (value === undefined || typeof value !== 'object' || value === null || !('tag' in value)) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return 'value' in value && typeof value.value === 'number' && Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return 'value' in value && typeof value.value === 'boolean' ? value.value : undefined
    case ValueTag.String:
      return 'value' in value && typeof value.value === 'string' ? value.value : undefined
    case ValueTag.Error:
      return 'code' in value && typeof value.code === 'number' ? formatErrorCode(value.code) : undefined
  }
}

function xlsxCacheStatusForInspection(
  cachedValue: RawCellContent | undefined,
  literalRecalculatedValue: RawCellContent | string | undefined,
): XlsxCacheStatus {
  if (cachedValue === undefined) {
    return 'missing-cache'
  }
  if (literalRecalculatedValue === undefined) {
    return 'unsupported-recalculation'
  }
  return literalValuesEqual(cachedValue, literalRecalculatedValue) ? 'fresh' : 'stale'
}

function staleCachedValueForXlsxCacheInspection(cacheStatus: XlsxCacheStatus): XlsxCacheFormulaInspection['staleCachedValue'] {
  switch (cacheStatus) {
    case 'stale':
      return true
    case 'fresh':
      return false
    case 'missing-cache':
    case 'unsupported-recalculation':
      return null
  }
}

function buildXlsxCacheStatusSummary(formulas: readonly XlsxCacheFormulaInspection[]): XlsxCacheStatusSummary {
  return {
    inspected: formulas.length,
    stale: formulas.filter((formula) => formula.cacheStatus === 'stale').length,
    fresh: formulas.filter((formula) => formula.cacheStatus === 'fresh').length,
    missingCache: formulas.filter((formula) => formula.cacheStatus === 'missing-cache').length,
    unsupportedRecalculation: formulas.filter((formula) => formula.cacheStatus === 'unsupported-recalculation').length,
  }
}

function literalValuesEqual(left: RawCellContent, right: RawCellContent | string): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function snapshotWithFormulaCachedValues(
  workbook: WorkPaperInstance,
  snapshot: WorkbookSnapshot,
): { readonly snapshot: WorkbookSnapshot; readonly errorCaches: readonly FormulaErrorCache[] } {
  const next = structuredClone(snapshot)
  const errorCaches: FormulaErrorCache[] = []
  for (const sheet of next.sheets) {
    const sheetId = workbook.getSheetId(sheet.name)
    if (sheetId === undefined) {
      continue
    }
    for (const cell of sheet.cells) {
      if (typeof cell.formula !== 'string' || cell.formula.trim().length === 0) {
        continue
      }
      const address = parseA1CellReference(cell.address)
      const cellValue = workbook.getCellValue({ sheet: sheetId, row: address.row, col: address.col })
      const cachedValue = literalInputForFormulaCache(cellValue)
      if (cachedValue !== undefined) {
        cell.value = cachedValue
        continue
      }
      const cachedError = errorInputForFormulaCache(cellValue)
      if (cachedError !== undefined) {
        delete cell.value
        errorCaches.push({ sheetName: sheet.name, address: cell.address, value: cachedError })
      }
    }
  }
  return { snapshot: next, errorCaches }
}

function literalInputForFormulaCache(value: XlsxFormulaRecalcCellValue): string | number | boolean | null | undefined {
  if (typeof value !== 'object' || value === null || !('tag' in value)) {
    return undefined
  }
  if (value.tag === ValueTag.Empty) {
    return null
  }
  if (value.tag === ValueTag.Number && 'value' in value && typeof value.value === 'number' && Number.isFinite(value.value)) {
    return value.value
  }
  if (value.tag === ValueTag.Boolean && 'value' in value && typeof value.value === 'boolean') {
    return value.value
  }
  if (value.tag === ValueTag.String && 'value' in value && typeof value.value === 'string') {
    return value.value
  }
  return undefined
}

function errorInputForFormulaCache(value: XlsxFormulaRecalcCellValue): string | undefined {
  if (typeof value !== 'object' || value === null || !('tag' in value) || value.tag !== ValueTag.Error || !('code' in value)) {
    return undefined
  }
  if (typeof value.code !== 'number') {
    return undefined
  }
  // Desktop Excel displays #FIELD! for linked-data field failures but saves #VALUE! in worksheet formula caches.
  return value.code === ErrorCode.Field ? '#VALUE!' : formatErrorCode(value.code)
}

function addFormulaErrorCachesToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  errorCaches: readonly FormulaErrorCache[],
): Uint8Array {
  if (errorCaches.length === 0) {
    return bytes
  }
  const zip = unzipSync(bytes)
  const cachesBySheetName = new Map<string, FormulaErrorCache[]>()
  for (const cache of errorCaches) {
    const sheetCaches = cachesBySheetName.get(cache.sheetName) ?? []
    sheetCaches.push(cache)
    cachesBySheetName.set(cache.sheetName, sheetCaches)
  }
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  for (let index = 0; index < orderedSheets.length; index += 1) {
    const sheet = orderedSheets[index]!
    const sheetCaches = cachesBySheetName.get(sheet.name)
    if (!sheetCaches || sheetCaches.length === 0) {
      continue
    }
    const sheetPath = `xl/worksheets/sheet${String(index + 1)}.xml`
    const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
    if (sheetXml.length === 0) {
      continue
    }
    let nextSheetXml = sheetXml
    for (const cache of sheetCaches) {
      nextSheetXml = replaceFormulaCellErrorCache(nextSheetXml, cache.address, cache.value)
    }
    if (nextSheetXml !== sheetXml) {
      zip[sheetPath] = strToU8(nextSheetXml)
    }
  }
  return zipSync(zip)
}

function replaceFormulaCellErrorCache(sheetXml: string, address: string, value: string): string {
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${escapeRegExp(address)}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(pattern, (cellXml) => writeCellErrorCache(cellXml, value))
}

function writeCellErrorCache(cellXml: string, value: string): string {
  const escapedValue = escapeXmlText(value)
  const typedCellXml = cellXml.replace(/^<c\b([^>]*)>/u, (_match: string, attributes: string) => {
    const nextAttributes = attributes.replace(/\s+t="[^"]*"/u, '')
    return `<c${nextAttributes} t="e">`
  })
  if (/<v>[\s\S]*?<\/v>/u.test(typedCellXml)) {
    return typedCellXml.replace(/<v>[\s\S]*?<\/v>/u, `<v>${escapedValue}</v>`)
  }
  if (/<\/f>/u.test(typedCellXml)) {
    return typedCellXml.replace(/<\/f>/u, `</f><v>${escapedValue}</v>`)
  }
  return typedCellXml.replace(/<\/c>$/u, `<v>${escapedValue}</v></c>`)
}

function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function snapshotForFreshFormulaRecalculation(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const next = structuredClone(snapshot)
  const calculationSettings = next.workbook.metadata?.calculationSettings
  if (calculationSettings === undefined) {
    return next
  }
  if (calculationSettings.mode !== 'manual' && calculationSettings.fullCalcOnLoad !== false) {
    return next
  }

  next.workbook.metadata ??= {}
  next.workbook.metadata.calculationSettings = {
    ...calculationSettings,
    mode: 'automatic',
    fullCalcOnLoad: true,
  }
  return next
}

function restoreOutputCalculationSettings(
  snapshot: WorkbookSnapshot,
  originalCalculationSettings: WorkbookCalculationSettings | undefined,
): WorkbookSnapshot {
  if (originalCalculationSettings === undefined) {
    return snapshot
  }
  const next = structuredClone(snapshot)
  next.workbook.metadata ??= {}
  next.workbook.metadata.calculationSettings = calculationSettingsAfterExplicitRecalculation(originalCalculationSettings)
  return next
}

function calculationSettingsAfterExplicitRecalculation(settings: WorkbookCalculationSettings): WorkbookCalculationSettings {
  const recalculated = structuredClone(settings)
  delete recalculated.calcCompleted
  delete recalculated.calcOnSave
  delete recalculated.forceFullCalc
  delete recalculated.fullCalcOnLoad
  return recalculated
}

export function parseQualifiedCellTarget(workbook: WorkPaperInstance, target: string): WorkPaperCellAddress {
  const parsed = parseQualifiedA1(target)
  const sheet = workbook.getSheetId(parsed.sheetName)
  if (sheet === undefined) {
    throw new Error(`Unknown sheet in XLSX formula recalculation target: ${parsed.sheetName}`)
  }
  return {
    sheet,
    row: parsed.row,
    col: parsed.col,
  }
}

export function parseQualifiedA1(target: string): { sheetName: string; row: number; col: number } {
  const trimmed = target.trim()
  const separator = findSheetSeparator(trimmed)
  if (separator <= 0 || separator >= trimmed.length - 1) {
    throw new Error(`Expected a sheet-qualified A1 target such as Inputs!B2, received: ${target}`)
  }

  const sheetName = unquoteSheetName(trimmed.slice(0, separator))
  const a1 = trimmed
    .slice(separator + 1)
    .replace(/\$/gu, '')
    .toUpperCase()
  const match = /^(?<col>[A-Z]+)(?<row>[1-9][0-9]*)$/u.exec(a1)
  if (!match?.groups) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  const row = match.groups['row']
  const col = match.groups['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  return {
    sheetName,
    ...parseA1Parts(row, col),
  }
}

function parseA1CellReference(address: string): { row: number; col: number } {
  const match = /^\$?(?<col>[A-Z]+)\$?(?<row>[1-9][0-9]*)$/u.exec(address.trim().toUpperCase())
  if (!match?.groups) {
    throw new Error(`Expected a single A1 cell reference, received: ${address}`)
  }

  const row = match.groups['row']
  const col = match.groups['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference, received: ${address}`)
  }

  return parseA1Parts(row, col)
}

function parseA1Parts(row: string, col: string): { row: number; col: number } {
  return {
    row: Number.parseInt(row, 10) - 1,
    col: columnLettersToIndex(col),
  }
}

function findSheetSeparator(target: string): number {
  let inQuote = false
  for (let index = 0; index < target.length; index += 1) {
    const char = target[index]
    if (char === "'") {
      if (inQuote && target[index + 1] === "'") {
        index += 1
      } else {
        inQuote = !inQuote
      }
      continue
    }
    if (char === '!' && !inQuote) {
      return index
    }
  }
  return -1
}

function unquoteSheetName(rawSheetName: string): string {
  const trimmed = rawSheetName.trim()
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/gu, "'")
  }
  return trimmed
}

function columnLettersToIndex(letters: string): number {
  let index = 0
  for (const char of letters) {
    index = index * 26 + (char.charCodeAt(0) - 64)
  }
  return index - 1
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}
