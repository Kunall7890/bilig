import * as XLSX from 'xlsx'

import type { LiteralInput, WorkbookExternalWorkbookReferenceSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import type {
  XlsxExternalWorkbookHydrationDiagnostics,
  XlsxExternalWorkbookHydrationMatchKind,
  XlsxExternalWorkbookInput,
} from './xlsx-import-limits.js'

interface ExternalCachedNumber {
  readonly kind: 'number'
  readonly value: number
}

interface ExternalCachedString {
  readonly kind: 'string'
  readonly value: string
}

interface ExternalCachedBoolean {
  readonly kind: 'boolean'
  readonly value: boolean
}

interface ExternalCachedError {
  readonly kind: 'error'
  readonly value: string
}

interface ExternalCachedBlank {
  readonly kind: 'blank'
}

export type ExternalCachedValue =
  | ExternalCachedNumber
  | ExternalCachedString
  | ExternalCachedBoolean
  | ExternalCachedError
  | ExternalCachedBlank
export type ExternalCachedCells = Map<string, ExternalCachedValue>
export interface ExternalCachedSheet {
  readonly sheetName: string
  readonly cells: ExternalCachedCells
}
export type ExternalCachedSheets = Map<string, ExternalCachedSheet>
export type ImportedExternalLinkCaches = Map<number, ExternalCachedSheets>

export interface ImportedExternalLinkCacheRefreshResult {
  readonly caches: ImportedExternalLinkCaches
  readonly artifactCaches: ImportedExternalLinkCaches
  readonly refreshedBookIndices: ReadonlySet<number>
  readonly diagnostics?: XlsxExternalWorkbookHydrationDiagnostics
}

export type ImportedExternalLinkCacheUsage = Map<number, Map<string, Set<string>>>
type ImportedExternalWorkbookReferences = Map<number, WorkbookExternalWorkbookReferenceSnapshot>

interface SheetJsReadableCell {
  readonly t?: unknown
  readonly v?: unknown
  readonly w?: unknown
}

export type ImportedExternalCacheSheetMap = ReadonlyMap<string, string>

export interface ImportedExternalCacheSheetSnapshot {
  readonly key: string
  readonly name: string
  readonly cells: WorkbookSnapshot['sheets'][number]['cells']
}

export interface ImportedExternalCacheSheetPlan {
  readonly sheetsByExternalSheet: ReadonlyMap<string, ImportedExternalCacheSheetSnapshot>
  readonly sheetNamesByExternalSheet: ImportedExternalCacheSheetMap
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeSheetName(sheetName: string): string {
  return sheetName.toLocaleLowerCase('en-US')
}

function normalizeCellAddress(address: string): string {
  return address.replaceAll('$', '').toUpperCase()
}

interface ParsedCellAddress {
  readonly row: number
  readonly col: number
}

function parseCachedCellAddress(address: string): ParsedCellAddress | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/u.exec(normalizeCellAddress(address))
  if (!match) {
    return null
  }
  let col = 0
  for (const character of match[1]!) {
    col = col * 26 + character.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return col > 0 && Number.isSafeInteger(row) ? { row: row - 1, col: col - 1 } : null
}

function toUint8Array(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input)
}

function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;').replace(/"/gu, '&quot;')
}

function normalizedWorkbookTarget(value: string | undefined): string | null {
  if (!value || value.trim().length === 0) {
    return null
  }
  const withoutFragment = value.trim().split(/[?#]/u)[0] ?? value.trim()
  const normalized = withoutFragment.replace(/\\/gu, '/').replace(/^file:\/+/iu, '/')
  try {
    return decodeURIComponent(normalized).toLocaleLowerCase('en-US')
  } catch {
    return normalized.toLocaleLowerCase('en-US')
  }
}

function workbookIdentityFromName(value: string | undefined): string | null {
  const target = normalizedWorkbookTarget(value)
  if (!target) {
    return null
  }
  const segments = target.split('/').filter((segment) => segment.length > 0)
  return segments.at(-1) ?? target
}

function externalWorkbookInputIdentityNames(input: XlsxExternalWorkbookInput): Set<string> {
  return new Set(
    [input.workbookName, input.fileName, input.target]
      .map(workbookIdentityFromName)
      .filter((value): value is string => value !== null && value.length > 0),
  )
}

function externalReferenceIdentityNames(reference: WorkbookExternalWorkbookReferenceSnapshot): Set<string> {
  return new Set(
    [reference.workbookName, reference.target]
      .map(workbookIdentityFromName)
      .filter((value): value is string => value !== null && value.length > 0),
  )
}

function externalWorkbookInputMatchesReferenceTarget(
  input: XlsxExternalWorkbookInput,
  reference: WorkbookExternalWorkbookReferenceSnapshot,
): boolean {
  const inputTarget = normalizedWorkbookTarget(input.target)
  const referenceTarget = normalizedWorkbookTarget(reference.target)
  return Boolean(inputTarget && referenceTarget && inputTarget === referenceTarget)
}

function externalWorkbookInputMatchesReferenceIdentity(
  input: XlsxExternalWorkbookInput,
  reference: WorkbookExternalWorkbookReferenceSnapshot,
): boolean {
  const inputTarget = normalizedWorkbookTarget(input.target)
  const referenceTarget = normalizedWorkbookTarget(reference.target)
  if (inputTarget && referenceTarget) {
    return false
  }
  const inputNames = externalWorkbookInputIdentityNames(input)
  if (inputNames.size === 0) {
    return false
  }
  for (const referenceName of externalReferenceIdentityNames(reference)) {
    if (inputNames.has(referenceName)) {
      return true
    }
  }
  return false
}

function externalReferenceIdentityCandidateCount(
  references: ImportedExternalWorkbookReferences,
  reference: WorkbookExternalWorkbookReferenceSnapshot,
): number {
  const referenceNames = externalReferenceIdentityNames(reference)
  if (referenceNames.size === 0) {
    return 0
  }
  let count = 0
  for (const candidate of references.values()) {
    for (const referenceName of externalReferenceIdentityNames(candidate)) {
      if (referenceNames.has(referenceName)) {
        count += 1
        break
      }
    }
  }
  return count
}

export interface ResolvedExternalWorkbookInputForReference {
  readonly input?: XlsxExternalWorkbookInput
  readonly status: 'matched' | 'skipped-no-match' | 'skipped-ambiguous-match'
  readonly candidateCount: number
  readonly referenceCandidateCount?: number
  readonly matchKind?: XlsxExternalWorkbookHydrationMatchKind
}

export function resolveExternalWorkbookInputForReference(
  inputs: readonly XlsxExternalWorkbookInput[],
  references: ImportedExternalWorkbookReferences,
  reference: WorkbookExternalWorkbookReferenceSnapshot,
): ResolvedExternalWorkbookInputForReference {
  const exactTargetCandidates = inputs.filter((input) => externalWorkbookInputMatchesReferenceTarget(input, reference))
  if (exactTargetCandidates.length === 1) {
    return {
      input: exactTargetCandidates[0]!,
      status: 'matched',
      candidateCount: 1,
      matchKind: 'exact-target',
    }
  }
  if (exactTargetCandidates.length > 1) {
    return {
      status: 'skipped-ambiguous-match',
      candidateCount: exactTargetCandidates.length,
      matchKind: 'exact-target',
    }
  }

  const identityCandidates = inputs.filter((input) => externalWorkbookInputMatchesReferenceIdentity(input, reference))
  if (identityCandidates.length === 0) {
    return {
      status: 'skipped-no-match',
      candidateCount: 0,
    }
  }
  const referenceCandidateCount = externalReferenceIdentityCandidateCount(references, reference)
  if (identityCandidates.length > 1 || referenceCandidateCount !== 1) {
    return {
      status: 'skipped-ambiguous-match',
      candidateCount: identityCandidates.length,
      referenceCandidateCount,
      matchKind: 'unique-workbook-identity',
    }
  }
  return {
    input: identityCandidates[0]!,
    status: 'matched',
    candidateCount: 1,
    referenceCandidateCount,
    matchKind: 'unique-workbook-identity',
  }
}

function sheetJsErrorLiteral(value: unknown, formattedValue: unknown): string {
  if (typeof formattedValue === 'string' && formattedValue.startsWith('#')) {
    return formattedValue
  }
  if (typeof value === 'string') {
    return value
  }
  const errorLiteralByCode = new Map<number, string>([
    [0, '#NULL!'],
    [7, '#DIV/0!'],
    [15, '#VALUE!'],
    [23, '#REF!'],
    [29, '#NAME?'],
    [36, '#NUM!'],
    [42, '#N/A'],
    [43, '#GETTING_DATA'],
  ])
  if (typeof value === 'number') {
    return errorLiteralByCode.get(value) ?? String(value)
  }
  return typeof formattedValue === 'string' ? formattedValue : externalCachedStringLiteral(value)
}

function externalCachedStringLiteral(value: unknown): string {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'bigint') {
    return String(value)
  }
  const serialized = JSON.stringify(value)
  return serialized ?? ''
}

function sheetJsCellToExternalCachedValue(cell: SheetJsReadableCell): ExternalCachedValue | null {
  const value = cell.v
  if (value === undefined || value === null) {
    return null
  }
  const cellType = typeof cell.t === 'string' ? cell.t : undefined
  if (cellType === 'b' || typeof value === 'boolean') {
    return { kind: 'boolean', value: value === true || value === 1 || value === '1' }
  }
  if (cellType === 'e') {
    return { kind: 'error', value: sheetJsErrorLiteral(value, cell.w) }
  }
  if (cellType === 's' || cellType === 'str') {
    return { kind: 'string', value: externalCachedStringLiteral(value) }
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return { kind: 'number', value }
  }
  return { kind: 'string', value: externalCachedStringLiteral(value) }
}

function sheetJsSheetName(workbook: XLSX.WorkBook, sheetName: string): string | undefined {
  return workbook.SheetNames.find((candidate) => normalizeSheetName(candidate) === normalizeSheetName(sheetName))
}

function readExternalWorkbookSheetCache(sheet: XLSX.WorkSheet, usedAddresses?: ReadonlySet<string>): ExternalCachedCells {
  const cells: ExternalCachedCells = new Map()
  for (const [address, rawCell] of Object.entries(sheet)) {
    if (address.startsWith('!') || !isRecord(rawCell)) {
      continue
    }
    const normalizedAddress = normalizeCellAddress(address)
    if (!parseCachedCellAddress(normalizedAddress)) {
      continue
    }
    if (usedAddresses && !usedAddresses.has(normalizedAddress)) {
      continue
    }
    const cachedValue = sheetJsCellToExternalCachedValue(rawCell)
    if (cachedValue) {
      cells.set(normalizedAddress, cachedValue)
    }
  }
  for (const usedAddress of usedAddresses ?? []) {
    if (parseCachedCellAddress(usedAddress) && !cells.has(usedAddress)) {
      cells.set(usedAddress, { kind: 'blank' })
    }
  }
  return cells
}

export function readExternalWorkbookCacheFromInput(
  input: XlsxExternalWorkbookInput,
  reference: WorkbookExternalWorkbookReferenceSnapshot,
  usage: ImportedExternalLinkCacheUsage | undefined,
): ExternalCachedSheets {
  const workbook = XLSX.read(toUint8Array(input.bytes), {
    type: 'array',
    cellFormula: true,
    cellNF: false,
    cellStyles: false,
    cellText: false,
    cellDates: false,
  })
  const referencedSheetNames = reference.sheetNames ?? []
  const sheetNames = referencedSheetNames.length > 0 ? referencedSheetNames : workbook.SheetNames
  const sheets: ExternalCachedSheets = new Map()
  const bookUsage = usage?.get(reference.bookIndex)
  for (const linkedSheetName of sheetNames) {
    const normalizedSheetName = normalizeSheetName(linkedSheetName)
    const usedAddresses = bookUsage?.get(normalizedSheetName)
    if (bookUsage && !usedAddresses) {
      continue
    }
    const workbookSheetName = sheetJsSheetName(workbook, linkedSheetName)
    const worksheet = workbookSheetName ? workbook.Sheets[workbookSheetName] : undefined
    if (!worksheet) {
      continue
    }
    const cells = readExternalWorkbookSheetCache(worksheet, usedAddresses)
    if (cells.size > 0) {
      sheets.set(normalizedSheetName, { sheetName: linkedSheetName, cells })
    }
  }
  return sheets
}

export function externalCachedValueToLiteralInput(value: ExternalCachedValue): LiteralInput | undefined {
  switch (value.kind) {
    case 'number':
    case 'string':
    case 'boolean':
      return value.value
    case 'blank':
      return null
    case 'error':
      return undefined
  }
}

function externalCachedErrorFormula(value: string): string {
  return ['#NULL!', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#N/A', '#NUM!', '#SPILL!', '#BLOCKED!', '#CYCLE!', '#FIELD!'].includes(value)
    ? value
    : '#VALUE!'
}

function sheetNameKey(sheetName: string): string {
  return sheetName.toLocaleLowerCase('en-US')
}

function sanitizeWorksheetNamePart(value: string): string {
  const sanitized = value
    .replace(/[:\\/?*[\]]/gu, '_')
    .replace(/'/gu, '_')
    .replace(/\s+/gu, '_')
    .replace(/_+/gu, '_')
    .replace(/^_+|_+$/gu, '')
  return sanitized.length > 0 ? sanitized : 'Sheet'
}

function uniqueExternalCacheSheetName(bookIndex: number, sheetName: string, usedSheetNames: Set<string>): string {
  const prefix = `__bilig_ext_${bookIndex}_`
  const suffix = sanitizeWorksheetNamePart(sheetName)
  const base = `${prefix}${suffix}`.slice(0, 31)
  let candidate = base.length > 0 ? base : '__bilig_ext'
  let disambiguator = 2
  while (usedSheetNames.has(sheetNameKey(candidate))) {
    const suffixText = `_${disambiguator}`
    candidate = `${base.slice(0, Math.max(0, 31 - suffixText.length))}${suffixText}`
    disambiguator += 1
  }
  usedSheetNames.add(sheetNameKey(candidate))
  return candidate
}

function compareExternalCacheCellAddresses(left: string, right: string): number {
  const leftAddress = parseCachedCellAddress(left)
  const rightAddress = parseCachedCellAddress(right)
  if (!leftAddress || !rightAddress) {
    return left.localeCompare(right)
  }
  return leftAddress.row - rightAddress.row || leftAddress.col - rightAddress.col
}

function externalCachedValueXml(value: ExternalCachedValue): { readonly typeAttribute: string; readonly value: string } | null {
  switch (value.kind) {
    case 'number':
      return { typeAttribute: '', value: String(value.value) }
    case 'boolean':
      return { typeAttribute: ' t="b"', value: value.value ? '1' : '0' }
    case 'error':
      return { typeAttribute: ' t="e"', value: escapeXmlText(value.value) }
    case 'string':
      return { typeAttribute: ' t="str"', value: escapeXmlText(value.value) }
    case 'blank':
      return null
  }
}

function buildExternalLinkSheetDataXml(sheetId: number, sheet: ExternalCachedSheet): string | null {
  const rows = new Map<number, { readonly address: string; readonly col: number; readonly value: ExternalCachedValue }[]>()
  for (const [address, value] of sheet.cells) {
    const parsed = parseCachedCellAddress(address)
    if (!parsed) {
      continue
    }
    const rowCells = rows.get(parsed.row) ?? []
    rowCells.push({ address: normalizeCellAddress(address), col: parsed.col, value })
    rows.set(parsed.row, rowCells)
  }
  if (rows.size === 0) {
    return null
  }

  const rowXml = [...rows.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .map(([rowIndex, cells]) => {
      const cellXml = cells
        .toSorted((left, right) => left.col - right.col)
        .flatMap((cell) => {
          const serialized = externalCachedValueXml(cell.value)
          if (!serialized) {
            return []
          }
          return `<cell r="${escapeXmlText(cell.address)}"${serialized.typeAttribute}><v>${serialized.value}</v></cell>`
        })
        .join('')
      if (cellXml.length === 0) {
        return null
      }
      return `<row r="${rowIndex + 1}">${cellXml}</row>`
    })
    .filter((value): value is string => value !== null)
    .join('')
  return rowXml.length > 0 ? `<sheetData sheetId="${sheetId}">${rowXml}</sheetData>` : null
}

export function buildExternalLinkSheetDataSetXml(sheetNames: readonly string[], sheets: ExternalCachedSheets): string | null {
  const sheetDataXml = sheetNames
    .map((sheetName, sheetId) => {
      const sheet = sheets.get(normalizeSheetName(sheetName))
      return sheet ? buildExternalLinkSheetDataXml(sheetId, sheet) : null
    })
    .filter((value): value is string => value !== null)
    .join('')
  return sheetDataXml.length > 0 ? `<sheetDataSet>${sheetDataXml}</sheetDataSet>` : null
}

export function externalCacheSheetKey(bookIndex: number, sheetName: string): string {
  return `${bookIndex}\u0000${normalizeSheetName(sheetName)}`
}

function buildImportedExternalCacheSheetCells(sheet: ExternalCachedSheet): WorkbookSnapshot['sheets'][number]['cells'] {
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  for (const [address, value] of sheet.cells) {
    const parsed = parseCachedCellAddress(address)
    if (!parsed) {
      continue
    }
    if (value.kind === 'error') {
      cells.push({
        address: normalizeCellAddress(address),
        row: parsed.row,
        col: parsed.col,
        formula: externalCachedErrorFormula(value.value),
      })
      continue
    }
    const literal = externalCachedValueToLiteralInput(value)
    if (literal === undefined) {
      continue
    }
    cells.push({
      address: normalizeCellAddress(address),
      row: parsed.row,
      col: parsed.col,
      value: literal,
    })
  }
  return cells.toSorted((left, right) => compareExternalCacheCellAddresses(left.address, right.address))
}

export function buildImportedExternalCacheSheetPlan(
  caches: ImportedExternalLinkCaches,
  existingSheetNames: readonly string[],
): ImportedExternalCacheSheetPlan {
  const usedSheetNames = new Set(existingSheetNames.map(sheetNameKey))
  const sheetsByExternalSheet = new Map<string, ImportedExternalCacheSheetSnapshot>()
  const sheetNamesByExternalSheet = new Map<string, string>()

  for (const [bookIndex, sheets] of [...caches.entries()].toSorted((left, right) => left[0] - right[0])) {
    for (const sheet of [...sheets.values()].toSorted((left, right) => left.sheetName.localeCompare(right.sheetName))) {
      const cells = buildImportedExternalCacheSheetCells(sheet)
      if (cells.length === 0) {
        continue
      }
      const key = externalCacheSheetKey(bookIndex, sheet.sheetName)
      const name = uniqueExternalCacheSheetName(bookIndex, sheet.sheetName, usedSheetNames)
      const snapshot: ImportedExternalCacheSheetSnapshot = { key, name, cells }
      sheetsByExternalSheet.set(key, snapshot)
      sheetNamesByExternalSheet.set(key, name)
    }
  }

  return { sheetsByExternalSheet, sheetNamesByExternalSheet }
}
