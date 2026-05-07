import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  LiteralInput,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookSnapshot,
  WorkbookValidationComparisonOperator,
  WorkbookValidationErrorStyle,
  WorkbookValidationListSourceSnapshot,
} from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const worksheetDataValidationTailElements = [
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'pivotTableDefinition',
] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
}

function stringChild(value: Record<string, unknown>, key: string): string | null {
  return stringValue(value[key])
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function normalizeAddress(address: string): string | null {
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address.replaceAll('$', '')))
  } catch {
    return null
  }
}

function absoluteAddress(address: string): string | null {
  const normalized = normalizeAddress(address)
  if (!normalized) {
    return null
  }
  const decoded = XLSX.utils.decode_cell(normalized)
  return `$${XLSX.utils.encode_col(decoded.c)}$${decoded.r + 1}`
}

function rangeRefA1(range: CellRangeRef): string | null {
  try {
    const decoded = XLSX.utils.decode_range(`${range.startAddress}:${range.endAddress}`.replaceAll('$', ''))
    return XLSX.utils.encode_range(decoded)
  } catch {
    return null
  }
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function formatExportSheetReference(
  sheetName: string,
  reference: string,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  const exportSheetName = exportSheetNamesByOriginalName.get(sheetName)
  return exportSheetName ? `${quoteSheetName(exportSheetName)}!${reference}` : null
}

function formatValidationListSource(
  source: WorkbookValidationListSourceSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  switch (source.kind) {
    case 'named-range':
      return source.name.trim().length > 0 ? source.name.trim() : null
    case 'cell-ref': {
      const address = absoluteAddress(source.address)
      return address ? formatExportSheetReference(source.sheetName, address, exportSheetNamesByOriginalName) : null
    }
    case 'range-ref': {
      const startAddress = absoluteAddress(source.startAddress)
      const endAddress = absoluteAddress(source.endAddress)
      return startAddress && endAddress
        ? formatExportSheetReference(source.sheetName, `${startAddress}:${endAddress}`, exportSheetNamesByOriginalName)
        : null
    }
    case 'structured-ref':
      return source.columnName.trim().length > 0 ? `${source.tableName}[${source.columnName}]` : source.tableName
  }
}

function formatListLiteral(value: LiteralInput): string {
  if (value === null) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function formatListFormula(
  rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }>,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  if (rule.values && rule.values.length > 0) {
    return `"${rule.values.map(formatListLiteral).join(',').replaceAll('"', '""')}"`
  }
  return rule.source ? formatValidationListSource(rule.source, exportSheetNamesByOriginalName) : null
}

function formatScalarFormulaValue(value: LiteralInput): string | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(value) : null
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (trimmed.startsWith('=')) {
      return trimmed.slice(1).trim()
    }
    return `"${value.replaceAll('"', '""')}"`
  }
  return '""'
}

function buildDataValidationAttributes(validation: WorkbookDataValidationSnapshot, type: string, operator?: string): string {
  const attributes = [`type="${escapeXml(type)}"`]
  if (operator) {
    attributes.push(`operator="${escapeXml(operator)}"`)
  }
  if (validation.allowBlank !== undefined) {
    attributes.push(`allowBlank="${validation.allowBlank ? '1' : '0'}"`)
  }
  if (validation.showDropdown !== undefined) {
    attributes.push(`showDropDown="${validation.showDropdown ? '0' : '1'}"`)
  }
  if (validation.errorStyle !== undefined) {
    attributes.push(`errorStyle="${escapeXml(validation.errorStyle)}"`)
  }
  if (validation.promptTitle !== undefined || validation.promptMessage !== undefined) {
    attributes.push('showInputMessage="1"')
  }
  if (validation.errorStyle !== undefined || validation.errorTitle !== undefined || validation.errorMessage !== undefined) {
    attributes.push('showErrorMessage="1"')
  }
  if (validation.promptTitle !== undefined) {
    attributes.push(`promptTitle="${escapeXml(validation.promptTitle)}"`)
  }
  if (validation.promptMessage !== undefined) {
    attributes.push(`prompt="${escapeXml(validation.promptMessage)}"`)
  }
  if (validation.errorTitle !== undefined) {
    attributes.push(`errorTitle="${escapeXml(validation.errorTitle)}"`)
  }
  if (validation.errorMessage !== undefined) {
    attributes.push(`error="${escapeXml(validation.errorMessage)}"`)
  }
  const rangeRef = rangeRefA1(validation.range)
  if (rangeRef) {
    attributes.push(`sqref="${escapeXml(rangeRef)}"`)
  }
  return attributes.join(' ')
}

function buildExportDataValidationXml(
  validation: WorkbookDataValidationSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): string | null {
  const rangeRef = rangeRefA1(validation.range)
  if (!rangeRef) {
    return null
  }
  const rule = validation.rule
  if (rule.kind === 'checkbox') {
    return null
  }
  if (rule.kind === 'list') {
    const formula = formatListFormula(rule, exportSheetNamesByOriginalName)
    if (!formula) {
      return null
    }
    return `<dataValidation ${buildDataValidationAttributes(validation, 'list')}><formula1>${escapeXml(formula)}</formula1></dataValidation>`
  }

  const formulas = rule.values
    .slice(0, 2)
    .map(formatScalarFormulaValue)
    .flatMap((value, index) => (value === null ? [] : [`<formula${String(index + 1)}>${escapeXml(value)}</formula${String(index + 1)}>`]))
  return `<dataValidation ${buildDataValidationAttributes(validation, rule.kind, rule.operator)}>${formulas.join('')}</dataValidation>`
}

function updateCountAttribute(attributes: string, nextCount: number): string {
  if (/\bcount="/u.test(attributes)) {
    return attributes.replace(/\bcount="[^"]*"/u, `count="${String(nextCount)}"`)
  }
  return `${attributes} count="${String(nextCount)}"`
}

function insertWorksheetDataValidations(sheetXml: string, validationXml: readonly string[]): string {
  const existing = /<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/u.exec(sheetXml)
  if (existing) {
    const existingCount = Number(/\bcount="([^"]*)"/u.exec(existing[1] ?? '')?.[1] ?? '0')
    const nextCount = Number.isFinite(existingCount) ? existingCount + validationXml.length : validationXml.length
    const attributes = updateCountAttribute(existing[1] ?? '', nextCount)
    return sheetXml.replace(existing[0], `<dataValidations${attributes}>${existing[2] ?? ''}${validationXml.join('')}</dataValidations>`)
  }

  const block = `<dataValidations count="${String(validationXml.length)}">${validationXml.join('')}</dataValidations>`
  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetDataValidationTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${block}${sheetXml.slice(insertIndex)}`
}

export function addExportDataValidationsToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.validations ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      if (!exportSheetNamesByOriginalName.has(sheet.name)) {
        return
      }
      const validations = (sheet.metadata?.validations ?? [])
        .filter((validation) => validation.range.sheetName === sheet.name)
        .map((validation) => buildExportDataValidationXml(validation, exportSheetNamesByOriginalName))
        .filter((validation): validation is string => validation !== null)
      if (validations.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetDataValidations(sheetXml, validations))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

function parseBooleanAttribute(value: unknown): boolean | undefined {
  if (value === true || value === '1' || value === 'true') {
    return true
  }
  if (value === false || value === '0' || value === 'false') {
    return false
  }
  return undefined
}

function parseComparisonOperator(value: unknown): WorkbookValidationComparisonOperator | null {
  switch (value) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual':
      return value
    default:
      return null
  }
}

function parseErrorStyle(value: unknown): WorkbookValidationErrorStyle | undefined {
  switch (value) {
    case 'stop':
    case 'warning':
    case 'information':
      return value
    default:
      return undefined
  }
}

function parseSqrefRange(sheetName: string, value: string): CellRangeRef | null {
  try {
    const decoded = XLSX.utils.decode_range(value.replaceAll('$', ''))
    return {
      sheetName,
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function parseQuotedSheetReference(value: string): { sheetName: string; reference: string } | null {
  if (!value.startsWith("'")) {
    return null
  }
  let sheetName = ''
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'" && value[index + 1] === "'") {
      sheetName += "'"
      index += 1
      continue
    }
    if (character === "'" && value[index + 1] === '!') {
      const reference = value.slice(index + 2).trim()
      return sheetName.trim().length > 0 && reference.length > 0 ? { sheetName, reference } : null
    }
    sheetName += character
  }
  return null
}

function parseSheetReference(value: string): { sheetName: string; reference: string } | null {
  const quoted = parseQuotedSheetReference(value)
  if (quoted) {
    return quoted
  }
  const separatorIndex = value.indexOf('!')
  if (separatorIndex <= 0) {
    return null
  }
  const sheetName = value.slice(0, separatorIndex).trim()
  const reference = value.slice(separatorIndex + 1).trim()
  return sheetName.length > 0 && reference.length > 0 ? { sheetName, reference } : null
}

function parseSourceReference(sheetName: string, reference: string): WorkbookValidationListSourceSnapshot | null {
  const parts = reference.split(':')
  if (parts.length === 1) {
    const address = normalizeAddress(parts[0] ?? '')
    return address ? { kind: 'cell-ref', sheetName, address } : null
  }
  if (parts.length === 2) {
    const startAddress = normalizeAddress(parts[0] ?? '')
    const endAddress = normalizeAddress(parts[1] ?? '')
    return startAddress && endAddress ? { kind: 'range-ref', sheetName, startAddress, endAddress } : null
  }
  return null
}

function parseStructuredReference(value: string): WorkbookValidationListSourceSnapshot | null {
  const match = /^([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]$/u.exec(value)
  if (!match) {
    return null
  }
  return {
    kind: 'structured-ref',
    tableName: match[1] ?? '',
    columnName: match[2] ?? '',
  }
}

function parseListFormula(
  formula: string,
): Pick<Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }>, 'values' | 'source'> | null {
  const trimmed = formula.trim()
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return {
      values: trimmed.slice(1, -1).replaceAll('""', '"').split(','),
    }
  }
  const expression = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  const sheetReference = parseSheetReference(expression)
  if (sheetReference) {
    const source = parseSourceReference(sheetReference.sheetName, sheetReference.reference)
    return source ? { source } : null
  }
  const structured = parseStructuredReference(expression)
  if (structured) {
    return { source: structured }
  }
  return /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(expression) ? { source: { kind: 'named-range', name: expression } } : null
}

function parseScalarFormulaValue(formula: string): LiteralInput {
  const trimmed = formula.trim()
  if (/^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/u.test(trimmed)) {
    const numberValue = Number(trimmed)
    if (Number.isFinite(numberValue)) {
      return numberValue
    }
  }
  if (/^TRUE$/iu.test(trimmed)) {
    return true
  }
  if (/^FALSE$/iu.test(trimmed)) {
    return false
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replaceAll('""', '"')
  }
  return trimmed
}

function parseValidationRule(entry: Record<string, unknown>): WorkbookDataValidationRuleSnapshot | null {
  const type = stringChild(entry, 'type') ?? 'any'
  const formula1 = stringChild(entry, 'formula1')
  const formula2 = stringChild(entry, 'formula2')
  switch (type) {
    case 'list': {
      if (!formula1) {
        return null
      }
      const listFormula = parseListFormula(formula1)
      return listFormula ? { kind: 'list', ...listFormula } : null
    }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength': {
      const operator = parseComparisonOperator(entry['operator']) ?? 'between'
      const values = [formula1, formula2].flatMap((formula) => (formula === null ? [] : [parseScalarFormulaValue(formula)]))
      return { kind: type, operator, values }
    }
    default:
      return null
  }
}

function parseDataValidationEntry(sheetName: string, entry: unknown): WorkbookDataValidationSnapshot[] {
  if (!isRecord(entry)) {
    return []
  }
  const rule = parseValidationRule(entry)
  const sqref = stringChild(entry, 'sqref')
  if (!rule || !sqref) {
    return []
  }

  return sqref
    .trim()
    .split(/\s+/u)
    .flatMap((rangeRef) => {
      const range = parseSqrefRange(sheetName, rangeRef)
      if (!range) {
        return []
      }
      const validation: WorkbookDataValidationSnapshot = {
        range,
        rule: structuredClone(rule),
      }
      const allowBlank = parseBooleanAttribute(entry['allowBlank'])
      if (allowBlank !== undefined) {
        validation.allowBlank = allowBlank
      }
      const showDropDown = parseBooleanAttribute(entry['showDropDown'])
      if (showDropDown !== undefined) {
        validation.showDropdown = !showDropDown
      }
      const promptTitle = stringChild(entry, 'promptTitle')
      if (promptTitle !== null) {
        validation.promptTitle = promptTitle
      }
      const promptMessage = stringChild(entry, 'prompt')
      if (promptMessage !== null) {
        validation.promptMessage = promptMessage
      }
      const errorStyle = parseErrorStyle(entry['errorStyle'])
      if (errorStyle !== undefined) {
        validation.errorStyle = errorStyle
      }
      const errorTitle = stringChild(entry, 'errorTitle')
      if (errorTitle !== null) {
        validation.errorTitle = errorTitle
      }
      const errorMessage = stringChild(entry, 'error')
      if (errorMessage !== null) {
        validation.errorMessage = errorMessage
      }
      return [validation]
    })
}

export function readImportedWorkbookDataValidations(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookDataValidationSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const validationsBySheet = new Map<string, WorkbookDataValidationSnapshot[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<dataValidations\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const entries = asArray(recordChild(recordChild(parsed, 'worksheet'), 'dataValidations')?.['dataValidation'])
    const validations = entries
      .flatMap((entry) => parseDataValidationEntry(sheetName, entry))
      .toSorted((left, right) =>
        `${left.range.sheetName}:${left.range.startAddress}:${left.range.endAddress}`.localeCompare(
          `${right.range.sheetName}:${right.range.startAddress}:${right.range.endAddress}`,
        ),
      )
    if (validations.length > 0) {
      validationsBySheet.set(sheetName, validations)
    }
  })

  return validationsBySheet
}
