import { XMLParser } from 'fast-xml-parser'

import type {
  WorkbookCalculationSettingsSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookFormulaAuditEntrySnapshot,
  WorkbookFormulaAuditSnapshot,
  WorkbookFormulaDiagnosticSnapshot,
} from '@bilig/protocol'
import type { WorksheetFormulaCell } from './xlsx-formulas.js'
import { getZipText, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { workbookSheetPath } from './xlsx-workbook-sheet-paths.js'

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

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

function formulaWithoutLeadingEquals(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function formulaLooksLikeR1C1(formula: string): boolean {
  return /(?:^|[^A-Za-z0-9_])R(?:[1-9][0-9]*|\[[-+]?[0-9]+\])?C(?:[1-9][0-9]*|\[[-+]?[0-9]+\])?(?:$|[^A-Za-z0-9_])/u.test(
    formulaWithoutLeadingEquals(formula),
  )
}

function formulaLooksLikeRangeReference(formula: string): boolean {
  return /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*\s*:\s*\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$/u.test(formulaWithoutLeadingEquals(formula).trim())
}

function worksheetFormulaEntry(
  sheetName: string,
  cell: WorksheetFormulaCell,
  settings: WorkbookCalculationSettingsSnapshot | undefined,
): WorkbookFormulaAuditEntrySnapshot {
  const hasStaleRisk =
    settings?.mode === 'manual' ||
    settings?.forceFullCalc === true ||
    settings?.fullCalcOnLoad === true ||
    settings?.calcCompleted === false ||
    settings?.fullPrecision === false
  const attributes = {
    ...(cell.aca !== null ? { aca: cell.aca } : {}),
    ...(cell.bx !== null ? { bx: cell.bx } : {}),
    ...(cell.ca !== null ? { ca: cell.ca } : {}),
    ...(cell.xmlSpace ? { xmlSpace: cell.xmlSpace } : {}),
  }
  return {
    context: 'worksheet-cell',
    clause: '18.3.1.40',
    sheetName,
    address: cell.address,
    formula: cell.formula,
    ...(cell.formulaType ? { formulaType: cell.formulaType } : {}),
    ...(cell.sharedIndex ? { sharedIndex: cell.sharedIndex } : {}),
    ...(cell.ref ? { ref: cell.ref } : {}),
    ...(cell.cellValueType ? { cellValueType: cell.cellValueType } : {}),
    ...(cell.cachedValue !== undefined ? { cachedValue: cell.cachedValue } : {}),
    ...(cell.cachedValueRaw !== null ? { cachedValueRaw: cell.cachedValueRaw } : {}),
    cacheStatus: cell.cachedValue !== undefined ? (hasStaleRisk ? 'staleRisk' : 'trustedCached') : 'missing',
    rawFormulaXml: cell.rawFormulaXml,
    ...(Object.keys(attributes).length > 0 ? { attributes } : {}),
  }
}

function definedNameFormulaEntry(definedName: WorkbookDefinedNameSnapshot): WorkbookFormulaAuditEntrySnapshot | null {
  const value = definedName.value
  if (!isRecord(value) || value['kind'] !== 'formula' || typeof value['formula'] !== 'string') {
    return null
  }
  return {
    context: 'defined-name',
    clause: '3.2.3.1',
    name: definedName.name,
    ...(definedName.scopeSheetName ? { sheetName: definedName.scopeSheetName } : {}),
    formula: formulaWithoutLeadingEquals(value['formula']),
  }
}

function readSheetFormulaTags(sheetXml: string | null): {
  readonly conditionalFormats: WorkbookFormulaAuditEntrySnapshot[]
  readonly dataValidations: WorkbookFormulaAuditEntrySnapshot[]
} {
  if (!sheetXml) {
    return { conditionalFormats: [], dataValidations: [] }
  }
  const parsed: unknown = xmlParser.parse(sheetXml)
  const worksheet = recordChild(parsed, 'worksheet')
  const conditionalFormats: WorkbookFormulaAuditEntrySnapshot[] = []
  for (const conditionalFormatting of asArray(worksheet?.['conditionalFormatting'])) {
    if (!isRecord(conditionalFormatting)) {
      continue
    }
    const sqref = stringValue(conditionalFormatting['sqref'])
    for (const rule of asArray(conditionalFormatting['cfRule'])) {
      if (!isRecord(rule)) {
        continue
      }
      for (const formula of asArray(rule['formula'])) {
        const text = stringValue(formula)
        if (text === null || text.trim().length === 0) {
          continue
        }
        conditionalFormats.push({
          context: 'conditional-format',
          clause: '3.2.3.1',
          formula: text,
          ...(sqref ? { sqref } : {}),
        })
      }
    }
  }
  const dataValidations: WorkbookFormulaAuditEntrySnapshot[] = []
  for (const validation of asArray(recordChild(worksheet, 'dataValidations')?.['dataValidation'])) {
    if (!isRecord(validation)) {
      continue
    }
    const sqref = stringValue(validation['sqref'])
    for (const key of ['formula1', 'formula2']) {
      const text = stringValue(validation[key])
      if (text === null || text.trim().length === 0) {
        continue
      }
      dataValidations.push({
        context: 'data-validation',
        clause: '3.2.3.1',
        formula: text,
        ...(sqref ? { sqref } : {}),
      })
    }
  }
  return { conditionalFormats, dataValidations }
}

function formulaDiagnostics(entry: WorkbookFormulaAuditEntrySnapshot): WorkbookFormulaDiagnosticSnapshot[] {
  if (entry.context === 'defined-name' && formulaLooksLikeR1C1(entry.formula)) {
    return [
      {
        code: 'r1c1-reference',
        context: entry.context,
        clause: '3.2.3.1',
        message: 'Excel supports R1C1 formula references in contexts not specified by ISO/IEC 29500.',
        formula: entry.formula,
        ...(entry.name ? { name: entry.name } : {}),
        ...(entry.sheetName ? { sheetName: entry.sheetName } : {}),
      },
    ]
  }
  if (entry.context === 'conditional-format' && formulaLooksLikeRangeReference(entry.formula)) {
    return [
      {
        code: 'conditional-format-range-reference',
        context: entry.context,
        clause: '3.2.3.1',
        message: 'Conditional-format formulas use Excel-specific context restrictions beyond generic formulas.',
        formula: entry.formula,
        ...(entry.sheetName ? { sheetName: entry.sheetName } : {}),
      },
    ]
  }
  if (entry.context === 'data-validation' && formulaLooksLikeR1C1(entry.formula)) {
    return [
      {
        code: 'data-validation-r1c1-reference',
        context: entry.context,
        clause: '3.2.3.1',
        message: 'Data-validation formulas use Excel-specific context restrictions beyond generic formulas.',
        formula: entry.formula,
        ...(entry.sheetName ? { sheetName: entry.sheetName } : {}),
      },
    ]
  }
  return []
}

function calculationDiagnostics(settings: WorkbookCalculationSettingsSnapshot | undefined): WorkbookFormulaDiagnosticSnapshot[] {
  if (!settings) {
    return []
  }
  const diagnostics: WorkbookFormulaDiagnosticSnapshot[] = []
  if (settings.mode === 'manual') {
    diagnostics.push({
      code: 'manual-calc-mode',
      context: 'calculation',
      clause: '18.2.2',
      message: 'Manual calculation mode is preserved; imported cached formula values can be stale.',
    })
  }
  if (settings.forceFullCalc === true) {
    diagnostics.push({
      code: 'force-full-calc',
      context: 'calculation',
      clause: '18.2.2',
      message: 'Workbook requests a full recalculation on open; imported cached formula values can be stale.',
    })
  }
  if (settings.calcCompleted === false) {
    diagnostics.push({
      code: 'calc-not-completed',
      context: 'calculation',
      clause: '18.2.2',
      message: 'Workbook calculation was not marked completed; imported cached formula values can be stale.',
    })
  }
  return diagnostics
}

function readCalcChain(source: XlsxZipSource, sheetNames: readonly string[]): WorkbookFormulaAuditSnapshot['calcChain'] | undefined {
  const zip = readXlsxZipEntries(source)
  const xml = getZipText(zip, 'xl/calcChain.xml')
  if (!xml) {
    return undefined
  }
  const sheetNamesByCalcChainSheetId = readCalcChainSheetNamesBySheetId(zip, sheetNames)
  const parsed: unknown = xmlParser.parse(xml)
  let currentSheetIndex = 1
  const cells = asArray(recordChild(parsed, 'calcChain')?.['c']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const address = stringValue(entry['r'])
    if (!address || address.trim().length === 0) {
      return []
    }
    const explicitSheetIndex = Number(stringValue(entry['i']))
    if (Number.isSafeInteger(explicitSheetIndex) && explicitSheetIndex > 0) {
      currentSheetIndex = explicitSheetIndex
    }
    const sheetName = sheetNamesByCalcChainSheetId.get(currentSheetIndex)
    return [
      {
        sheetIndex: currentSheetIndex,
        address,
        ...(sheetName ? { sheetName } : {}),
        ...(stringValue(entry['s']) === '1' ? { childChain: true } : {}),
        ...(stringValue(entry['l']) === '1' ? { newDependencyLevel: true } : {}),
      },
    ]
  })
  return cells.length > 0 ? { packagePath: 'xl/calcChain.xml', cells } : undefined
}

function readCalcChainSheetNamesBySheetId(
  zip: ReturnType<typeof readXlsxZipEntries>,
  fallbackSheetNames: readonly string[],
): Map<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return fallbackCalcChainSheetNamesBySheetId(fallbackSheetNames)
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const sheets = asArray(recordChild(recordChild(parsed, 'workbook'), 'sheets')?.['sheet'])
  const sheetNamesById = new Map<number, string>()
  for (const sheet of sheets) {
    if (!isRecord(sheet)) {
      continue
    }
    const sheetId = Number(stringValue(sheet['sheetId']))
    const name = stringValue(sheet['name'])
    if (Number.isSafeInteger(sheetId) && sheetId > 0 && name && name.length > 0) {
      sheetNamesById.set(sheetId, name)
    }
  }
  return sheetNamesById.size > 0 ? sheetNamesById : fallbackCalcChainSheetNamesBySheetId(fallbackSheetNames)
}

function fallbackCalcChainSheetNamesBySheetId(sheetNames: readonly string[]): Map<number, string> {
  return new Map(sheetNames.map((sheetName, index) => [index + 1, sheetName]))
}

export function readImportedWorkbookFormulaAudit(args: {
  readonly source: XlsxZipSource
  readonly sheetNames: readonly string[]
  readonly sheetPathsByName: ReadonlyMap<string, string>
  readonly fallbackSheetPaths: readonly string[]
  readonly worksheetFormulasBySheet: ReadonlyMap<string, ReadonlyMap<string, WorksheetFormulaCell>>
  readonly definedNames: readonly WorkbookDefinedNameSnapshot[]
  readonly calculationSettings?: WorkbookCalculationSettingsSnapshot
}): WorkbookFormulaAuditSnapshot | undefined {
  const formulas: WorkbookFormulaAuditEntrySnapshot[] = []
  const diagnostics: WorkbookFormulaDiagnosticSnapshot[] = []
  for (const sheetName of args.sheetNames) {
    for (const cell of args.worksheetFormulasBySheet.get(sheetName)?.values() ?? []) {
      const entry = worksheetFormulaEntry(sheetName, cell, args.calculationSettings)
      formulas.push(entry)
      diagnostics.push(...formulaDiagnostics(entry))
    }
  }

  for (const definedName of args.definedNames) {
    const entry = definedNameFormulaEntry(definedName)
    if (!entry) {
      continue
    }
    formulas.push(entry)
    diagnostics.push(...formulaDiagnostics(entry))
  }

  const zip = readXlsxZipEntries(args.source)
  args.sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(args.sheetPathsByName, args.fallbackSheetPaths, sheetName, index)
    const taggedFormulas = readSheetFormulaTags(sheetPath ? getZipText(zip, sheetPath) : null)
    for (const entry of [...taggedFormulas.conditionalFormats, ...taggedFormulas.dataValidations]) {
      const sheetEntry = { ...entry, sheetName }
      formulas.push(sheetEntry)
      diagnostics.push(...formulaDiagnostics(sheetEntry))
    }
  })

  diagnostics.push(...calculationDiagnostics(args.calculationSettings))
  const calcChain = readCalcChain(args.source, args.sheetNames)
  if (formulas.length === 0 && diagnostics.length === 0 && !calcChain) {
    return undefined
  }
  return {
    formulas,
    ...(diagnostics.length > 0 ? { diagnostics } : {}),
    ...(calcChain ? { calcChain } : {}),
  }
}
