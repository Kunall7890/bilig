import { readFileSync } from 'node:fs'
import { basename } from 'node:path'

import {
  getZipText,
  readXlsxWorkbookCells,
  readXlsxZipEntriesLazy,
  writeSimpleXlsxWorkbook,
  type ImportedWorkbookDiagnostics,
  type XlsxExternalWorkbookInput,
  type XlsxWorkbookCells,
  type XlsxZipEntries,
} from '@bilig/xlsx'
import type { XlsxFormulaRecalcEdit } from './types.js'

export const workbookCompatibilityReportSchemaVersion = 'bilig-workbook-compatibility-report.v1'

export type XlsxCacheInspectionLimit = number | 'all'
export type WorkbookCompatibilityRiskLevel = 'low' | 'medium' | 'high'

export interface WorkbookCompatibilityNamedCount {
  readonly name: string
  readonly count: number
}

export interface WorkbookCompatibilityReportOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly inspectLimit?: XlsxCacheInspectionLimit
}

export interface WorkbookCompatibilityReport {
  readonly schemaVersion: typeof workbookCompatibilityReportSchemaVersion
  readonly verified: true
  readonly input: {
    readonly fileName: string
    readonly externalWorkbookCount: number
    readonly inspectLimit: XlsxCacheInspectionLimit
  }
  readonly workbook: {
    readonly sheetCount: number
    readonly sheetNames: readonly string[]
    readonly nonEmptyCellCount: number
    readonly formulaCellCount: number
    readonly definedNameCount: number
    readonly tableCount: number
    readonly pivotTableCount: number
    readonly chartCount: number
    readonly macroModuleCount: number
  }
  readonly findings: {
    readonly unsupportedFunctions: readonly WorkbookCompatibilityNamedCount[]
    readonly externalLinks: {
      readonly count: number
      readonly unresolvedCount: number
      readonly refreshedCount: number
    }
    readonly macroModules: {
      readonly count: number
      readonly byteLength: number
    }
    readonly volatileFunctions: readonly WorkbookCompatibilityNamedCount[]
    readonly pivotTables: {
      readonly count: number
      readonly unsupportedCount: number
      readonly cacheOnlyCount: number
    }
    readonly staleCachedFormulas: {
      readonly count: number
    }
    readonly missingCachedFormulaValues: {
      readonly count: number
    }
    readonly unsupportedRecalculations: {
      readonly count: number
    }
    readonly warnings: readonly string[]
  }
  readonly risk: {
    readonly level: WorkbookCompatibilityRiskLevel
    readonly reasons: readonly string[]
  }
  readonly cacheInspection: {
    readonly inspectedFormulaCellCount: number
    readonly uninspectedFormulaCellCount: number
    readonly inspectionLimit: XlsxCacheInspectionLimit
    readonly suggestedReads: readonly string[]
  }
  readonly diagnostics?: ImportedWorkbookDiagnostics
  readonly commandSucceeded: true
  readonly inspectionCompleted: true
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
  readonly limitations: readonly string[]
  readonly next: {
    readonly docs: string
    readonly command: string
  }
}

interface CliExternalWorkbook {
  readonly path: string
  readonly target?: string
}

interface WorkbookCompatibilityCliOptions {
  readonly mode: 'file' | 'demo'
  readonly inputPath: string | undefined
  readonly externalWorkbooks: readonly CliExternalWorkbook[]
  readonly inspectLimit: XlsxCacheInspectionLimit
  readonly json: boolean
}

export interface WorkbookCompatibilityReportCliContext {
  readonly commandName?: string
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

const defaultInspectLimit: XlsxCacheInspectionLimit = 'all'
const docsUrl = 'https://proompteng.github.io/bilig/workbook-compatibility-report.html'
const volatileFunctionNames = ['TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT', 'SUBTOTAL', 'AGGREGATE'] as const
const knownUnsupportedRiskFunctions = new Set([
  'CALL',
  'COPILOT',
  'CUBEKPIMEMBER',
  'CUBEMEMBER',
  'CUBEMEMBERPROPERTY',
  'CUBERANKEDMEMBER',
  'CUBESET',
  'CUBESETCOUNT',
  'CUBEVALUE',
  'DDE',
  'DETECTLANGUAGE',
  'FILTERXML',
  'GOOGLEFINANCE',
  'IMAGE',
  'IMPORTDATA',
  'IMPORTFEED',
  'IMPORTHTML',
  'IMPORTRANGE',
  'IMPORTXML',
  'INFO',
  'PY',
  'REGISTER.ID',
  'RTD',
  'SQL.REQUEST',
  'STOCKHISTORY',
  'TRANSLATE',
  'WEBSERVICE',
])

export function buildWorkbookCompatibilityReport(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: WorkbookCompatibilityReportOptions = {},
): WorkbookCompatibilityReport {
  const bytes = toUint8Array(input)
  const fileName = options.fileName ?? 'workbook.xlsx'
  const externalWorkbooks = options.externalWorkbooks ?? []
  const zip = readXlsxZipEntriesLazy(bytes)
  const workbook = readXlsxWorkbookCells(zip)
  const workbookParts = scanWorkbookPackageParts(zip)
  const cacheInspection = inspectWorkbookFormulaCaches(workbook, options.inspectLimit ?? defaultInspectLimit)
  const unsupportedFunctions = countNamedValues(
    cacheInspection.formulas.flatMap((entry) => knownUnsupportedFunctionNamesFromFormula(entry.formula)),
  )
  const volatileFunctions = countNamedValues(cacheInspection.formulas.flatMap((entry) => volatileFunctionNamesFromFormula(entry.formula)))
  const report: Omit<WorkbookCompatibilityReport, 'risk'> = {
    schemaVersion: workbookCompatibilityReportSchemaVersion,
    verified: true,
    input: {
      fileName,
      externalWorkbookCount: externalWorkbooks.length,
      inspectLimit: cacheInspection.inspectionLimit,
    },
    workbook: {
      sheetCount: workbook.sheets.length,
      sheetNames: workbook.sheets.map((sheet) => sheet.name),
      nonEmptyCellCount: workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0),
      formulaCellCount: cacheInspection.formulaCellCount,
      definedNameCount: workbookParts.definedNameCount,
      tableCount: workbookParts.tableCount,
      pivotTableCount: workbookParts.pivotTableCount,
      chartCount: workbookParts.chartCount,
      macroModuleCount: workbookParts.macroModuleCount,
    },
    findings: {
      unsupportedFunctions,
      externalLinks: {
        count: workbookParts.externalLinkCount,
        unresolvedCount: 0,
        refreshedCount: 0,
      },
      macroModules: {
        count: workbookParts.macroModuleCount,
        byteLength: workbookParts.macroByteLength,
      },
      volatileFunctions,
      pivotTables: {
        count: workbookParts.pivotTableCount,
        unsupportedCount: 0,
        cacheOnlyCount: 0,
      },
      staleCachedFormulas: {
        count: cacheInspection.staleCachedFormulaCount,
      },
      missingCachedFormulaValues: {
        count: cacheInspection.cacheStatusSummary.missingCache,
      },
      unsupportedRecalculations: {
        count: cacheInspection.cacheStatusSummary.unsupportedRecalculation,
      },
      warnings: cacheInspection.warnings,
    },
    cacheInspection: {
      inspectedFormulaCellCount: cacheInspection.inspectedFormulaCellCount,
      uninspectedFormulaCellCount: cacheInspection.uninspectedFormulaCellCount,
      inspectionLimit: cacheInspection.inspectionLimit,
      suggestedReads: cacheInspection.suggestedReads,
    },
    commandSucceeded: true,
    inspectionCompleted: true,
    recalculationCompleted: false,
    excelParity: 'not_proven',
    limitations: [
      'This report identifies workbook features that may require investigation before using Bilig in a service or agent workflow.',
      'It is not an Excel compatibility certification.',
      'It scans workbook package metadata and formula caches; use xlsx-cache-doctor for native recalculation proof.',
      'It does not execute VBA, refresh pivots, refresh external data sources, or prove desktop Excel UI behavior.',
    ],
    next: {
      docs: docsUrl,
      command: `workbook-compatibility-report ${shellQuote(fileName)} --json`,
    },
  }

  return {
    ...report,
    risk: buildRisk(report),
  }
}

export function runWorkbookCompatibilityReportCli(args: readonly string[], context: WorkbookCompatibilityReportCliContext = {}): number {
  const commandName = context.commandName ?? 'workbook-compatibility-report'
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
      writeStdout(renderHelp(commandName))
      return 0
    }
    const options = parseCliArgs(args, commandName)
    const input = options.mode === 'demo' ? buildWorkbookCompatibilityDemoBytes() : readFileSync(requireInputPath(options))
    const fileName = options.mode === 'demo' ? 'bilig-workbook-compatibility-demo.xlsx' : basename(requireInputPath(options))
    const externalWorkbooks = readExternalWorkbookInputs(options.externalWorkbooks)
    const report = buildWorkbookCompatibilityReport(input, {
      fileName,
      ...(externalWorkbooks.length > 0 ? { externalWorkbooks } : {}),
      inspectLimit: options.inspectLimit,
    })

    if (options.json) {
      writeStdout(`${JSON.stringify(report, null, 2)}\n`)
      return report.verified ? 0 : 1
    }

    writeStdout(renderHumanReport(report))
    return report.verified ? 0 : 1
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

export function buildWorkbookCompatibilityDemoBytes(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Inputs',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'Metric' },
          { address: 'B1', row: 0, col: 1, value: 'Value' },
          { address: 'A2', row: 1, col: 0, value: 'Units' },
          { address: 'B2', row: 1, col: 1, value: 40 },
          { address: 'A3', row: 2, col: 0, value: 'Price' },
          { address: 'B3', row: 2, col: 1, value: 1200 },
        ],
      },
      {
        name: 'Summary',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'Metric' },
          { address: 'B1', row: 0, col: 1, value: 'Value' },
          { address: 'A2', row: 1, col: 0, value: 'Revenue' },
          { address: 'B2', row: 1, col: 1, formula: 'Inputs!B2*Inputs!B3', value: 60_000 },
          { address: 'A3', row: 2, col: 0, value: 'GeneratedAt' },
          { address: 'B3', row: 2, col: 1, formula: 'NOW()', value: 45_123 },
          { address: 'A4', row: 3, col: 0, value: 'CubeSales' },
          { address: 'B4', row: 3, col: 1, formula: 'CUBEVALUE("ThisWorkbookDataModel","[Measures].[Sales]")' },
        ],
      },
    ],
  })
}

interface WorkbookPackagePartScan {
  readonly definedNameCount: number
  readonly tableCount: number
  readonly pivotTableCount: number
  readonly chartCount: number
  readonly macroModuleCount: number
  readonly macroByteLength: number
  readonly externalLinkCount: number
}

interface WorkbookFormulaCacheInspection {
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: XlsxCacheInspectionLimit
  readonly suggestedReads: readonly string[]
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: {
    readonly missingCache: number
    readonly unsupportedRecalculation: number
  }
  readonly warnings: readonly string[]
  readonly formulas: readonly {
    readonly target: string
    readonly formula: string
    readonly hasCachedValue: boolean
  }[]
}

function inspectWorkbookFormulaCaches(workbook: XlsxWorkbookCells, inspectLimit: XlsxCacheInspectionLimit): WorkbookFormulaCacheInspection {
  const formulas = workbook.sheets.flatMap((sheet) =>
    sheet.cells.flatMap((cell) =>
      cell.formula && cell.formula.trim().length > 0
        ? [
            {
              target: `${sheet.name}!${cell.address}`,
              formula: cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`,
              hasCachedValue: cell.hasValue,
            },
          ]
        : [],
    ),
  )
  const normalizedLimit = normalizeInspectLimit(inspectLimit)
  const inspected = normalizedLimit === 'all' ? formulas : formulas.slice(0, normalizedLimit)
  const missingCache = inspected.filter((entry) => !entry.hasCachedValue).length
  const unsupportedRecalculation = inspected.filter((entry) => knownUnsupportedFunctionNamesFromFormula(entry.formula).length > 0).length
  const staleCachedFormulaCount = inspected.filter(
    (entry) => entry.hasCachedValue && volatileFunctionNamesFromFormula(entry.formula).length > 0,
  ).length
  const volatileFormulaCount = inspected.filter((entry) => volatileFunctionNamesFromFormula(entry.formula).length > 0).length
  const warnings = [
    ...(volatileFormulaCount > 0
      ? ['Volatile formulas were detected; cached formula values may depend on workbook calculation time.']
      : []),
    ...(unsupportedRecalculation > 0
      ? ['Unsupported formula families were detected; use xlsx-cache-doctor or an oracle harness before trusting cached values.']
      : []),
  ].toSorted()
  return {
    formulaCellCount: formulas.length,
    inspectedFormulaCellCount: inspected.length,
    uninspectedFormulaCellCount: formulas.length - inspected.length,
    inspectionLimit: normalizedLimit,
    suggestedReads: inspected.map((entry) => entry.target),
    staleCachedFormulaCount,
    cacheStatusSummary: {
      missingCache,
      unsupportedRecalculation,
    },
    warnings,
    formulas: inspected,
  }
}

function normalizeInspectLimit(limit: XlsxCacheInspectionLimit): XlsxCacheInspectionLimit {
  if (limit === 'all') {
    return limit
  }
  if (!Number.isInteger(limit) || limit < 1) {
    return defaultInspectLimit
  }
  return limit
}

function scanWorkbookPackageParts(zip: XlsxZipEntries): WorkbookPackagePartScan {
  const paths = Object.keys(zip)
  const workbookXml = getZipText(zip, 'xl/workbook.xml') ?? ''
  const externalReferenceCount = workbookXml.match(/<(?:[A-Za-z_][\w.-]*:)?externalReference\b/gu)?.length ?? 0
  const externalLinkPartCount = paths.filter((path) => /^xl\/externalLinks\/externalLink[0-9]+\.xml$/u.test(path)).length
  const macroPaths = paths.filter((path) => /(?:^|\/)vbaProject\.bin$/iu.test(path))
  return {
    definedNameCount: workbookXml.match(/<(?:[A-Za-z_][\w.-]*:)?definedName\b/gu)?.length ?? 0,
    tableCount: paths.filter((path) => /^xl\/tables\/[^/]+\.xml$/u.test(path)).length,
    pivotTableCount: paths.filter((path) => /^xl\/pivotTables\/[^/]+\.xml$/u.test(path)).length,
    chartCount: paths.filter((path) => /^xl\/charts\/[^/]+\.xml$/u.test(path) || /^xl\/chartsheets\/[^/]+\.xml$/u.test(path)).length,
    macroModuleCount: macroPaths.length,
    macroByteLength: macroPaths.reduce((sum, path) => sum + (zip[path]?.byteLength ?? 0), 0),
    externalLinkCount: Math.max(externalReferenceCount, externalLinkPartCount),
  }
}

function buildRisk(report: Omit<WorkbookCompatibilityReport, 'risk'>): WorkbookCompatibilityReport['risk'] {
  const highReasons: string[] = []
  const mediumReasons: string[] = []
  if (report.findings.unsupportedFunctions.length > 0) {
    highReasons.push(`unsupported functions: ${formatNamedCounts(report.findings.unsupportedFunctions)}`)
  }
  if (report.findings.externalLinks.count > 0) {
    highReasons.push(`external workbook links: ${report.findings.externalLinks.count.toString()}`)
  }
  if (report.findings.macroModules.count > 0) {
    highReasons.push(`VBA macro payloads preserved but not executed: ${report.findings.macroModules.count.toString()}`)
  }
  if (report.findings.pivotTables.unsupportedCount > 0) {
    highReasons.push(`unsupported pivot tables: ${report.findings.pivotTables.unsupportedCount.toString()}`)
  }
  if (report.findings.pivotTables.count > 0) {
    mediumReasons.push(`pivot tables require review: ${report.findings.pivotTables.count.toString()}`)
  }
  if (report.findings.volatileFunctions.length > 0) {
    mediumReasons.push(`volatile functions: ${formatNamedCounts(report.findings.volatileFunctions)}`)
  }
  if (report.findings.staleCachedFormulas.count > 0) {
    mediumReasons.push(`stale cached formulas: ${report.findings.staleCachedFormulas.count.toString()}`)
  }
  if (report.findings.missingCachedFormulaValues.count > 0) {
    mediumReasons.push(`missing cached formula values: ${report.findings.missingCachedFormulaValues.count.toString()}`)
  }
  if (report.findings.unsupportedRecalculations.count > 0) {
    mediumReasons.push(`unsupported recalculation results: ${report.findings.unsupportedRecalculations.count.toString()}`)
  }
  if (report.cacheInspection.uninspectedFormulaCellCount > 0) {
    mediumReasons.push(`uninspected formula cells: ${report.cacheInspection.uninspectedFormulaCellCount.toString()}`)
  }
  if (report.findings.warnings.length > 0) {
    mediumReasons.push(`import warnings: ${report.findings.warnings.length.toString()}`)
  }
  const reasons = highReasons.length > 0 ? highReasons : mediumReasons
  if (reasons.length === 0) {
    return {
      level: 'low',
      reasons: ['No known workbook compatibility risk signals were detected by this report.'],
    }
  }
  return {
    level: highReasons.length > 0 ? 'high' : 'medium',
    reasons,
  }
}

function parseCliArgs(args: readonly string[], commandName: string): WorkbookCompatibilityCliOptions {
  const demo = args.includes('--demo')
  const inputPath = demo ? undefined : args[0]
  if (!demo && (!inputPath || inputPath.startsWith('-'))) {
    throw new Error('Expected input XLSX path or --demo')
  }

  const externalWorkbooks: CliExternalWorkbook[] = []
  let inspectLimit: XlsxCacheInspectionLimit = defaultInspectLimit
  let json = false

  for (let index = demo ? 0 : 1; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error(`Unexpected missing ${commandName} argument`)
    }
    switch (arg) {
      case '--demo':
        break
      case '--external-workbook':
        externalWorkbooks.push({ path: requireNextArg(args, index, '--external-workbook') })
        index += 1
        break
      case '--external-workbook-target':
        externalWorkbooks.push({
          path: requireNextArg(args, index, '--external-workbook-target'),
          target: requireNextArg(args, index + 1, '--external-workbook-target target'),
        })
        index += 2
        break
      case '--inspect-limit':
        inspectLimit = parseInspectLimit(requireNextArg(args, index, '--inspect-limit'))
        index += 1
        break
      case '--json':
        json = true
        break
      default:
        throw new Error(`Unknown ${commandName} option: ${arg}`)
    }
  }

  return {
    mode: demo ? 'demo' : 'file',
    inputPath,
    externalWorkbooks,
    inspectLimit,
    json,
  }
}

function readExternalWorkbookInputs(workbooks: readonly CliExternalWorkbook[]): XlsxExternalWorkbookInput[] {
  return workbooks.map((workbook) => ({
    bytes: readFileSync(workbook.path),
    fileName: basename(workbook.path),
    ...(workbook.target ? { target: workbook.target } : {}),
  }))
}

function knownUnsupportedFunctionNamesFromFormula(formula: string): readonly string[] {
  return extractFormulaFunctionNames(formula).filter(isKnownUnsupportedRiskFunction)
}

function volatileFunctionNamesFromFormula(formula: string): readonly string[] {
  const volatileNames = new Set<string>(volatileFunctionNames)
  return extractFormulaFunctionNames(formula).filter((name) => volatileNames.has(name))
}

function extractFormulaFunctionNames(formula: string): readonly string[] {
  const stripped = formulaWithoutStringLiterals(formula)
  const names = new Set<string>()
  for (const match of stripped.matchAll(/(?:^|[^A-Za-z0-9_.])([_A-Za-z][A-Za-z0-9_.]*)\s*\(/gu)) {
    const rawName = match[1]
    if (rawName) {
      names.add(normalizeFormulaFunctionName(rawName))
    }
  }
  return [...names].toSorted()
}

function formulaWithoutStringLiterals(formula: string): string {
  let stripped = ''
  let index = 0
  while (index < formula.length) {
    if (formula[index] !== '"') {
      stripped += formula[index]
      index += 1
      continue
    }
    stripped += ' '
    index += 1
    while (index < formula.length) {
      stripped += ' '
      if (formula[index] === '"' && formula[index + 1] === '"') {
        index += 2
        continue
      }
      if (formula[index] === '"') {
        index += 1
        break
      }
      index += 1
    }
  }
  return stripped
}

function normalizeFormulaFunctionName(name: string): string {
  return name
    .replace(/^_xlfn\./iu, '')
    .replace(/^_xludf\./iu, '')
    .toUpperCase()
}

function isKnownUnsupportedRiskFunction(name: string): boolean {
  return knownUnsupportedRiskFunctions.has(name) || name.startsWith('_XLD') || name.startsWith('_XLUDF') || name.startsWith('_XLL')
}

function countNamedValues(values: readonly string[]): readonly WorkbookCompatibilityNamedCount[] {
  const counts = new Map<string, number>()
  for (const value of values) {
    counts.set(value, (counts.get(value) ?? 0) + 1)
  }
  return [...counts.entries()]
    .map(([name, count]) => ({ name, count }))
    .toSorted((left, right) => right.count - left.count || left.name.localeCompare(right.name))
}

function formatNamedCounts(counts: readonly WorkbookCompatibilityNamedCount[]): string {
  return counts.map((entry) => `${entry.name} (${entry.count.toString()})`).join(', ')
}

function renderHumanReport(report: WorkbookCompatibilityReport): string {
  return [
    `Workbook analyzed. Risk level: ${report.risk.level.toUpperCase()}`,
    'Findings:',
    `- Unsupported functions: ${report.findings.unsupportedFunctions.length > 0 ? formatNamedCounts(report.findings.unsupportedFunctions) : '0'}`,
    `- External links: ${report.findings.externalLinks.count.toString()}`,
    `- Macro modules: ${report.findings.macroModules.count.toString()}`,
    `- Pivot tables: ${report.findings.pivotTables.count.toString()}`,
    `- Volatile functions: ${report.findings.volatileFunctions.length > 0 ? formatNamedCounts(report.findings.volatileFunctions) : '0'}`,
    `- Formula cells: ${report.workbook.formulaCellCount.toString()}`,
    `- Stale cached formulas: ${report.findings.staleCachedFormulas.count.toString()}`,
    `- Missing cached formula values: ${report.findings.missingCachedFormulaValues.count.toString()}`,
    'This report identifies workbook features that may require investigation before using Bilig in a service or agent workflow. It is not an Excel compatibility certification.',
    '',
  ].join('\n')
}

function renderHelp(commandName: string): string {
  return `Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json]

Inspect workbook features that may require investigation before using Bilig in a
Node service or agent workflow. This is not an Excel compatibility certificate.

Options:
  --demo                  Generate an intentionally risky workbook and report on it.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectLimit}.
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --json                  Print the machine-readable JSON report.
  --help, -h              Show this help.
`
}

function parseInspectLimit(raw: string): XlsxCacheInspectionLimit {
  if (raw === 'all') {
    return raw
  }
  const value = Number(raw)
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`Expected --inspect-limit to be "all" or a positive integer, received: ${raw}`)
  }
  return value
}

function requireNextArg(args: readonly string[], index: number, option: string): string {
  const value = args[index + 1]
  if (!value || value.startsWith('--')) {
    throw new Error(`Expected value after ${option}`)
  }
  return value
}

function requireInputPath(options: WorkbookCompatibilityCliOptions): string {
  if (!options.inputPath) {
    throw new Error('Expected input XLSX path')
  }
  return options.inputPath
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}
