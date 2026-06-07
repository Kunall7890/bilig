import { readFileSync } from 'node:fs'
import { basename } from 'node:path'

import { WorkPaper } from '@bilig/headless'
import type { ImportedWorkbookDiagnostics, XlsxExternalWorkbookInput, XlsxImportOptions } from '@bilig/headless/xlsx'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'
import { replaceXlsxWorksheetCellXml } from '@bilig/xlsx'
import { inspectXlsxCache, type XlsxCacheInspectionLimit, type XlsxFormulaRecalcEdit } from './legacy-workpaper.js'

export const workbookCompatibilityReportSchemaVersion = 'bilig-workbook-compatibility-report.v1'

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
  readonly recalculationCompleted: true
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
  const imported = importXlsx(bytes, fileName, workbookCompatibilityImportOptions(externalWorkbooks))
  const metadata = imported.snapshot.workbook.metadata
  const cacheInspection = inspectXlsxCache(bytes, {
    fileName,
    ...(externalWorkbooks.length > 0 ? { externalWorkbooks } : {}),
    ...(options.edits ? { edits: options.edits } : {}),
    inspectLimit: options.inspectLimit ?? defaultInspectLimit,
  })
  const formulaAuditEntries = metadata?.formulaAudit?.formulas ?? []
  const unsupportedFunctions = countNamedValues([
    ...cacheInspection.formulas.flatMap((entry) => knownUnsupportedFunctionNamesFromFormula(entry.formula)),
    ...formulaAuditEntries.flatMap((entry) =>
      unsupportedAuditFunctionNamesFromFormula(entry.formula, entry.cacheStatus === 'unsupportedCached'),
    ),
  ])
  const volatileFunctions = countNamedValues(formulaAuditEntries.flatMap((entry) => volatileFunctionNamesFromFormula(entry.formula)))
  const macroModules = metadata?.macroPayloads ?? []
  const pivots = metadata?.pivots ?? []
  const unsupportedPivots = metadata?.unsupportedPivots ?? []
  const unsupportedDependencies = metadata?.unsupportedFormulaDependencies ?? []
  const externalWorkbookReferences = metadata?.externalWorkbookReferences ?? []
  const externalWorkbookHydration = imported.diagnostics?.externalWorkbookHydration
  const report: Omit<WorkbookCompatibilityReport, 'risk'> = {
    schemaVersion: workbookCompatibilityReportSchemaVersion,
    verified: true,
    input: {
      fileName,
      externalWorkbookCount: externalWorkbooks.length,
      inspectLimit: cacheInspection.inspectionLimit,
    },
    workbook: {
      sheetCount: imported.sheetNames.length,
      sheetNames: imported.sheetNames,
      nonEmptyCellCount: imported.snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0),
      formulaCellCount: cacheInspection.formulaCellCount,
      definedNameCount: metadata?.definedNames?.length ?? 0,
      tableCount: metadata?.tables?.length ?? 0,
      pivotTableCount: pivots.length + unsupportedPivots.length,
      chartCount:
        (metadata?.charts?.length ?? 0) + (metadata?.chartArtifacts?.parts.length ?? 0) + (metadata?.chartSheetArtifacts?.length ?? 0),
      macroModuleCount: macroModules.length,
    },
    findings: {
      unsupportedFunctions,
      externalLinks: {
        count: externalWorkbookReferences.length,
        unresolvedCount: unsupportedDependencies.reduce((sum, dependency) => sum + dependency.unresolvedExternalReferenceCount, 0),
        refreshedCount: numberValue(externalWorkbookHydration, 'refreshedBookIndices') ?? 0,
      },
      macroModules: {
        count: macroModules.length,
        byteLength: macroModules.reduce((sum, payload) => sum + payload.byteLength, 0),
      },
      volatileFunctions,
      pivotTables: {
        count: pivots.length + unsupportedPivots.length,
        unsupportedCount: unsupportedPivots.length,
        cacheOnlyCount: pivots.filter((pivot) => pivot.cacheOnly === true || pivot.sourceKind === 'external-cache-only').length,
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
      warnings: [...new Set([...imported.warnings, ...cacheInspection.warnings])].toSorted(),
    },
    cacheInspection: {
      inspectedFormulaCellCount: cacheInspection.inspectedFormulaCellCount,
      uninspectedFormulaCellCount: cacheInspection.uninspectedFormulaCellCount,
      inspectionLimit: cacheInspection.inspectionLimit,
      suggestedReads: cacheInspection.suggestedReads,
    },
    ...(imported.diagnostics || cacheInspection.diagnostics ? { diagnostics: cacheInspection.diagnostics ?? imported.diagnostics } : {}),
    commandSucceeded: true,
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
    limitations: [
      'This report identifies workbook features that may require investigation before using Bilig in a service or agent workflow.',
      'It is not an Excel compatibility certification.',
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

function workbookCompatibilityImportOptions(externalWorkbooks: readonly XlsxExternalWorkbookInput[]): XlsxImportOptions {
  return externalWorkbooks.length > 0
    ? { externalWorkbooks, externalLinkCacheArtifactMode: 'replace-refreshed' }
    : { preferNativeSimpleImport: true }
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
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['Price', 1200],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Revenue', '=Inputs!B2*Inputs!B3'],
      ['GeneratedAt', '=NOW()'],
      ['CubeSales', '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Sales]")'],
    ],
  })
  try {
    let bytes = exportXlsx(workbook.exportSnapshot())
    bytes = replaceWorksheetCellXml(bytes, 'xl/worksheets/sheet2.xml', 'B3', '<c r="B3"><f>NOW()</f><v>45123</v></c>')
    return replaceWorksheetCellXml(
      bytes,
      'xl/worksheets/sheet2.xml',
      'B4',
      '<c r="B4"><f>CUBEVALUE(&quot;ThisWorkbookDataModel&quot;,&quot;[Measures].[Sales]&quot;)</f><v>60000</v></c>',
    )
  } finally {
    workbook.dispose()
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

function unsupportedAuditFunctionNamesFromFormula(formula: string, fromUnsupportedAudit: boolean): readonly string[] {
  if (!fromUnsupportedAudit) {
    return []
  }
  return extractFormulaFunctionNames(formula)
    .filter((name) => !isKnownUnsupportedRiskFunction(name))
    .slice(0, 1)
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

function numberValue(source: unknown, key: string): number | undefined {
  if (key === 'refreshedBookIndices' && isRecord(source)) {
    const value = source[key]
    return Array.isArray(value) ? value.length : undefined
  }
  if (!isRecord(source)) {
    return undefined
  }
  const value = source[key]
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
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

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  return replaceXlsxWorksheetCellXml(bytes, {
    path,
    address,
    replacement,
    missingMessage: `Demo XLSX is missing ${path} ${address}`,
  })
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
