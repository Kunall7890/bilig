import { readFileSync } from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import type { RawCellContent } from '@bilig/headless'
import type { XlsxExternalWorkbookInput } from '@bilig/headless/xlsx'
import { ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import {
  exportXlsx,
  inspectXlsxCache,
  recalculateXlsxToFile,
  WorkPaper,
  type XlsxFormulaRecalcCellValue,
  type XlsxFormulaRecalcEdit,
} from './index.js'

interface CliExternalWorkbook {
  readonly path: string
  readonly target?: string
}

interface CliOptions {
  readonly mode: 'file' | 'demo'
  readonly inputPath: string | undefined
  readonly outputPath: string
  readonly edits: readonly XlsxFormulaRecalcEdit[]
  readonly reads: readonly string[]
  readonly externalWorkbooks: readonly CliExternalWorkbook[]
  readonly inspect: boolean
  readonly inspectLimit: number | 'all'
  readonly timeoutMs?: number
  readonly json: boolean
}

interface GithubActionWorkflowOptions {
  readonly workbooks: string
  readonly changedFilesOnly: boolean
  readonly failOnStale: boolean
  readonly inspectLimit: CliOptions['inspectLimit']
  readonly jsonOutput: string
  readonly markdownOutput: string
  readonly packageVersion: string
  readonly workflowName: string
}

export interface XlsxFormulaRecalcCliContext {
  readonly commandName?: string
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

const defaultInspectFormulaLimit = 'all'
const cacheDoctorCommandName = 'xlsx-cache-doctor'
const printGithubActionOption = '--print-github-action'
const defaultGithubActionPackageVersion = readPackageVersion()

export function runXlsxFormulaRecalcCli(args: readonly string[], context: XlsxFormulaRecalcCliContext = {}): number {
  const commandName = context.commandName ?? 'xlsx-recalc'
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
      printHelp(commandName, writeStdout)
      return 0
    }

    if (commandName === cacheDoctorCommandName && args.includes(printGithubActionOption)) {
      printGithubActionWorkflow(parseGithubActionWorkflowArgs(args, commandName), writeStdout)
      return 0
    }

    const options = parseCliArgs(normalizeCliArgsForCommand(args, commandName), commandName)
    const input = inputBytesForCli(options, commandName)
    const inputName =
      options.mode === 'demo'
        ? commandName === cacheDoctorCommandName
          ? 'bilig-cache-doctor-stale-demo.xlsx'
          : 'bilig-formula-recalc-demo.xlsx'
        : basename(requireInputPath(options))
    const externalWorkbooks = readExternalWorkbookInputs(options.externalWorkbooks)
    if (options.inspect) {
      printInspectionSummary({ input, inputName, externalWorkbooks, options, writeStdout })
      return 0
    }
    const result = recalculateXlsxToFile(input, {
      fileName: inputName,
      ...(externalWorkbooks.length > 0 ? { externalWorkbooks } : {}),
      edits: options.edits,
      reads: options.reads,
      ...(options.timeoutMs === undefined ? {} : { config: { evaluationTimeoutMs: options.timeoutMs } }),
      outputPath: options.outputPath,
    })

    const summary = {
      mode: options.mode,
      input: options.inputPath ?? 'generated demo workbook',
      output: options.outputPath,
      edits: options.edits.length,
      externalWorkbooks: externalWorkbooks.length,
      reads: result.reads,
      warnings: result.warnings,
      ...(result.diagnostics ? { diagnostics: result.diagnostics } : {}),
      commandSucceeded: true,
      recalculationCompleted: true,
      excelParity: 'not_proven',
      ...(options.mode === 'demo'
        ? {
            expectedReadback: { 'Summary!B2': 72_000 },
            expectedValueMatched: numericReadValue(result.reads['Summary!B2']) === 72_000,
          }
        : {}),
    }

    if (options.json) {
      writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
    } else {
      writeStdout(`Recalculated ${summary.input} -> ${options.outputPath}\n`)
      for (const [target, value] of Object.entries(result.reads)) {
        writeStdout(`${target}: ${JSON.stringify(value)}\n`)
      }
      if (result.warnings.length > 0) {
        writeStdout(`Warnings: ${result.warnings.length.toString()}\n`)
      }
    }
    return 0
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

function readPackageVersion(): string {
  const packageJsonPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json')
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`Expected package.json object at ${packageJsonPath}`)
  }
  const version = (parsed as { readonly version?: unknown }).version
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error(`Expected package.json version at ${packageJsonPath}`)
  }
  return version
}

function parseGithubActionWorkflowArgs(args: readonly string[], commandName: string): GithubActionWorkflowOptions {
  let workbooks: string | undefined
  let changedFilesOnly = true
  let failOnStale = false
  let inspectLimit: CliOptions['inspectLimit'] = defaultInspectFormulaLimit
  let jsonOutput = '${{ runner.temp }}/xlsx-cache-doctor.json'
  let markdownOutput = '${{ runner.temp }}/xlsx-cache-doctor.md'
  let packageVersion = defaultGithubActionPackageVersion
  let workflowName = 'xlsx-cache-doctor'

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error(`Unexpected missing ${commandName} argument`)
    }
    switch (arg) {
      case printGithubActionOption:
        break
      case '--workbook':
      case '--workbooks':
        workbooks = requireNextArg(args, index, arg)
        index += 1
        break
      case '--changed-files-only':
        changedFilesOnly = parseBooleanOption(requireNextArg(args, index, '--changed-files-only'), '--changed-files-only')
        index += 1
        break
      case '--fail-on-stale':
        failOnStale = parseBooleanOption(requireNextArg(args, index, '--fail-on-stale'), '--fail-on-stale')
        index += 1
        break
      case '--inspect-limit':
        inspectLimit = parseInspectLimit(requireNextArg(args, index, '--inspect-limit'))
        index += 1
        break
      case '--json-output':
        jsonOutput = requireNextArg(args, index, '--json-output')
        index += 1
        break
      case '--markdown-output':
        markdownOutput = requireNextArg(args, index, '--markdown-output')
        index += 1
        break
      case '--package-version':
        packageVersion = requireNextArg(args, index, '--package-version')
        index += 1
        break
      case '--workflow-name':
        workflowName = requireNextArg(args, index, '--workflow-name')
        index += 1
        break
      default:
        if (arg.startsWith('-')) {
          throw new Error(`Unknown ${commandName} option for ${printGithubActionOption}: ${arg}`)
        }
        if (workbooks !== undefined) {
          throw new Error(`Unexpected extra workbook glob for ${printGithubActionOption}: ${arg}`)
        }
        workbooks = arg
    }
  }

  if (!workbooks) {
    throw new Error(`Expected workbook path or glob after ${printGithubActionOption}`)
  }

  return {
    workbooks,
    changedFilesOnly,
    failOnStale,
    inspectLimit,
    jsonOutput,
    markdownOutput,
    packageVersion,
    workflowName,
  }
}

function printGithubActionWorkflow(options: GithubActionWorkflowOptions, writeStdout: (text: string) => void): void {
  writeStdout(
    [
      `name: ${yamlDoubleQuote(options.workflowName)}`,
      '',
      'on:',
      '  pull_request:',
      '    paths:',
      '      - "**/*.xlsx"',
      '  workflow_dispatch:',
      '',
      'permissions:',
      '  contents: read',
      '',
      'jobs:',
      '  inspect-xlsx-formula-caches:',
      '    runs-on: ubuntu-latest',
      '    steps:',
      '      - uses: actions/checkout@v5',
      '        with:',
      '          fetch-depth: 0',
      '',
      '      - uses: actions/setup-node@v6',
      '        with:',
      '          node-version: "22"',
      '          package-manager-cache: false',
      '',
      '      - id: cache-doctor',
      '        uses: proompteng/bilig@v1',
      '        with:',
      `          workbooks: ${yamlDoubleQuote(options.workbooks)}`,
      `          changed-files-only: ${yamlDoubleQuote(String(options.changedFilesOnly))}`,
      `          package-version: ${yamlDoubleQuote(options.packageVersion)}`,
      `          inspect-limit: ${yamlDoubleQuote(String(options.inspectLimit))}`,
      `          json-output: ${yamlDoubleQuote(options.jsonOutput)}`,
      `          markdown-output: ${yamlDoubleQuote(options.markdownOutput)}`,
      `          fail-on-stale: ${yamlDoubleQuote(String(options.failOnStale))}`,
      '',
      '      - uses: actions/upload-artifact@v4',
      '        if: always()',
      '        with:',
      '          name: xlsx-cache-doctor-report',
      '          path: |',
      '            ${{ steps.cache-doctor.outputs.json }}',
      '            ${{ steps.cache-doctor.outputs.markdown }}',
      '',
    ].join('\n'),
  )
}

function parseBooleanOption(raw: string, option: string): boolean {
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  throw new Error(`Expected ${option} to be "true" or "false", received: ${raw}`)
}

function yamlDoubleQuote(value: string): string {
  return JSON.stringify(value)
}

function normalizeCliArgsForCommand(args: readonly string[], commandName: string): readonly string[] {
  if (commandName !== cacheDoctorCommandName || args.includes('--inspect') || hasExplicitRecalcIntent(args)) {
    return args
  }
  return [...args, '--inspect']
}

function hasExplicitRecalcIntent(args: readonly string[]): boolean {
  return args.includes('--read') || args.includes('--out') || args.includes('-o')
}

function parseCliArgs(args: readonly string[], commandName: string): CliOptions {
  const demo = args.includes('--demo')
  const inputPath = demo ? undefined : args[0]
  if (!demo && (!inputPath || inputPath.startsWith('-'))) {
    throw new Error('Expected input XLSX path or --demo')
  }

  const edits: XlsxFormulaRecalcEdit[] = []
  const reads: string[] = []
  const externalWorkbooks: CliExternalWorkbook[] = []
  let outputPath: string | undefined
  let inspect = false
  let inspectLimit: CliOptions['inspectLimit'] = defaultInspectFormulaLimit
  let timeoutMs: number | undefined
  let json = false

  for (let index = demo ? 0 : 1; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error(`Unexpected missing ${commandName} argument`)
    }
    switch (arg) {
      case '--demo':
        break
      case '--set':
        edits.push(parseEdit(requireNextArg(args, index, '--set')))
        index += 1
        break
      case '--read':
        reads.push(requireNextArg(args, index, '--read'))
        index += 1
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
      case '--inspect':
        inspect = true
        break
      case '--inspect-limit':
        inspectLimit = parseInspectLimit(requireNextArg(args, index, '--inspect-limit'))
        index += 1
        break
      case '--timeout-ms':
        timeoutMs = parsePositiveIntegerOption(requireNextArg(args, index, '--timeout-ms'), '--timeout-ms')
        index += 1
        break
      case '--out':
      case '-o':
        outputPath = requireNextArg(args, index, arg)
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
    outputPath:
      outputPath ?? (demo ? 'bilig-formula-recalc-demo.xlsx' : defaultOutputPath(requireDefined(inputPath, 'Expected input XLSX path'))),
    edits: edits.length > 0 ? edits : demoDefaultEdits(demo),
    reads: reads.length > 0 ? reads : demoDefaultReads(demo),
    externalWorkbooks,
    inspect,
    inspectLimit,
    ...(timeoutMs === undefined ? {} : { timeoutMs }),
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

interface PrintInspectionSummaryInput {
  readonly input: Uint8Array | Buffer
  readonly inputName: string
  readonly externalWorkbooks: readonly XlsxExternalWorkbookInput[]
  readonly options: CliOptions
  readonly writeStdout: (text: string) => void
}

function printInspectionSummary(args: PrintInspectionSummaryInput): void {
  const inspection = inspectXlsxCache(args.input, {
    fileName: args.inputName,
    ...(args.externalWorkbooks.length > 0 ? { externalWorkbooks: args.externalWorkbooks } : {}),
    edits: args.options.edits,
    inspectLimit: args.options.inspectLimit,
    ...(args.options.timeoutMs === undefined ? {} : { config: { evaluationTimeoutMs: args.options.timeoutMs } }),
  })
  const summary = {
    schemaVersion: inspection.schemaVersion,
    mode: args.options.mode,
    input: args.options.inputPath ?? 'generated demo workbook',
    edits: args.options.edits.length,
    externalWorkbooks: args.externalWorkbooks.length,
    sheetNames: inspection.sheetNames,
    formulaCellCount: inspection.formulaCellCount,
    inspectedFormulaCellCount: inspection.inspectedFormulaCellCount,
    uninspectedFormulaCellCount: inspection.uninspectedFormulaCellCount,
    inspectionLimit: inspection.inspectionLimit,
    staleCachedFormulaCount: inspection.staleCachedFormulaCount,
    cacheStatusSummary: inspection.cacheStatusSummary,
    suggestedReads: inspection.suggestedReads,
    formulas: inspection.formulas,
    warnings: inspection.warnings,
    ...(inspection.diagnostics ? { diagnostics: inspection.diagnostics } : {}),
    commandSucceeded: true,
    inspectionCompleted: inspection.inspectionCompleted,
    recalculationCompleted: inspection.recalculationCompleted,
    excelParity: inspection.excelParity,
  }

  if (args.options.json) {
    args.writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
    return
  }

  args.writeStdout(`Inspected ${summary.input}\n`)
  args.writeStdout(`Sheets: ${summary.sheetNames.join(', ')}\n`)
  args.writeStdout(`Formula cells: ${summary.formulaCellCount.toString()}\n`)
  args.writeStdout(`Inspected formula cells: ${summary.inspectedFormulaCellCount.toString()}\n`)
  args.writeStdout(`Uninspected formula cells: ${summary.uninspectedFormulaCellCount.toString()}\n`)
  args.writeStdout(`Stale cached formula cells: ${summary.staleCachedFormulaCount.toString()}\n`)
  args.writeStdout(`Fresh cached formula cells: ${summary.cacheStatusSummary.fresh.toString()}\n`)
  args.writeStdout(`Missing cached formula values: ${summary.cacheStatusSummary.missingCache.toString()}\n`)
  args.writeStdout(`Unsupported recalculation results: ${summary.cacheStatusSummary.unsupportedRecalculation.toString()}\n`)
  if (summary.suggestedReads.length > 0) {
    args.writeStdout(`Suggested reads: ${summary.suggestedReads.join(', ')}\n`)
    args.writeStdout(`${suggestedRecalcCommand(args.options, summary.suggestedReads)}\n`)
  }
  if (summary.warnings.length > 0) {
    args.writeStdout(`Warnings: ${summary.warnings.length.toString()}\n`)
  }
  args.writeStdout('If the important formula is missing, rerun with --inspect-limit all and --read <Sheet!Cell>.\n')
}

function numericReadValue(value: XlsxFormulaRecalcCellValue | undefined): number | undefined {
  return value !== undefined &&
    typeof value === 'object' &&
    value !== null &&
    'tag' in value &&
    value.tag === ValueTag.Number &&
    'value' in value
    ? value.value
    : undefined
}

function suggestedRecalcCommand(options: CliOptions, reads: readonly string[]): string {
  const command = ['xlsx-recalc']
  if (options.mode === 'demo') {
    command.push('--demo')
  } else if (options.inputPath) {
    command.push(shellQuote(options.inputPath))
  }
  for (const edit of options.edits) {
    command.push('--set', shellQuote(`${edit.target}=${String(edit.value)}`))
  }
  for (const read of reads.slice(0, 5)) {
    command.push('--read', shellQuote(read))
  }
  command.push('--json')
  return command.join(' ')
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function inputBytesForCli(options: CliOptions, commandName: string): Uint8Array {
  if (options.mode !== 'demo') {
    return readFileSync(requireInputPath(options))
  }
  return commandName === cacheDoctorCommandName && options.inspect ? buildStaleCacheDoctorDemoWorkbookBytes() : buildDemoWorkbookBytes()
}

function buildDemoWorkbookBytes(): Uint8Array {
  const sourceWorkbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['Price', 1200],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Revenue', '=Inputs!B2*Inputs!B3'],
    ],
  })
  try {
    return exportXlsx(sourceWorkbook.exportSnapshot())
  } finally {
    sourceWorkbook.dispose()
  }
}

function buildStaleCacheDoctorDemoWorkbookBytes(): Uint8Array {
  return replaceWorksheetCellXml(
    buildDemoWorkbookBytes(),
    'xl/worksheets/sheet2.xml',
    'B2',
    '<c r="B2"><f>Inputs!B2*Inputs!B3</f><v>60000</v></c>',
  )
}

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  const nextXml = xml.replace(new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'), replacement)
  if (nextXml === xml) {
    throw new Error(`Demo XLSX is missing ${path} ${address}`)
  }
  zip[path] = strToU8(nextXml)
  return zipSync(zip)
}

function demoDefaultEdits(enabled: boolean): readonly XlsxFormulaRecalcEdit[] {
  return enabled
    ? [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ]
    : []
}

function demoDefaultReads(enabled: boolean): readonly string[] {
  return enabled ? ['Summary!B2'] : []
}

function parseEdit(raw: string): XlsxFormulaRecalcEdit {
  const separator = raw.indexOf('=')
  if (separator <= 0) {
    throw new Error(`Expected --set value in Target=Value form, received: ${raw}`)
  }
  return {
    target: raw.slice(0, separator),
    value: parseRawCellContent(raw.slice(separator + 1)),
  }
}

function parseRawCellContent(raw: string): RawCellContent {
  if (raw === 'null') {
    return null
  }
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  if (/^-?(?:\d+|\d*\.\d+)(?:e[+-]?\d+)?$/iu.test(raw)) {
    return Number(raw)
  }
  return raw
}

function parseInspectLimit(raw: string): CliOptions['inspectLimit'] {
  if (raw === 'all') {
    return raw
  }
  const value = Number(raw)
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`Expected --inspect-limit to be "all" or a positive integer, received: ${raw}`)
  }
  return value
}

function parsePositiveIntegerOption(raw: string, option: string): number {
  const value = Number(raw)
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`Expected ${option} to be a positive integer, received: ${raw}`)
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

function requireInputPath(options: CliOptions): string {
  return requireDefined(options.inputPath, 'Expected input XLSX path')
}

function requireDefined(value: string | undefined, message: string): string {
  if (!value) {
    throw new Error(message)
  }
  return value
}

function defaultOutputPath(inputPath: string): string {
  const extension = extname(inputPath)
  const base = extension.length > 0 ? basename(inputPath, extension) : basename(inputPath)
  return join(dirname(inputPath), `${base}.recalculated${extension || '.xlsx'}`)
}

function printHelp(commandName: string, writeStdout: (text: string) => void): void {
  if (commandName === cacheDoctorCommandName) {
    writeStdout(`Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json]

Diagnose stale cached formula values without writing an output XLSX. This is
the memorable alias for: xlsx-recalc <input.xlsx> --inspect.

Options:
  --demo                  Generate a tiny workbook and inspect its formula cache.
  ${printGithubActionOption} <workbook-glob>
                          Print a ready-to-commit GitHub Actions workflow that uses proompteng/bilig@v1.
  --set <Sheet!A1=value>  Edit an input cell before diagnosis. Repeatable.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectFormulaLimit}.
  --timeout-ms <n>        Formula evaluation timeout in milliseconds.
  --fail-on-stale <true|false>
                          With ${printGithubActionOption}, decide whether the generated workflow fails pull requests. Defaults to false.
  --changed-files-only <true|false>
                          With ${printGithubActionOption}, inspect only changed XLSX files. Defaults to true.
  --json-output <path>    With ${printGithubActionOption}, set the JSON report path. Defaults to \${{ runner.temp }}/xlsx-cache-doctor.json.
  --markdown-output <path>
                          With ${printGithubActionOption}, set the Markdown report path. Defaults to \${{ runner.temp }}/xlsx-cache-doctor.md.
  --package-version <version>
                          With ${printGithubActionOption}, pin @bilig/xlsx-formula-recalc in the generated workflow. Defaults to ${defaultGithubActionPackageVersion}.
  --workflow-name <name>  With ${printGithubActionOption}, set the generated workflow name.
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
    return
  }

  writeStdout(`Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json] [--out demo.recalculated.xlsx]

Options:
  --demo                  Generate a tiny workbook, edit inputs, recalculate, and write proof XLSX.
  --set <Sheet!A1=value>  Edit an input cell before recalculation. Repeatable.
  --read <Sheet!A1>       Read a recalculated cell after edits. Repeatable.
  --inspect               Inspect formula cells, stale cached values, and suggested --read targets.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectFormulaLimit}.
  --timeout-ms <n>        Formula evaluation timeout in milliseconds.
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --out, -o <path>        Output XLSX path. Defaults to <input>.recalculated.xlsx.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
}
