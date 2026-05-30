import { readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'

import type { RawCellContent } from '@bilig/headless'
import type { XlsxExternalWorkbookInput } from '@bilig/headless/xlsx'
import { formatErrorCode, ValueTag } from '@bilig/protocol'
import { exportXlsx, importXlsx, recalculateXlsx, WorkPaper, type XlsxFormulaRecalcCellValue, type XlsxFormulaRecalcEdit } from './index.js'

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
  readonly json: boolean
}

export interface XlsxFormulaRecalcCliContext {
  readonly commandName?: string
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

const repoStarUrl = 'https://github.com/proompteng/bilig/stargazers'
const releaseWatchUrl = 'https://github.com/proompteng/bilig/subscription'
const adoptionBlockerUrl = 'https://github.com/proompteng/bilig/discussions/new?category=general'
const defaultInspectFormulaLimit = 'all'
const cacheDoctorCommandName = 'xlsx-cache-doctor'

export function runXlsxFormulaRecalcCli(args: readonly string[], context: XlsxFormulaRecalcCliContext = {}): number {
  const commandName = context.commandName ?? 'xlsx-recalc'
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
      printHelp(commandName, writeStdout)
      return 0
    }

    const options = parseCliArgs(normalizeCliArgsForCommand(args, commandName), commandName)
    const input = options.mode === 'demo' ? buildDemoWorkbookBytes() : readFileSync(requireInputPath(options))
    const inputName = options.mode === 'demo' ? 'bilig-formula-recalc-demo.xlsx' : basename(requireInputPath(options))
    const externalWorkbooks = readExternalWorkbookInputs(options.externalWorkbooks)
    if (options.inspect) {
      printInspectionSummary({ input, inputName, externalWorkbooks, options, writeStdout })
      return 0
    }
    const result = recalculateXlsx(input, {
      fileName: inputName,
      ...(externalWorkbooks.length > 0 ? { externalWorkbooks } : {}),
      edits: options.edits,
      reads: options.reads,
    })
    writeFileSync(options.outputPath, result.xlsx)

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
      writeStdout(`If this fixed your stale XLSX formula path, star or bookmark Bilig: ${repoStarUrl}\n`)
      writeStdout(`If it almost worked, open the concrete workbook blocker: ${adoptionBlockerUrl}\n`)
      writeStdout(`Watch formula and compatibility releases: ${releaseWatchUrl}\n`)
    }
    return 0
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
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
  const importOptions =
    args.externalWorkbooks.length > 0
      ? {
          externalWorkbooks: args.externalWorkbooks,
          externalLinkCacheArtifactMode: 'replace-refreshed' as const,
        }
      : {}
  const imported = importXlsx(args.input, args.inputName, importOptions)
  const formulaCells = collectFormulaCells(imported.snapshot)
  const inspectedFormulaCells = args.options.inspectLimit === 'all' ? formulaCells : formulaCells.slice(0, args.options.inspectLimit)
  const uninspectedFormulaCellCount = formulaCells.length - inspectedFormulaCells.length
  const suggestedReads = inspectedFormulaCells.map((cell) => cell.target)
  const recalculated = recalculateXlsx(args.input, {
    fileName: args.inputName,
    ...(args.externalWorkbooks.length > 0 ? { externalWorkbooks: args.externalWorkbooks } : {}),
    edits: args.options.edits,
    reads: suggestedReads,
  })
  const formulas = inspectedFormulaCells.map((cell) => {
    const recalculatedValue = recalculated.reads[cell.target]
    const literalRecalculatedValue = literalValueForInspection(recalculatedValue)
    const staleCachedValue =
      cell.cachedValue === undefined || literalRecalculatedValue === undefined
        ? null
        : !literalValuesEqual(cell.cachedValue, literalRecalculatedValue)
    return {
      target: cell.target,
      formula: cell.formula,
      ...(cell.cachedValue !== undefined ? { cachedValue: cell.cachedValue } : {}),
      recalculatedValue,
      ...(literalRecalculatedValue !== undefined ? { literalRecalculatedValue } : {}),
      staleCachedValue,
    }
  })
  const staleCachedFormulaCount = formulas.filter((formula) => formula.staleCachedValue === true).length
  const summary = {
    mode: args.options.mode,
    input: args.options.inputPath ?? 'generated demo workbook',
    edits: args.options.edits.length,
    externalWorkbooks: args.externalWorkbooks.length,
    sheetNames: imported.sheetNames,
    formulaCellCount: formulaCells.length,
    inspectedFormulaCellCount: inspectedFormulaCells.length,
    uninspectedFormulaCellCount,
    inspectionLimit: args.options.inspectLimit,
    staleCachedFormulaCount,
    suggestedReads,
    formulas,
    warnings: recalculated.warnings,
    ...(recalculated.diagnostics ? { diagnostics: recalculated.diagnostics } : {}),
    commandSucceeded: true,
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
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
  if (suggestedReads.length > 0) {
    args.writeStdout(`Suggested reads: ${suggestedReads.join(', ')}\n`)
    args.writeStdout(`${suggestedRecalcCommand(args.options, suggestedReads)}\n`)
  }
  if (summary.warnings.length > 0) {
    args.writeStdout(`Warnings: ${summary.warnings.length.toString()}\n`)
  }
  args.writeStdout(`If the important formula is missing or wrong, open the reduced workbook blocker: ${adoptionBlockerUrl}\n`)
}

interface FormulaInspectionCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: RawCellContent
}

type ImportedWorkbookSnapshot = Parameters<typeof WorkPaper.buildFromSnapshot>[0]

function collectFormulaCells(snapshot: ImportedWorkbookSnapshot): FormulaInspectionCell[] {
  const cells: FormulaInspectionCell[] = []
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

function literalValueForInspection(value: XlsxFormulaRecalcCellValue | undefined): RawCellContent | string | undefined {
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

function literalValuesEqual(left: RawCellContent, right: RawCellContent | string): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
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
  --set <Sheet!A1=value>  Edit an input cell before diagnosis. Repeatable.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectFormulaLimit}.
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
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --out, -o <path>        Output XLSX path. Defaults to <input>.recalculated.xlsx.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
}
