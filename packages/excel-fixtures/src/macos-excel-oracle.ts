import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { basename, extname, join } from 'node:path'

import { ErrorCode } from '@bilig/protocol'

import type { NormalizedFormulaValue } from './oracle-harness.js'

export const defaultMacosExcelAppPath = '/Applications/Microsoft Excel.app' as const

export interface MacosExcelOracleFormulaCell {
  readonly address: string
  readonly formula: string
}

export type MacosExcelLinkUpdateMode = 'all' | 'external' | 'never' | 'remote'
export type MacosExcelAutoFilterOperator = 'and' | 'filterByValue' | 'or'

export interface MacosExcelRecalculationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly valueCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
}

export interface MacosExcelRecalculationOracleResult {
  readonly excelVersion: string
  readonly rawValues: readonly string[]
  readonly values: readonly NormalizedFormulaValue[]
}

export interface MacosExcelInspectionOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly inspectCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
  readonly refreshAll?: boolean
}

export type MacosExcelStructuralOperation =
  | { readonly kind: 'insertRows'; readonly range: string }
  | { readonly kind: 'insertColumns'; readonly range: string }
  | { readonly kind: 'deleteRows'; readonly range: string }
  | { readonly kind: 'deleteColumns'; readonly range: string }
  | { readonly kind: 'setCellValue'; readonly address: string; readonly value: string | number | boolean }
  | { readonly kind: 'clearCell'; readonly address: string }
  | { readonly kind: 'moveRows'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'moveColumns'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'createDataTable'; readonly range: string; readonly rowInput?: string; readonly columnInput?: string }
  | {
      readonly kind: 'applyAutoFilter'
      readonly range: string
      readonly field: number
      readonly criteria1?: string | number | boolean
      readonly operator?: MacosExcelAutoFilterOperator
      readonly criteria2?: string | number | boolean
      readonly visibleDropDown?: boolean
    }

export interface MacosExcelStructuralOperationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly operations: readonly MacosExcelStructuralOperation[]
  readonly inspectCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly formulaCells?: readonly MacosExcelOracleFormulaCell[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
  readonly refreshAll?: boolean
}

export interface MacosExcelCellInspection {
  readonly address: string
  readonly formula?: string
  readonly rawValue: string
  readonly value: NormalizedFormulaValue
}

export interface MacosExcelInspectionOracleResult {
  readonly cells: readonly MacosExcelCellInspection[]
  readonly excelVersion: string
}

export function isMacosExcelInstalled(appPath: string = defaultMacosExcelAppPath): boolean {
  return existsSync(appPath)
}

export function runMacosExcelRecalculationOracle(request: MacosExcelRecalculationOracleRequest): MacosExcelRecalculationOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-')
  const scriptPath = join(tempDir, 'recalculate.scpt')
  try {
    const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
    writeFileSync(scriptPath, createMacosExcelRecalculationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelRecalculationOutput(rawOutput, request.valueCells.length)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

export function runMacosExcelInspectionOracle(request: MacosExcelInspectionOracleRequest): MacosExcelInspectionOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-inspect-')
  const scriptPath = join(tempDir, 'inspect.scpt')
  try {
    const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
    writeFileSync(scriptPath, createMacosExcelInspectionAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelInspectionOutput(rawOutput, request.inspectCells)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

export function runMacosExcelStructuralOperationOracle(
  request: MacosExcelStructuralOperationOracleRequest,
): MacosExcelInspectionOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-structure-')
  const scriptPath = join(tempDir, 'structure.scpt')
  try {
    const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
    writeFileSync(scriptPath, createMacosExcelStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelInspectionOutput(rawOutput, request.inspectCells)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function createMacosExcelOracleTempDir(prefix: string): string {
  const root = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp/bilig-excel-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function stageWorkbookForMacosExcelOracle(workbookPath: string, tempDir: string): string {
  const workbookBasename = basename(workbookPath)
  const extension = extname(workbookBasename)
  const baseNameWithoutExtension = extension ? workbookBasename.slice(0, -extension.length) : workbookBasename
  const stagedWorkbookPath = join(tempDir, `${baseNameWithoutExtension}-${String(process.pid)}-${String(Date.now())}${extension}`)
  copyFileSync(workbookPath, stagedWorkbookPath)
  return stagedWorkbookPath
}

function copySavedWorkbookFromMacosExcelOracle(
  request:
    | Pick<MacosExcelRecalculationOracleRequest, 'saveWorkbook' | 'workbookPath'>
    | Pick<MacosExcelInspectionOracleRequest, 'saveWorkbook' | 'workbookPath'>
    | Pick<MacosExcelStructuralOperationOracleRequest, 'saveWorkbook' | 'workbookPath'>,
  stagedWorkbookPath: string,
): void {
  if (request.saveWorkbook === true) {
    copyFileSync(stagedWorkbookPath, request.workbookPath)
  }
}

function updateLinksAppleScript(mode: MacosExcelLinkUpdateMode): string {
  switch (mode) {
    case 'never':
      return ''
    case 'all':
    case 'external':
    case 'remote':
      return `      try
        update link (targetWorkbook)
      end try`
  }
}

function workbookOpenAppleScriptHelpers(): string {
  return `on workbookNameFromPath(workbookPath)
  set previousDelimiters to AppleScript's text item delimiters
  set AppleScript's text item delimiters to "/"
  set workbookName to last text item of workbookPath
  set AppleScript's text item delimiters to previousDelimiters
  return workbookName
end workbookNameFromPath

on openWorkbookForBiligOracle(workbookPath)
  set workbookName to my workbookNameFromPath(workbookPath)
  do shell script "open -b com.microsoft.Excel " & quoted form of workbookPath
  tell application "Microsoft Excel"
    repeat with openAttempt from 1 to 100
      try
        if (name of active workbook) is workbookName then
          return active workbook
        end if
      end try
      delay 0.1
    end repeat
  end tell
  error "Microsoft Excel did not open workbook " & workbookName number -1728
end openWorkbookForBiligOracle
`
}

function openCompanionWorkbooksAppleScript(): string {
  return `      if (count of argv) > 1 then
        repeat with companionIndex from 2 to count of argv
          set companionPath to item companionIndex of argv
          set companionWorkbook to my openWorkbookForBiligOracle(companionPath)
          set companionWorkbooks to companionWorkbooks & {companionWorkbook}
        end repeat
      end if`
}

function closeCompanionWorkbooksAppleScript(): string {
  return `      repeat with companionWorkbook in companionWorkbooks
        try
          close companionWorkbook saving no
        end try
      end repeat`
}

export function createMacosExcelRecalculationAppleScript(
  request: Pick<MacosExcelRecalculationOracleRequest, 'formulaCells' | 'saveWorkbook' | 'updateLinks' | 'valueCells' | 'worksheetName'>,
): string {
  if (request.valueCells.length === 0) {
    throw new Error('macOS Excel oracle request must read at least one value cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinks = updateLinksAppleScript(request.updateLinks ?? 'never')
  const formulaCells = request.formulaCells
    .map(
      (cell) =>
        `      set formula of range ${toAppleScriptString(cell.address)} of worksheet ${toAppleScriptString(
          request.worksheetName,
        )} of (targetWorkbook) to ${toAppleScriptString(cell.formula)}`,
    )
    .join('\n')
  const valueReads = request.valueCells
    .map(
      (address) =>
        `      set output to output & linefeed & my typedCellValue(value of range ${toAppleScriptString(
          address,
        )} of worksheet ${toAppleScriptString(request.worksheetName)} of (targetWorkbook), string value of range ${toAppleScriptString(
          address,
        )} of worksheet ${toAppleScriptString(request.worksheetName)} of (targetWorkbook))`,
    )
    .join('\n')

  return `${workbookOpenAppleScriptHelpers()}
on run argv
  set workbookPath to item 1 of argv
  tell application "Microsoft Excel"
    set targetWorkbook to missing value
    set companionWorkbooks to {}
    set output to ""
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to my openWorkbookForBiligOracle(workbookPath)
${updateLinks}
${formulaCells}
      calculate full rebuild
      set output to "version=" & (version as string)
${valueReads}
      close (targetWorkbook) saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      try
        close (targetWorkbook) saving no
      end try
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
  return output
end run

on typedCellValue(cellValue, renderedValue)
  if cellValue is missing value then
    if my isExcelErrorDisplayText(renderedValue) then
      return "error" & (ASCII character 9) & renderedValue
    end if
    return "blank" & (ASCII character 9)
  end if
  set valueClass to class of cellValue
  if valueClass is boolean then
    if cellValue then
      return "boolean" & (ASCII character 9) & "true"
    end if
    return "boolean" & (ASCII character 9) & "false"
  end if
  if valueClass is integer or valueClass is real then
    return "number" & (ASCII character 9) & (cellValue as string)
  end if
  return "string" & (ASCII character 9) & (cellValue as string)
end typedCellValue

on isExcelErrorDisplayText(displayText)
  if displayText is "#DIV/0!" then return true
  if displayText is "#REF!" then return true
  if displayText is "#VALUE!" then return true
  if displayText is "#NAME?" then return true
  if displayText is "#N/A" then return true
  if displayText is "#SPILL!" then return true
  if displayText is "#BLOCKED!" then return true
  if displayText is "#NUM!" then return true
  if displayText is "#NULL!" then return true
  if displayText is "#CALC!" then return true
  if displayText is "#FIELD!" then return true
  if displayText is "#UNKNOWN!" then return true
  if displayText is "#GETTING_DATA" then return true
  return false
end isExcelErrorDisplayText
`
}

export function createMacosExcelInspectionAppleScript(
  request: Pick<
    MacosExcelInspectionOracleRequest,
    'formulaCells' | 'inspectCells' | 'refreshAll' | 'saveWorkbook' | 'updateLinks' | 'worksheetName'
  >,
): string {
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel inspection oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinks = updateLinksAppleScript(request.updateLinks ?? 'never')
  const refreshAll = request.refreshAll === true ? '      refresh all (targetWorkbook)' : ''
  const formulaCells = request.formulaCells
    .map(
      (cell) =>
        `      set formula of range ${toAppleScriptString(cell.address)} of worksheet ${toAppleScriptString(
          request.worksheetName,
        )} of (targetWorkbook) to ${toAppleScriptString(cell.formula)}`,
    )
    .join('\n')
  const inspectionReads = request.inspectCells
    .map((address) => {
      const escapedAddress = toAppleScriptString(address)
      const escapedSheet = toAppleScriptString(request.worksheetName)
      return `      set inspectedRange to range ${escapedAddress} of worksheet ${escapedSheet} of (targetWorkbook)
      set output to output & linefeed & ${escapedAddress} & (ASCII character 9) & my formulaText(formula of inspectedRange) & (ASCII character 9) & my typedCellValue(value of inspectedRange, string value of inspectedRange)`
    })
    .join('\n')

  return `${workbookOpenAppleScriptHelpers()}
on run argv
  set workbookPath to item 1 of argv
  tell application "Microsoft Excel"
    set targetWorkbook to missing value
    set companionWorkbooks to {}
    set output to ""
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to my openWorkbookForBiligOracle(workbookPath)
${updateLinks}
${formulaCells}
${refreshAll}
      calculate full rebuild
      set output to "version=" & (version as string)
${inspectionReads}
      close (targetWorkbook) saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      try
        close (targetWorkbook) saving no
      end try
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
  return output
end run

${cellValueAppleScriptHelpers()}`
}

export function createMacosExcelStructuralOperationAppleScript(
  request: Pick<
    MacosExcelStructuralOperationOracleRequest,
    'formulaCells' | 'inspectCells' | 'operations' | 'refreshAll' | 'saveWorkbook' | 'worksheetName' | 'updateLinks'
  >,
): string {
  if (request.operations.length === 0) {
    throw new Error('macOS Excel structural oracle request must apply at least one operation')
  }
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel structural oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinks = updateLinksAppleScript(request.updateLinks ?? 'never')
  const refreshAll = request.refreshAll === true ? '      refresh all (targetWorkbook)' : ''
  const formulaCells = (request.formulaCells ?? [])
    .map(
      (cell) =>
        `      set formula of range ${toAppleScriptString(cell.address)} of targetWorksheet to ${toAppleScriptString(cell.formula)}`,
    )
    .join('\n')
  const operations = request.operations.map((operation) => `      ${structuralOperationAppleScript(operation)}`).join('\n')
  const inspectionReads = request.inspectCells
    .map((address) => {
      const escapedAddress = toAppleScriptString(address)
      return `      set inspectedRange to range ${escapedAddress} of targetWorksheet
      set output to output & linefeed & ${escapedAddress} & (ASCII character 9) & my formulaText(formula of inspectedRange) & (ASCII character 9) & my typedCellValue(value of inspectedRange, string value of inspectedRange)`
    })
    .join('\n')

  return `${workbookOpenAppleScriptHelpers()}
on run argv
  set workbookPath to item 1 of argv
  tell application "Microsoft Excel"
    set targetWorkbook to missing value
    set companionWorkbooks to {}
    set output to ""
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to my openWorkbookForBiligOracle(workbookPath)
${updateLinks}
      set targetWorksheet to worksheet ${toAppleScriptString(request.worksheetName)} of (targetWorkbook)
${formulaCells}
${operations}
${refreshAll}
      calculate full rebuild
      set output to "version=" & (version as string)
${inspectionReads}
      close (targetWorkbook) saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      try
        close (targetWorkbook) saving no
      end try
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
  return output
end run

${cellValueAppleScriptHelpers()}`
}

export function parseMacosExcelRecalculationOutput(rawOutput: string, expectedValueCount: number): MacosExcelRecalculationOracleResult {
  const lines = rawOutput.split(/\r?\n/u)
  const versionLine = lines[0]
  if (!versionLine?.startsWith('version=')) {
    throw new Error(`Unexpected Microsoft Excel oracle output header: ${versionLine ?? '<empty>'}`)
  }
  const rawValues = lines.slice(1)
  if (rawValues.length !== expectedValueCount) {
    throw new Error(`Expected ${String(expectedValueCount)} Excel oracle values, received ${String(rawValues.length)}`)
  }
  return {
    excelVersion: versionLine.slice('version='.length),
    rawValues,
    values: rawValues.map(parseTypedExcelValue),
  }
}

export function parseMacosExcelInspectionOutput(rawOutput: string, expectedAddresses: readonly string[]): MacosExcelInspectionOracleResult {
  const lines = rawOutput.split(/\r?\n/u)
  const versionLine = lines[0]
  if (!versionLine?.startsWith('version=')) {
    throw new Error(`Unexpected Microsoft Excel inspection oracle output header: ${versionLine ?? '<empty>'}`)
  }
  const cellLines = lines.slice(1)
  if (cellLines.length !== expectedAddresses.length) {
    throw new Error(`Expected ${String(expectedAddresses.length)} Excel inspection cells, received ${String(cellLines.length)}`)
  }
  return {
    excelVersion: versionLine.slice('version='.length),
    cells: cellLines.map((line, index) => parseInspectionCell(line, expectedAddresses[index]!)),
  }
}

function parseTypedExcelValue(rawValue: string): NormalizedFormulaValue {
  const separatorIndex = rawValue.indexOf('\t')
  const kind = separatorIndex === -1 ? rawValue : rawValue.slice(0, separatorIndex)
  const value = separatorIndex === -1 ? '' : rawValue.slice(separatorIndex + 1)
  switch (kind) {
    case 'blank':
      return { kind: 'blank' }
    case 'boolean':
      if (value === 'true') {
        return { kind: 'boolean', value: true }
      }
      if (value === 'false') {
        return { kind: 'boolean', value: false }
      }
      throw new Error(`Unexpected Microsoft Excel boolean oracle value: ${value}`)
    case 'number': {
      const numberValue = Number(value)
      if (!Number.isFinite(numberValue)) {
        throw new Error(`Unexpected Microsoft Excel number oracle value: ${value}`)
      }
      return { kind: 'number', value: numberValue }
    }
    case 'string':
      return { kind: 'string', value }
    case 'error':
      if (value.length === 0) {
        throw new Error('Unexpected empty Microsoft Excel error oracle value')
      }
      return { kind: 'error', value: normalizeExcelErrorValue(value) }
    default:
      throw new Error(`Unexpected Microsoft Excel oracle value kind: ${kind}`)
  }
}

function parseInspectionCell(rawLine: string, expectedAddress: string): MacosExcelCellInspection {
  const [address, rawFormula, kind, value] = rawLine.split('\t')
  if (address !== expectedAddress) {
    throw new Error(`Expected Excel inspection cell ${expectedAddress}, received ${address ?? '<missing>'}`)
  }
  if (!kind) {
    throw new Error(`Missing Excel inspection value kind for ${expectedAddress}`)
  }
  const rawValue = `${kind}\t${value ?? ''}`
  const formula = rawFormula === undefined || rawFormula.length === 0 ? undefined : rawFormula
  return {
    address,
    ...(formula !== undefined ? { formula } : {}),
    rawValue,
    value: parseTypedExcelValue(rawValue),
  }
}

function structuralOperationAppleScript(operation: MacosExcelStructuralOperation): string {
  switch (operation.kind) {
    case 'insertRows':
      return `insert into range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift down`
    case 'insertColumns':
      return `insert into range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift to right`
    case 'deleteRows':
      return `delete range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift up`
    case 'deleteColumns':
      return `delete range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift to left`
    case 'setCellValue':
      return `set value of range ${toAppleScriptString(operation.address)} of targetWorksheet to ${toAppleScriptValue(operation.value)}`
    case 'clearCell':
      return `clear contents range ${toAppleScriptString(operation.address)} of targetWorksheet`
    case 'moveRows':
      return [
        `cut range (range ${toAppleScriptString(operation.sourceRange)} of targetWorksheet)`,
        `insert into range (range ${toAppleScriptString(operation.destinationRange)} of targetWorksheet) shift shift down`,
      ].join('\n      ')
    case 'moveColumns':
      return [
        `cut range (range ${toAppleScriptString(operation.sourceRange)} of targetWorksheet)`,
        `insert into range (range ${toAppleScriptString(operation.destinationRange)} of targetWorksheet) shift shift to right`,
      ].join('\n      ')
    case 'createDataTable':
      if (!operation.rowInput && !operation.columnInput) {
        throw new Error('macOS Excel data table operation requires a row input, column input, or both')
      }
      return [
        `data table (range ${toAppleScriptString(operation.range)} of targetWorksheet)`,
        operation.rowInput ? `row input (range ${toAppleScriptString(operation.rowInput)} of targetWorksheet)` : '',
        operation.columnInput ? `column input (range ${toAppleScriptString(operation.columnInput)} of targetWorksheet)` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
    case 'applyAutoFilter':
      if (!Number.isSafeInteger(operation.field) || operation.field <= 0) {
        throw new Error('macOS Excel AutoFilter operation requires a positive one-based field')
      }
      return [
        `autofilter range (range ${toAppleScriptString(operation.range)} of targetWorksheet)`,
        `field ${String(operation.field)}`,
        operation.criteria1 !== undefined ? `criteria1 ${toAppleScriptValue(operation.criteria1)}` : '',
        operation.operator ? `operator ${autoFilterOperatorAppleScript(operation.operator)}` : '',
        operation.criteria2 !== undefined ? `criteria2 ${toAppleScriptValue(operation.criteria2)}` : '',
        operation.visibleDropDown !== undefined ? `visible drop down ${toAppleScriptValue(operation.visibleDropDown)}` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
  }
}

function autoFilterOperatorAppleScript(operator: MacosExcelAutoFilterOperator): string {
  switch (operator) {
    case 'and':
      return 'autofilter and'
    case 'or':
      return 'autofilter or'
    case 'filterByValue':
      return 'filter by value'
  }
}

function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}

function toAppleScriptValue(value: string | number | boolean): string {
  if (typeof value === 'string') {
    return toAppleScriptString(value)
  }
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false'
  }
  return String(value)
}

const excelErrorCodeByDisplay = new Map<string, ErrorCode>([
  ['#DIV/0!', ErrorCode.Div0],
  ['#REF!', ErrorCode.Ref],
  ['#VALUE!', ErrorCode.Value],
  ['#NAME?', ErrorCode.Name],
  ['#N/A', ErrorCode.NA],
  ['#SPILL!', ErrorCode.Spill],
  ['#BLOCKED!', ErrorCode.Blocked],
  ['#NUM!', ErrorCode.Num],
])

function normalizeExcelErrorValue(value: string): string {
  return String(excelErrorCodeByDisplay.get(value.toUpperCase()) ?? value)
}

function cellValueAppleScriptHelpers(): string {
  return `on formulaText(cellFormula)
  if cellFormula is missing value then
    return ""
  end if
  return cellFormula as string
end formulaText

on typedCellValue(cellValue, renderedValue)
  if cellValue is missing value then
    if my isExcelErrorDisplayText(renderedValue) then
      return "error" & (ASCII character 9) & renderedValue
    end if
    return "blank" & (ASCII character 9)
  end if
  set valueClass to class of cellValue
  if valueClass is boolean then
    if cellValue then
      return "boolean" & (ASCII character 9) & "true"
    end if
    return "boolean" & (ASCII character 9) & "false"
  end if
  if valueClass is integer or valueClass is real then
    return "number" & (ASCII character 9) & (cellValue as string)
  end if
  return "string" & (ASCII character 9) & (cellValue as string)
end typedCellValue

on isExcelErrorDisplayText(displayText)
  if displayText is "#DIV/0!" then return true
  if displayText is "#REF!" then return true
  if displayText is "#VALUE!" then return true
  if displayText is "#NAME?" then return true
  if displayText is "#N/A" then return true
  if displayText is "#SPILL!" then return true
  if displayText is "#BLOCKED!" then return true
  if displayText is "#NUM!" then return true
  if displayText is "#NULL!" then return true
  if displayText is "#CALC!" then return true
  if displayText is "#FIELD!" then return true
  if displayText is "#UNKNOWN!" then return true
  if displayText is "#GETTING_DATA" then return true
  return false
end isExcelErrorDisplayText
`
}
