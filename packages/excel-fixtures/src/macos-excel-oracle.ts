import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import {
  parseMacosExcelInspectionOutput,
  parseMacosExcelPackageOpenSaveOutput,
  parseMacosExcelRecalculationOutput,
  parseMacosExcelRejectedStructuralOperationOutput,
} from './macos-excel-oracle-output.js'
import type { NormalizedFormulaValue } from './oracle-harness.js'

export {
  parseMacosExcelInspectionOutput,
  parseMacosExcelPackageOpenSaveOutput,
  parseMacosExcelRecalculationOutput,
  parseMacosExcelRejectedStructuralOperationOutput,
} from './macos-excel-oracle-output.js'

export const defaultMacosExcelAppPath = '/Applications/Microsoft Excel.app' as const

export interface MacosExcelOracleFormulaCell {
  readonly address: string
  readonly formula: string
}

export type MacosExcelLinkUpdateMode = 'all' | 'external' | 'never' | 'remote'
export type MacosExcelSortHeader = 'guess' | 'no' | 'yes'
export type MacosExcelSortOrder = 'ascending' | 'descending'
export type MacosExcelSortOrientation = 'columns' | 'rows'
export type MacosExcelAutoFilterOperator = 'autofilter and' | 'autofilter or' | 'filter by value'

export interface MacosExcelSortKey {
  readonly key: string
  readonly order?: MacosExcelSortOrder
}

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
  readonly refreshWorkbook?: boolean
}

export interface MacosExcelPackageOpenSaveOracleRequest {
  readonly workbookPath: string
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
  readonly refreshWorkbook?: boolean
}

export type MacosExcelStructuralOperation =
  | { readonly kind: 'insertRows'; readonly range: string }
  | { readonly kind: 'insertColumns'; readonly range: string }
  | { readonly kind: 'deleteRows'; readonly range: string }
  | { readonly kind: 'deleteColumns'; readonly range: string }
  | { readonly kind: 'setCellValue'; readonly address: string; readonly value: string | number | boolean }
  | { readonly kind: 'clearCell'; readonly address: string }
  | { readonly kind: 'createSheet'; readonly name: string }
  | { readonly kind: 'renameSheet'; readonly newName: string }
  | { readonly kind: 'deleteSheet'; readonly name: string }
  | { readonly kind: 'moveSheet'; readonly name: string; readonly before?: string; readonly after?: string }
  | { readonly kind: 'moveRows'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'moveColumns'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'createDataTable'; readonly range: string; readonly rowInput?: string; readonly columnInput?: string }
  | {
      readonly kind: 'applySort'
      readonly range: string
      readonly keys: readonly MacosExcelSortKey[]
      readonly header?: MacosExcelSortHeader
      readonly orientation?: MacosExcelSortOrientation
    }
  | {
      readonly kind: 'applyTableSort'
      readonly tableName: string
      readonly keys: readonly MacosExcelSortKey[]
      readonly header?: MacosExcelSortHeader
      readonly orientation?: MacosExcelSortOrientation
    }
  | {
      readonly kind: 'applyTableAutoFilter'
      readonly tableName: string
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
}

export interface MacosExcelRejectedStructuralOperationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly operation: MacosExcelStructuralOperation
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
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

export interface MacosExcelPackageOpenSaveOracleResult {
  readonly excelVersion: string
}

export interface MacosExcelRejectedStructuralOperationOracleResult {
  readonly excelVersion: string
  readonly errorMessage: string
  readonly errorNumber: number
  readonly sheetNames: readonly string[]
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
    openWorkbooksForMacosExcelOracle(appPath, [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], request.timeoutMs)
    writeFileSync(scriptPath, createMacosExcelRecalculationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelRecalculationOutput(rawOutput, request.valueCells.length)
  } finally {
    removeMacosExcelOracleTempDir(tempDir)
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
    openWorkbooksForMacosExcelOracle(appPath, [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], request.timeoutMs)
    writeFileSync(scriptPath, createMacosExcelInspectionAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelInspectionOutput(rawOutput, request.inspectCells)
  } finally {
    removeMacosExcelOracleTempDir(tempDir)
  }
}

export function runMacosExcelPackageOpenSaveOracle(request: MacosExcelPackageOpenSaveOracleRequest): MacosExcelPackageOpenSaveOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-package-')
  const scriptPath = join(tempDir, 'package-open-save.scpt')
  try {
    const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
    openWorkbooksForMacosExcelOracle(appPath, [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], request.timeoutMs)
    writeFileSync(scriptPath, createMacosExcelPackageOpenSaveAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelPackageOpenSaveOutput(rawOutput)
  } finally {
    removeMacosExcelOracleTempDir(tempDir)
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
    openWorkbooksForMacosExcelOracle(appPath, [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], request.timeoutMs)
    writeFileSync(scriptPath, createMacosExcelStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    copySavedWorkbookFromMacosExcelOracle(request, stagedWorkbookPath)
    return parseMacosExcelInspectionOutput(rawOutput, request.inspectCells)
  } finally {
    removeMacosExcelOracleTempDir(tempDir)
  }
}

export function runMacosExcelRejectedStructuralOperationOracle(
  request: MacosExcelRejectedStructuralOperationOracleRequest,
): MacosExcelRejectedStructuralOperationOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-rejected-structure-')
  const scriptPath = join(tempDir, 'rejected-structure.scpt')
  try {
    const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
    openWorkbooksForMacosExcelOracle(appPath, [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], request.timeoutMs)
    writeFileSync(scriptPath, createMacosExcelRejectedStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    return parseMacosExcelRejectedStructuralOperationOutput(rawOutput)
  } finally {
    removeMacosExcelOracleTempDir(tempDir)
  }
}

function removeMacosExcelOracleTempDir(dirPath: string): void {
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      rmSync(dirPath, { recursive: true, force: true })
      return
    } catch (error) {
      if (isRecord(error) && ['EBUSY', 'EINTR', 'ENOTEMPTY'].includes(String(error['code']))) {
        sleepSync(100)
        continue
      }
      throw error
    }
  }
  rmSync(dirPath, { recursive: true, force: true })
}

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function createMacosExcelOracleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-excel-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function openWorkbooksForMacosExcelOracle(appPath: string, workbookPaths: readonly string[], timeoutMs: number | undefined): void {
  for (const workbookPath of workbookPaths) {
    execFileSync('open', ['-a', appPath, workbookPath], {
      timeout: Math.min(timeoutMs ?? 60_000, 30_000),
    })
    waitForMacosExcelWorkbookOpen(workbookPath, timeoutMs)
  }
}

function waitForMacosExcelWorkbookOpen(workbookPath: string, timeoutMs: number | undefined): void {
  const deadline = Date.now() + Math.min(timeoutMs ?? 60_000, 30_000)
  let lastError: unknown

  while (Date.now() < deadline) {
    try {
      execFileSync('osascript', ['-e', macosExcelWorkbookOpenCheckAppleScript(), workbookPath], {
        encoding: 'utf8',
        timeout: 5_000,
      })
      return
    } catch (error) {
      lastError = error
      sleepSync(250)
    }
  }

  throw new Error(`Timed out waiting for Microsoft Excel to open ${workbookPath}: ${String(lastError)}`)
}

function macosExcelWorkbookOpenCheckAppleScript(): string {
  return `on run argv
  set expectedPath to item 1 of argv
  tell application "Microsoft Excel"
    try
      if (full name of active workbook as string) is expectedPath then return "ready"
    end try
  end tell
  error "Workbook is not open: " & expectedPath number -2700
end run`
}

function stageWorkbookForMacosExcelOracle(workbookPath: string, tempDir: string): string {
  const stagedWorkbookPath = join(tempDir, basename(workbookPath))
  copyFileSync(workbookPath, stagedWorkbookPath)
  return stagedWorkbookPath
}

function copySavedWorkbookFromMacosExcelOracle(
  request:
    | Pick<MacosExcelRecalculationOracleRequest, 'saveWorkbook' | 'workbookPath'>
    | Pick<MacosExcelInspectionOracleRequest, 'saveWorkbook' | 'workbookPath'>
    | Pick<MacosExcelPackageOpenSaveOracleRequest, 'saveWorkbook' | 'workbookPath'>
    | Pick<MacosExcelStructuralOperationOracleRequest, 'saveWorkbook' | 'workbookPath'>,
  stagedWorkbookPath: string,
): void {
  if (request.saveWorkbook === true) {
    copyFileSync(stagedWorkbookPath, request.workbookPath)
  }
}

function macosExcelUpdateLinksModeAppleScript(mode: MacosExcelLinkUpdateMode): string {
  switch (mode) {
    case 'all':
      return 'update remote and external links'
    case 'external':
      return 'update external links only'
    case 'never':
      return 'do not update links'
    case 'remote':
      return 'update remote links only'
  }
}

function openCompanionWorkbooksAppleScript(): string {
  return `      if (count of argv) > 1 then
        repeat with companionIndex from 2 to count of argv
          set companionPath to item companionIndex of argv
          set companionWorkbook to open workbook workbook file name companionPath update links do not update links
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

function macroPromptDisablerAppleScript(): string {
  const command = `/usr/bin/osascript <<'BILIG_MACRO_PROMPT' >/dev/null 2>&1 &
repeat 120 times
  delay 0.25
  tell application "System Events"
    tell process "Microsoft Excel"
      if exists button "Disable Macros" of window 1 then
        click button "Disable Macros" of window 1
        return
      end if
    end tell
  end tell
end repeat
BILIG_MACRO_PROMPT`

  return `on startMacroPromptDisabler()
  do shell script ${toAppleScriptString(command)}
end startMacroPromptDisabler`
}

export function createMacosExcelRecalculationAppleScript(
  request: Pick<MacosExcelRecalculationOracleRequest, 'formulaCells' | 'saveWorkbook' | 'updateLinks' | 'valueCells' | 'worksheetName'>,
): string {
  if (request.valueCells.length === 0) {
    throw new Error('macOS Excel oracle request must read at least one value cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
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

  return `on run argv
  set workbookPath to item 1 of argv
  set targetWorkbook to missing value
  set companionWorkbooks to {}
  set output to ""
  tell application "Microsoft Excel"
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to open workbook workbook file name workbookPath update links ${updateLinksMode}
${formulaCells}
      calculate full rebuild
      set output to "version=" & (version as string)
${valueReads}
      close targetWorkbook saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
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
    'formulaCells' | 'inspectCells' | 'refreshWorkbook' | 'saveWorkbook' | 'updateLinks' | 'worksheetName'
  >,
): string {
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel inspection oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const formulaCells = request.formulaCells
    .map(
      (cell) =>
        `      set formula of range ${toAppleScriptString(cell.address)} of worksheet ${toAppleScriptString(
          request.worksheetName,
        )} of (targetWorkbook) to ${toAppleScriptString(cell.formula)}`,
    )
    .join('\n')
  const refreshWorkbook = request.refreshWorkbook === true ? '      refresh all targetWorkbook' : ''
  const inspectionReads = request.inspectCells
    .map((address) => {
      const escapedAddress = toAppleScriptString(address)
      const escapedSheet = toAppleScriptString(request.worksheetName)
      return `      set inspectedRange to range ${escapedAddress} of worksheet ${escapedSheet} of (targetWorkbook)
      set output to output & linefeed & ${escapedAddress} & (ASCII character 9) & my formulaText(formula of inspectedRange) & (ASCII character 9) & my typedCellValue(value of inspectedRange, string value of inspectedRange)`
    })
    .join('\n')

  return `on run argv
  set workbookPath to item 1 of argv
  set targetWorkbook to missing value
  set companionWorkbooks to {}
  set output to ""
  tell application "Microsoft Excel"
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to open workbook workbook file name workbookPath update links ${updateLinksMode}
${formulaCells}
${refreshWorkbook}
      calculate full rebuild
      set output to "version=" & (version as string)
${inspectionReads}
      close targetWorkbook saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
  return output
end run

${cellValueAppleScriptHelpers()}`
}

export function createMacosExcelPackageOpenSaveAppleScript(
  request: Pick<MacosExcelPackageOpenSaveOracleRequest, 'refreshWorkbook' | 'saveWorkbook' | 'updateLinks'>,
): string {
  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const refreshWorkbook = request.refreshWorkbook === true ? '      refresh all targetWorkbook' : ''

  return `on run argv
  set workbookPath to item 1 of argv
  set targetWorkbook to missing value
  set companionWorkbooks to {}
  set output to ""
  set priorAutomationSecurity to missing value
  tell application "Microsoft Excel"
    try
      set priorAutomationSecurity to automation security
      set automation security to msoAutomationSecurityForceDisable
${openCompanionWorkbooksAppleScript()}
      my startMacroPromptDisabler()
      set targetWorkbook to open workbook workbook file name workbookPath update links ${updateLinksMode}
${refreshWorkbook}
      calculate full rebuild
      set output to "version=" & (version as string)
      close targetWorkbook saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
      if priorAutomationSecurity is not missing value then
        set automation security to priorAutomationSecurity
      end if
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
${closeCompanionWorkbooksAppleScript()}
      if priorAutomationSecurity is not missing value then
        try
          set automation security to priorAutomationSecurity
        end try
      end if
      error errMsg number errNum
    end try
  end tell
  return output
end run

${macroPromptDisablerAppleScript()}`
}

export function createMacosExcelStructuralOperationAppleScript(
  request: Pick<
    MacosExcelStructuralOperationOracleRequest,
    'formulaCells' | 'inspectCells' | 'operations' | 'saveWorkbook' | 'worksheetName' | 'updateLinks'
  >,
): string {
  if (request.operations.length === 0) {
    throw new Error('macOS Excel structural oracle request must apply at least one operation')
  }
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel structural oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
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

  return `on run argv
  set workbookPath to item 1 of argv
  set targetWorkbook to missing value
  set companionWorkbooks to {}
  set output to ""
  tell application "Microsoft Excel"
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to open workbook workbook file name workbookPath update links ${updateLinksMode}
      set targetWorksheet to worksheet ${toAppleScriptString(request.worksheetName)} of targetWorkbook
${formulaCells}
${operations}
      calculate full rebuild
      set output to "version=" & (version as string)
${inspectionReads}
      close targetWorkbook saving ${closeSavingMode}
${closeCompanionWorkbooksAppleScript()}
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
  return output
end run

${cellValueAppleScriptHelpers()}`
}

export function createMacosExcelRejectedStructuralOperationAppleScript(
  request: Pick<MacosExcelRejectedStructuralOperationOracleRequest, 'operation' | 'updateLinks' | 'worksheetName'>,
): string {
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const operation = structuralOperationAppleScript(request.operation)

  return `on run argv
  set workbookPath to item 1 of argv
  set targetWorkbook to missing value
  set companionWorkbooks to {}
  tell application "Microsoft Excel"
    try
${openCompanionWorkbooksAppleScript()}
      set targetWorkbook to open workbook workbook file name workbookPath update links ${updateLinksMode}
      set targetWorksheet to worksheet ${toAppleScriptString(request.worksheetName)} of targetWorkbook
      set output to "version=" & (version as string)
      try
        ${operation}
      on error operationErrorMessage number operationErrorNumber
        set output to output & linefeed & "operation=rejected"
        set output to output & linefeed & "errorNumber=" & (operationErrorNumber as string)
        set output to output & linefeed & "errorMessage=" & (operationErrorMessage as string)
        set output to output & my workbookSheetNames(targetWorkbook)
        close targetWorkbook saving no
${closeCompanionWorkbooksAppleScript()}
        return output
      end try
      set output to output & linefeed & "operation=applied" & my workbookSheetNames(targetWorkbook)
      close targetWorkbook saving no
${closeCompanionWorkbooksAppleScript()}
      error "Expected Microsoft Excel to reject structural workbook operation" number -2700
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
${closeCompanionWorkbooksAppleScript()}
      error errMsg number errNum
    end try
  end tell
end run

on workbookSheetNames(targetWorkbook)
  set output to ""
  tell application "Microsoft Excel"
    repeat with sheetIndex from 1 to count of worksheets of targetWorkbook
      set output to output & linefeed & "sheet=" & (name of worksheet (sheetIndex as integer) of targetWorkbook as string)
    end repeat
  end tell
  return output
end workbookSheetNames`
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
    case 'createSheet':
      return [
        'set createdWorksheet to make new worksheet at after worksheet (count of worksheets of targetWorkbook) of targetWorkbook',
        `set name of createdWorksheet to ${toAppleScriptString(operation.name)}`,
      ].join('\n      ')
    case 'renameSheet':
      return `set name of targetWorksheet to ${toAppleScriptString(operation.newName)}`
    case 'deleteSheet':
      return [
        'try',
        `  delete worksheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        'on error',
        `  delete chart sheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        'end try',
      ].join('\n      ')
    case 'moveSheet':
      return moveSheetAppleScript(operation)
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
    case 'applySort':
      if (operation.keys.length === 0 || operation.keys.length > 3) {
        throw new Error('macOS Excel sort operation requires one to three sort keys')
      }
      return [
        `sort (range ${toAppleScriptString(operation.range)} of targetWorksheet)`,
        ...operation.keys.flatMap((key, index) => {
          const position = String(index + 1)
          return [
            `key${position} (range ${toAppleScriptString(key.key)} of targetWorksheet)`,
            `order${position} ${sortOrderAppleScript(key.order ?? 'ascending')}`,
          ]
        }),
        operation.header ? `header ${sortHeaderAppleScript(operation.header)}` : '',
        operation.orientation ? `orientation ${sortOrientationAppleScript(operation.orientation)}` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
    case 'applyTableSort':
      if (operation.keys.length === 0 || operation.keys.length > 3) {
        throw new Error('macOS Excel table sort operation requires one to three sort keys')
      }
      return [
        `set tableSort to sort object of list object ${toAppleScriptString(operation.tableName)} of targetWorksheet`,
        `clear sortfieldset (sortfieldset of tableSort)`,
        ...operation.keys.map((key) =>
          [
            `add sortfield (sortfieldset of tableSort)`,
            `key (range ${toAppleScriptString(key.key)} of targetWorksheet)`,
            `order ${sortOrderAppleScript(key.order ?? 'ascending')}`,
          ].join(' '),
        ),
        operation.header ? `set sort header of tableSort to ${sortHeaderAppleScript(operation.header)}` : '',
        operation.orientation ? `set sort orientation of tableSort to ${sortOrientationAppleScript(operation.orientation)}` : '',
        `apply sort tableSort`,
      ]
        .filter((part) => part.length > 0)
        .join('\n      ')
    case 'applyTableAutoFilter':
      if (!Number.isSafeInteger(operation.field) || operation.field <= 0) {
        throw new Error('macOS Excel table AutoFilter operation requires a positive one-based field')
      }
      return [
        `autofilter range (range object of autofilter object of list object ${toAppleScriptString(operation.tableName)} of targetWorksheet)`,
        `field ${String(operation.field)}`,
        operation.criteria1 !== undefined ? `criteria1 ${toAppleScriptValue(operation.criteria1)}` : '',
        operation.operator ? `operator ${operation.operator}` : '',
        operation.criteria2 !== undefined ? `criteria2 ${toAppleScriptValue(operation.criteria2)}` : '',
        operation.visibleDropDown !== undefined ? `visible drop down ${toAppleScriptValue(operation.visibleDropDown)}` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
  }
}

function sortHeaderAppleScript(header: MacosExcelSortHeader): string {
  switch (header) {
    case 'guess':
      return 'header guess'
    case 'no':
      return 'header no'
    case 'yes':
      return 'header yes'
  }
}

function moveSheetAppleScript(operation: Extract<MacosExcelStructuralOperation, { readonly kind: 'moveSheet' }>): string {
  if ((operation.before === undefined) === (operation.after === undefined)) {
    throw new Error('macOS Excel moveSheet operation requires exactly one before or after anchor')
  }
  const anchor = operation.before
    ? `to before worksheet ${toAppleScriptString(operation.before)} of targetWorkbook`
    : `to after worksheet ${toAppleScriptString(operation.after!)} of targetWorkbook`
  return `move worksheet ${toAppleScriptString(operation.name)} of targetWorkbook ${anchor}`
}

function sortOrderAppleScript(order: MacosExcelSortOrder): string {
  switch (order) {
    case 'ascending':
      return 'sort ascending'
    case 'descending':
      return 'sort descending'
  }
}

function sortOrientationAppleScript(orientation: MacosExcelSortOrientation): string {
  switch (orientation) {
    case 'columns':
      return 'sort rows'
    case 'rows':
      return 'sort columns'
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
