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
import { cellValueAppleScriptHelpers } from './macos-excel-oracle-cell-value-helpers.js'

import type {
  MacosExcelCalculationPolicy,
  MacosExcelInspectionOracleRequest,
  MacosExcelInspectionOracleResult,
  MacosExcelLinkUpdateMode,
  MacosExcelPackageOpenSaveOracleRequest,
  MacosExcelPackageOpenSaveOracleResult,
  MacosExcelRecalculationOracleRequest,
  MacosExcelRecalculationOracleResult,
  MacosExcelRejectedStructuralOperationOracleRequest,
  MacosExcelRejectedStructuralOperationOracleResult,
  MacosExcelSortHeader,
  MacosExcelSortOrder,
  MacosExcelSortOrientation,
  MacosExcelStructuralOperation,
  MacosExcelStructuralOperationOracleRequest,
} from './macos-excel-oracle-types.js'

export {
  parseMacosExcelInspectionOutput,
  parseMacosExcelPackageOpenSaveOutput,
  parseMacosExcelRecalculationOutput,
  parseMacosExcelRejectedStructuralOperationOutput,
} from './macos-excel-oracle-output.js'
export type {
  MacosExcelAutoFilterOperator,
  MacosExcelCalculationPolicy,
  MacosExcelCellInspection,
  MacosExcelInspectionOracleRequest,
  MacosExcelInspectionOracleResult,
  MacosExcelLinkUpdateMode,
  MacosExcelOracleFormulaCell,
  MacosExcelPackageOpenSaveOracleRequest,
  MacosExcelPackageOpenSaveOracleResult,
  MacosExcelRecalculationOracleRequest,
  MacosExcelRecalculationOracleResult,
  MacosExcelRejectedStructuralOperationOracleRequest,
  MacosExcelRejectedStructuralOperationOracleResult,
  MacosExcelSortHeader,
  MacosExcelSortKey,
  MacosExcelSortOrder,
  MacosExcelSortOrientation,
  MacosExcelStructuralOperation,
  MacosExcelStructuralOperationOracleRequest,
} from './macos-excel-oracle-types.js'

export const defaultMacosExcelAppPath = '/Applications/Microsoft Excel.app' as const

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
    openWorkbooksForMacosExcelOracle(
      appPath,
      [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])],
      request.timeoutMs,
      request.updateLinks ?? 'never',
    )
    writeFileSync(scriptPath, createMacosExcelRecalculationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
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
    openWorkbooksForMacosExcelOracle(
      appPath,
      [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])],
      request.timeoutMs,
      request.updateLinks ?? 'never',
    )
    writeFileSync(scriptPath, createMacosExcelInspectionAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
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
    openWorkbooksForMacosExcelOracle(
      appPath,
      [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])],
      request.timeoutMs,
      request.updateLinks ?? 'never',
    )
    writeFileSync(scriptPath, createMacosExcelPackageOpenSaveAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
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
    openWorkbooksForMacosExcelOracle(
      appPath,
      [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])],
      request.timeoutMs,
      request.updateLinks ?? 'never',
    )
    writeFileSync(scriptPath, createMacosExcelStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
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
    openWorkbooksForMacosExcelOracle(
      appPath,
      [stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])],
      request.timeoutMs,
      request.updateLinks ?? 'never',
    )
    writeFileSync(scriptPath, createMacosExcelRejectedStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, stagedWorkbookPath, ...(request.companionWorkbookPaths ?? [])], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
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

function openWorkbooksForMacosExcelOracle(
  appPath: string,
  workbookPaths: readonly string[],
  timeoutMs: number | undefined,
  updateLinks: MacosExcelLinkUpdateMode,
): void {
  closeStaleMacosExcelOracleWorkbooks()
  startMacosExcelLinkUpdatePromptHandler(updateLinks)
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
      execFileSync('osascript', ['-e', macosExcelWorkbookOpenCheckAppleScript(), workbookPath, basename(workbookPath)], {
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
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
  set expectedName to item 2 of argv
  set normalizedExpectedPath to my normalizedMacosPath(expectedPath)
  tell application "Microsoft Excel"
    repeat with workbookIndex from 1 to (count workbooks)
      set candidateWorkbook to workbook workbookIndex
      try
        if (name of candidateWorkbook as string) is expectedName then return "ready"
      end try
      try
        if my normalizedMacosPath(full name of candidateWorkbook as string) is normalizedExpectedPath then return "ready"
      end try
    end repeat
  end tell
  error "Workbook is not open: " & expectedPath number -2700
end run

on normalizedMacosPath(rawPath)
  if rawPath starts with "file://" then
    set rawPath to text 8 thru -1 of rawPath
  end if
  if rawPath starts with "/private/" then
    return text 9 thru -1 of rawPath
  end if
  return rawPath
end normalizedMacosPath`
}

function closeStaleMacosExcelOracleWorkbooks(): void {
  try {
    execFileSync('osascript', ['-e', macosExcelCloseStaleOracleWorkbooksAppleScript()], {
      stdio: 'ignore',
      timeout: 5_000,
    })
  } catch {
    // A modal prompt can block stale workbook cleanup; the prompt handler below will clear it.
  }
}

function macosExcelCloseStaleOracleWorkbooksAppleScript(): string {
  return `tell application "Microsoft Excel"
  repeat with workbookIndex from (count workbooks) to 1 by -1
    set candidateWorkbook to workbook workbookIndex
    try
      set candidatePath to full name of candidateWorkbook as string
      if my isBiligOracleWorkbookPath(candidatePath) then
        close candidateWorkbook saving no
      end if
    end try
  end repeat
end tell

on isBiligOracleWorkbookPath(candidatePath)
  if candidatePath contains "/bilig-excel-oracle/" then return true
  if candidatePath contains "/bilig-headless-excel-" then return true
  return false
end isBiligOracleWorkbookPath`
}

function startMacosExcelLinkUpdatePromptHandler(updateLinks: MacosExcelLinkUpdateMode): void {
  const buttonName = updateLinks === 'never' ? "Don't Update" : 'Update'
  const command = `/usr/bin/osascript <<'BILIG_LINK_UPDATE_PROMPT' >/dev/null 2>&1 &
repeat 120 times
  delay 0.25
  tell application "System Events"
    tell process "Microsoft Excel"
      if exists button ${toAppleScriptString(buttonName)} of window 1 then
        click button ${toAppleScriptString(buttonName)} of window 1
        return
      end if
    end tell
  end tell
end repeat
BILIG_LINK_UPDATE_PROMPT`

  execFileSync('/bin/sh', ['-c', command], { timeout: 5_000 })
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

function macosExcelCalculationAppleScript(policy: MacosExcelCalculationPolicy | undefined): string {
  switch (policy ?? 'fullRebuild') {
    case 'fullRebuild':
      return '      calculate full rebuild'
    case 'none':
      return ''
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
  request: Pick<
    MacosExcelRecalculationOracleRequest,
    'calculationPolicy' | 'formulaCells' | 'saveWorkbook' | 'updateLinks' | 'valueCells' | 'worksheetName'
  >,
): string {
  if (request.valueCells.length === 0) {
    throw new Error('macOS Excel oracle request must read at least one value cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const calculationCommand = macosExcelCalculationAppleScript(request.calculationPolicy)
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
${calculationCommand}
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

${cellValueAppleScriptHelpers()}`
}

export function createMacosExcelInspectionAppleScript(
  request: Pick<
    MacosExcelInspectionOracleRequest,
    'calculationPolicy' | 'formulaCells' | 'inspectCells' | 'refreshWorkbook' | 'saveWorkbook' | 'updateLinks' | 'worksheetName'
  >,
): string {
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel inspection oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const calculationCommand = macosExcelCalculationAppleScript(request.calculationPolicy)
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
${calculationCommand}
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
  request: Pick<MacosExcelPackageOpenSaveOracleRequest, 'calculationPolicy' | 'refreshWorkbook' | 'saveWorkbook' | 'updateLinks'>,
): string {
  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
  const updateLinksMode = macosExcelUpdateLinksModeAppleScript(request.updateLinks ?? 'never')
  const refreshWorkbook = request.refreshWorkbook === true ? '      refresh all targetWorkbook' : ''
  const calculationCommand = macosExcelCalculationAppleScript(request.calculationPolicy)

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
${calculationCommand}
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
    'calculationPolicy' | 'formulaCells' | 'inspectCells' | 'operations' | 'saveWorkbook' | 'worksheetName' | 'updateLinks'
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
  const calculationCommand = macosExcelCalculationAppleScript(request.calculationPolicy)
  const sheetDeletePromptHandler = request.operations.some((operation) => operation.kind === 'deleteSheet')
    ? '      my startSheetDeletePromptHandler()'
    : ''
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
${sheetDeletePromptHandler}
${operations}
${calculationCommand}
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

${cellValueAppleScriptHelpers()}
${sheetDeletePromptHandlerAppleScript()}`
}

function sheetDeletePromptHandlerAppleScript(): string {
  const command = `/usr/bin/osascript <<'BILIG_SHEET_DELETE_PROMPT' >/dev/null 2>&1 &
repeat 120 times
  delay 0.25
  tell application "System Events"
    tell process "Microsoft Excel"
      if exists button "Delete" of window 1 then
        click button "Delete" of window 1
        return
      end if
    end tell
  end tell
end repeat
BILIG_SHEET_DELETE_PROMPT`

  return `on startSheetDeletePromptHandler()
  do shell script ${toAppleScriptString(command)}
end startSheetDeletePromptHandler`
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
        `if exists worksheet ${toAppleScriptString(operation.name)} of targetWorkbook then`,
        `  delete worksheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        `else if exists chart sheet ${toAppleScriptString(operation.name)} of targetWorkbook then`,
        `  delete chart sheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        'else',
        `  error ${toAppleScriptString(`Sheet not found: ${operation.name}`)} number -1728`,
        'end if',
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
    case 'deleteTable':
      return `delete list object ${toAppleScriptString(operation.tableName)} of targetWorksheet`
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
