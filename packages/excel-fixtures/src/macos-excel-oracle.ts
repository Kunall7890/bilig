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
import { withMacosExcelOracleLock } from './macos-excel-oracle-lock.js'
import { cellValueAppleScriptHelpers } from './macos-excel-oracle-cell-value-helpers.js'
import { sheetDeletePromptHandlerAppleScript, toAppleScriptString } from './macos-excel-oracle-applescript-helpers.js'
import { structuralOperationAppleScript } from './macos-excel-oracle-structural-operation-applescript.js'

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

  return withMacosExcelOracleLock({ label: 'recalculation', timeoutMs: request.timeoutMs }, () => {
    const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-')
    const scriptPath = join(tempDir, 'recalculate.scpt')
    try {
      const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
      openWorkbooksForMacosExcelOracle(
        appPath,
        macosExcelPreOpenWorkbookPaths(stagedWorkbookPath, request.companionWorkbookPaths),
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
  })
}

export function runMacosExcelInspectionOracle(request: MacosExcelInspectionOracleRequest): MacosExcelInspectionOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  return withMacosExcelOracleLock({ label: 'inspection', timeoutMs: request.timeoutMs }, () => {
    const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-inspect-')
    const scriptPath = join(tempDir, 'inspect.scpt')
    try {
      const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
      openWorkbooksForMacosExcelOracle(
        appPath,
        macosExcelPreOpenWorkbookPaths(stagedWorkbookPath, request.companionWorkbookPaths),
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
  })
}

export function runMacosExcelPackageOpenSaveOracle(request: MacosExcelPackageOpenSaveOracleRequest): MacosExcelPackageOpenSaveOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  return withMacosExcelOracleLock({ label: 'package-open-save', timeoutMs: request.timeoutMs }, () => {
    const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-package-')
    const scriptPath = join(tempDir, 'package-open-save.scpt')
    try {
      const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
      openWorkbooksForMacosExcelOracle(
        appPath,
        macosExcelPreOpenWorkbookPaths(stagedWorkbookPath, request.companionWorkbookPaths),
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
  })
}

export function runMacosExcelStructuralOperationOracle(
  request: MacosExcelStructuralOperationOracleRequest,
): MacosExcelInspectionOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  return withMacosExcelOracleLock({ label: 'structural-operation', timeoutMs: request.timeoutMs }, () => {
    const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-structure-')
    const scriptPath = join(tempDir, 'structure.scpt')
    try {
      const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
      openWorkbooksForMacosExcelOracle(
        appPath,
        macosExcelPreOpenWorkbookPaths(stagedWorkbookPath, request.companionWorkbookPaths),
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
  })
}

export function runMacosExcelRejectedStructuralOperationOracle(
  request: MacosExcelRejectedStructuralOperationOracleRequest,
): MacosExcelRejectedStructuralOperationOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  return withMacosExcelOracleLock({ label: 'rejected-structural-operation', timeoutMs: request.timeoutMs }, () => {
    const tempDir = createMacosExcelOracleTempDir('bilig-macos-excel-oracle-rejected-structure-')
    const scriptPath = join(tempDir, 'rejected-structure.scpt')
    try {
      const stagedWorkbookPath = stageWorkbookForMacosExcelOracle(request.workbookPath, tempDir)
      openWorkbooksForMacosExcelOracle(
        appPath,
        macosExcelPreOpenWorkbookPaths(stagedWorkbookPath, request.companionWorkbookPaths),
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
  })
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

function macosExcelPreOpenWorkbookPaths(
  stagedWorkbookPath: string,
  companionWorkbookPaths: readonly string[] | undefined,
): readonly string[] {
  return [...(companionWorkbookPaths ?? []), stagedWorkbookPath]
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
    openWorkbookForMacosExcelOracle(appPath, workbookPath, timeoutMs, updateLinks)
  }
}

function openWorkbookForMacosExcelOracle(
  appPath: string,
  workbookPath: string,
  timeoutMs: number | undefined,
  updateLinks: MacosExcelLinkUpdateMode,
): void {
  startMacosExcelMacroPromptHandler()
  execFileSync('open', ['-a', appPath, workbookPath], {
    timeout: Math.min(timeoutMs ?? 60_000, 30_000),
  })
  try {
    waitForMacosExcelWorkbookOpen(workbookPath, timeoutMs)
  } catch (openError) {
    if (restartMacosExcelForOracleOpenRecovery()) {
      startMacosExcelLinkUpdatePromptHandler(updateLinks)
      startMacosExcelMacroPromptHandler()
      execFileSync('open', ['-a', appPath, workbookPath], {
        timeout: Math.min(timeoutMs ?? 60_000, 30_000),
      })
      waitForMacosExcelWorkbookOpen(workbookPath, timeoutMs)
      return
    }
    throw openError
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
    try
      set workbookCount to count of workbooks
    on error
      set workbookCount to 0
    end try
    repeat with workbookIndex from 1 to workbookCount
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

function restartMacosExcelForOracleOpenRecovery(): boolean {
  try {
    const result = execFileSync('osascript', ['-e', macosExcelQuitIfNoWorkbooksAppleScript()], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      timeout: 5_000,
    }).trim()
    return result === 'quit' && waitForMacosExcelProcessToExit()
  } catch {
    return false
  }
}

function waitForMacosExcelProcessToExit(): boolean {
  const deadline = Date.now() + 10_000
  while (Date.now() < deadline) {
    try {
      const isRunning = execFileSync('osascript', ['-e', 'tell application "System Events" to return exists process "Microsoft Excel"'], {
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
        timeout: 2_000,
      }).trim()
      if (isRunning === 'false') {
        return true
      }
    } catch {
      return true
    }
    sleepSync(250)
  }
  return false
}

function macosExcelQuitIfNoWorkbooksAppleScript(): string {
  return `tell application "Microsoft Excel"
  try
    set workbookCount to count of workbooks
  on error
    set workbookCount to 0
  end try
  if workbookCount is 0 then
    quit
    return "quit"
  end if
  return "busy"
end tell`
}

function macosExcelCloseStaleOracleWorkbooksAppleScript(): string {
  return `tell application "Microsoft Excel"
  try
    set workbookCount to count of workbooks
  on error
    return
  end try
  repeat with workbookIndex from workbookCount to 1 by -1
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
  execFileSync('/bin/sh', ['-c', macosExcelButtonPromptHandlerCommand('BILIG_LINK_UPDATE_PROMPT', buttonName)], { timeout: 5_000 })
}

function startMacosExcelMacroPromptHandler(): void {
  execFileSync('/bin/sh', ['-c', macosExcelButtonPromptHandlerCommand('BILIG_MACRO_PROMPT', 'Disable Macros')], { timeout: 5_000 })
}

function macosExcelButtonPromptHandlerCommand(heredocName: string, buttonName: string): string {
  return `/usr/bin/osascript <<'${heredocName}' >/dev/null 2>&1 &
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
${heredocName}`
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
  const command = macosExcelButtonPromptHandlerCommand('BILIG_MACRO_PROMPT', 'Disable Macros')

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
