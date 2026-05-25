import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import type {
  MacosExcelInspectionOracleRequest,
  MacosExcelLinkUpdateMode,
  MacosExcelPackageOpenSaveOracleRequest,
  MacosExcelRecalculationOracleRequest,
  MacosExcelStructuralOperationOracleRequest,
} from './macos-excel-oracle-types.js'

export function removeMacosExcelOracleTempDir(dirPath: string): void {
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

export function createMacosExcelOracleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-excel-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

export function macosExcelPreOpenWorkbookPaths(
  stagedWorkbookPath: string,
  companionWorkbookPaths: readonly string[] | undefined,
): readonly string[] {
  return [...(companionWorkbookPaths ?? []), stagedWorkbookPath]
}

export function openWorkbooksForMacosExcelOracle(
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

export function stageWorkbookForMacosExcelOracle(workbookPath: string, tempDir: string): string {
  const stagedWorkbookPath = join(tempDir, basename(workbookPath))
  copyFileSync(workbookPath, stagedWorkbookPath)
  return stagedWorkbookPath
}

export function copySavedWorkbookFromMacosExcelOracle(
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

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
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

export function macosExcelButtonPromptHandlerCommand(heredocName: string, buttonName: string): string {
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

function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}
