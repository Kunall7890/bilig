import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import type { NormalizedFormulaValue } from './oracle-harness.js'

export const defaultMacosExcelAppPath = '/Applications/Microsoft Excel.app' as const

export interface MacosExcelOracleFormulaCell {
  readonly address: string
  readonly formula: string
}

export interface MacosExcelRecalculationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly valueCells: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
}

export interface MacosExcelRecalculationOracleResult {
  readonly excelVersion: string
  readonly rawValues: readonly string[]
  readonly values: readonly NormalizedFormulaValue[]
}

export function isMacosExcelInstalled(appPath: string = defaultMacosExcelAppPath): boolean {
  return existsSync(appPath)
}

export function runMacosExcelRecalculationOracle(request: MacosExcelRecalculationOracleRequest): MacosExcelRecalculationOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-macos-excel-oracle-'))
  const scriptPath = join(tempDir, 'recalculate.scpt')
  try {
    writeFileSync(scriptPath, createMacosExcelRecalculationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, request.workbookPath], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    return parseMacosExcelRecalculationOutput(rawOutput, request.valueCells.length)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

export function createMacosExcelRecalculationAppleScript(
  request: Pick<MacosExcelRecalculationOracleRequest, 'formulaCells' | 'saveWorkbook' | 'valueCells' | 'worksheetName'>,
): string {
  if (request.valueCells.length === 0) {
    throw new Error('macOS Excel oracle request must read at least one value cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
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
        )} of worksheet ${toAppleScriptString(request.worksheetName)} of (targetWorkbook))`,
    )
    .join('\n')

  return `on run argv
  set workbookPath to POSIX file (item 1 of argv)
  set targetWorkbook to missing value
  set output to ""
  tell application "Microsoft Excel"
    set display alerts to false
    set screen updating to false
    try
      set targetWorkbook to open workbook workbook file name workbookPath
${formulaCells}
      calculate full rebuild
      set output to "version=" & (version as string)
${valueReads}
      close targetWorkbook saving ${closeSavingMode}
      set screen updating to true
    on error errMsg number errNum
      if targetWorkbook is not missing value then
        try
          close targetWorkbook saving no
        end try
      end if
      set screen updating to true
      error errMsg number errNum
    end try
  end tell
  return output
end run

on typedCellValue(cellValue)
  if cellValue is missing value then
    return "blank" & tab
  end if
  set valueClass to class of cellValue
  if valueClass is boolean then
    if cellValue then
      return "boolean" & tab & "true"
    end if
    return "boolean" & tab & "false"
  end if
  if valueClass is integer or valueClass is real then
    return "number" & tab & (cellValue as string)
  end if
  return "string" & tab & (cellValue as string)
end typedCellValue
`
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
    default:
      throw new Error(`Unexpected Microsoft Excel oracle value kind: ${kind}`)
  }
}

function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}
