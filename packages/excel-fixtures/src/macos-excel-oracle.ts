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

export interface MacosExcelInspectionOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly inspectCells: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
}

export type MacosExcelStructuralOperation =
  | { readonly kind: 'insertRows'; readonly range: string }
  | { readonly kind: 'insertColumns'; readonly range: string }
  | { readonly kind: 'deleteRows'; readonly range: string }
  | { readonly kind: 'deleteColumns'; readonly range: string }
  | { readonly kind: 'moveRows'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'moveColumns'; readonly sourceRange: string; readonly destinationRange: string }

export interface MacosExcelStructuralOperationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly operations: readonly MacosExcelStructuralOperation[]
  readonly inspectCells: readonly string[]
  readonly formulaCells?: readonly MacosExcelOracleFormulaCell[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
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

export function runMacosExcelInspectionOracle(request: MacosExcelInspectionOracleRequest): MacosExcelInspectionOracleResult {
  const appPath = request.appPath ?? defaultMacosExcelAppPath
  if (!isMacosExcelInstalled(appPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${appPath}`)
  }

  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-macos-excel-oracle-inspect-'))
  const scriptPath = join(tempDir, 'inspect.scpt')
  try {
    writeFileSync(scriptPath, createMacosExcelInspectionAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, request.workbookPath], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
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

  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-macos-excel-oracle-structure-'))
  const scriptPath = join(tempDir, 'structure.scpt')
  try {
    writeFileSync(scriptPath, createMacosExcelStructuralOperationAppleScript(request))
    const rawOutput = execFileSync('osascript', [scriptPath, request.workbookPath], {
      encoding: 'utf8',
      timeout: request.timeoutMs ?? 60_000,
    }).trim()
    return parseMacosExcelInspectionOutput(rawOutput, request.inspectCells)
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
`
}

export function createMacosExcelInspectionAppleScript(
  request: Pick<MacosExcelInspectionOracleRequest, 'formulaCells' | 'inspectCells' | 'saveWorkbook' | 'worksheetName'>,
): string {
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel inspection oracle request must inspect at least one cell')
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
  const inspectionReads = request.inspectCells
    .map((address) => {
      const escapedAddress = toAppleScriptString(address)
      const escapedSheet = toAppleScriptString(request.worksheetName)
      return `      set inspectedRange to range ${escapedAddress} of worksheet ${escapedSheet} of (targetWorkbook)
      set output to output & linefeed & ${escapedAddress} & (ASCII character 9) & my formulaText(formula of inspectedRange) & (ASCII character 9) & my typedCellValue(value of inspectedRange)`
    })
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
${inspectionReads}
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

${cellValueAppleScriptHelpers()}`
}

export function createMacosExcelStructuralOperationAppleScript(
  request: Pick<
    MacosExcelStructuralOperationOracleRequest,
    'formulaCells' | 'inspectCells' | 'operations' | 'saveWorkbook' | 'worksheetName'
  >,
): string {
  if (request.operations.length === 0) {
    throw new Error('macOS Excel structural oracle request must apply at least one operation')
  }
  if (request.inspectCells.length === 0) {
    throw new Error('macOS Excel structural oracle request must inspect at least one cell')
  }

  const closeSavingMode = request.saveWorkbook === true ? 'yes' : 'no'
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
      set output to output & linefeed & ${escapedAddress} & (ASCII character 9) & my formulaText(formula of inspectedRange) & (ASCII character 9) & my typedCellValue(value of inspectedRange)`
    })
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
      set targetWorksheet to worksheet ${toAppleScriptString(request.worksheetName)} of targetWorkbook
${formulaCells}
${operations}
      calculate full rebuild
      set output to "version=" & (version as string)
${inspectionReads}
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
  }
}

function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}

function cellValueAppleScriptHelpers(): string {
  return `on formulaText(cellFormula)
  if cellFormula is missing value then
    return ""
  end if
  return cellFormula as string
end formulaText

on typedCellValue(cellValue)
  if cellValue is missing value then
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
`
}
