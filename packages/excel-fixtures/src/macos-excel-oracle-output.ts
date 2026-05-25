import { ErrorCode } from '@bilig/protocol'

import type { NormalizedFormulaValue } from './oracle-harness.js'
import type {
  MacosExcelCellInspection,
  MacosExcelInspectionOracleResult,
  MacosExcelPackageOpenSaveOracleResult,
  MacosExcelRecalculationOracleResult,
  MacosExcelRejectedStructuralOperationOracleResult,
} from './macos-excel-oracle.js'

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

export function parseMacosExcelPackageOpenSaveOutput(rawOutput: string): MacosExcelPackageOpenSaveOracleResult {
  const lines = rawOutput.split(/\r?\n/u)
  const versionLine = lines[0]
  if (!versionLine?.startsWith('version=')) {
    throw new Error(`Unexpected Microsoft Excel package oracle output header: ${versionLine ?? '<empty>'}`)
  }
  if (lines.length !== 1) {
    throw new Error(`Unexpected Microsoft Excel package oracle output lines: ${String(lines.length)}`)
  }
  return {
    excelVersion: versionLine.slice('version='.length),
  }
}

export function parseMacosExcelRejectedStructuralOperationOutput(rawOutput: string): MacosExcelRejectedStructuralOperationOracleResult {
  const lines = rawOutput.split(/\r?\n/u)
  const versionLine = lines[0]
  if (!versionLine?.startsWith('version=')) {
    throw new Error(`Unexpected Microsoft Excel rejected-operation oracle output header: ${versionLine ?? '<empty>'}`)
  }
  if (lines[1] !== 'operation=rejected') {
    throw new Error(`Expected Microsoft Excel to reject structural operation, received: ${lines[1] ?? '<missing>'}`)
  }
  const errorNumberLine = lines[2]
  if (!errorNumberLine?.startsWith('errorNumber=')) {
    throw new Error(`Missing Microsoft Excel rejected-operation error number: ${errorNumberLine ?? '<empty>'}`)
  }
  const errorNumber = Number(errorNumberLine.slice('errorNumber='.length))
  if (!Number.isFinite(errorNumber)) {
    throw new Error(`Invalid Microsoft Excel rejected-operation error number: ${errorNumberLine}`)
  }
  const errorMessageLine = lines[3]
  if (!errorMessageLine?.startsWith('errorMessage=')) {
    throw new Error(`Missing Microsoft Excel rejected-operation error message: ${errorMessageLine ?? '<empty>'}`)
  }
  return {
    excelVersion: versionLine.slice('version='.length),
    errorNumber,
    errorMessage: errorMessageLine.slice('errorMessage='.length),
    sheetNames: lines.slice(4).map((line) => {
      if (!line.startsWith('sheet=')) {
        throw new Error(`Unexpected Microsoft Excel rejected-operation sheet line: ${line}`)
      }
      return line.slice('sheet='.length)
    }),
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

const excelErrorCodeByDisplay = new Map<string, ErrorCode>([
  ['#NULL!', ErrorCode.Null],
  ['#DIV/0!', ErrorCode.Div0],
  ['#REF!', ErrorCode.Ref],
  ['#VALUE!', ErrorCode.Value],
  ['#NAME?', ErrorCode.Name],
  ['#N/A', ErrorCode.NA],
  ['#SPILL!', ErrorCode.Spill],
  ['#BLOCKED!', ErrorCode.Blocked],
  ['#NUM!', ErrorCode.Num],
  ['#FIELD!', ErrorCode.Field],
])

function normalizeExcelErrorValue(value: string): string {
  return String(excelErrorCodeByDisplay.get(value.toUpperCase()) ?? value)
}
