import { parseCellAddress, parseRangeAddress, type FormulaNode, type ParsedCellReferenceInfo } from '@bilig/formula'
import { ErrorCode } from '@bilig/protocol'
import type { RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand } from '../runtime-state.js'

export interface DirectScalarWorkbook {
  readonly getSheet: (sheetName: string) => { readonly id: number } | undefined
}

export interface DirectScalarCompiledFormula {
  readonly optimizedAst: FormulaNode
  readonly astMatchesSource?: boolean
  readonly symbolicRefs: readonly string[]
  readonly symbolicNames: readonly string[]
  readonly symbolicTables: readonly string[]
  readonly symbolicSpills: readonly string[]
  readonly parsedSymbolicRefs?: readonly ParsedCellReferenceInfo[]
}

export function buildDirectScalarOperand(args: {
  readonly node: FormulaNode
  readonly ownerSheetName: string
  readonly ownerSheetId: number | undefined
  readonly workbook: DirectScalarWorkbook
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly nextTranslatedCellRef?: () =>
    | {
        readonly sheetName: string | undefined
        readonly address: string
        readonly row: number | undefined
        readonly col: number | undefined
      }
    | undefined
}): RuntimeDirectScalarOperand | undefined {
  if (args.node.kind === 'NumberLiteral') {
    return { kind: 'literal-number', value: args.node.value }
  }
  if (args.node.kind === 'CellRef') {
    const translated = args.nextTranslatedCellRef?.()
    const sheetName = translated?.sheetName ?? args.node.sheetName ?? args.ownerSheetName
    const sheetId = sheetName === args.ownerSheetName ? args.ownerSheetId : args.workbook.getSheet(sheetName)?.id
    if (sheetId === undefined) {
      return {
        kind: 'error',
        code: ErrorCode.Ref,
      }
    }
    if (translated?.row !== undefined && translated.col !== undefined) {
      return {
        kind: 'cell',
        cellIndex: args.ensureCellTrackedByCoords(sheetId, translated.row, translated.col),
      }
    }
    let parsed: ReturnType<typeof parseCellAddress>
    try {
      parsed = parseCellAddress(translated?.address ?? args.node.ref, sheetName)
    } catch {
      return {
        kind: 'error',
        code: ErrorCode.Ref,
      }
    }
    return {
      kind: 'cell',
      cellIndex: args.ensureCellTracked(parsed.sheetName ?? sheetName, parsed.text),
    }
  }
  return undefined
}

export function unwrapDirectScalarBinaryNode(node: FormulaNode): { readonly node: FormulaNode; readonly resultOffset: number | undefined } {
  if (
    node.kind === 'BinaryExpr' &&
    node.operator === '+' &&
    node.right.kind === 'NumberLiteral' &&
    Number.isFinite(node.right.value) &&
    node.left.kind === 'BinaryExpr' &&
    (node.left.operator === '+' || node.left.operator === '-' || node.left.operator === '*' || node.left.operator === '/')
  ) {
    return {
      node: node.left,
      resultOffset: node.right.value,
    }
  }
  return {
    node,
    resultOffset: undefined,
  }
}

export function tryParseDependencyRangeAddress(
  address: string,
  currentSheetName?: string,
): ReturnType<typeof parseRangeAddress> | undefined {
  try {
    return parseRangeAddress(address, currentSheetName)
  } catch {
    return undefined
  }
}

export function tryParseDependencyCellAddress(address: string, currentSheetName?: string): ReturnType<typeof parseCellAddress> | undefined {
  try {
    return parseCellAddress(address, currentSheetName)
  } catch {
    return undefined
  }
}

export function buildDirectScalarDescriptor(args: {
  readonly compiled: DirectScalarCompiledFormula
  readonly ownerSheetName: string
  readonly ownerSheetId: number | undefined
  readonly workbook: DirectScalarWorkbook
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
}): RuntimeDirectScalarDescriptor | undefined {
  if (args.compiled.symbolicNames.length > 0 || args.compiled.symbolicTables.length > 0 || args.compiled.symbolicSpills.length > 0) {
    return undefined
  }
  let translatedCellRefIndex = 0
  const nextTranslatedCellRef =
    args.compiled.astMatchesSource === false
      ? ():
          | {
              readonly sheetName: string | undefined
              readonly address: string
              readonly row: number | undefined
              readonly col: number | undefined
            }
          | undefined => {
          const parsed = args.compiled.parsedSymbolicRefs?.[translatedCellRefIndex]
          const address = parsed?.address ?? args.compiled.symbolicRefs[translatedCellRefIndex]
          translatedCellRefIndex += 1
          return address ? { sheetName: parsed?.sheetName, address, row: parsed?.row, col: parsed?.col } : undefined
        }
      : undefined
  const unwrapped = unwrapDirectScalarBinaryNode(args.compiled.optimizedAst)
  const node = unwrapped.node
  if (node.kind === 'BinaryExpr' && (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/')) {
    const left = buildDirectScalarOperand({
      node: node.left,
      ownerSheetName: args.ownerSheetName,
      ownerSheetId: args.ownerSheetId,
      workbook: args.workbook,
      ensureCellTracked: args.ensureCellTracked,
      ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
      ...(nextTranslatedCellRef ? { nextTranslatedCellRef } : {}),
    })
    const right = buildDirectScalarOperand({
      node: node.right,
      ownerSheetName: args.ownerSheetName,
      ownerSheetId: args.ownerSheetId,
      workbook: args.workbook,
      ensureCellTracked: args.ensureCellTracked,
      ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
      ...(nextTranslatedCellRef ? { nextTranslatedCellRef } : {}),
    })
    if (left && right) {
      return {
        kind: 'binary',
        operator: node.operator,
        left,
        right,
        ...(unwrapped.resultOffset !== undefined ? { resultOffset: unwrapped.resultOffset } : {}),
      }
    }
  }
  if (node.kind === 'CallExpr' && node.callee.trim().toUpperCase() === 'ABS' && node.args.length === 1) {
    const operand = buildDirectScalarOperand({
      node: node.args[0]!,
      ownerSheetName: args.ownerSheetName,
      ownerSheetId: args.ownerSheetId,
      workbook: args.workbook,
      ensureCellTracked: args.ensureCellTracked,
      ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
      ...(nextTranslatedCellRef ? { nextTranslatedCellRef } : {}),
    })
    if (operand) {
      return {
        kind: 'abs',
        operand,
      }
    }
  }
  return undefined
}
