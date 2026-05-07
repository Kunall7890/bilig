import { describe, expect, it, vi } from 'vitest'
import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ErrorCode } from '@bilig/protocol'
import { WorkbookStore } from '../workbook-store.js'
import {
  buildDirectScalarDescriptor,
  buildDirectScalarOperand,
  tryParseDependencyCellAddress,
  tryParseDependencyRangeAddress,
  unwrapDirectScalarBinaryNode,
  type DirectScalarCompiledFormula,
} from '../engine/services/formula-binding-direct-scalar.js'

function compiled(optimizedAst: FormulaNode, overrides: Partial<DirectScalarCompiledFormula> = {}): DirectScalarCompiledFormula {
  return {
    optimizedAst,
    symbolicRefs: [],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    ...overrides,
  }
}

describe('formula binding direct scalar helpers', () => {
  it('builds binary and ABS direct scalar descriptors', () => {
    const workbook = new WorkbookStore('scalar')
    const sheet = workbook.createSheet('Sheet1')
    const ensureCellTracked = vi.fn(() => 11)

    expect(
      buildDirectScalarDescriptor({
        compiled: compiled(parseFormula('A1+2+5')),
        ownerSheetName: 'Sheet1',
        ownerSheetId: sheet.id,
        workbook,
        ensureCellTracked,
        ensureCellTrackedByCoords: vi.fn(() => 99),
      }),
    ).toEqual({
      kind: 'binary',
      operator: '+',
      left: { kind: 'cell', cellIndex: 11 },
      right: { kind: 'literal-number', value: 2 },
      resultOffset: 5,
    })

    expect(
      buildDirectScalarDescriptor({
        compiled: compiled(parseFormula('ABS(A1)')),
        ownerSheetName: 'Sheet1',
        ownerSheetId: sheet.id,
        workbook,
        ensureCellTracked,
        ensureCellTrackedByCoords: vi.fn(() => 99),
      }),
    ).toEqual({ kind: 'abs', operand: { kind: 'cell', cellIndex: 11 } })
  })

  it('uses translated parsed refs when the optimized AST no longer matches source metadata', () => {
    const workbook = new WorkbookStore('scalar')
    const sheet = workbook.createSheet('Sheet1')
    const ensureByCoords = vi.fn(() => 44)

    expect(
      buildDirectScalarDescriptor({
        compiled: compiled(parseFormula('A1+2'), {
          astMatchesSource: false,
          symbolicRefs: ['C3'],
          parsedSymbolicRefs: [{ kind: 'cell', address: 'C3', sheetName: 'Sheet1', row: 2, col: 2 }],
        }),
        ownerSheetName: 'Sheet1',
        ownerSheetId: sheet.id,
        workbook,
        ensureCellTracked: vi.fn(() => 11),
        ensureCellTrackedByCoords: ensureByCoords,
      }),
    ).toEqual({
      kind: 'binary',
      operator: '+',
      left: { kind: 'cell', cellIndex: 44 },
      right: { kind: 'literal-number', value: 2 },
    })
    expect(ensureByCoords).toHaveBeenCalledWith(sheet.id, 2, 2)
  })

  it('returns direct scalar error operands for unresolved sheet references', () => {
    expect(
      buildDirectScalarOperand({
        node: parseFormula('Other!A1'),
        ownerSheetName: 'Sheet1',
        ownerSheetId: 1,
        workbook: new WorkbookStore('scalar'),
        ensureCellTracked: vi.fn(() => 11),
        ensureCellTrackedByCoords: vi.fn(() => 99),
      }),
    ).toEqual({ kind: 'error', code: ErrorCode.Ref })
  })

  it('parses dependency addresses without throwing on invalid references', () => {
    expect(tryParseDependencyCellAddress('A1', 'Sheet1')?.text).toBe('A1')
    expect(tryParseDependencyCellAddress('not a cell', 'Sheet1')).toBeUndefined()
    expect(tryParseDependencyRangeAddress('A1:B2', 'Sheet1')?.kind).toBe('cells')
    expect(tryParseDependencyRangeAddress('not a range', 'Sheet1')).toBeUndefined()
    expect(unwrapDirectScalarBinaryNode(parseFormula('A1+2+5'))).toMatchObject({ resultOffset: 5 })
  })
})
