import { describe, expect, it } from 'vitest'
import { CapturedCellMutationKind, materializeCapturedCellMutationRestores } from '../engine/services/mutation-captured-cell-restores.js'

describe('mutation captured cell restores', () => {
  it('materializes compact restore buffers into engine mutation refs', () => {
    const refs = materializeCapturedCellMutationRestores({
      sheetIds: Uint32Array.of(1, 1, 2, 2, 3, 3),
      cellIndexPlusOnes: Uint32Array.of(11, 0, 7, 8, 0, 10),
      rows: Uint32Array.of(0, 1, 2, 3, 4, 5),
      cols: Uint32Array.of(1, 2, 3, 4, 5, 6),
      kinds: Uint8Array.of(
        CapturedCellMutationKind.NumberValue,
        CapturedCellMutationKind.BooleanValue,
        CapturedCellMutationKind.LiteralValue,
        CapturedCellMutationKind.Formula,
        CapturedCellMutationKind.NullValue,
        CapturedCellMutationKind.Clear,
      ),
      numbers: Float64Array.of(42, 1, 0, 0, 0, 0),
      values: [undefined, undefined, 'text'],
      formulas: [undefined, undefined, undefined, 'A1+B1'],
      potentialNewCells: 0,
    })

    expect(refs).toEqual([
      { sheetId: 1, cellIndex: 10, mutation: { kind: 'setCellValue', row: 0, col: 1, value: 42 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 2, value: true } },
      { sheetId: 2, cellIndex: 6, mutation: { kind: 'setCellValue', row: 2, col: 3, value: 'text' } },
      { sheetId: 2, cellIndex: 7, mutation: { kind: 'setCellFormula', row: 3, col: 4, formula: 'A1+B1' } },
      { sheetId: 3, mutation: { kind: 'setCellValue', row: 4, col: 5, value: null } },
      { sheetId: 3, cellIndex: 9, mutation: { kind: 'clearCell', row: 5, col: 6 } },
    ])
  })
})
