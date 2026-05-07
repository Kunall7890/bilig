import { describe, expect, it } from 'vitest'

import { compileFormula } from '../compiler.js'
import { parseFormula } from '../parser.js'

describe('Excel 3D sheet range references', () => {
  it('parses a sheet span qualified single-cell reference as a range reference', () => {
    expect(parseFormula('SUM(Jan:Mar!B2)')).toEqual({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [
        {
          kind: 'RangeRef',
          refKind: 'cells',
          sheetName: 'Jan',
          sheetEndName: 'Mar',
          start: 'B2',
          end: 'B2',
        },
      ],
    })
  })

  it('parses quoted sheet spans with cell ranges', () => {
    expect(parseFormula("SUM('Jan 2026':'Mar 2026'!B2:C2)")).toEqual({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [
        {
          kind: 'RangeRef',
          refKind: 'cells',
          sheetName: 'Jan 2026',
          sheetEndName: 'Mar 2026',
          start: 'B2',
          end: 'C2',
        },
      ],
    })
  })

  it('keeps sheet ranges on the JS plan and dependency metadata', () => {
    const compiled = compileFormula('SUM(Jan:Mar!B2:C2)')

    expect(compiled.deps).toEqual(['Jan:Mar!B2:C2'])
    expect(compiled.parsedDeps).toEqual([
      expect.objectContaining({
        kind: 'range',
        refKind: 'cells',
        sheetName: 'Jan',
        sheetEndName: 'Mar',
        startAddress: 'B2',
        endAddress: 'C2',
      }),
    ])
    expect(compiled.jsPlan).toContainEqual({
      opcode: 'push-range',
      sheetName: 'Jan',
      sheetEndName: 'Mar',
      start: 'B2',
      end: 'C2',
      refKind: 'cells',
    })
  })
})
