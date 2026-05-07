import { describe, expect, it } from 'vitest'
import { parseFormula } from '@bilig/formula'
import {
  collectDirectApproximateLookupCandidates,
  collectIndexedExactLookupCandidates,
  hasDirectApproximateLookupCandidate,
  hasIndexedExactLookupCandidate,
  staticIntegerValue,
} from '../engine/services/formula-binding-lookup-candidates.js'

describe('formula binding lookup candidates', () => {
  it('reads integer literals including negative unary literals', () => {
    expect(staticIntegerValue(parseFormula('1'))).toBe(1)
    expect(staticIntegerValue(parseFormula('-1'))).toBe(-1)
    expect(staticIntegerValue(parseFormula('1.5'))).toBeUndefined()
    expect(staticIntegerValue(parseFormula('A1'))).toBeUndefined()
  })

  it('collects exact lookup candidates from MATCH and XMATCH calls', () => {
    expect(collectIndexedExactLookupCandidates(parseFormula('MATCH(A1,A2:A4,0)'))).toEqual([
      { start: 'A2', end: 'A4', startRow: 1, endRow: 3, startCol: 0, endCol: 0 },
    ])
    expect(collectIndexedExactLookupCandidates(parseFormula('XMATCH(A1,Sheet2!B1:B5,0,-1)'))).toEqual([
      { sheetName: 'Sheet2', start: 'B1', end: 'B5', startRow: 0, endRow: 4, startCol: 1, endCol: 1 },
    ])
    expect(hasIndexedExactLookupCandidate(parseFormula('SUM(MATCH(A1,A2:A4,0),1)'))).toBe(true)
    expect(hasIndexedExactLookupCandidate(parseFormula('MATCH(A1,A2:A4,1)'))).toBe(false)
  })

  it('ignores omitted lookup arguments while walking nested formulas', () => {
    const ast = parseFormula('MATCH(INDEX(A1:F1,MATCH(TRUE,INDEX(A1:F1<>0,),0)),A1:F1,0)')

    expect(collectIndexedExactLookupCandidates(ast)).toEqual([{ start: 'A1', end: 'F1', startRow: 0, endRow: 0, startCol: 0, endCol: 5 }])
    expect(collectDirectApproximateLookupCandidates(ast)).toEqual([])
  })

  it('collects approximate lookup candidates only for supported forward search modes', () => {
    expect(collectDirectApproximateLookupCandidates(parseFormula('MATCH(A1,A2:A4,1)'))).toEqual([
      { start: 'A2', end: 'A4', startRow: 1, endRow: 3, startCol: 0, endCol: 0 },
    ])
    expect(collectDirectApproximateLookupCandidates(parseFormula('XMATCH(A1,Sheet2!B1:B5,-1,1)'))).toEqual([
      { sheetName: 'Sheet2', start: 'B1', end: 'B5', startRow: 0, endRow: 4, startCol: 1, endCol: 1 },
    ])
    expect(hasDirectApproximateLookupCandidate(parseFormula('XMATCH(A1,A2:A4,-1,-1)'))).toBe(false)
    expect(hasDirectApproximateLookupCandidate(parseFormula('XMATCH(A1,A2:A4,0,1)'))).toBe(false)
  })
})
