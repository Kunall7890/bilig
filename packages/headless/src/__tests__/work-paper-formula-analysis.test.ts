import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  calculateWorkPaperFormula,
  getWorkPaperNamedExpressionsFromFormula,
  normalizeWorkPaperFormula,
  validateWorkPaperFormula,
  type WorkPaperFormulaAnalysisHooks,
} from '../work-paper-formula-analysis.js'
import { WorkPaperNotAFormulaError, WorkPaperParseError } from '../work-paper-errors.js'

const hooks: WorkPaperFormulaAnalysisHooks = {
  messageOf: (error, fallback) => (error instanceof Error && error.message ? error.message : fallback),
}

describe('work paper formula analysis', () => {
  it('normalizes valid formula text with a leading equals sign', () => {
    expect(normalizeWorkPaperFormula('=sum(a1, 2)', hooks)).toBe('=SUM(A1,2)')
  })

  it('requires formula text for normalization and named-expression inspection', () => {
    expect(() => normalizeWorkPaperFormula('SUM(A1)', hooks)).toThrow(WorkPaperNotAFormulaError)
    expect(() => getWorkPaperNamedExpressionsFromFormula('Alpha+Beta', hooks)).toThrow(WorkPaperNotAFormulaError)
  })

  it('wraps parser failures with the public parse error', () => {
    expect(() => normalizeWorkPaperFormula('=SUM(', hooks)).toThrow(WorkPaperParseError)
    expect(() => getWorkPaperNamedExpressionsFromFormula('=SUM(', hooks)).toThrow(WorkPaperParseError)
  })

  it('collects unique named-expression references in sorted order', () => {
    expect(getWorkPaperNamedExpressionsFromFormula('=Beta+Alpha+Beta', hooks)).toEqual(['Alpha', 'Beta'])
  })

  it('delegates valid formula calculation to the scratch workbook helper', () => {
    const applied: string[] = []

    expect(
      calculateWorkPaperFormula({
        createWorkbook: () => ({
          engine: {
            createSheet: (sheetName) => {
              applied.push(`sheet:${sheetName}`)
            },
            getSpillRanges: () => [],
          },
          registerNamedExpression: () => {},
          requireSheetId: () => 1,
          replaceSheetContent: () => {},
          clearHistoryStacks: () => {},
          applyRawContent: (_address, content) => {
            applied.push(`content:${content}`)
          },
          getRangeValues: () => [],
          getCellValue: () => ({ tag: ValueTag.Number, value: 42 }),
          dispose: () => {},
        }),
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=SUM(1,2)',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(applied).toContain('content:=SUM(1,2)')
  })

  it('disposes the scratch workbook when formula calculation fails', () => {
    let disposed = false

    expect(() =>
      calculateWorkPaperFormula({
        createWorkbook: () => ({
          engine: {
            createSheet: () => {},
            getSpillRanges: () => [],
          },
          registerNamedExpression: () => {},
          requireSheetId: () => 1,
          replaceSheetContent: () => {},
          clearHistoryStacks: () => {},
          applyRawContent: () => {
            throw new Error('scratch failed')
          },
          getRangeValues: () => [],
          getCellValue: () => ({ tag: ValueTag.Number, value: 42 }),
          dispose: () => {
            disposed = true
          },
        }),
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=SUM(1,2)',
        messageOf: hooks.messageOf,
      }),
    ).toThrow(WorkPaperParseError)
    expect(disposed).toBe(true)
  })

  it('validates formulas without throwing', () => {
    expect(validateWorkPaperFormula('=SUM(1,2)')).toBe(true)
    expect(validateWorkPaperFormula('SUM(1,2)')).toBe(false)
    expect(validateWorkPaperFormula('=SUM(')).toBe(false)
  })
})
