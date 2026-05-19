import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { formulaCompatibilityRegistry } from '../compatibility.js'
import { evaluateAst } from '../js-evaluator.js'
import type { EvaluationContext } from '../js-evaluator-types.js'
import { parseFormula } from '../parser.js'
import { serializeFormula } from '../translation.js'
import { evaluableFormulaArbitrary, evaluationContext } from './formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const inventoryEvaluationTemplates = [
  { inventoryId: 'aggregation:sum-range', formula: 'SUM($A$1:$A$6)' },
  { inventoryId: 'aggregation:counta-range', formula: 'COUNTA($A$1:$A$6)' },
  { inventoryId: 'aggregation:countblank-range', formula: 'COUNTBLANK($A$1:$A$6)' },
  { inventoryId: 'logical:iferror-catches-any-error', formula: 'IFERROR(1/0,TaxRate)' },
  { inventoryId: 'logical:ifna-catches-na-only', formula: 'IFNA(NA(),Label)' },
  { inventoryId: 'information:n-basic', formula: 'N(Flag)+N($A$1)' },
  { inventoryId: 'information:t-basic', formula: 'T(Label)&"!"' },
  { inventoryId: 'text:len-basic', formula: 'LEN(Label)' },
  { inventoryId: 'arithmetic:cross-sheet-multiply', formula: "'My Sheet'!$A$1+N('My Sheet'!$A$2)" },
].map((template) => {
  const entry = formulaCompatibilityRegistry.find((candidate) => candidate.id === template.inventoryId)
  if (!entry) {
    throw new Error(`Missing compatibility inventory entry ${template.inventoryId}`)
  }
  return {
    inventoryId: template.inventoryId,
    formula: template.formula,
    family: entry.family,
  }
})

const generatedCellValueArbitrary = fc.oneof<CellValue>(
  fc.integer({ min: -20, max: 20 }).map((value) => ({ tag: ValueTag.Number, value })),
  fc.constantFrom('', 'alpha', '42', '-7').map((value) => ({ tag: ValueTag.String, value, stringId: 0 })),
  fc.boolean().map((value) => ({ tag: ValueTag.Boolean, value })),
  fc.constant({ tag: ValueTag.Empty }),
  fc.constantFrom(ErrorCode.Div0, ErrorCode.NA, ErrorCode.Value).map((code) => ({ tag: ValueTag.Error, code })),
)

const inventoryEvaluationCaseArbitrary = fc.record({
  template: fc.constantFrom(...inventoryEvaluationTemplates),
  sheetValues: fc.array(generatedCellValueArbitrary, { minLength: 6, maxLength: 6 }),
  quotedSheetValues: fc.array(generatedCellValueArbitrary, { minLength: 6, maxLength: 6 }),
  taxRate: generatedCellValueArbitrary,
  label: generatedCellValueArbitrary,
  flag: generatedCellValueArbitrary,
})

describe('formula evaluation fuzz', () => {
  it('keeps JS evaluation stable across canonicalization for coercion-heavy formulas', async () => {
    await runProperty({
      suite: 'formula/evaluation/canonicalization-stability',
      arbitrary: evaluableFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula))
        expect(evaluateAst(parseFormula(formula), evaluationContext)).toEqual(evaluateAst(parseFormula(canonical), evaluationContext))
      },
    })
  })

  it('keeps inventory-derived mixed-value evaluation stable across canonicalization', async () => {
    await runProperty({
      suite: 'formula/evaluation/inventory-mixed-canonicalization',
      arbitrary: inventoryEvaluationCaseArbitrary,
      parameters: { numRuns: 140 },
      predicate: ({ template, sheetValues, quotedSheetValues, taxRate, label, flag }) => {
        const context = buildInventoryEvaluationContext({ sheetValues, quotedSheetValues, taxRate, label, flag })
        const canonical = serializeFormula(parseFormula(template.formula))
        expect(evaluateAst(parseFormula(template.formula), context)).toEqual(evaluateAst(parseFormula(canonical), context))
      },
    })
  })
})

function buildInventoryEvaluationContext(args: {
  readonly sheetValues: readonly CellValue[]
  readonly quotedSheetValues: readonly CellValue[]
  readonly taxRate: CellValue
  readonly label: CellValue
  readonly flag: CellValue
}): EvaluationContext {
  return {
    sheetName: 'Sheet1',
    currentAddress: 'D9',
    resolveCell: (sheetName, address) => {
      const index = addressToGeneratedIndex(address)
      if (index === undefined) {
        return empty()
      }
      if (sheetName === 'Sheet1') {
        return args.sheetValues[index] ?? empty()
      }
      if (sheetName === 'My Sheet') {
        return args.quotedSheetValues[index] ?? empty()
      }
      return empty()
    },
    resolveRange: (sheetName, start, end, refKind) => {
      if (refKind !== 'cells' || normalizeAddress(start) !== 'A1' || normalizeAddress(end) !== 'A6') {
        return []
      }
      return sheetName === 'My Sheet' ? [...args.quotedSheetValues] : [...args.sheetValues]
    },
    resolveName: (name) => {
      switch (name.toUpperCase()) {
        case 'TAXRATE':
          return args.taxRate
        case 'LABEL':
          return args.label
        case 'FLAG':
          return args.flag
        default:
          return { tag: ValueTag.Error, code: ErrorCode.Name }
      }
    },
    listSheetNames: () => ['Sheet1', 'My Sheet'],
  }
}

function addressToGeneratedIndex(address: string): number | undefined {
  const normalized = normalizeAddress(address)
  const match = /^A([1-6])$/.exec(normalized)
  return match ? Number.parseInt(match[1], 10) - 1 : undefined
}

function normalizeAddress(address: string): string {
  return address.replaceAll('$', '').toUpperCase()
}

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}
