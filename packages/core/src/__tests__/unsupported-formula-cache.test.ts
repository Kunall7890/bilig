import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  collectDefinedFormulaNames,
  formulaMayContainFullRecalcPreservableUnavailableFormulaCall,
  formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc,
  formulaShouldUseCachedUnsupportedFunctionValue,
} from '../snapshot/unsupported-formula-cache.js'

function snapshotWithDefinedNames(names: readonly string[]): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'unsupported-formula-cache-test',
      metadata: {
        definedNames: names.map((name) => ({ name, value: { kind: 'formula', formula: '=1' } })),
      },
    },
    sheets: [],
  }
}

describe('unsupported formula cache', () => {
  it('normalizes defined formula names from workbook metadata', () => {
    expect(collectDefinedFormulaNames(snapshotWithDefinedNames([' TaxRate ', '', 'custom.fx']))).toEqual(new Set(['TAXRATE', 'CUSTOM.FX']))
    expect(collectDefinedFormulaNames({ version: 1, workbook: { name: 'empty' }, sheets: [] })).toEqual(new Set())
  })

  it('detects unavailable calls without flagging builtins, defined names, or scoped locals', () => {
    const definedNames = collectDefinedFormulaNames(snapshotWithDefinedNames(['ExternalRate']))

    expect(formulaShouldUseCachedUnsupportedFunctionValue('SUM(A1:A2)', definedNames)).toBe(false)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('ExternalRate(A1)', definedNames)).toBe(false)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('MissingFn(A1)', definedNames)).toBe(true)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('LAMBDA(x, x + 1)(1)', definedNames)).toBe(false)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('LAMBDA(x, MissingFn(x))(1)', definedNames)).toBe(true)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('LET(local, 1, local + 1)', definedNames)).toBe(false)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('LET(local, MissingFn(), local)', definedNames)).toBe(true)
    expect(formulaShouldUseCachedUnsupportedFunctionValue('(', definedNames)).toBe(false)
  })

  it('uses a cheap marker gate for full-recalc-preservable cached formulas', () => {
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('SUM(A1:A10)')).toBe(false)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('UNKNOWNFUNC(A1)')).toBe(false)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_FV(A1,"Industry")')).toBe(true)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_fv(A1,"Industry")')).toBe(true)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_xldudf_WISEPRICE(B1,"Shares Outstanding")')).toBe(true)
  })

  it('preserves only supported imported-cache markers during full recalculation', () => {
    const definedNames = new Set<string>()

    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('SUM(1,2)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('UNKNOWNFUNC(42)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_FV(A1,"Industry")', definedNames)).toBe(true)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_fv(A1,"Industry")', definedNames)).toBe(true)
    expect(
      formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_xldudf_WISEPRICE(B1,"Shares Outstanding")', definedNames),
    ).toBe(true)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('MissingFn(A1)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('(', definedNames)).toBe(false)
  })

  it('uses the AST walk as the authority when marker text is present', () => {
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('"_FV(A1)"', new Set())).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_FV(A1)', new Set(['_FV']))).toBe(false)
  })
})
