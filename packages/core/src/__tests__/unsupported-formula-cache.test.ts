import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  collectDefinedFormulaNames,
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

  it('only preserves known full-recalc-safe unavailable functions', () => {
    const definedNames = new Set<string>()

    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_xldudf_MyAddin(A1)', definedNames)).toBe(true)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_FV(A1)', definedNames)).toBe(true)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('MissingFn(A1)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('SUM(A1:A2)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('(', definedNames)).toBe(false)
  })
})
