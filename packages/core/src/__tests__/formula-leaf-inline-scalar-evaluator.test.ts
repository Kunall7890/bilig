import { compileFormula, type CompiledFormula, type JsPlanInstruction } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import {
  INLINE_SCALAR_FAST_PLAN_ARITHMETIC,
  INLINE_SCALAR_FAST_PLAN_CONCAT,
  INLINE_SCALAR_FAST_PLAN_IF_STRING,
  INLINE_SCALAR_FAST_PLAN_LEN,
  INLINE_SCALAR_FAST_PLAN_MIN_MAX,
  INLINE_SCALAR_FAST_PLAN_PMT,
  INLINE_SCALAR_FAST_PLAN_ROUND_SQRT,
  type RuntimeFormula,
} from '../engine/runtime-state.js'
import {
  buildInlineScalarPlanCellIndices,
  classifyInlineScalarFastPlan,
  tryEvaluateFormulaLeafInlineScalar,
} from '../engine/services/formula-leaf-inline-scalar-evaluator.js'
import { StringPool } from '../string-pool.js'

type TestCell = number | string | boolean | null | CellValue

function makeState(cells: Record<number, TestCell>) {
  const strings = new StringPool()
  const tags: number[] = []
  const numbers: number[] = []
  const stringIds: number[] = []

  for (const [rawIndex, rawValue] of Object.entries(cells)) {
    const index = Number(rawIndex)
    const value = typeof rawValue === 'object' && rawValue !== null && 'tag' in rawValue ? rawValue : literalCell(rawValue)
    tags[index] = value.tag
    if (value.tag === ValueTag.Number) {
      numbers[index] = value.value
    } else if (value.tag === ValueTag.Boolean) {
      numbers[index] = value.value ? 1 : 0
    } else if (value.tag === ValueTag.String) {
      stringIds[index] = strings.intern(value.value)
    } else if (value.tag === ValueTag.Error) {
      numbers[index] = value.code
    }
  }

  return {
    workbook: {
      cellStore: {
        tags,
        numbers,
        stringIds,
        getValue: (cellIndex: number, getString: (stringId: number) => string): CellValue => {
          const tag = (tags[cellIndex] ?? ValueTag.Empty) as ValueTag
          switch (tag) {
            case ValueTag.Number:
              return { tag: ValueTag.Number, value: numbers[cellIndex] ?? 0 }
            case ValueTag.Boolean:
              return { tag: ValueTag.Boolean, value: (numbers[cellIndex] ?? 0) !== 0 }
            case ValueTag.String: {
              const stringId = stringIds[cellIndex] ?? 0
              return { tag: ValueTag.String, value: getString(stringId), stringId }
            }
            case ValueTag.Error:
              return { tag: ValueTag.Error, code: (numbers[cellIndex] ?? ErrorCode.Value) as ErrorCode }
            case ValueTag.Empty:
              return { tag: ValueTag.Empty }
          }
        },
      },
      getCalculationSettings: () => ({}),
    },
    strings,
  }
}

function literalCell(value: TestCell): CellValue {
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'string') {
    return { tag: ValueTag.String, value, stringId: 0 }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  if (value === null) {
    return { tag: ValueTag.Empty }
  }
  return value
}

function runtimeFormula(source: string, dependencyIndices: readonly number[], inline = true): RuntimeFormula {
  const compiled = compileFormula(source)
  return runtimeFormulaFromCompiled(compiled, dependencyIndices, inline)
}

function runtimeFormulaFromCompiled(compiled: CompiledFormula, dependencyIndices: readonly number[], inline = true): RuntimeFormula {
  const dependencyIndexArray = Uint32Array.from(dependencyIndices)
  const emptySlice = { ptr: -1, len: 0, cap: 0 }
  return {
    cellIndex: 0,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source: 'test',
    compiled,
    plan: { id: 0, source: 'test', compiled },
    dependencyIndices: dependencyIndexArray,
    dependencyEntities: emptySlice,
    rangeDependencies: new Uint32Array(0),
    graphRangeDependencies: new Uint32Array(0),
    runtimeProgram: new Uint32Array(0),
    constants: new Float64Array(0),
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
    ...(inline ? { inlineScalarPlanCellIndices: buildInlineScalarPlanCellIndices(compiled, dependencyIndexArray) } : {}),
  }
}

function manualRuntimeFormula(plan: readonly JsPlanInstruction[]): RuntimeFormula {
  const cellIndices = new Uint32Array(plan.length)
  cellIndices.fill(0xffffffff)
  const compiled = {
    ...compileFormula('A1+1'),
    jsPlan: [...plan],
    parsedDeps: [],
    producesSpill: false,
    astMatchesSource: true,
  }
  return {
    ...runtimeFormulaFromCompiled(compiled, [], false),
    inlineScalarPlanCellIndices: cellIndices,
  }
}

function evaluateManual(plan: readonly JsPlanInstruction[]): CellValue | undefined {
  return tryEvaluateFormulaLeafInlineScalar({
    state: makeState({}),
    formula: manualRuntimeFormula(plan),
  })
}

describe('formula leaf inline scalar evaluator', () => {
  it('classifies and evaluates the optimized inline scalar plan families', () => {
    const state = makeState({
      10: 3,
      11: 4,
      12: 8,
      20: 'north',
      21: 'west',
      30: 10,
      40: 0.12,
      41: 12,
      42: 1200,
      50: 7,
    })

    const cases = [
      {
        source: 'A1+B1*2',
        deps: [10, 11],
        kind: INLINE_SCALAR_FAST_PLAN_ARITHMETIC,
        expected: { tag: ValueTag.Number, value: 11 },
      },
      {
        source: 'LEN(A1)+LEN(B1)',
        deps: [20, 21],
        kind: INLINE_SCALAR_FAST_PLAN_LEN,
        expected: { tag: ValueTag.Number, value: 9 },
      },
      {
        source: 'CONCATENATE(A1,"-",B1)',
        deps: [20, 21],
        kind: INLINE_SCALAR_FAST_PLAN_CONCAT,
        expected: { tag: ValueTag.String, value: 'north-west', stringId: 0 },
      },
      {
        source: 'MIN(A1,B1,C1)+MAX(A1,B1,C1)',
        deps: [10, 11, 12],
        kind: INLINE_SCALAR_FAST_PLAN_MIN_MAX,
        expected: { tag: ValueTag.Number, value: 11 },
      },
      {
        source: 'ROUND(SQRT(A1),2)',
        deps: [30],
        kind: INLINE_SCALAR_FAST_PLAN_ROUND_SQRT,
        expected: { tag: ValueTag.Number, value: 3.16 },
      },
      {
        source: 'IF(A1>5,"high","low")',
        deps: [50],
        kind: INLINE_SCALAR_FAST_PLAN_IF_STRING,
        expected: { tag: ValueTag.String, value: 'high', stringId: 0 },
      },
    ]

    for (const testCase of cases) {
      const formula = runtimeFormula(testCase.source, testCase.deps)
      expect(classifyInlineScalarFastPlan(formula.compiled)).toBe(testCase.kind)
      expect(tryEvaluateFormulaLeafInlineScalar({ state, formula })).toEqual(testCase.expected)
    }

    const pmtFormula = runtimeFormula('PMT(A1/12,B1,C1)', [40, 41, 42])
    expect(classifyInlineScalarFastPlan(pmtFormula.compiled)).toBe(INLINE_SCALAR_FAST_PLAN_PMT)
    const pmt = tryEvaluateFormulaLeafInlineScalar({ state, formula: pmtFormula })
    expect(pmt?.tag).toBe(ValueTag.Number)
    expect(pmt?.tag === ValueTag.Number ? pmt.value : 0).toBeCloseTo(-106.6185, 4)
  })

  it('runs the interpreter path with parsed dependencies when no inline index plan is cached', () => {
    const state = makeState({
      10: 'north',
      11: 4,
      12: true,
      13: { tag: ValueTag.Error, code: ErrorCode.Ref },
    })

    expect(tryEvaluateFormulaLeafInlineScalar({ state, formula: runtimeFormula('A1&"!"', [10], false) })).toEqual({
      tag: ValueTag.String,
      value: 'north!',
      stringId: 0,
    })
    expect(tryEvaluateFormulaLeafInlineScalar({ state, formula: runtimeFormula('-A1', [11], false) })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(tryEvaluateFormulaLeafInlineScalar({ state, formula: runtimeFormula('IF(A1,"yes","no")', [12], false) })).toEqual({
      tag: ValueTag.String,
      value: 'yes',
      stringId: 0,
    })
    expect(tryEvaluateFormulaLeafInlineScalar({ state, formula: runtimeFormula('A1+1', [13], false) })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('rejects plans that are too large, spilling, unsupported, or missing dependency metadata', () => {
    const compiled = compileFormula('A1+1')
    const unsupported = runtimeFormulaFromCompiled(
      {
        ...compiled,
        jsPlan: [{ opcode: 'push-range', start: 'A1', end: 'A2', refKind: 'cells' }, { opcode: 'return' }],
      },
      [10],
    )

    expect(classifyInlineScalarFastPlan({ ...compiled, jsPlan: [], producesSpill: false, astMatchesSource: true })).toBeUndefined()
    expect(classifyInlineScalarFastPlan({ ...compiled, producesSpill: true })).toBeUndefined()
    expect(classifyInlineScalarFastPlan({ ...compiled, astMatchesSource: false })).toBeUndefined()
    expect(buildInlineScalarPlanCellIndices({ ...compiled, parsedDeps: undefined }, Uint32Array.of(10))).toBeUndefined()
    expect(tryEvaluateFormulaLeafInlineScalar({ state: makeState({ 10: 1 }), formula: unsupported })).toBeUndefined()
  })

  it('covers interpreter arithmetic, comparisons, and stack guards', () => {
    const binary = (operator: Extract<JsPlanInstruction, { opcode: 'binary' }>['operator'], left = 8, right = 2) =>
      evaluateManual([
        { opcode: 'push-number', value: left },
        { opcode: 'push-number', value: right },
        { opcode: 'binary', operator },
        { opcode: 'return' },
      ])

    expect(binary('+')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(binary('-')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(binary('*')).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(binary('/')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(binary('^', 2, 3)).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(binary('/', 8, 0)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(binary(':')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(binary('=')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(binary('<>')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(binary('>')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(binary('>=')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(binary('<')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(binary('<=')).toEqual({ tag: ValueTag.Boolean, value: false })

    expect(
      evaluateManual([
        ...Array.from({ length: 17 }, (_, value): JsPlanInstruction => ({ opcode: 'push-number', value })),
        { opcode: 'return' },
      ]),
    ).toBeUndefined()
  })

  it('covers direct interpreter call paths and error edges', () => {
    const call = (callee: string, argc: number, values: readonly JsPlanInstruction[]) =>
      evaluateManual([...values, { opcode: 'call', callee, argc }, { opcode: 'return' }])

    expect(
      call('concatenate', 3, [
        { opcode: 'push-string', value: 'north' },
        { opcode: 'push-string', value: '-' },
        { opcode: 'push-string', value: 'west' },
      ]),
    ).toEqual({ tag: ValueTag.String, value: 'north-west', stringId: 0 })
    expect(call('CONCATENATE', 0, [])).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(call('CONCATENATE', 1, [{ opcode: 'push-error', code: ErrorCode.Ref }])).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      call('LEN', 2, [
        { opcode: 'push-string', value: 'abc' },
        { opcode: 'push-string', value: 'd' },
      ]),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      call('MAX', 3, [
        { opcode: 'push-number', value: 2 },
        { opcode: 'push-string', value: 'skip' },
        { opcode: 'push-number', value: 7 },
      ]),
    ).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(
      call('MIN', 2, [
        { opcode: 'push-number', value: 2 },
        { opcode: 'push-number', value: 7 },
      ]),
    ).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(
      call('POWER', 2, [
        { opcode: 'push-number', value: 2 },
        { opcode: 'push-number', value: 4 },
      ]),
    ).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    expect(call('POWER', 1, [{ opcode: 'push-number', value: 2 }])).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(call('ROUND', 1, [{ opcode: 'push-number', value: 1.23 }])).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(call('SQRT', 1, [{ opcode: 'push-number', value: -1 }])).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      call('SQRT', 2, [
        { opcode: 'push-number', value: 4 },
        { opcode: 'push-number', value: 2 },
      ]),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      call('PMT', 2, [
        { opcode: 'push-number', value: 0 },
        { opcode: 'push-number', value: 12 },
      ]),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      call('PMT', 3, [
        { opcode: 'push-number', value: 0 },
        { opcode: 'push-number', value: 12 },
        { opcode: 'push-number', value: 1200 },
      ]),
    ).toEqual({
      tag: ValueTag.Number,
      value: -100,
    })
  })
})
