import { describe, expect, it, vi } from 'vitest'
import {
  defineModel,
  findRange,
  formula,
  runWorkbookAction,
  runWorkbookPlan,
  verifyWorkbookReadbacks,
  type WorkbookModel,
  type WorkbookRunAdapter,
} from '../index.js'

function valueModel(): WorkbookModel<{ readonly output: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'run-value-model',

    find(workbook) {
      return {
        output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
      }
    },

    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.output, 12)
        workbook.check.valueEquals(refs.output, 12)
      },
    },
  })
}

function first<T>(values: readonly T[]): T {
  const [value] = values
  if (value === undefined) {
    throw new Error('expected at least one value')
  }
  return value
}

describe('@bilig/workbook run api', () => {
  it('plans, verifies, applies, reads back, and returns done for value checks', async () => {
    const model = valueModel()
    const apply = vi.fn<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>['apply']>(() => ({
      status: 'applied',
      undo: { id: 'undo-1' },
    }))
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      {
        target: first(targets),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(model, 'write', { apply, read })

    expect(apply).toHaveBeenCalledTimes(1)
    expect(read).toHaveBeenCalledTimes(1)
    expect(read.mock.calls[0]?.[0]).toEqual([expect.objectContaining({ label: 'Sheet1!B2' })])
    expect(result).toEqual({
      status: 'done',
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        {
          status: 'passed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      undo: { id: 'undo-1' },
    })
  })

  it('passes formula readback checks with exact normalized formula text', async () => {
    const model = defineModel({
      name: 'run-formula-model',

      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          rate: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const expected = formula.multiply(refs.amount, refs.rate)
          workbook.writeFormula(refs.result, expected)
          workbook.check.formulaEquals(refs.result, expected)
        },
      },
    })
    const source = '(Sheet1!A2)*(Sheet1!B2)'

    const result = await runWorkbookAction(model, 'calculate', {
      apply: () => ({ status: 'applied' }),
      read: (targets) => [
        {
          target: first(targets),
          formula: source,
        },
      ],
    })

    expect(result.status).toBe('done')
    expect(result.checks).toEqual([
      {
        status: 'passed',
        kind: 'formulaEquals',
        target: expect.objectContaining({ label: 'Sheet1!C2' }),
        message: `Sheet1!C2 formula equals ${source}`,
        expectation: {
          kind: 'formulaEquals',
          formula: source,
          inputs: [expect.objectContaining({ label: 'Sheet1!A2' }), expect.objectContaining({ label: 'Sheet1!B2' })],
        },
      },
    ])
  })

  it('does not apply when action planning fails', async () => {
    const model = valueModel()
    const apply = vi.fn<WorkbookRunAdapter['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'missing', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model run-value-model does not define action missing',
        },
      ],
      checks: [],
    })
  })

  it('does not apply when static plan verification fails', async () => {
    const hidden = findRange({ sheetName: 'Sheet1', address: 'Z9' })
    const model = defineModel({
      name: 'invalid-run-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.raw('Sheet1!Z9', { inputs: [hidden] }))
        },
      },
    })
    const apply = vi.fn<WorkbookRunAdapter['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'calculate', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'formula_input_not_resolved',
          message: 'Sheet1!Z9 is used as a formula input but is missing from refsUsed',
        },
      ],
      checks: [],
    })
  })

  it('returns failed when the adapter apply step fails', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected the plan',
          },
        ],
      }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected the plan',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('returns failed when an expected readback is missing', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'applied' }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_missing',
          message: 'Sheet1!B2 has no readback',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('returns failed when value readback mismatches', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'applied' }),
      read: (targets) => [
        {
          target: first(targets),
          value: 13,
        },
      ],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'value_mismatch',
          message: 'Sheet1!B2 expected value 12 but read 13',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('returns failed when formula readback mismatches', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
          },
        },
      ],
      [{ target, formula: '=A2+B2' }],
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [
        {
          status: 'failed',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
          },
        },
      ],
      issues: [
        {
          code: 'formula_mismatch',
          check: expect.objectContaining({ kind: 'formulaEquals' }),
          target,
          expected: 'A2+B2',
          actual: '=A2+B2',
          message: 'Sheet1!C2 expected formula A2+B2 but read =A2+B2',
        },
      ],
    })
  })

  it('supports async apply and async read adapters', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: async () => ({ status: 'applied' }),
      read: async (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(planned.status).toBe('done')
  })

  it('runs an already planned action without model access', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'failed' }),
    })

    expect(planned.status).toBe('failed')
    const planResult = await runWorkbookPlan(
      {
        modelName: 'empty-run-plan',
        actionName: 'noop',
        refs: {},
        refsUsed: [],
        commands: [],
        ops: [],
        changed: [],
        checks: [],
      },
      {
        apply: () => ({ status: 'applied' }),
      },
    )

    expect(planResult).toEqual({
      status: 'done',
      changed: [],
      checks: [],
    })
  })
})
