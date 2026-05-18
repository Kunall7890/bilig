import { describe, expect, it } from 'vitest'
import { parseFormula } from '@bilig/formula'
import {
  buildWorkbookActionPlan,
  collectWorkbookRefs,
  inspectModel,
  isWorkbookRef,
  planWorkbookAction,
  defineModel,
  formula,
} from '../index.js'

describe('@bilig/workbook model api', () => {
  it('preserves model metadata, refs, checks, commands, and concrete workbook ops', () => {
    const model = defineModel({
      name: 'custom-model',

      find(workbook) {
        const table = workbook.findTable({ headers: ['Base', 'Rate', 'Result'] })

        return {
          table,
          base: table.column('Base'),
          rate: table.column('Rate'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table), workbook.check.noFormulaErrors(refs.result)]
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
          workbook.check.noFormulaErrors(refs.result)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')

    expect(plan.modelName).toBe('custom-model')
    expect(plan.actionName).toBe('calculate')
    expect(plan.refs.table.headers).toEqual(['Base', 'Rate', 'Result'])
    expect(plan.refsUsed).toEqual([plan.refs.table, plan.refs.base, plan.refs.rate, plan.refs.result])
    expect(plan.commands).toEqual([
      {
        kind: 'writeFormula',
        target: plan.refs.result,
        formula: '(__bilig_ref_table_Base_Rate_Result_Base)*(__bilig_ref_table_Base_Rate_Result_Rate)',
        inputs: [plan.refs.base, plan.refs.rate],
      },
    ])
    expect(plan.ops).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'C2',
        formula: '(__bilig_ref_table_Base_Rate_Result_Base)*(__bilig_ref_table_Base_Rate_Result_Rate)',
      },
    ])
    expect(plan.changed).toEqual([
      {
        kind: 'writeFormula',
        target: plan.refs.result,
        message: 'Write formula to Sheet1!C2',
      },
    ])
    expect(plan.checks.map((check) => check.kind)).toEqual(['exists', 'noFormulaErrors', 'noFormulaErrors'])
    const [op] = plan.ops
    expect(op?.kind).toBe('setCellFormula')
    if (op?.kind === 'setCellFormula') {
      parseFormula(op.formula)
    }
  })

  it('creates formula helpers that normalize through the formula parser', () => {
    const amount = formula.raw('Sheet1!A1')
    const rate = formula.raw('Sheet1!B1')
    const source = formula.source(formula.sum(formula.multiply(amount, rate), 10))

    expect(source).toBe('SUM((Sheet1!A1)*(Sheet1!B1),10)')
    parseFormula(source)
  })

  it('tracks formula inputs separately from formula text', () => {
    const model = defineModel({
      name: 'formula-input-model',

      find(workbook) {
        const inputs = workbook.findTable({ name: 'Inputs' })
        return {
          amount: inputs.column('Amount'),
          rate: inputs.column('Rate'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.sum(formula.multiply(refs.amount, refs.rate), refs.amount))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')
    const [command] = plan.commands

    expect(command).toEqual({
      kind: 'writeFormula',
      target: plan.refs.result,
      formula: 'SUM((Inputs[Amount])*(Inputs[Rate]),Inputs[Amount])',
      inputs: [plan.refs.amount, plan.refs.rate],
    })
  })

  it('collects workbook refs from arbitrary consumer ref shapes', () => {
    const model = defineModel({
      name: 'nested-ref-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        const amount = table.column('Amount')
        return {
          groups: [
            {
              table,
              amount,
              duplicate: amount,
            },
          ],
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'E2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, refs.groups[0]?.amount ?? formula.raw('0'))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')

    expect(collectWorkbookRefs(plan.refs)).toEqual([plan.refs.groups[0]?.table, plan.refs.groups[0]?.amount, plan.refs.result])
    expect(plan.refsUsed).toEqual([plan.refs.groups[0]?.table, plan.refs.groups[0]?.amount, plan.refs.result])
    expect(isWorkbookRef(plan.refs.result)).toBe(true)
  })

  it('collects workbook refs safely from cyclic objects', () => {
    const model = defineModel({
      name: 'cyclic-ref-model',

      find(workbook) {
        const result = workbook.findRange({ sheetName: 'Sheet1', address: 'F2' })
        const refs: { result: typeof result; self?: unknown } = { result }
        refs.self = refs
        return refs
      },

      actions: {
        clear({ refs, workbook }) {
          workbook.clear(refs.result)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'clear')

    expect(plan.refsUsed).toEqual([plan.refs.result])
  })

  it('rejects invalid formulas before they become workbook actions', () => {
    expect(() => formula.raw('SUM(')).toThrowError()
  })

  it('rejects models that cannot do anything', () => {
    expect(() =>
      defineModel({
        name: 'empty-model',
        find() {
          return {}
        },
        actions: {},
      }),
    ).toThrowError('Workbook model empty-model must define at least one action')
  })

  it('describes model actions without running find or actions', () => {
    const model = defineModel({
      name: 'inspectable-model',
      find() {
        throw new Error('find should not run during inspection')
      },
      checks() {
        throw new Error('checks should not run during inspection')
      },
      actions: {
        calculate() {
          throw new Error('action should not run during inspection')
        },
        reset() {
          throw new Error('action should not run during inspection')
        },
      },
    })

    expect(inspectModel(model)).toEqual({
      name: 'inspectable-model',
      actions: ['calculate', 'reset'],
      hasChecks: true,
    })
  })

  it('returns structured planning failures instead of forcing agents to catch exceptions', () => {
    const model = defineModel({
      name: 'failing-model',
      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: {
        calculate() {
          throw new Error('formula target was not resolved')
        },
      },
    })

    expect(planWorkbookAction(model, 'missing')).toEqual({
      status: 'failed',
      checks: [],
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model failing-model does not define action missing',
        },
      ],
    })

    expect(planWorkbookAction(model, 'calculate')).toEqual({
      status: 'failed',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'formula target was not resolved',
        },
      ],
    })
  })

  it('keeps planned checks when action planning fails', () => {
    const model = defineModel({
      name: 'checkable-failure-model',
      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },
      actions: {
        calculate() {
          throw new Error('cannot write without a result target')
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('failed')
    if (result.status === 'failed') {
      expect(result.checks).toEqual([
        {
          status: 'planned',
          kind: 'exists',
          target: expect.objectContaining({
            kind: 'table',
            name: 'Inputs',
          }),
          message: 'Inputs exists',
        },
      ])
      expect(result.errors).toEqual([
        {
          code: 'action_failed',
          message: 'cannot write without a result target',
        },
      ])
    }
  })
})
