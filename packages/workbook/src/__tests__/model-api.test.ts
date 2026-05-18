import { describe, expect, it } from 'vitest'
import { parseFormula } from '@bilig/formula'
import {
  buildWorkbookActionPlan,
  check,
  collectWorkbookRefs,
  describePlan,
  describePlanResult,
  describeRef,
  findColumn,
  findName,
  findRange,
  findRows,
  findTable,
  inspectModel,
  isWorkbookRef,
  planWorkbookAction,
  defineModel,
  formula,
  verifyModel,
  verifyPlan,
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
    expect(plan.checks.map((plannedCheck) => plannedCheck.kind)).toEqual(['exists', 'noFormulaErrors', 'noFormulaErrors'])
    expect(verifyPlan(plan)).toEqual({
      status: 'valid',
      modelName: 'custom-model',
      actionName: 'calculate',
      issues: [],
    })
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

  it('exports simple top-level find helpers for generic refs', () => {
    const table = findTable({ name: 'Inputs', sheetName: 'Model', headers: ['Amount', 'Rate'] })
    const amount = findColumn({ table, name: 'Amount' })
    const amountViaTable = table.column('Amount')
    const result = findRange({ sheetName: 'Model', address: 'C2' })
    const namedRate = findName('Rate')
    const rows = findRows({
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Active',
      },
    })

    expect(table).toEqual({
      kind: 'table',
      id: 'table_Model_Inputs_Amount_Rate',
      label: 'Inputs',
      name: 'Inputs',
      sheetName: 'Model',
      headers: ['Amount', 'Rate'],
      column: expect.any(Function),
    })
    expect(amount).toEqual(amountViaTable)
    expect(result).toEqual({
      kind: 'range',
      id: 'range_Model_C2_C2',
      label: 'Model!C2',
      range: {
        sheetName: 'Model',
        startAddress: 'C2',
        endAddress: 'C2',
      },
    })
    expect(namedRate).toEqual({
      kind: 'name',
      id: 'name_Rate',
      label: 'Rate',
      name: 'Rate',
    })
    expect(rows).toEqual({
      kind: 'rows',
      id: 'table_Model_Inputs_Amount_Rate_Status_eq',
      label: 'table_Model_Inputs_Amount_Rate rows where Status eq',
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Active',
      },
    })
  })

  it('exports simple top-level check helpers for generic refs', () => {
    const result = findRange({ sheetName: 'Model', address: 'C2' })

    expect(check.exists(result)).toEqual({
      status: 'planned',
      kind: 'exists',
      target: result,
      message: 'Model!C2 exists',
    })
    expect(check.noFormulaErrors(result)).toEqual({
      status: 'planned',
      kind: 'noFormulaErrors',
      target: result,
      message: 'Model!C2 has no formula errors',
    })
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

  it('describes plans as JSON-safe agent-readable intent', () => {
    const model = defineModel({
      name: 'described-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          table,
          amount: table.column('Amount'),
          rate: table.column('Rate'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.amount, refs.rate))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')
    const described = describePlan(plan)
    const tableDescription = described.refsUsed[0]
    const table = {
      kind: 'table',
      id: 'table_Inputs',
      label: 'Inputs',
      name: 'Inputs',
    } as const
    const amount = {
      kind: 'column',
      id: 'table_Inputs_Amount',
      label: 'Inputs.Amount',
      table,
      name: 'Amount',
    } as const
    const rate = {
      kind: 'column',
      id: 'table_Inputs_Rate',
      label: 'Inputs.Rate',
      table,
      name: 'Rate',
    } as const
    const result = {
      kind: 'range',
      id: 'range_Sheet1_D2_D2',
      label: 'Sheet1!D2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'D2',
        endAddress: 'D2',
      },
    } as const

    expect(describeRef(plan.refs.table)).toEqual(table)
    expect(tableDescription).toBeDefined()
    if (tableDescription === undefined) {
      throw new Error('expected first ref description')
    }
    expect('column' in tableDescription).toBe(false)
    expect(described).toEqual({
      modelName: 'described-model',
      actionName: 'calculate',
      refsUsed: [table, amount, rate, result],
      commands: [
        {
          kind: 'writeFormula',
          target: result,
          formula: '(Inputs[Amount])*(Inputs[Rate])',
          inputs: [amount, rate],
        },
      ],
      ops: [
        {
          kind: 'setCellFormula',
          sheetName: 'Sheet1',
          address: 'D2',
          formula: '(Inputs[Amount])*(Inputs[Rate])',
        },
      ],
      changed: [
        {
          kind: 'writeFormula',
          target: result,
          message: 'Write formula to Sheet1!D2',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'exists',
          target: table,
          message: 'Inputs exists',
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
    expect('refs' in described).toBe(false)
  })

  it('verifies that planned intent only uses resolved refs', () => {
    const model = defineModel({
      name: 'verify-missing-ref-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const hiddenInput = workbook.findRange({ sheetName: 'Sheet1', address: 'B2' })
          workbook.writeFormula(refs.result, formula.add(hiddenInput, 1))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')

    expect(verifyPlan(plan)).toEqual({
      status: 'invalid',
      modelName: 'verify-missing-ref-model',
      actionName: 'calculate',
      issues: [
        {
          code: 'formula_input_not_resolved',
          path: 'commands[0].inputs[0]',
          ref: {
            kind: 'range',
            id: 'range_Sheet1_B2_B2',
            label: 'Sheet1!B2',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B2',
            },
          },
          message: 'Sheet1!B2 is used as a formula input but is missing from refsUsed',
        },
      ],
    })
  })

  it('verifies parseable formulas and matching concrete ops', () => {
    const model = defineModel({
      name: 'verify-broken-plan-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.raw('1+1'))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')
    const [command] = plan.commands
    if (command === undefined || command.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const brokenPlan = {
      ...plan,
      commands: [{ ...command, formula: 'SUM(' }],
      ops: [],
    }

    expect(verifyPlan(brokenPlan).issues).toEqual([
      expect.objectContaining({
        code: 'invalid_formula',
        path: 'commands[0].formula',
      }),
      {
        code: 'missing_concrete_op',
        path: 'commands[0]',
        ref: {
          kind: 'range',
          id: 'range_Sheet1_D2_D2',
          label: 'Sheet1!D2',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D2',
            endAddress: 'D2',
          },
        },
        message: 'Sheet1!D2 has no matching concrete workbook op',
      },
    ])
  })

  it('verifies every model action with JSON-safe planning results', () => {
    const model = defineModel({
      name: 'whole-model-verification',

      find(workbook) {
        return {
          input: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.add(refs.input, 1))
        },
        broken({ refs, workbook }) {
          const hiddenInput = workbook.findRange({ sheetName: 'Sheet1', address: 'B2' })
          workbook.writeFormula(refs.result, formula.add(hiddenInput, 1))
        },
        fail() {
          throw new Error('missing workbook target')
        },
      },
    })

    const verification = verifyModel(model)

    expect(verification.status).toBe('invalid')
    expect(verification.modelName).toBe('whole-model-verification')
    expect(
      verification.actions.map((action) => ({
        actionName: action.actionName,
        planning: action.planning.status,
        verification: action.verification?.status,
      })),
    ).toEqual([
      {
        actionName: 'broken',
        planning: 'planned',
        verification: 'invalid',
      },
      {
        actionName: 'calculate',
        planning: 'planned',
        verification: 'valid',
      },
      {
        actionName: 'fail',
        planning: 'failed',
        verification: undefined,
      },
    ])
    expect(verification.actions[0]?.verification?.issues.map((issue) => issue.code)).toEqual(['formula_input_not_resolved'])
    expect(verification.actions[2]?.planning).toEqual({
      status: 'failed',
      modelName: 'whole-model-verification',
      actionName: 'fail',
      errors: [
        {
          code: 'action_failed',
          message: 'missing workbook target',
        },
      ],
      checks: [],
    })
    expect(JSON.parse(JSON.stringify(verification))).toEqual(verification)
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
      modelName: 'failing-model',
      actionName: 'missing',
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
      modelName: 'failing-model',
      actionName: 'calculate',
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
      expect(result.modelName).toBe('checkable-failure-model')
      expect(result.actionName).toBe('calculate')
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

  it('describes failed plan results without raw workbook refs', () => {
    const model = defineModel({
      name: 'described-failure-model',

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
          throw new Error('missing output target')
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    const described = describePlanResult(result)

    expect(described).toEqual({
      status: 'failed',
      modelName: 'described-failure-model',
      actionName: 'calculate',
      errors: [
        {
          code: 'action_failed',
          message: 'missing output target',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'exists',
          target: {
            kind: 'table',
            id: 'table_Inputs',
            label: 'Inputs',
            name: 'Inputs',
          },
          message: 'Inputs exists',
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
  })
})
