import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  checkPlanData,
  collectWorkbookRefData,
  collectWorkbookRefs,
  defineModel,
  describePlan,
  describeRuntimeRequirements,
  findRows,
  findTable,
  findRange,
  formula,
  hydratePlanData,
  hydrateWorkbookRef,
  hydrateWorkbookRefs,
  isPlanData,
  isWorkbookRef,
  isWorkbookRefData,
  runWorkbookPlan,
  toPlanData,
  toWorkbookRefData,
  verifyPlanData,
  type WorkbookActionPlanDescription,
} from '../index.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isPlanDescription(value: unknown): value is WorkbookActionPlanDescription {
  return (
    isRecord(value) &&
    typeof value['modelName'] === 'string' &&
    typeof value['actionName'] === 'string' &&
    Array.isArray(value['refsUsed']) &&
    Array.isArray(value['commands']) &&
    Array.isArray(value['ops']) &&
    Array.isArray(value['changed']) &&
    Array.isArray(value['checks'])
  )
}

function accessorArray(get: () => unknown): unknown[] {
  const value = Array.from<unknown>({ length: 1 })
  Object.defineProperty(value, '0', {
    enumerable: true,
    get,
  })
  return value
}

function nonEnumerableArray(value: unknown): unknown[] {
  const values: unknown[] = []
  Object.defineProperty(values, '0', {
    enumerable: false,
    value,
  })
  values.length = 1
  return values
}

describe('@bilig/workbook transport api', () => {
  it('round-trips refs as plain JSON data and hydrates ergonomic helpers back', () => {
    const table = findTable({ name: 'Inputs', sheetName: 'Model', headers: ['Amount', 'Status'] })
    const rows = findRows({ table, where: { column: 'Status', op: 'eq', value: 'ready' } })
    const amount = rows.column('Amount')
    const data = toWorkbookRefData(amount)
    const parsed = JSON.parse(JSON.stringify(data)) as unknown
    if (data.kind !== 'column' || data.rows === undefined) {
      throw new Error('expected row-filtered column data')
    }

    expect(isWorkbookRefData(parsed)).toBe(true)
    expect(isWorkbookRef(parsed)).toBe(false)
    expect(collectWorkbookRefData({ amount: parsed })).toEqual([data, data.rows, data.table])

    if (!isWorkbookRefData(parsed)) {
      throw new Error('expected parsed workbook ref data')
    }
    const hydrated = hydrateWorkbookRef(parsed)

    expect(isWorkbookRef(hydrated)).toBe(true)
    expect(hydrated).toEqual(amount)
    if (hydrated.kind !== 'column' || hydrated.rows === undefined) {
      throw new Error('expected hydrated row-filtered column')
    }
    expect(hydrated.table.column('Amount')).toEqual(table.column('Amount'))
    expect(hydrated.rows.column('Amount')).toEqual(amount)
  })

  it('hydrates arbitrary consumer ref shapes without requiring hidden methods in transport data', () => {
    const table = findTable({ name: 'Inputs' })
    const rows = findRows({ table, where: { column: 'Status', op: 'eq', value: 'ready' } })
    const result = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const transport = JSON.parse(
      JSON.stringify({
        groups: [
          {
            table,
            amount: rows.column('Amount'),
          },
        ],
        result,
      }),
    ) as unknown

    expect(collectWorkbookRefs(transport)).toEqual([result])

    const hydrated = hydrateWorkbookRefs(transport)
    if (!isRecord(hydrated) || !Array.isArray(hydrated['groups']) || !isWorkbookRef(hydrated['result'])) {
      throw new Error('expected hydrated transport shape')
    }
    const [group] = hydrated['groups']
    if (!isRecord(group) || !isWorkbookRef(group['table']) || !isWorkbookRef(group['amount']) || group['amount'].kind !== 'column') {
      throw new Error('expected hydrated transport group')
    }

    expect(collectWorkbookRefs(hydrated)).toEqual([group['table'], group['amount'], group['amount'].rows, hydrated['result']])
    if (group['table'].kind !== 'table') {
      throw new Error('expected table ref')
    }
    expect(group['table'].column('Amount')).toEqual(table.column('Amount'))
  })

  it('does not execute accessors while collecting or hydrating transported refs', () => {
    const result = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const resultData = toWorkbookRefData(result)
    const transport = {
      result: resultData,
    }
    Object.defineProperty(transport, 'hidden', {
      enumerable: true,
      get() {
        throw new Error('transport accessor should not run')
      },
    })

    expect(collectWorkbookRefData(transport)).toEqual([resultData])

    const hydrated = hydrateWorkbookRefs(transport)
    expect(isRecord(hydrated)).toBe(true)
    if (!isRecord(hydrated)) {
      throw new Error('expected hydrated object')
    }
    expect(Object.hasOwn(hydrated, 'hidden')).toBe(false)
    expect(hydrated['result']).toEqual(result)
    expect(collectWorkbookRefs(hydrated)).toEqual([result])

    let arrayGetterInvoked = false
    const arrayTransport = {
      refs: accessorArray(() => {
        arrayGetterInvoked = true
        throw new Error('array accessor should not run')
      }),
    }

    expect(collectWorkbookRefData(arrayTransport)).toEqual([])
    const hydratedArrayTransport = hydrateWorkbookRefs(arrayTransport)
    expect(isRecord(hydratedArrayTransport)).toBe(true)
    if (!isRecord(hydratedArrayTransport) || !Array.isArray(hydratedArrayTransport['refs'])) {
      throw new Error('expected hydrated array transport shape')
    }
    expect(hydratedArrayTransport['refs']).toEqual([])
    expect(arrayGetterInvoked).toBe(false)
  })

  it('rejects accessor-backed ref data arrays without invoking getters', () => {
    let headerGetterInvoked = false
    const tableData = {
      kind: 'table',
      id: 'table_Inputs',
      label: 'Inputs',
      headers: accessorArray(() => {
        headerGetterInvoked = true
        throw new Error('header getter should not run')
      }),
    }

    expect(isWorkbookRefData(tableData)).toBe(false)
    expect(headerGetterInvoked).toBe(false)
  })

  it('copies known nested ref fields without invoking extra accessors', () => {
    const range = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const data = toWorkbookRefData(range)
    if (data.kind !== 'range') {
      throw new Error('expected range ref data')
    }

    let extraGetterInvoked = false
    Object.defineProperty(data.range, 'extra', {
      enumerable: true,
      get() {
        extraGetterInvoked = true
        throw new Error('extra getter should not run')
      },
    })

    expect(toWorkbookRefData(data)).toEqual({
      kind: 'range',
      id: 'range_Sheet1_D2_D2',
      label: 'Sheet1!D2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'D2',
        endAddress: 'D2',
      },
    })
    expect(hydrateWorkbookRef(data)).toEqual(range)
    expect(extraGetterInvoked).toBe(false)
  })

  it('verifies JSON-safe plan descriptions after transport round-trip', () => {
    const model = defineModel({
      name: 'transport-plan-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          amount: table.column('Amount'),
          rate: table.column('Rate'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.amount, refs.rate))
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'calculate')
    const transported: unknown = JSON.parse(JSON.stringify(describePlan(plan)))
    if (!isPlanDescription(transported)) {
      throw new Error('expected transported plan description')
    }

    expect(verifyPlanData(transported)).toEqual({
      status: 'valid',
      modelName: 'transport-plan-model',
      actionName: 'calculate',
      issues: [],
    })

    const [command] = transported.commands
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected transported formula command')
    }
    const hidden = toWorkbookRefData(findRange({ sheetName: 'Sheet1', address: 'Z9' }))
    const commands = [...transported.commands]
    commands[0] = Object.assign({}, command, {
      inputs: [...command.inputs, hidden],
    })
    const broken: WorkbookActionPlanDescription = {
      ...transported,
      commands,
    }

    expect(verifyPlanData(broken).issues).toEqual([
      expect.objectContaining({
        code: 'formula_input_not_resolved',
        path: 'commands[0].inputs[2]',
      }),
      expect.objectContaining({
        code: 'formula_input_not_labeled',
        path: 'commands[0].inputs[2]',
      }),
    ])
  })

  it('returns structured plan data issues before hydration', () => {
    const model = defineModel({
      name: 'transport-check-plan-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'write')
    const parsedData: unknown = JSON.parse(JSON.stringify(toPlanData(plan)))
    if (!isRecord(parsedData)) {
      throw new Error('expected plan data object')
    }
    const data = parsedData

    expect(checkPlanData(data)).toEqual({
      status: 'valid',
      plan: data,
      issues: [],
    })
    expect(checkPlanData(null)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })

    const broken = {
      ...data,
      input: {
        rows: [1, Number.NaN],
      },
      modelName: 12,
      refsUsed: [{ kind: 'range' }],
      commands: [{ kind: 'writeValue', target: data['refsUsed'], value: undefined }],
      changed: [{ kind: 'writeValue' }],
      checks: [{ status: 'planned', kind: 'exists', message: 'Exists', proof: { when: new Date('2026-05-23T00:00:00Z') } }],
    }

    expect(checkPlanData(broken)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'modelName',
          message: 'Workbook plan data modelName must be a string',
        },
        {
          code: 'invalid_plan_data',
          path: 'input.rows[1]',
          message: 'Workbook plan data input must be JSON-safe: Action input at input.rows[1] must be a finite number',
        },
        {
          code: 'invalid_plan_data',
          path: 'refsUsed[0]',
          message: 'Workbook plan data ref at refsUsed[0] is invalid',
        },
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
        {
          code: 'invalid_plan_data',
          path: 'changed[0]',
          message: 'Workbook plan data change at changed[0] is invalid',
        },
        {
          code: 'invalid_plan_data',
          path: 'checks[0].proof.when',
          message: 'Workbook plan data check proof must be JSON-safe: Action input at input.when must be a plain JSON object, not Date',
        },
      ],
    })
    expect(() => hydratePlanData(broken)).toThrow('Workbook plan data is invalid: Workbook plan data modelName must be a string')
  })

  it('treats transported plan fields as own data, not inherited prototype data', () => {
    const model = defineModel({
      name: 'transport-own-data-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'write')
    const data = structuredClone(toPlanData(plan))
    const inheritedData = Object.create(data) as unknown

    expect(isPlanData(inheritedData)).toBe(false)
    expect(checkPlanData(inheritedData)).toEqual({
      status: 'invalid',
      issues: expect.arrayContaining([
        {
          code: 'invalid_plan_data',
          path: 'modelName',
          message: 'Workbook plan data modelName must be a string',
        },
        {
          code: 'invalid_plan_data',
          path: 'actionName',
          message: 'Workbook plan data actionName must be a string',
        },
        {
          code: 'invalid_plan_data',
          path: 'refsUsed',
          message: 'Workbook plan data refsUsed must be an array',
        },
      ]),
    })
  })

  it('rejects accessor-backed transported plan arrays without invoking getters', () => {
    const model = defineModel({
      name: 'transport-array-data-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'write')
    const data = structuredClone(toPlanData(plan))

    let commandGetterInvoked = false
    const accessorCommands = accessorArray(() => {
      commandGetterInvoked = true
      throw new Error('getter must not run')
    })
    const accessorData = {
      ...data,
      commands: accessorCommands,
    }

    expect(isPlanData(accessorData)).toBe(false)
    expect(checkPlanData(accessorData)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
      ],
    })
    expect(() => hydratePlanData(accessorData)).toThrow(
      'Workbook plan data is invalid: Workbook plan data command at commands[0] is invalid',
    )
    expect(commandGetterInvoked).toBe(false)

    let refGetterInvoked = false
    const accessorRefs = accessorArray(() => {
      refGetterInvoked = true
      throw new Error('getter must not run')
    })

    expect(
      checkPlanData({
        ...data,
        refsUsed: accessorRefs,
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'refsUsed[0]',
          message: 'Workbook plan data ref at refsUsed[0] is invalid',
        },
      ],
    })
    expect(refGetterInvoked).toBe(false)

    const firstChange = (data['changed'] as readonly unknown[])[0]
    expect(
      checkPlanData({
        ...data,
        changed: nonEnumerableArray(firstChange),
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'changed[0]',
          message: 'Workbook plan data change at changed[0] is invalid',
        },
      ],
    })
  })

  it('runs transported plan data without the consumer refs object', async () => {
    const model = defineModel({
      name: 'transport-executable-plan-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          amount: table.column('Amount'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const expected = formula.raw('amount_token*2', {
            labels: [{ name: 'amount_token', ref: refs.amount }],
          })
          workbook.writeFormula(refs.result, expected)
          workbook.check.formulaEquals(refs.result, expected)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'calculate')
    const transported: unknown = JSON.parse(JSON.stringify(toPlanData(plan)))

    expect(isPlanData(transported)).toBe(true)
    if (!isPlanData(transported)) {
      throw new Error('expected executable transported plan data')
    }

    const hydrated = hydratePlanData(transported)
    expect(hydrated.refs).toEqual({ refsUsed: hydrated.refsUsed })
    expect(verifyPlanData(transported).status).toBe('valid')
    expect(describeRuntimeRequirements(transported).requirements.map((entry) => entry.capability)).toEqual(['writeFormula', 'read'])

    const result = await runWorkbookPlan(transported, {
      apply(receivedPlan) {
        expect(receivedPlan.refs).toEqual({ refsUsed: receivedPlan.refsUsed })
        return {
          status: 'applied',
          previewOps: receivedPlan.ops,
          appliedOps: receivedPlan.ops,
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'amount_token*2',
        }))
      },
    })

    expect(result).toMatchObject({
      status: 'done',
      checks: [
        {
          status: 'passed',
          kind: 'formulaEquals',
          proof: {
            source: 'readback',
            formula: 'amount_token*2',
          },
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(describePlan(hydrated)))).toEqual(transported)
  })
})
