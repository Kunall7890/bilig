import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  collectWorkbookRefData,
  collectWorkbookRefs,
  defineModel,
  describePlan,
  findRows,
  findTable,
  findRange,
  formula,
  hydrateWorkbookRef,
  hydrateWorkbookRefs,
  isWorkbookRef,
  isWorkbookRefData,
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
})
