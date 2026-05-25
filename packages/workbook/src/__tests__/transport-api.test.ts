import { describe, expect, it, vi } from 'vitest'
import {
  buildWorkbookActionPlan,
  checkPlanData,
  checkWorkbookRef,
  checkWorkbookRefData,
  collectWorkbookRefData,
  collectWorkbookRefs,
  createWorkbookFindApi,
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
  normalizeRangeRef,
  runWorkbookPlan,
  toPlanData,
  toWorkbookRefData,
  verifyPlanData,
  workbookPlanId,
  type WorkbookActionPlanDescription,
} from '../index.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function mutableRecord(value: unknown): Record<string, unknown> {
  if (!isRecord(value) || Array.isArray(value)) {
    throw new Error('expected object record')
  }
  return value
}

function mutableRecordArray(value: unknown): Record<string, unknown>[] {
  if (!Array.isArray(value)) {
    throw new Error('expected object record array')
  }
  return value.map((entry) => mutableRecord(entry))
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

function arrayBackedRecord(fields: Record<string, unknown>): unknown[] {
  const value: unknown[] = []
  for (const [key, entry] of Object.entries(fields)) {
    Object.defineProperty(value, key, {
      enumerable: true,
      value: entry,
    })
  }
  return value
}

function customPrototype(value: object): unknown {
  const custom = new (class {
    readonly inherited = true
  })()
  Object.defineProperties(custom, Object.getOwnPropertyDescriptors(value))
  return custom
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
    expect(Object.isFrozen(data)).toBe(true)
    expect(Object.isFrozen(data.table)).toBe(true)
    expect(Object.isFrozen(data.table.headers)).toBe(true)
    expect(Object.isFrozen(data.rows)).toBe(true)
    expect(Object.isFrozen(data.rows.where)).toBe(true)
    expect(Object.isFrozen(collectWorkbookRefData({ amount: parsed }))).toBe(true)

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

  it('checks transported refs with structured frozen issues', () => {
    const table = findTable({ name: 'Inputs', headers: ['Amount'] })
    const rows = findRows({ table, where: { column: 'Status', op: 'eq', value: 'ready' } })
    const amount = rows.column('Amount')
    const data = toWorkbookRefData(amount)
    const valid = checkWorkbookRefData(JSON.parse(JSON.stringify(data)))

    expect(valid).toEqual({
      status: 'valid',
      ref: data,
      issues: [],
    })
    expect(Object.isFrozen(valid)).toBe(true)
    expect(Object.isFrozen(valid.issues)).toBe(true)
    if (valid.status !== 'valid') {
      throw new Error('expected valid ref')
    }
    expect(Object.isFrozen(valid.ref)).toBe(true)
    expect(Object.isFrozen(valid.ref.kind === 'column' ? valid.ref.table : valid.ref)).toBe(true)

    const invalid = checkWorkbookRefData({
      kind: 'rows',
      id: 'rows',
      label: 'Rows',
      where: {
        column: 'Status',
        op: 'contains',
        value: 12,
      },
    })

    expect(invalid).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'ref.where.value',
          message: 'Workbook ref data ref.where.value is not compatible with row operator contains',
        },
      ],
    })
    expect(Object.isFrozen(invalid)).toBe(true)
    expect(Object.isFrozen(invalid.issues)).toBe(true)
    expect(Object.isFrozen(invalid.issues[0])).toBe(true)
  })

  it('checks live refs with structured helper issues before transport', () => {
    const table = findTable({ name: 'Inputs' })
    const amount = table.column('Amount')
    const valid = checkWorkbookRef(amount)

    expect(valid).toEqual({
      status: 'valid',
      ref: amount,
      data: toWorkbookRefData(amount),
      issues: [],
    })
    expect(Object.isFrozen(valid)).toBe(true)
    expect(Object.isFrozen(valid.issues)).toBe(true)
    if (valid.status !== 'valid') {
      throw new Error('expected valid ref')
    }
    expect(Object.isFrozen(valid.data)).toBe(true)

    const dataOnlyColumn = toWorkbookRefData(amount)
    expect(isWorkbookRefData(dataOnlyColumn)).toBe(true)
    expect(isWorkbookRef(dataOnlyColumn)).toBe(false)
    expect(checkWorkbookRef(dataOnlyColumn)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_field',
          path: 'ref.table.column',
          message: 'Workbook ref ref.table.column is required',
        },
      ],
    })
  })

  it('checks accessor-backed transported refs without invoking getters', () => {
    let getterInvoked = false
    const ref = {
      kind: 'range',
      id: 'range',
      label: 'Range',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    Object.defineProperty(ref.range, 'startAddress', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('startAddress getter must not run')
      },
    })

    expect(isWorkbookRefData(ref)).toBe(false)
    expect(checkWorkbookRefData(ref)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'ref.range.startAddress',
          message: 'Workbook ref data ref.range.startAddress must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects transported column refs whose nested refs have the wrong kind', () => {
    const table = toWorkbookRefData(findTable({ name: 'Inputs' }))
    const rows = toWorkbookRefData(findRows({ table: findTable({ name: 'Inputs' }), where: { column: 'Status', op: 'eq', value: 'ok' } }))

    expect(
      checkWorkbookRefData({
        kind: 'column',
        id: 'bad-column',
        label: 'Bad',
        table: rows,
        rows: table,
        name: 'Amount',
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'ref.table.kind',
          message: 'Workbook ref data ref.table must be a table ref',
        },
        {
          code: 'invalid_field',
          path: 'ref.rows.kind',
          message: 'Workbook ref data ref.rows must be a rows ref',
        },
      ],
    })
  })

  it('returns frozen selector and find helper data for stable agent handoff', () => {
    const api = createWorkbookFindApi()
    const normalizedRange = normalizeRangeRef({ sheetName: 'Sheet1', address: 'A1' })
    const ref = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const refData = toWorkbookRefData(ref)
    const refs = collectWorkbookRefs({ ref })

    expect(Object.isFrozen(api)).toBe(true)
    expect(Object.isFrozen(normalizedRange)).toBe(true)
    expect(Object.isFrozen(refData)).toBe(true)
    if (refData.kind !== 'range') {
      throw new Error('expected range ref data')
    }
    expect(Object.isFrozen(refData.range)).toBe(true)
    expect(Object.isFrozen(refs)).toBe(true)
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

  it('rejects transported row refs with incompatible operator values', () => {
    const table = findTable({ name: 'Inputs' })
    const rows = findRows({ table, where: { column: 'Status', op: 'contains', value: 'ready' } })
    const data = toWorkbookRefData(rows)
    const invalidRowsData = {
      ...data,
      where: {
        ...data.where,
        value: 12,
      },
    }

    expect(isWorkbookRefData(invalidRowsData)).toBe(false)
    expect(collectWorkbookRefData({ rows: invalidRowsData })).toEqual([toWorkbookRefData(table)])
    expect(() => hydrateWorkbookRef(invalidRowsData)).toThrowError('Workbook ref data is invalid')
    expect(() => toWorkbookRefData(invalidRowsData)).toThrowError('Workbook ref data is invalid')

    const invalidColumnData = {
      kind: 'column',
      id: 'bad-column',
      label: 'Bad column',
      table: toWorkbookRefData(table),
      rows: invalidRowsData,
      name: 'Amount',
    }

    expect(isWorkbookRefData(invalidColumnData)).toBe(false)
    expect(collectWorkbookRefData({ amount: invalidColumnData })).toEqual([toWorkbookRefData(table)])
  })

  it('copies known nested ref fields without invoking extra accessors', () => {
    const range = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const frozenData = toWorkbookRefData(range)
    if (frozenData.kind !== 'range') {
      throw new Error('expected range ref data')
    }
    expect(Object.isFrozen(frozenData.range)).toBe(true)
    const data = {
      ...frozenData,
      range: { ...frozenData.range },
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

    const valid = checkPlanData(data)
    expect(valid).toEqual({
      status: 'valid',
      plan: data,
      issues: [],
    })
    expect(valid.status === 'valid' ? valid.plan : undefined).not.toBe(data)
    expect(Object.isFrozen(valid)).toBe(true)
    expect(valid.status === 'valid' ? Object.isFrozen(valid.plan) : false).toBe(true)
    expect(Object.isFrozen(valid.issues)).toBe(true)

    const invalidNull = checkPlanData(null)
    expect(invalidNull).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })
    expect(Object.isFrozen(invalidNull)).toBe(true)
    expect(Object.isFrozen(invalidNull.issues)).toBe(true)
    expect(Object.isFrozen(invalidNull.issues[0])).toBe(true)

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

  it('canonicalizes transported plan data before ids and hydration', () => {
    const model = defineModel({
      name: 'transport-canonical-plan-model',

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
    const noisy = mutableRecord(structuredClone(data))
    noisy['agentScratchpad'] = {
      ignored: true,
    }
    const commands = mutableRecordArray(noisy['commands'])
    const firstCommand = commands[0]
    if (firstCommand === undefined) {
      throw new Error('expected command')
    }
    firstCommand['agentScratchpad'] = 'ignored'
    const target = mutableRecord(firstCommand['target'])
    target['agentScratchpad'] = 'ignored'
    const range = mutableRecord(target['range'])
    range['agentScratchpad'] = 'ignored'

    let extraGetterInvoked = false
    Object.defineProperty(noisy, 'hiddenScratchpad', {
      enumerable: true,
      get() {
        extraGetterInvoked = true
        throw new Error('extra getter must not run')
      },
    })

    const check = checkPlanData(noisy)
    expect(check).toEqual({
      status: 'valid',
      plan: data,
      issues: [],
    })
    expect(isPlanData(noisy)).toBe(true)
    if (!isPlanData(noisy)) {
      throw new Error('expected noisy plan data')
    }
    expect(workbookPlanId(noisy)).toBe(workbookPlanId(data))
    expect(describePlan(hydratePlanData(noisy))).toEqual(data)
    expect(extraGetterInvoked).toBe(false)
  })

  it('returns failed run results for invalid transported plan data without applying', async () => {
    let commandGetterInvoked = false
    const commands = accessorArray(() => {
      commandGetterInvoked = true
      throw new Error('command getter must not run')
    })
    const apply = vi.fn(() => ({ status: 'applied' as const }))

    const result = await runWorkbookPlan(
      {
        modelName: 12,
        actionName: 'write',
        refsUsed: [],
        commands,
        ops: [],
        changed: [],
        checks: [],
      },
      { apply },
    )

    expect(apply).not.toHaveBeenCalled()
    expect(commandGetterInvoked).toBe(false)
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_plan_data',
          path: 'modelName',
          issueCode: 'invalid_plan_data',
          message: 'Workbook plan data modelName must be a string',
        },
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          issueCode: 'invalid_plan_data',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
      ],
      changed: [],
      checks: [],
    })
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
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })
  })

  it('rejects custom-prototype transported plan data', () => {
    const model = defineModel({
      name: 'transport-custom-prototype-plan-model',
      find() {
        return {}
      },
      actions: {
        seed({ workbook }) {
          workbook.writeValue(findRange({ sheetName: 'Sheet1', address: 'A1' }), 1)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'seed')

    expect(checkPlanData(customPrototype(describePlan(plan)))).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })
  })

  it('hydrates transported plan data without invoking inherited optional getters', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    let getterInvoked = false
    const failOnAccess = (): never => {
      getterInvoked = true
      throw new Error('inherited optional getter must not run')
    }
    const planPrototype = Object.defineProperty({}, 'input', {
      get: failOnAccess,
    })
    const commandPrototype = Object.defineProperties(
      {},
      {
        style: {
          get: failOnAccess,
        },
        numberFormat: {
          get: failOnAccess,
        },
      },
    )
    const changePrototype = Object.defineProperty({}, 'target', {
      get: failOnAccess,
    })
    const checkPrototype = Object.defineProperties(
      {},
      {
        target: {
          get: failOnAccess,
        },
        refs: {
          get: failOnAccess,
        },
        expectation: {
          get: failOnAccess,
        },
        proof: {
          get: failOnAccess,
        },
      },
    )
    const command = Object.setPrototypeOf(
      {
        kind: 'format',
        target,
        numberFormat: '0.00',
      },
      commandPrototype,
    )
    const change = Object.setPrototypeOf(
      {
        kind: 'format',
        message: 'formatted range',
      },
      changePrototype,
    )
    const check = Object.setPrototypeOf(
      {
        status: 'planned',
        kind: 'exists',
        message: 'range exists',
      },
      checkPrototype,
    )
    const data = Object.setPrototypeOf(
      {
        modelName: 'transport-inherited-optional-plan-model',
        actionName: 'format',
        refsUsed: [target],
        commands: [command],
        ops: [],
        changed: [change],
        checks: [check],
      },
      planPrototype,
    )

    expect(checkPlanData(data)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })
    expect(() => hydratePlanData(data)).toThrow('Workbook plan data is invalid: Workbook plan data must be an object')
    expect(getterInvoked).toBe(false)
  })

  it('rejects transported format commands with empty style patches', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const data = {
      modelName: 'transport-empty-style-plan-model',
      actionName: 'format',
      refsUsed: [target],
      commands: [
        {
          kind: 'format',
          target,
          style: { font: {} },
        },
      ],
      ops: [],
      changed: [],
      checks: [],
    }

    expect(isPlanData(data)).toBe(false)
    expect(checkPlanData(data)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
      ],
    })
  })

  it('rejects array-backed transported plan roots as uninspectable handoff data', () => {
    const model = defineModel({
      name: 'transport-array-backed-root-model',

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
    const data = mutableRecord(structuredClone(toPlanData(plan)))
    const arrayBackedData = arrayBackedRecord(data)

    expect(isPlanData(arrayBackedData)).toBe(false)
    expect(checkPlanData(arrayBackedData)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data must be an object',
        },
      ],
    })
    expect(() => hydratePlanData(arrayBackedData)).toThrow('Workbook plan data is invalid: Workbook plan data must be an object')
  })

  it('rejects array-backed transported plan entries as uninspectable handoff data', () => {
    const model = defineModel({
      name: 'transport-array-backed-entry-model',

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
    const data = mutableRecord(structuredClone(toPlanData(plan)))
    const rawCommands = data['commands']
    if (!Array.isArray(rawCommands)) {
      throw new Error('expected command array')
    }
    const firstCommand = mutableRecord(rawCommands[0])
    if (firstCommand === undefined) {
      throw new Error('expected command')
    }
    rawCommands[0] = arrayBackedRecord(firstCommand)

    expect(isPlanData(data)).toBe(false)
    expect(checkPlanData(data)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
      ],
    })
    expect(() => hydratePlanData(data)).toThrow('Workbook plan data is invalid: Workbook plan data command at commands[0] is invalid')
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

  it('rejects accessor-backed transported low-level ops without invoking getters', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const planData = {
      modelName: 'transport-op-accessor-plan-model',
      actionName: 'op',
      refsUsed: [target],
      commands: [
        {
          kind: 'op',
          op: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 1,
          },
        },
      ],
      ops: [
        {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 1,
        },
      ],
      changed: [],
      checks: [],
    }

    const opArrayPlanData = structuredClone(planData)

    let commandOpGetterInvoked = false
    const [command] = mutableRecordArray(planData.commands)
    if (command === undefined) {
      throw new Error('expected command')
    }
    Object.defineProperty(mutableRecord(command['op']), 'extra', {
      enumerable: true,
      get() {
        commandOpGetterInvoked = true
        throw new Error('command op getter must not run')
      },
    })

    expect(isPlanData(planData)).toBe(false)
    expect(checkPlanData(planData)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'commands[0]',
          message: 'Workbook plan data command at commands[0] is invalid',
        },
      ],
    })
    expect(() => hydratePlanData(planData)).toThrow('Workbook plan data is invalid: Workbook plan data command at commands[0] is invalid')
    expect(commandOpGetterInvoked).toBe(false)

    let opGetterInvoked = false
    Object.defineProperty(mutableRecordArray(opArrayPlanData.ops)[0], 'extra', {
      enumerable: true,
      get() {
        opGetterInvoked = true
        throw new Error('op getter must not run')
      },
    })

    expect(isPlanData(opArrayPlanData)).toBe(false)
    expect(checkPlanData(opArrayPlanData)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'ops[0]',
          message: 'Workbook plan data op at ops[0] is invalid',
        },
      ],
    })
    expect(() => hydratePlanData(opArrayPlanData)).toThrow('Workbook plan data is invalid: Workbook plan data op at ops[0] is invalid')
    expect(opGetterInvoked).toBe(false)
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
