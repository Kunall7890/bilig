import { describe, expect, it } from 'vitest'
import { parseFormula } from '@bilig/formula'
import {
  buildWorkbookActionPlan,
  check,
  collectWorkbookRefs,
  describeModel,
  describePlan,
  describePlanResult,
  describeRef,
  findColumn,
  find,
  findName,
  findRange,
  findRows,
  findTable,
  hydrateWorkbookRefs,
  inspectModel,
  isWorkbookRef,
  isWorkbookRefKind,
  isWorkbookRowOperator,
  isWorkbookRowValueCompatible,
  planWorkbookAction,
  defineModel,
  formula,
  verifyModel,
  verifyPlan,
  workbookRefKinds,
  workbookRowOperators,
  workbookRowOperatorValueTypes,
  type WorkbookAction,
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
    expect(Object.isFrozen(plan)).toBe(true)
    expect(Object.isFrozen(plan.refsUsed)).toBe(true)
    expect(Object.isFrozen(plan.commands)).toBe(true)
    expect(Object.isFrozen(plan.commands[0])).toBe(true)
    const [plannedCommand] = plan.commands
    if (plannedCommand?.kind !== 'writeFormula') {
      throw new Error('expected writeFormula command')
    }
    expect(Object.isFrozen(plannedCommand.inputs)).toBe(true)
    expect(Object.isFrozen(plannedCommand.labels)).toBe(true)
    expect(Object.isFrozen(plan.ops)).toBe(true)
    expect(Object.isFrozen(plan.ops[0])).toBe(true)
    expect(Object.isFrozen(plan.changed)).toBe(true)
    expect(Object.isFrozen(plan.changed[0])).toBe(true)
    expect(Object.isFrozen(plan.checks)).toBe(true)
    expect(Object.isFrozen(plan.checks[0])).toBe(true)
    expect(plan.refs.table.headers).toEqual(['Base', 'Rate', 'Result'])
    expect(plan.refsUsed).toEqual([plan.refs.table, plan.refs.base, plan.refs.rate, plan.refs.result])
    expect(plan.commands).toEqual([
      {
        kind: 'writeFormula',
        target: plan.refs.result,
        formula: '(__bilig_ref_table_Base_Rate_Result_Base)*(__bilig_ref_table_Base_Rate_Result_Rate)',
        inputs: [plan.refs.base, plan.refs.rate],
        labels: [
          { name: '__bilig_ref_table_Base_Rate_Result_Base', ref: plan.refs.base },
          { name: '__bilig_ref_table_Base_Rate_Result_Rate', ref: plan.refs.rate },
        ],
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

  it('keeps find, check, and action workbook phases scoped', () => {
    const seen: string[][] = []
    const model = defineModel({
      name: 'phase-scoped-model',

      find(workbook) {
        seen.push(Object.keys(workbook).toSorted())
        expect(Object.isFrozen(workbook)).toBe(true)
        expect('check' in workbook).toBe(false)
        expect('writeValue' in workbook).toBe(false)
        return {
          table: workbook.findTable({ name: 'Inputs' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      checks({ refs, workbook }) {
        seen.push(Object.keys(workbook).toSorted())
        expect(Object.isFrozen(workbook)).toBe(true)
        expect('writeValue' in workbook).toBe(false)
        expect('addOp' in workbook).toBe(false)
        const hidden = workbook.findRange({ sheetName: 'Sheet1', address: 'A1' })
        return [
          workbook.check.exists(refs.table),
          workbook.check.custom({ kind: 'hiddenSupport', refs: [hidden], message: 'Hidden support ref' }),
        ]
      },

      actions: {
        write({ refs, workbook }) {
          seen.push(Object.keys(workbook).toSorted())
          expect(Object.isFrozen(workbook)).toBe(true)
          workbook.writeValue(refs.output, 12)
          workbook.check.valueEquals(refs.output, 12)
        },
      },
    })

    const result = planWorkbookAction(model, 'write')

    expect(result.status).toBe('planned')
    expect(seen).toEqual([
      ['findColumn', 'findName', 'findRange', 'findRows', 'findTable'],
      ['check', 'findColumn', 'findName', 'findRange', 'findRows', 'findTable'],
      ['addOp', 'check', 'clear', 'findColumn', 'findName', 'findRange', 'findRows', 'findTable', 'format', 'writeFormula', 'writeValue'],
    ])
    if (result.status === 'planned') {
      expect(result.plan.commands).toEqual([
        {
          kind: 'writeValue',
          target: result.plan.refs.output,
          value: 12,
        },
      ])
      expect(verifyPlan(result.plan).issues.map((issue) => issue.code)).toEqual(['check_ref_not_resolved'])
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
    })
    expect(table.column).toEqual(expect.any(Function))
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
      id: 'table_Model_Inputs_Amount_Rate_Status_eq_string__22Active_22',
      label: 'Inputs rows where Status eq "Active"',
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Active',
      },
    })
    expect(rows.column).toEqual(expect.any(Function))
    expect(rows.column('Amount')).toEqual({
      kind: 'column',
      id: 'table_Model_Inputs_Amount_Rate_Status_eq_string__22Active_22_Amount',
      label: 'Inputs rows where Status eq "Active".Amount',
      table,
      rows,
      name: 'Amount',
    })
    expect(() =>
      findRows({
        sheetName: 'Model',
        where: {
          column: 'Status',
          op: 'eq',
          value: 'Active',
        },
      }).column('Amount'),
    ).toThrowError('Rows column selection requires a table-backed row selector')
  })

  it('exports a frozen find namespace and keeps ref helpers off enumerable data', () => {
    const table = find.table({ name: 'Inputs', sheetName: 'Model', headers: ['Amount'] })
    const rows = find.rows({ table, where: { column: 'Status', op: 'eq', value: 'Active' } })
    const column = rows.column('Amount')
    const range = find.range({ sheetName: 'Sheet1', address: 'A1' })

    expect(Object.isFrozen(find)).toBe(true)
    expect(find.findTable).toBe(findTable)
    expect(find.table).toBe(findTable)
    expect(Object.keys(table)).toEqual(['kind', 'id', 'label', 'name', 'sheetName', 'headers'])
    expect(Object.keys(rows)).toEqual(['kind', 'id', 'label', 'table', 'where'])
    expect(table.column).toEqual(expect.any(Function))
    expect(rows.column).toEqual(expect.any(Function))
    expect(Object.isFrozen(table)).toBe(true)
    expect(Object.isFrozen(table.headers)).toBe(true)
    expect(Object.isFrozen(rows)).toBe(true)
    expect(Object.isFrozen(rows.where)).toBe(true)
    expect(Object.isFrozen(column)).toBe(true)
    expect(Object.isFrozen(range)).toBe(true)
    expect(Object.isFrozen(range.range)).toBe(true)
    expect(JSON.parse(JSON.stringify(table))).toEqual({
      kind: 'table',
      id: 'table_Model_Inputs_Amount',
      label: 'Inputs',
      name: 'Inputs',
      sheetName: 'Model',
      headers: ['Amount'],
    })
    expect(() => Object.defineProperty(table, 'label', { value: 'Changed' })).toThrowError(TypeError)
  })

  it('exports frozen ref kind and row operator contracts for agent tools', () => {
    expect(workbookRefKinds).toEqual(['range', 'name', 'table', 'column', 'rows'])
    expect(workbookRowOperators).toEqual(['eq', 'neq', 'contains', 'startsWith', 'gt', 'gte', 'lt', 'lte'])
    expect(workbookRowOperatorValueTypes).toEqual({
      eq: ['number', 'string', 'boolean', 'null'],
      neq: ['number', 'string', 'boolean', 'null'],
      contains: ['string'],
      startsWith: ['string'],
      gt: ['number', 'string'],
      gte: ['number', 'string'],
      lt: ['number', 'string'],
      lte: ['number', 'string'],
    })
    expect(Object.isFrozen(workbookRefKinds)).toBe(true)
    expect(Object.isFrozen(workbookRowOperators)).toBe(true)
    expect(Object.isFrozen(workbookRowOperatorValueTypes)).toBe(true)
    expect(Object.isFrozen(workbookRowOperatorValueTypes.eq)).toBe(true)

    expect(isWorkbookRefKind('table')).toBe(true)
    expect(isWorkbookRefKind('chart')).toBe(false)
    expect(isWorkbookRowOperator('gte')).toBe(true)
    expect(isWorkbookRowOperator('between')).toBe(false)
    expect(isWorkbookRowValueCompatible('contains', 'Active')).toBe(true)
    expect(isWorkbookRowValueCompatible('contains', 12)).toBe(false)
    expect(isWorkbookRowValueCompatible('gt', 12)).toBe(true)
    expect(isWorkbookRowValueCompatible('gt', 'M')).toBe(true)
    expect(isWorkbookRowValueCompatible('gt', true)).toBe(false)
    expect(isWorkbookRowValueCompatible('eq', null)).toBe(true)
  })

  it('keeps row selector refs distinct by predicate value', () => {
    const table = findTable({ name: 'Inputs' })
    const activeRows = findRows({
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Active',
      },
    })
    const inactiveRows = findRows({
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Inactive',
      },
    })

    expect(activeRows.id).toBe('table_Inputs_Status_eq_string__22Active_22')
    expect(inactiveRows.id).toBe('table_Inputs_Status_eq_string__22Inactive_22')
    expect(activeRows.label).toBe('Inputs rows where Status eq "Active"')
    expect(inactiveRows.label).toBe('Inputs rows where Status eq "Inactive"')
    expect(collectWorkbookRefs({ activeRows, inactiveRows })).toEqual([activeRows, table, inactiveRows])
  })

  it('describes row-filtered columns as JSON-safe dependent refs', () => {
    const table = findTable({ name: 'Inputs' })
    const activeRows = findRows({
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'Active',
      },
    })
    const activeAmount = activeRows.column('Amount')

    expect(collectWorkbookRefs({ activeAmount })).toEqual([activeAmount, activeRows, table])
    expect(formula.source(formula.ref(activeAmount))).toBe('__bilig_ref_table_Inputs_Status_eq_string__22Active_22_Amount')
    expect(describeRef(activeAmount)).toEqual({
      kind: 'column',
      id: 'table_Inputs_Status_eq_string__22Active_22_Amount',
      label: 'Inputs rows where Status eq "Active".Amount',
      table: {
        kind: 'table',
        id: 'table_Inputs',
        label: 'Inputs',
        name: 'Inputs',
      },
      rows: {
        kind: 'rows',
        id: 'table_Inputs_Status_eq_string__22Active_22',
        label: 'Inputs rows where Status eq "Active"',
        table: {
          kind: 'table',
          id: 'table_Inputs',
          label: 'Inputs',
          name: 'Inputs',
        },
        where: {
          column: 'Status',
          op: 'eq',
          value: 'Active',
        },
      },
      name: 'Amount',
    })
  })

  it('normalizes public find selectors before planning', () => {
    const table = findTable({ name: ' Inputs ', sheetName: ' Model ', headers: [' Rate ', 'Amount'] })
    const range = findRange({ sheetName: ' Sheet1 ', address: ' c2 ' })
    const rows = findRows({
      table,
      where: {
        column: ' Status ',
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
    })
    expect(table.column).toEqual(expect.any(Function))
    expect(table.column(' Amount ')).toEqual({
      kind: 'column',
      id: 'table_Model_Inputs_Amount_Rate_Amount',
      label: 'Inputs.Amount',
      table,
      name: 'Amount',
    })
    expect(range).toEqual({
      kind: 'range',
      id: 'range_Sheet1_C2_C2',
      label: 'Sheet1!C2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C2',
        endAddress: 'C2',
      },
    })
    expect(rows.where.column).toBe('Status')
    expect(rows.label).toBe('Inputs rows where Status eq "Active"')
  })

  it('rejects malformed public find selectors before planning', () => {
    expect(() => findName('   ')).toThrowError('Workbook selector name cannot be empty')
    expect(() => findTable({})).toThrowError('Workbook table selector needs a name, sheet name, or headers')
    expect(() => findTable({ headers: [] })).toThrowError('Workbook table headers cannot be empty')
    expect(() => findTable({ headers: ['Amount', ' '] })).toThrowError('Workbook selector table header cannot be empty')
    expect(() => findTable({ headers: [' Amount ', 'Amount'] })).toThrowError('Workbook table headers cannot contain duplicates: Amount')

    const table = findTable({ name: 'Inputs' })
    expect(() => findColumn({ table, name: ' ' })).toThrowError('Workbook selector column name cannot be empty')
    expect(() => findRange({ sheetName: 'Sheet1', address: 'not-a-cell' })).toThrowError('Workbook range address is invalid: not-a-cell')
    expect(() => findRange({ sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'A1' })).toThrowError(
      'Workbook range endAddress must not be before startAddress',
    )
    expect(() =>
      findRows({
        where: {
          column: 'Status',
          op: 'eq',
          value: 'Active',
        },
      }),
    ).toThrowError('Workbook rows selector requires a table or sheet name')
    expect(() =>
      Reflect.apply(findRows, undefined, [
        {
          table,
          where: {
            column: 'Status',
            op: 'bad',
            value: 'Active',
          },
        },
      ]),
    ).toThrowError('Unsupported workbook row operator: bad')
    expect(() =>
      Reflect.apply(findRows, undefined, [
        {
          table,
          where: {
            column: 'Status',
            op: 'eq',
            value: Number.NaN,
          },
        },
      ]),
    ).toThrowError('Workbook rows selector value must be a finite JSON literal')
    expect(() =>
      findRows({
        table,
        where: {
          column: 'Status',
          op: 'contains',
          value: 12,
        },
      }),
    ).toThrowError('Workbook rows selector operator contains requires a string value')
    expect(() =>
      findRows({
        table,
        where: {
          column: 'Status',
          op: 'gt',
          value: true,
        },
      }),
    ).toThrowError('Workbook rows selector operator gt requires a number or string value')
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
      labels: [
        { name: 'Inputs[Amount]', ref: plan.refs.amount },
        { name: 'Inputs[Rate]', ref: plan.refs.rate },
      ],
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

  it('does not treat inherited ref-looking properties as workbook refs', () => {
    const inherited = Object.create({
      kind: 'range',
      id: 'range_Polluted_A1_A1',
      label: 'Polluted!A1',
    }) as unknown

    expect(isWorkbookRef(inherited)).toBe(false)
    expect(collectWorkbookRefs({ inherited })).toEqual([])
  })

  it('hydrates refs inside generic containers without losing magic JSON keys', () => {
    const range = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const rangeData = describeRef(range)
    const input = JSON.parse(
      `{"__proto__":${JSON.stringify(rangeData)},"constructor":{"nested":${JSON.stringify(rangeData)}},"label":"agent-owned"}`,
    ) as unknown

    const hydrated = hydrateWorkbookRefs(input)

    expect(Object.getPrototypeOf(hydrated)).toBe(Object.prototype)
    expect(Object.hasOwn(hydrated ?? {}, '__proto__')).toBe(true)
    expect(Object.hasOwn(hydrated ?? {}, 'constructor')).toBe(true)
    if (typeof hydrated !== 'object' || hydrated === null || Array.isArray(hydrated)) {
      throw new Error('expected hydrated object')
    }
    const protoValue = Object.getOwnPropertyDescriptor(hydrated, '__proto__')?.value
    const constructorValue = Object.getOwnPropertyDescriptor(hydrated, 'constructor')?.value
    if (typeof constructorValue !== 'object' || constructorValue === null || Array.isArray(constructorValue)) {
      throw new Error('expected constructor payload object')
    }
    expect(isWorkbookRef(protoValue)).toBe(true)
    expect(isWorkbookRef(Object.getOwnPropertyDescriptor(constructorValue, 'nested')?.value)).toBe(true)
    expect(JSON.parse(JSON.stringify(hydrated))).toEqual(
      JSON.parse(`{"__proto__":${JSON.stringify(rangeData)},"constructor":{"nested":${JSON.stringify(rangeData)}},"label":"agent-owned"}`),
    )
  })

  it('does not treat incomplete JSON descriptions as live workbook refs', () => {
    const table = findTable({ name: 'Inputs' })
    const tableDescription = describeRef(table)
    const incompleteRange = {
      kind: 'range',
      id: 'range_Model_A1_A1',
      label: 'Model!A1',
    }
    const malformedRows = {
      kind: 'rows',
      id: 'rows_bad',
      label: 'bad rows',
      where: {
        column: 'Status',
        op: 'eq',
        value: Number.NaN,
      },
      column() {
        return table.column('Status')
      },
    }

    expect(isWorkbookRef(table)).toBe(true)
    expect(isWorkbookRef(tableDescription)).toBe(false)
    expect(isWorkbookRef(incompleteRange)).toBe(false)
    expect(isWorkbookRef(malformedRows)).toBe(false)
    expect(collectWorkbookRefs({ tableDescription, incompleteRange, malformedRows })).toEqual([])
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
          labels: [
            { name: 'Inputs[Amount]', ref: amount },
            { name: 'Inputs[Rate]', ref: rate },
          ],
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

  it('rejects ambiguous model and action names before planning', () => {
    expect(() =>
      defineModel({
        name: ' ambiguous-model ',
        find() {
          return {}
        },
        actions: {
          calculate() {},
        },
      }),
    ).toThrowError('Workbook model name must not have leading or trailing whitespace')

    expect(() =>
      defineModel({
        name: 'ambiguous-action-model',
        find() {
          return {}
        },
        actions: {
          ' calculate '() {},
        },
      }),
    ).toThrowError('Workbook model ambiguous-action-model action name must not have leading or trailing whitespace')

    expect(() =>
      defineModel({
        name: 'empty-action-model',
        find() {
          return {}
        },
        actions: {
          '   '() {},
        },
      }),
    ).toThrowError('Workbook model empty-action-model action name cannot be empty')
  })

  it('keeps model actions generic without trusting action-map prototypes', () => {
    type PrototypeRefs = { readonly target: ReturnType<typeof findRange> }
    const inheritedActions: Record<string, WorkbookAction<PrototypeRefs>> = {}
    Object.setPrototypeOf(inheritedActions, {
      inherited({ refs, workbook }: Parameters<WorkbookAction<PrototypeRefs>>[0]) {
        workbook.writeValue(refs.target, 404)
      },
    })
    inheritedActions.calculate = ({ refs, workbook }) => {
      workbook.writeValue(refs.target, 12)
    }

    const model = defineModel({
      name: 'own-action-model',
      find(workbook) {
        return {
          target: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: inheritedActions,
    })

    expect(Object.getPrototypeOf(model.actions)).toBe(null)
    expect(inspectModel(model).actions).toEqual(['calculate'])
    expect(planWorkbookAction(model, 'inherited')).toEqual({
      status: 'failed',
      modelName: 'own-action-model',
      actionName: 'inherited',
      checks: [],
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model own-action-model does not define action inherited',
        },
      ],
    })

    expect(buildWorkbookActionPlan(model, 'calculate').commands).toEqual([
      {
        kind: 'writeValue',
        target: expect.objectContaining({ label: 'Sheet1!A1' }),
        value: 12,
      },
    ])
  })

  it('allows own action names that look like object prototype fields', () => {
    type PrototypeNameRefs = { readonly target: ReturnType<typeof findRange> }
    const actions: Record<string, WorkbookAction<PrototypeNameRefs>> = {}
    Object.setPrototypeOf(actions, null)
    actions.toString = ({ refs, workbook }) => {
      workbook.writeValue(refs.target, 1)
    }
    actions.constructor = ({ refs, workbook }) => {
      workbook.writeValue(refs.target, 2)
    }
    Object.defineProperty(actions, '__proto__', {
      enumerable: true,
      value({ refs, workbook }: Parameters<WorkbookAction<PrototypeNameRefs>>[0]) {
        workbook.writeValue(refs.target, 3)
      },
    })

    const model = defineModel({
      name: 'prototype-name-model',
      find(workbook) {
        return {
          target: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      actions,
    })

    expect(Object.getPrototypeOf(model.actions)).toBe(null)
    expect(inspectModel(model).actions).toEqual(['__proto__', 'constructor', 'toString'])
    expect(buildWorkbookActionPlan(model, 'toString').commands).toEqual([
      {
        kind: 'writeValue',
        target: expect.objectContaining({ label: 'Sheet1!B2' }),
        value: 1,
      },
    ])
    expect(buildWorkbookActionPlan(model, 'constructor').commands).toEqual([
      {
        kind: 'writeValue',
        target: expect.objectContaining({ label: 'Sheet1!B2' }),
        value: 2,
      },
    ])
    expect(buildWorkbookActionPlan(model, '__proto__').commands).toEqual([
      {
        kind: 'writeValue',
        target: expect.objectContaining({ label: 'Sheet1!B2' }),
        value: 3,
      },
    ])
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
      actionDetails: [{ name: 'calculate' }, { name: 'reset' }],
      hasChecks: true,
    })
  })

  it('describes models as JSON-safe manifests without running model code', () => {
    const model = defineModel({
      name: 'described-model-manifest',
      find() {
        throw new Error('find should not run during model description')
      },
      checks() {
        throw new Error('checks should not run during model description')
      },
      actions: {
        calculate() {
          throw new Error('action should not run during model description')
        },
        reset() {
          throw new Error('action should not run during model description')
        },
      },
    })

    const description = describeModel(model)

    expect(description).toEqual({
      name: 'described-model-manifest',
      actions: ['calculate', 'reset'],
      actionDetails: [{ name: 'calculate' }, { name: 'reset' }],
      hasChecks: true,
    })
    expect(JSON.parse(JSON.stringify(description))).toEqual(description)
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
