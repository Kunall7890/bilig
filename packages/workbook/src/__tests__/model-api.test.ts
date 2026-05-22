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
  inspectModel,
  isWorkbookPlanIssueCode,
  isWorkbookRef,
  planWorkbookAction,
  defineModel,
  formula,
  verifyModel,
  verifyPlan,
  workbookPlanIssueCodes,
} from '../index.js'

describe('@bilig/workbook model api', () => {
  it('exports stable inspectable plan issue codes', () => {
    expect(Object.isFrozen(workbookPlanIssueCodes)).toBe(true)
    expect(workbookPlanIssueCodes).toContain('invalid_action_input')
    expect(workbookPlanIssueCodes).toContain('formula_input_not_resolved')
    expect(workbookPlanIssueCodes).toContain('invalid_check_expectation_formula')
    expect(workbookPlanIssueCodes).toContain('missing_workbook_op')
    expect(new Set(workbookPlanIssueCodes).size).toBe(workbookPlanIssueCodes.length)
    expect(isWorkbookPlanIssueCode('check_ref_not_resolved')).toBe(true)
    expect(isWorkbookPlanIssueCode('custom_plan_issue')).toBe(false)
  })

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
        formula: '(__bilig_ref_table_p_Base_p_Rate_p_Result_p_Base)*(__bilig_ref_table_p_Base_p_Rate_p_Result_p_Rate)',
        inputs: [plan.refs.base, plan.refs.rate],
      },
    ])
    expect(plan.ops).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'C2',
        formula: '(__bilig_ref_table_p_Base_p_Rate_p_Result_p_Base)*(__bilig_ref_table_p_Base_p_Rate_p_Result_p_Rate)',
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

  it('freezes planned portable intent so agents can trust read-only snapshots', () => {
    const model = defineModel({
      name: 'frozen-plan-model',
      find(workbook) {
        return {
          input: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.valueEquals(refs.output, 2)]
      },
      actions: {
        calculate({ refs, workbook, input }) {
          workbook.writeFormula(refs.output, formula.add(refs.input, 1))
          workbook.writeValue(refs.input, input === undefined ? 1 : 2)
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate', { requestId: 'agent-1' })

    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      return
    }
    expect(Object.isFrozen(result.plan)).toBe(true)
    expect(Object.isFrozen(result.plan.input)).toBe(true)
    expect(Object.isFrozen(result.plan.refs)).toBe(true)
    expect(Object.isFrozen(result.plan.commands)).toBe(true)
    expect(Object.isFrozen(result.plan.commands[0])).toBe(true)
    const firstCommand = result.plan.commands[0]
    expect(firstCommand?.kind).toBe('writeFormula')
    if (firstCommand?.kind !== 'writeFormula') {
      return
    }
    expect(Object.isFrozen(firstCommand.inputs)).toBe(true)
    expect(Object.isFrozen(result.plan.ops)).toBe(true)
    expect(Object.isFrozen(result.plan.ops[0])).toBe(true)
    expect(Object.isFrozen(result.plan.changed)).toBe(true)
    expect(Object.isFrozen(result.plan.changed[0])).toBe(true)
    expect(Object.isFrozen(result.plan.checks)).toBe(true)
    expect(Object.isFrozen(result.plan.checks[0])).toBe(true)
    expect(Object.isFrozen(result.plan.refsUsed)).toBe(true)
  })

  it('freezes defined models so agents inspect a stable model contract', () => {
    const model = defineModel({
      name: 'frozen-model',
      description: 'Stable model contract',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: {
        write: {
          description: 'Write a value',
          input: {
            kind: 'object',
            fields: {
              value: { kind: 'number', required: true },
            },
          },
          run({ refs, workbook }) {
            workbook.writeValue(refs.output, 1)
          },
        },
      },
    })

    expect(Object.isFrozen(model)).toBe(true)
    expect(Object.isFrozen(model.actions)).toBe(true)
    expect(Object.isFrozen(model.actions.write)).toBe(true)
    if (typeof model.actions.write !== 'function') {
      expect(Object.isFrozen(model.actions.write.input)).toBe(true)
      expect(Object.isFrozen(model.actions.write.input?.fields?.value)).toBe(true)
    }
    expect(() => Object.defineProperty(model.actions, 'other', { value: () => undefined })).toThrowError(TypeError)
  })

  it('rejects malformed model configs from JavaScript callers with clear errors', () => {
    expect(() => Reflect.apply(defineModel, undefined, [null])).toThrowError('Workbook model config must be an object')
    expect(() =>
      Reflect.apply(defineModel, undefined, [
        {
          name: 7,
          find() {
            return {}
          },
          actions: {
            write() {
              return undefined
            },
          },
        },
      ]),
    ).toThrowError('Workbook model name must be a string')
    expect(() =>
      Reflect.apply(defineModel, undefined, [
        {
          name: 'bad-find',
          actions: {
            write() {
              return undefined
            },
          },
        },
      ]),
    ).toThrowError('Workbook model bad-find find must be a function')
    expect(() =>
      Reflect.apply(defineModel, undefined, [
        {
          name: 'bad-actions',
          find() {
            return {}
          },
          actions: [],
        },
      ]),
    ).toThrowError('Workbook model bad-actions actions must be an object')
  })

  it('rejects malformed high-level action commands from JavaScript callers', () => {
    const model = defineModel({
      name: 'runtime-guarded-actions',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      actions: {
        badValue({ refs, workbook }) {
          Reflect.apply(workbook.writeValue, workbook, [refs.output, Number.NaN])
        },
        badTarget({ workbook }) {
          Reflect.apply(workbook.clear, workbook, [{ kind: 'range' }])
        },
        badFormat({ refs, workbook }) {
          Reflect.apply(workbook.format, workbook, [refs.output, { numberFormat: 7 }])
        },
      },
    })

    expect(planWorkbookAction(model, 'badValue')).toEqual({
      status: 'failed',
      modelName: 'runtime-guarded-actions',
      actionName: 'badValue',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook write value must be a finite JSON literal',
        },
      ],
    })
    expect(planWorkbookAction(model, 'badTarget')).toEqual({
      status: 'failed',
      modelName: 'runtime-guarded-actions',
      actionName: 'badTarget',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook clear target must be a WorkbookRef',
        },
      ],
    })
    expect(planWorkbookAction(model, 'badFormat')).toEqual({
      status: 'failed',
      modelName: 'runtime-guarded-actions',
      actionName: 'badFormat',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook number format must be a string or null',
        },
      ],
    })
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
      id: 'table_p_Model_p_Inputs_p_Amount_p_Rate',
      label: 'Inputs',
      name: 'Inputs',
      sheetName: 'Model',
      headers: ['Amount', 'Rate'],
    })
    expect(table.column).toEqual(expect.any(Function))
    expect(amount).toEqual(amountViaTable)
    expect(result).toEqual({
      kind: 'range',
      id: 'range_p_Model_p_C2_p_C2',
      label: 'Model!C2',
      range: {
        sheetName: 'Model',
        startAddress: 'C2',
        endAddress: 'C2',
      },
    })
    expect(namedRate).toEqual({
      kind: 'name',
      id: 'name_p_Rate',
      label: 'Rate',
      name: 'Rate',
    })
    expect(rows).toEqual({
      kind: 'rows',
      id: 'table_p_Model_p_Inputs_p_Amount_p_Rate_p_Status_p_eq_p_string_p__x22_Active_x22_',
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
      id: 'table_p_Model_p_Inputs_p_Amount_p_Rate_p_Status_p_eq_p_string_p__x22_Active_x22__p_Amount',
      label: 'Inputs rows where Status eq "Active".Amount',
      table,
      rows,
      name: 'Amount',
    })
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
      id: 'table_p_Model_p_Inputs_p_Amount',
      label: 'Inputs',
      name: 'Inputs',
      sheetName: 'Model',
      headers: ['Amount'],
    })
    expect(() => Object.defineProperty(table, 'label', { value: 'Changed' })).toThrowError(TypeError)
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

    expect(activeRows.id).toBe('table_p_Inputs_p_Status_p_eq_p_string_p__x22_Active_x22_')
    expect(inactiveRows.id).toBe('table_p_Inputs_p_Status_p_eq_p_string_p__x22_Inactive_x22_')
    expect(activeRows.label).toBe('Inputs rows where Status eq "Active"')
    expect(inactiveRows.label).toBe('Inputs rows where Status eq "Inactive"')
    expect(collectWorkbookRefs({ activeRows, inactiveRows })).toEqual([activeRows, table, inactiveRows])
  })

  it('keeps punctuation-distinct selectors from collapsing to the same ref id', () => {
    const dashName = findName('A-B')
    const underscoreName = findName('A_B')
    const table = findTable({ headers: ['A-B', 'A_B'] })
    const dashColumn = table.column('A-B')
    const underscoreColumn = table.column('A_B')
    const dashRows = findRows({ table, where: { column: 'Kind', op: 'eq', value: 'A-B' } })
    const underscoreRows = findRows({ table, where: { column: 'Kind', op: 'eq', value: 'A_B' } })

    expect(dashName.id).not.toBe(underscoreName.id)
    expect(dashColumn.id).not.toBe(underscoreColumn.id)
    expect(dashRows.id).not.toBe(underscoreRows.id)
    expect(collectWorkbookRefs({ dashName, underscoreName, dashColumn, underscoreColumn, dashRows, underscoreRows })).toEqual([
      dashName,
      underscoreName,
      dashColumn,
      table,
      underscoreColumn,
      dashRows,
      underscoreRows,
    ])
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
    expect(formula.source(formula.ref(activeAmount))).toBe('__bilig_ref_table_p_Inputs_p_Status_p_eq_p_string_p__x22_Active_x22__p_Amount')
    expect(describeRef(activeAmount)).toEqual({
      kind: 'column',
      id: 'table_p_Inputs_p_Status_p_eq_p_string_p__x22_Active_x22__p_Amount',
      label: 'Inputs rows where Status eq "Active".Amount',
      table: {
        kind: 'table',
        id: 'table_p_Inputs',
        label: 'Inputs',
        name: 'Inputs',
      },
      rows: {
        kind: 'rows',
        id: 'table_p_Inputs_p_Status_p_eq_p_string_p__x22_Active_x22_',
        label: 'Inputs rows where Status eq "Active"',
        table: {
          kind: 'table',
          id: 'table_p_Inputs',
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
    const table = findTable({ name: ' Inputs ', sheetName: ' Model ', headers: [' Amount ', 'Rate'] })
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
      id: 'table_p_Model_p_Inputs_p_Amount_p_Rate',
      label: 'Inputs',
      name: 'Inputs',
      sheetName: 'Model',
      headers: ['Amount', 'Rate'],
    })
    expect(table.column).toEqual(expect.any(Function))
    expect(table.column(' Amount ')).toEqual({
      kind: 'column',
      id: 'table_p_Model_p_Inputs_p_Amount_p_Rate_p_Amount',
      label: 'Inputs.Amount',
      table,
      name: 'Amount',
    })
    expect(range).toEqual({
      kind: 'range',
      id: 'range_p_Sheet1_p_C2_p_C2',
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
    expect(() => Reflect.apply(findTable, undefined, [null])).toThrowError('Workbook table selector must be an object')
    expect(() => findName('   ')).toThrowError('Workbook selector name cannot be empty')
    expect(() => findTable({})).toThrowError('Workbook table selector needs a name, sheet name, or headers')
    expect(() => findTable({ headers: [] })).toThrowError('Workbook table headers cannot be empty')
    expect(() => findTable({ headers: ['Amount', ' '] })).toThrowError('Workbook selector table header cannot be empty')

    const table = findTable({ name: 'Inputs' })
    const otherTable = findTable({ name: 'OtherInputs' })
    const otherRows = findRows({ table: otherTable, where: { column: 'Status', op: 'eq', value: 'Active' } })
    expect(() => Reflect.apply(findColumn, undefined, [null])).toThrowError('Workbook column selector must be an object')
    expect(() => Reflect.apply(findColumn, undefined, [{ name: 'Amount' }])).toThrowError('Workbook column selector requires a table')
    expect(() => findColumn({ table, name: ' ' })).toThrowError('Workbook selector column name cannot be empty')
    expect(() => findColumn({ table, rows: otherRows, name: 'Amount' })).toThrowError(
      'Workbook column selector rows must belong to the table',
    )
    expect(() => Reflect.apply(findRange, undefined, [null])).toThrowError('Workbook range selector must be an object')
    expect(() => findRange({ sheetName: 'Sheet1', address: 'not-a-cell' })).toThrowError('Workbook range address is invalid: not-a-cell')
    expect(() => findRange({ sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'A1' })).toThrowError(
      'Workbook range endAddress must not be before startAddress',
    )
    expect(() => Reflect.apply(findRows, undefined, [null])).toThrowError('Workbook rows selector must be an object')
    expect(() =>
      Reflect.apply(findRows, undefined, [
        {
          where: {
            column: 'Status',
            op: 'eq',
            value: 'Active',
          },
        },
      ]),
    ).toThrowError('Workbook rows selector requires a table')
    expect(() => Reflect.apply(findRows, undefined, [{ table, where: null }])).toThrowError(
      'Workbook rows selector requires a where object',
    )
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

  it('does not treat inherited ref-looking properties as workbook refs', () => {
    const inherited = Object.create({
      kind: 'range',
      id: 'range_p_Polluted_p_A1_p_A1',
      label: 'Polluted!A1',
    }) as unknown

    expect(isWorkbookRef(inherited)).toBe(false)
    expect(collectWorkbookRefs({ inherited })).toEqual([])
  })

  it('does not treat incomplete JSON descriptions as live workbook refs', () => {
    const table = findTable({ name: 'Inputs' })
    const tableDescription = describeRef(table)
    const incompleteRange = {
      kind: 'range',
      id: 'range_p_Model_p_A1_p_A1',
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
      id: 'table_p_Inputs',
      label: 'Inputs',
      name: 'Inputs',
    } as const
    const amount = {
      kind: 'column',
      id: 'table_p_Inputs_p_Amount',
      label: 'Inputs.Amount',
      table,
      name: 'Amount',
    } as const
    const rate = {
      kind: 'column',
      id: 'table_p_Inputs_p_Rate',
      label: 'Inputs.Rate',
      table,
      name: 'Rate',
    } as const
    const result = {
      kind: 'range',
      id: 'range_p_Sheet1_p_D2_p_D2',
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
            id: 'range_p_Sheet1_p_B2_p_B2',
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
          id: 'range_p_Sheet1_p_D2_p_D2',
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
            id: 'table_p_Inputs',
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
