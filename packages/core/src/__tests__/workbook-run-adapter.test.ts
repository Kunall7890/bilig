import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { buildWorkbookActionPlan, defineModel, findRange, findTable, formula, runWorkbookAction, runWorkbookPlan } from '@bilig/workbook'
import { SpreadsheetEngine } from '../engine.js'
import { createWorkbookRunAdapter } from '../workbook-run-adapter.js'

describe('workbook run adapter', () => {
  it('applies agent workbook plans and verifies readback checks through the engine', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-readback' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)

    const model = defineModel({
      name: 'generic-calculation',
      find(workbook) {
        return {
          left: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          right: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      checks({ refs, workbook }) {
        return [
          workbook.check.exists(refs.output),
          workbook.check.formulaEquals(refs.output, formula.add(refs.left, refs.right)),
          workbook.check.valueEquals(refs.output, 5),
          workbook.check.noFormulaErrors(refs.output),
        ]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.add(refs.left, refs.right))
        },
      },
    })

    const adapter = createWorkbookRunAdapter(engine)
    const plan = buildWorkbookActionPlan(model, 'calculate')
    const preview = await adapter.preview?.(plan)

    expect(preview?.materializedOps).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'C1',
        formula: '(Sheet1!A1)+(Sheet1!B1)',
      },
    ])

    const result = await runWorkbookPlan(plan, adapter)

    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(result).toMatchObject({ status: 'done' })
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('(Sheet1!A1)+(Sheet1!B1)')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(result.undo?.id).toMatch(/^generic-calculation\.calculate\.undo\.\d+$/)
    expect(result.undo?.ops?.length).toBeGreaterThan(0)
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['formulaEquals', 'passed'],
      ['valueEquals', 'passed'],
      ['noFormulaErrors', 'passed'],
    ])
    expect(result.checks).toEqual([
      expect.objectContaining({
        kind: 'exists',
        proof: {
          kind: 'runtime',
          message: 'Runtime confirmed the reference exists',
          data: { exists: true, target: 'Sheet1!C1' },
        },
      }),
      expect.any(Object),
      expect.any(Object),
      expect.objectContaining({
        kind: 'noFormulaErrors',
        proof: {
          kind: 'runtime',
          message: 'Runtime confirmed no formula errors',
          data: { passed: true, target: 'Sheet1!C1' },
        },
      }),
    ])
    expect(result.applied).toEqual({
      opCount: 1,
      ops: preview?.materializedOps,
    })

    engine.applyOps(result.undo?.ops ?? [], { trusted: true })

    expect(engine.getCell('Sheet1', 'C1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Empty })
  })

  it('fails generic noFormulaErrors checks when the engine calculates an error cell', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-formula-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 0)

    const model = defineModel({
      name: 'formula-error-check',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.noFormulaErrors(refs.output)]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.raw('1/A1'))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_failed',
          message: 'Sheet1!B1 failed check noFormulaErrors: Sheet1!B1 has no formula errors',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'noFormulaErrors',
          target: findRange({ sheetName: 'Sheet1', address: 'B1' }),
          message: 'Sheet1!B1 has no formula errors',
          proof: {
            kind: 'runtime',
            message: 'Runtime found formula errors or could not resolve formulas',
            data: { passed: false, target: 'Sheet1!B1' },
          },
        },
      ],
      undo: {
        id: expect.stringMatching(/^formula-error-check\.calculate\.undo\.\d+$/),
        ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'B1' }],
      },
    })
    expect(engine.getCellValue('Sheet1', 'B1').tag).toBe(ValueTag.Error)
  })

  it('does not treat engine error cells as null readback values', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-error-readback' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 0)

    const model = defineModel({
      name: 'formula-error-readback',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.valueEquals(refs.output, null), workbook.check.valuesEqual(refs.output, [[null]])]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.raw('1/A1'))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result.status).toBe('failed')
    if (result.status !== 'failed') {
      return
    }
    expect(result.errors.map((error) => [error.code, error.message])).toEqual([
      ['value_mismatch', 'Sheet1!B1 has no value readback'],
      ['values_mismatch', 'Sheet1!B1 has no values readback'],
    ])
    expect(engine.getCellValue('Sheet1', 'B1').tag).toBe(ValueTag.Error)
  })

  it('materializes table column formula commands into row-wise engine formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-table-column-formulas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Base')
    engine.setCellValue('Sheet1', 'B1', 'Rate')
    engine.setCellValue('Sheet1', 'C1', 'Result')
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'B2', 3)
    engine.setCellValue('Sheet1', 'A3', 4)
    engine.setCellValue('Sheet1', 'B3', 5)
    engine.setTable({
      name: 'CalcInputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Base', 'Rate', 'Result'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-table-calculation',
      find(workbook) {
        const table = workbook.findTable({ headers: ['Base', 'Rate', 'Result'] })
        return {
          table,
          base: table.column('Base'),
          rate: table.column('Rate'),
          result: table.column('Result'),
        }
      },
      checks({ refs, workbook }) {
        return [
          workbook.check.exists(refs.result),
          workbook.check.noFormulaErrors(refs.result),
          workbook.check.valuesEqual(refs.result, [[6], [20]]),
          workbook.check.formulasEqual(refs.result, [['(Sheet1!A2)*(Sheet1!B2)'], ['(Sheet1!A3)*(Sheet1!B3)']]),
        ]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(result).toMatchObject({ status: 'done' })
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('(Sheet1!A2)*(Sheet1!B2)')
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCell('Sheet1', 'C3').formula).toBe('(Sheet1!A3)*(Sheet1!B3)')
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['noFormulaErrors', 'passed'],
      ['valuesEqual', 'passed'],
      ['formulasEqual', 'passed'],
    ])
    expect(result.applied).toEqual({
      opCount: 2,
      ops: [
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C2', formula: '(Sheet1!A2)*(Sheet1!B2)' },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C3', formula: '(Sheet1!A3)*(Sheet1!B3)' },
      ],
    })
  })

  it('materializes overlapping placeholder formula inputs without corrupting replacements', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-overlapping-placeholders' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Key')
    engine.setCellValue('Sheet1', 'B1', 'A')
    engine.setCellValue('Sheet1', 'C1', 'A_B')
    engine.setCellValue('Sheet1', 'D1', 'Result')
    engine.setCellValue('Sheet1', 'A2', 'row')
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D2',
      columnNames: ['Key', 'A', 'A_B', 'Result'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-overlapping-placeholders',
      find(workbook) {
        const table = workbook.findTable({ headers: ['Key', 'A', 'A_B', 'Result'] })
        return {
          table,
          a: table.column('A'),
          aB: table.column('A_B'),
          result: table.column('Result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.valueEquals(refs.result, 5), workbook.check.formulasEqual(refs.result, [['(Sheet1!B2)+(Sheet1!C2)']])]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.add(refs.a, refs.aB))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(result).toMatchObject({ status: 'done' })
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('(Sheet1!B2)+(Sheet1!C2)')
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('materializes formula inputs only at formula token boundaries', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-token-boundaries' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setDefinedName('Input', { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' })

    const model = defineModel({
      name: 'generic-token-boundaries',
      find(workbook) {
        return {
          input: workbook.findName('Input'),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.raw('InputValue+Input', { inputs: [refs.input] }))
        },
      },
    })

    const preview = await createWorkbookRunAdapter(engine).preview?.(buildWorkbookActionPlan(model, 'calculate'))

    expect(preview?.materializedOps).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'B1',
        formula: 'InputValue+Sheet1!A1',
      },
    ])
  })

  it('does not materialize declared formula inputs inside string literals', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-formula-strings' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Key')
    engine.setCellValue('Sheet1', 'B1', 'Amount')
    engine.setCellValue('Sheet1', 'C1', 'Result')
    engine.setCellValue('Sheet1', 'A2', 'row')
    engine.setCellValue('Sheet1', 'B2', 7)
    engine.setTable({
      name: 'Input',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C2',
      columnNames: ['Key', 'Amount', 'Result'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-formula-string-literals',
      find(workbook) {
        const table = workbook.findTable({ name: 'Input' })
        return {
          amount: table.column('Amount'),
          result: table.column('Result'),
        }
      },
      checks({ refs, workbook }) {
        return [
          workbook.check.noFormulaErrors(refs.result),
          workbook.check.formulasEqual(refs.result, [['T("Input[Amount]")&"-"&Sheet1!B2']]),
        ]
      },
      actions: {
        calculate({ refs, workbook }) {
          const token = formula.source(formula.ref(refs.amount))
          workbook.writeFormula(refs.result, formula.raw(`T("${token}")&"-"&${token}`, { inputs: [refs.amount] }))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('T("Input[Amount]")&"-"&Sheet1!B2')
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['noFormulaErrors', 'passed'],
      ['formulasEqual', 'passed'],
    ])
  })

  it('applies additional plan ops after command materialization instead of silently ignoring them', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-extra-ops' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const model = defineModel({
      name: 'generic-extra-op',
      find(workbook) {
        return {
          first: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.first, 1)
        },
      },
    })
    const plan = buildWorkbookActionPlan(model, 'write')
    const adapter = createWorkbookRunAdapter(engine)

    const result = await runWorkbookPlan(
      {
        ...plan,
        ops: [...plan.ops, { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 2 }],
      },
      adapter,
    )

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(result.applied).toEqual({
      opCount: 2,
      ops: [
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 2 },
      ],
    })
  })

  it('verifies exists checks for workbook tables without hardcoded model assumptions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-table-exists' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Base', 'Rate'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-table-check',
      find(workbook) {
        return {
          table: workbook.findTable({ headers: ['Base', 'Rate'] }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },
      actions: {
        inspect() {},
      },
    })

    const result = await runWorkbookAction(model, 'inspect', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(result.checks).toHaveLength(1)
    expect(result.checks[0]).toMatchObject({
      status: 'passed',
      kind: 'exists',
      target: {
        kind: 'table',
        id: findTable({ headers: ['Base', 'Rate'] }).id,
        label: 'table with Base, Rate',
        headers: ['Base', 'Rate'],
      },
      message: 'table with Base, Rate exists',
    })
  })

  it('applies high-level style and number format commands through engine metadata ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-format' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12.5)
    engine.setCellValue('Sheet1', 'B1', 25)

    const model = defineModel({
      name: 'generic-formatting',
      find(workbook) {
        return {
          range: workbook.findRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' }),
        }
      },
      actions: {
        format({ refs, workbook }) {
          workbook.format(refs.range, {
            style: { fill: { backgroundColor: '#dbeafe' }, font: { bold: true } },
            numberFormat: '0.00',
          })
        },
      },
    })

    const result = await runWorkbookAction(model, 'format', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    const firstStyle = engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)
    const secondStyle = engine.getCellStyle(engine.getCell('Sheet1', 'B1').styleId)
    expect(firstStyle).toMatchObject({ fill: { backgroundColor: '#dbeafe' }, font: { bold: true } })
    expect(secondStyle?.id).toBe(firstStyle?.id)
    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')
    expect(engine.getCell('Sheet1', 'B1').format).toBe('0.00')
  })

  it('materializes row-filtered table formulas only for matching rows', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-filtered-row-formulas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Kind')
    engine.setCellValue('Sheet1', 'B1', 'Base')
    engine.setCellValue('Sheet1', 'C1', 'Rate')
    engine.setCellValue('Sheet1', 'D1', 'Result')
    engine.setCellValue('Sheet1', 'A2', 'actual')
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellValue('Sheet1', 'A3', 'budget')
    engine.setCellValue('Sheet1', 'B3', 100)
    engine.setCellValue('Sheet1', 'C3', 100)
    engine.setCellValue('Sheet1', 'A4', 'actual')
    engine.setCellValue('Sheet1', 'B4', 4)
    engine.setCellValue('Sheet1', 'C4', 5)
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D4',
      columnNames: ['Kind', 'Base', 'Rate', 'Result'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-row-calculation',
      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        const actualRows = workbook.findRows({ table, where: { column: 'Kind', op: 'eq', value: 'actual' } })
        return {
          actualRows,
          base: actualRows.column('Base'),
          rate: actualRows.column('Rate'),
          result: actualRows.column('Result'),
        }
      },
      checks({ refs, workbook }) {
        return [
          workbook.check.exists(refs.actualRows),
          workbook.check.noFormulaErrors(refs.result),
          workbook.check.valuesEqual(refs.result, [[6], [20]]),
        ]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
        },
      },
    })

    const adapter = createWorkbookRunAdapter(engine)
    const plan = buildWorkbookActionPlan(model, 'calculate')
    const preview = await adapter.preview?.(plan)

    expect(preview?.materializedOps).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'D2',
        formula: '(Sheet1!B2)*(Sheet1!C2)',
      },
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'D4',
        formula: '(Sheet1!B4)*(Sheet1!C4)',
      },
    ])

    const result = await runWorkbookPlan(plan, adapter)

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('(Sheet1!B2)*(Sheet1!C2)')
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCell('Sheet1', 'D3').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCell('Sheet1', 'D4').formula).toBe('(Sheet1!B4)*(Sheet1!C4)')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['noFormulaErrors', 'passed'],
      ['valuesEqual', 'passed'],
    ])
    expect(result.applied).toEqual({
      opCount: 2,
      ops: preview?.materializedOps,
    })
  })

  it('fails exists checks for row selectors that match no rows', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-empty-row-selector' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Kind')
    engine.setCellValue('Sheet1', 'B1', 'Amount')
    engine.setCellValue('Sheet1', 'A2', 'budget')
    engine.setCellValue('Sheet1', 'B2', 10)
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Kind', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-empty-row-check',
      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          rows: workbook.findRows({ table, where: { column: 'Kind', op: 'eq', value: 'actual' } }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.rows)]
      },
      actions: {
        inspect() {},
      },
    })

    const result = await runWorkbookAction(model, 'inspect', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({
      status: 'failed',
      errors: [{ code: 'check_failed' }],
      checks: [expect.objectContaining({ status: 'failed', kind: 'exists' })],
    })
  })

  it('rejects ambiguous table selectors instead of silently choosing a table', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-ambiguous-tables' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setTable({
      name: 'First',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Input', 'Result'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setTable({
      name: 'Second',
      sheetName: 'Sheet1',
      startAddress: 'D1',
      endAddress: 'E2',
      columnNames: ['Input', 'Result'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-ambiguous-table',
      find(workbook) {
        const table = workbook.findTable({ headers: ['Input', 'Result'] })
        return {
          input: table.column('Input'),
          result: table.column('Result'),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.ref(refs.input))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: expect.stringContaining('Ambiguous table selector table with Input, Result matched First, Second'),
        },
      ],
    })
  })
})
