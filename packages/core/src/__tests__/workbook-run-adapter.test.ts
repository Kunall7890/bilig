import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { defineModel, findRange, findTable, formula, planWorkbookAction, runWorkbookAction, workbookPlanId } from '@bilig/workbook'
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

    const planned = planWorkbookAction(model, 'calculate')
    if (planned.status !== 'planned') {
      throw new Error(planned.errors.map((error) => error.message).join('\n'))
    }

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine), undefined, { strict: true })

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(result.apply).toMatchObject({
      matched: true,
      planId: workbookPlanId(planned.plan),
      commandReceipts: [
        {
          commandIndex: 0,
          commandKind: 'writeFormula',
          previewOps: [
            {
              kind: 'setCellFormula',
              sheetName: 'Sheet1',
              address: 'C1',
              formula: 'Sheet1!A1+Sheet1!B1',
            },
          ],
          appliedOps: [
            {
              kind: 'setCellFormula',
              sheetName: 'Sheet1',
              address: 'C1',
              formula: 'Sheet1!A1+Sheet1!B1',
            },
          ],
          resolvedRefs: {
            target: {
              kind: 'range',
              label: 'Sheet1!C1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'C1',
                endAddress: 'C1',
              },
            },
            inputs: [
              expect.objectContaining({
                kind: 'range',
                label: 'Sheet1!A1',
              }),
              expect.objectContaining({
                kind: 'range',
                label: 'Sheet1!B1',
              }),
            ],
          },
        },
      ],
      proof: {
        source: '@bilig/core',
        opCount: 1,
      },
    })
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('Sheet1!A1+Sheet1!B1')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(result.undo?.id).toMatch(/^generic-calculation\.calculate\.undo\.\d+$/)
    expect(result.undo?.ops?.length).toBeGreaterThan(0)
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['formulaEquals', 'passed'],
      ['valueEquals', 'passed'],
      ['noFormulaErrors', 'passed'],
    ])

    engine.applyOps(result.undo?.ops ?? [], { trusted: true })

    expect(engine.getCell('Sheet1', 'C1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Empty })
  })

  it('materializes formula labels by tokens and proves formula readback through labels', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-formula-labels' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)

    const model = defineModel({
      name: 'generic-token-formula',
      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          amountRate: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          const expression = formula.raw('amount_rate+amount', {
            labels: [
              { name: 'amount', ref: refs.amount },
              { name: 'amount_rate', ref: refs.amountRate },
            ],
          })
          workbook.writeFormula(refs.output, expression)
          workbook.check.formulaEquals(refs.output, expression)
          workbook.check.valueEquals(refs.output, 5)
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine), undefined, { strict: true })

    expect(result.status).toBe('done')
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('Sheet1!B1+Sheet1!A1')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(result.apply?.commandReceipts?.[0]).toMatchObject({
      commandKind: 'writeFormula',
      formulaLabels: [
        { name: 'amount', source: 'Sheet1!A1' },
        { name: 'amount_rate', source: 'Sheet1!B1' },
      ],
      previewOps: [
        {
          kind: 'setCellFormula',
          sheetName: 'Sheet1',
          address: 'C1',
          formula: 'Sheet1!B1+Sheet1!A1',
        },
      ],
    })
    expect(result.checks).toEqual([
      expect.objectContaining({
        status: 'passed',
        kind: 'formulaEquals',
        proof: {
          source: 'readback',
          formula: 'Sheet1!B1+Sheet1!A1',
          expectedFormula: 'amount_rate+amount',
          materializedFormula: 'Sheet1!B1+Sheet1!A1',
          formulaLabels: [
            { name: 'amount', source: 'Sheet1!A1' },
            { name: 'amount_rate', source: 'Sheet1!B1' },
          ],
        },
      }),
      expect.objectContaining({
        status: 'passed',
        kind: 'valueEquals',
      }),
    ])
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
      apply: {
        matched: true,
        proof: {
          source: '@bilig/core',
          opCount: 1,
        },
      },
      errors: [
        {
          code: 'check_failed',
          message: 'Sheet1!B1 failed check noFormulaErrors: Sheet1!B1 has no formula errors',
        },
      ],
      changed: [
        {
          kind: 'writeFormula',
          target: findRange({ sheetName: 'Sheet1', address: 'B1' }),
          message: 'Write formula to Sheet1!B1',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'noFormulaErrors',
          target: findRange({ sheetName: 'Sheet1', address: 'B1' }),
          message: 'Sheet1!B1 has no formula errors',
        },
      ],
      undo: {
        id: expect.stringMatching(/^formula-error-check\.calculate\.undo\.\d+$/),
      },
    })
    expect(result.undo?.ops?.length).toBeGreaterThan(0)
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
        return [workbook.check.exists(refs.result), workbook.check.noFormulaErrors(refs.result)]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('Sheet1!A2*Sheet1!B2')
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCell('Sheet1', 'C3').formula).toBe('Sheet1!A3*Sheet1!B3')
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['noFormulaErrors', 'passed'],
    ])
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
        return [workbook.check.exists(refs.actualRows), workbook.check.noFormulaErrors(refs.result)]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
        },
      },
    })

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('Sheet1!B2*Sheet1!C2')
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCell('Sheet1', 'D3').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCell('Sheet1', 'D4').formula).toBe('Sheet1!B4*Sheet1!C4')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['noFormulaErrors', 'passed'],
    ])
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
          code: 'apply_failed',
          message: expect.stringContaining('Ambiguous table selector table with Input, Result matched First, Second'),
        },
      ],
    })
  })
})
