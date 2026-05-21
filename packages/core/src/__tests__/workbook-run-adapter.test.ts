import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { defineModel, findRange, findTable, formula, runWorkbookAction } from '@bilig/workbook'
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

    const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
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

    expect(result).toEqual({
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
        },
      ],
    })
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
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('(Sheet1!A2)*(Sheet1!B2)')
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCell('Sheet1', 'C3').formula).toBe('(Sheet1!A3)*(Sheet1!B3)')
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

  it('verifies named cell and single table column readbacks through the adapter', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-named-readbacks' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'B1', 'Amount')
    engine.setCellValue('Sheet1', 'B2', 7)
    engine.setDefinedName('PinnedValue', { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' })
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'B1',
      endAddress: 'B2',
      columnNames: ['Amount'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-readback-check',
      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          pinned: workbook.findName('PinnedValue'),
          amount: table.column('Amount'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.pinned), workbook.check.valueEquals(refs.pinned, 42), workbook.check.valueEquals(refs.amount, 7)]
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
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['valueEquals', 'passed'],
      ['valueEquals', 'passed'],
    ])
  })

  it('applies write value, clear, format clear, and low-level op commands', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-command-kinds' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'stale')
    engine.setCellValue('Sheet1', 'B1', 12.5)
    const formattedRange = { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' } as const
    const initialFormat = engine.workbook.internCellNumberFormat('0.00')
    engine.applyOps(
      [
        { kind: 'upsertCellNumberFormat', format: initialFormat },
        { kind: 'setFormatRange', range: formattedRange, formatId: initialFormat.id },
      ],
      { trusted: true },
    )
    expect(engine.getCell('Sheet1', 'B1').format).toBe('0.00')

    const model = defineModel({
      name: 'generic-command-kinds',
      find(workbook) {
        return {
          stale: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          formatted: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'D1' }),
          opOutput: workbook.findRange({ sheetName: 'Sheet1', address: 'E1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.output), workbook.check.valueEquals(refs.opOutput, 99)]
      },
      actions: {
        update({ refs, workbook }) {
          workbook.writeValue(refs.output, 'filled')
          workbook.clear(refs.stale)
          workbook.format(refs.formatted, { numberFormat: null })
          workbook.addOp(
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'E1',
              value: 99,
            },
            {
              target: refs.opOutput,
              message: 'Seed op output value',
            },
          )
        },
      },
    })

    const result = await runWorkbookAction(model, 'update', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({ status: 'done' })
    if (result.status !== 'done') {
      throw new Error(result.errors.map((error) => error.message).join('\n'))
    }
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.String, value: 'filled', stringId: expect.any(Number) })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.String, value: 'filled', stringId: expect.any(Number) })
    expect(engine.getCell('Sheet1', 'B1').format).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(result.checks.map((check) => [check.kind, check.status])).toEqual([
      ['exists', 'passed'],
      ['valueEquals', 'passed'],
    ])
  })

  it('handles no-op commands, unresolved targets, and terminal checks predictably', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-edge-cases' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const noOpModel = defineModel({
      name: 'generic-noop-format',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        touch({ refs, workbook }) {
          workbook.format(refs.output, {})
        },
      },
    })

    const noOpResult = await runWorkbookAction(
      noOpModel,
      'touch',
      createWorkbookRunAdapter(engine, { captureUndo: false, potentialNewCells: 1 }),
    )

    expect(noOpResult).toMatchObject({ status: 'done' })
    if (noOpResult.status !== 'done') {
      throw new Error(noOpResult.errors.map((error) => error.message).join('\n'))
    }
    expect(noOpResult.undo).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Empty })

    const missingTargetModel = defineModel({
      name: 'generic-missing-target',
      find(workbook) {
        return {
          missing: workbook.findName('MissingName'),
        }
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.missing, 1)
        },
      },
    })

    const missingResult = await runWorkbookAction(missingTargetModel, 'write', createWorkbookRunAdapter(engine))

    expect(missingResult).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'engine_apply_failed',
          message: 'Cannot resolve MissingName for writeValue',
        },
      ],
    })

    const terminalCheck = {
      status: 'failed',
      kind: 'exists',
      message: 'already failed before adapter verification',
    } as const
    const adapter = createWorkbookRunAdapter(engine)
    expect(
      adapter.verifyChecks?.([terminalCheck], {
        modelName: 'manual',
        actionName: 'inspect',
        refs: {},
        refsUsed: [],
        commands: [],
        ops: [],
        changed: [],
        checks: [terminalCheck],
      }),
    ).toEqual([terminalCheck])
  })

  it('fails unresolved noFormulaErrors checks instead of leaving them planned', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-run-adapter-unresolved-formula-check' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setTable({
      name: 'Inputs',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Kind', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    const model = defineModel({
      name: 'generic-row-check',
      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          rows: workbook.findRows({ table, where: { column: 'Kind', op: 'eq', value: 'actual' } }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.noFormulaErrors(refs.rows)]
      },
      actions: {
        inspect() {},
      },
    })

    const result = await runWorkbookAction(model, 'inspect', createWorkbookRunAdapter(engine))

    expect(result).toMatchObject({
      status: 'failed',
      errors: [{ code: 'check_failed' }],
      checks: [expect.objectContaining({ status: 'failed', kind: 'noFormulaErrors' })],
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
          code: 'engine_apply_failed',
          message: expect.stringContaining('Ambiguous table selector table with Input, Result matched First, Second'),
        },
      ],
    })
  })
})
