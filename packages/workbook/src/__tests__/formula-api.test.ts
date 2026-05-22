import { describe, expect, it } from 'vitest'
import { defineModel, describePlan, findRange, formula, planWorkbookAction, verifyPlan } from '../index.js'

describe('@bilig/workbook formula api', () => {
  it('keeps explicit raw formula inputs inspectable', () => {
    const model = defineModel({
      name: 'raw-formula-input-model',
      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          rate: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(
            refs.result,
            formula.raw('SUM(Sheet1!A1,Sheet1!B1)', {
              inputs: [refs.amount, refs.rate, refs.amount],
            }),
          )
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')

    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      return
    }
    const [command] = result.plan.commands
    expect(command).toMatchObject({
      kind: 'writeFormula',
      formula: 'SUM(Sheet1!A1,Sheet1!B1)',
      inputs: [result.plan.refs.amount, result.plan.refs.rate],
    })
    expect(describePlan(result.plan).commands[0]).toMatchObject({
      kind: 'writeFormula',
      inputs: [
        { kind: 'range', label: 'Sheet1!A1' },
        { kind: 'range', label: 'Sheet1!B1' },
      ],
    })
    expect(verifyPlan(result.plan)).toEqual({
      status: 'valid',
      modelName: 'raw-formula-input-model',
      actionName: 'calculate',
      issues: [],
    })
  })

  it('keeps raw formula source behavior backward compatible', () => {
    expect(formula.raw('=1+1')).toEqual({
      kind: 'formula',
      source: '1+1',
      inputs: [],
    })
  })

  it('dedupes raw formula inputs without mutating the caller array', () => {
    const first = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const second = findRange({ sheetName: 'Sheet1', address: 'B1' })
    const inputs = [first, second, first]

    const expression = formula.raw('SUM(Sheet1!A1,Sheet1!B1)', { inputs })

    expect(expression.inputs).toEqual([first, second])
    expect(expression.inputs).not.toBe(inputs)
    expect(inputs).toEqual([first, second, first])
  })

  it('freezes formula expressions and inspectable input lists', () => {
    const first = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const second = findRange({ sheetName: 'Sheet1', address: 'B1' })

    const expression = formula.add(first, second)
    const inputs = formula.inputs(first)

    expect(Object.isFrozen(expression)).toBe(true)
    expect(Object.isFrozen(expression.inputs)).toBe(true)
    expect(Object.isFrozen(inputs)).toBe(true)
    expect(inputs).toEqual([first])
  })

  it('inspects declared formula inputs and parser-discovered dependencies separately', () => {
    const amount = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const rate = findRange({ sheetName: 'Sheet1', address: 'B1' })
    const expression = formula.raw('SUM(Sheet1!A1,Data!B2:C4,Total,Items[Amount],A1#)', {
      inputs: [amount, rate],
    })

    const inspection = formula.inspect(expression)

    expect(Object.isFrozen(inspection)).toBe(true)
    expect(Object.isFrozen(inspection.inputs)).toBe(true)
    expect(Object.isFrozen(inspection.dependencies)).toBe(true)
    expect(Object.isFrozen(inspection.names)).toBe(true)
    expect(inspection).toMatchObject({
      source: 'SUM(Sheet1!A1,Data!B2:C4,Total,Items[Amount],A1#)',
      inputs: [amount, rate],
      dependencies: [
        {
          kind: 'range',
          address: 'Sheet1!A1:A1',
          refKind: 'cells',
          sheetName: 'Sheet1',
          explicitSheet: true,
          startAddress: 'A1',
          endAddress: 'A1',
          startRow: 0,
          endRow: 0,
          startCol: 0,
          endCol: 0,
          startRowAbsolute: false,
          endRowAbsolute: false,
          startColAbsolute: false,
          endColAbsolute: false,
        },
        {
          kind: 'range',
          address: 'Data!B2:C4',
          refKind: 'cells',
          sheetName: 'Data',
          explicitSheet: true,
          startAddress: 'B2',
          endAddress: 'C4',
          startRow: 1,
          endRow: 3,
          startCol: 1,
          endCol: 2,
          startRowAbsolute: false,
          endRowAbsolute: false,
          startColAbsolute: false,
          endColAbsolute: false,
        },
      ],
      names: ['Total'],
      tables: ['Items'],
      spills: ['A1'],
      volatile: false,
      producesSpill: false,
    })
    expect(formula.dependencies(expression)).toEqual(inspection.dependencies)
  })

  it('flags raw formula inputs that are not part of resolved model refs', () => {
    const hidden = findRange({ sheetName: 'Sheet1', address: 'Z9' })
    const model = defineModel({
      name: 'raw-formula-hidden-input-model',
      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.raw('Sheet1!Z9', { inputs: [hidden] }))
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')

    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      return
    }
    expect(verifyPlan(result.plan)).toEqual({
      status: 'invalid',
      modelName: 'raw-formula-hidden-input-model',
      actionName: 'calculate',
      issues: [
        {
          code: 'formula_input_not_resolved',
          path: 'commands[0].inputs[0]',
          ref: {
            kind: 'range',
            id: 'range_p_Sheet1_p_Z9_p_Z9',
            label: 'Sheet1!Z9',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'Z9',
              endAddress: 'Z9',
            },
          },
          message: 'Sheet1!Z9 is used as a formula input but is missing from refsUsed',
        },
      ],
    })
  })

  it('rejects malformed formula operands and raw formula inputs before planning', () => {
    const input = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const malformedRef: unknown = JSON.parse('{"kind":"range","id":"missing-range-shape","label":"bad ref"}')
    const malformedOperand: unknown = JSON.parse('{"kind":"table","id":"bad-table","label":"bad table"}')

    expect(() => {
      // @ts-expect-error exercising runtime validation for plain JS callers
      formula.raw('Sheet1!A1', { inputs: [malformedRef] })
    }).toThrowError('Formula input at inputs[0] must be a WorkbookRef')

    expect(() => {
      // @ts-expect-error exercising runtime validation for plain JS callers
      formula.raw('Sheet1!A1', { inputs: {} })
    }).toThrowError('Formula inputs must be an array')

    expect(() => {
      // @ts-expect-error exercising runtime validation for plain JS callers
      formula.add(input, malformedOperand)
    }).toThrowError('Formula operand must be a formula expression, WorkbookRef, string, finite number, or boolean')
  })
})
