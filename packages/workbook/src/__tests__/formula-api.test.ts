import { describe, expect, it } from 'vitest'
import { defineModel, describePlan, findRange, formula, planWorkbookAction, verifyPlan } from '../index.js'

describe('@bilig/workbook formula api', () => {
  it('freezes the public formula namespace and helper arrays', () => {
    const input = findRange({ sheetName: 'Sheet1', address: 'A1' })

    expect(Object.isFrozen(formula)).toBe(true)
    expect(formula.inputs(input)).toEqual([input])
    expect(Object.isFrozen(formula.inputs(input))).toBe(true)
    expect(formula.labels(input)).toEqual([{ name: 'Sheet1!A1', ref: input }])
    expect(Object.isFrozen(formula.labels(input))).toBe(true)
    expect(Object.isFrozen(formula.labels(input)[0])).toBe(true)
    expect(formula.inputs(1)).toEqual([])
    expect(Object.isFrozen(formula.inputs(1))).toBe(true)
    expect(formula.labels(false)).toEqual([])
    expect(Object.isFrozen(formula.labels(false))).toBe(true)
  })

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
      labels: [
        { name: 'Sheet1!A1', ref: result.plan.refs.amount },
        { name: 'Sheet1!B1', ref: result.plan.refs.rate },
      ],
    })
    expect(describePlan(result.plan).commands[0]).toMatchObject({
      kind: 'writeFormula',
      inputs: [
        { kind: 'range', label: 'Sheet1!A1' },
        { kind: 'range', label: 'Sheet1!B1' },
      ],
      labels: [
        { name: 'Sheet1!A1', ref: { kind: 'range', label: 'Sheet1!A1' } },
        { name: 'Sheet1!B1', ref: { kind: 'range', label: 'Sheet1!B1' } },
      ],
    })
    expect(verifyPlan(result.plan)).toEqual({
      status: 'valid',
      modelName: 'raw-formula-input-model',
      actionName: 'calculate',
      issues: [],
    })
  })

  it('keeps custom raw formula labels executable and verifiable', () => {
    const model = defineModel({
      name: 'raw-formula-label-model',
      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(
            refs.result,
            formula.raw('amount_token*2', {
              labels: [{ name: 'amount_token', ref: refs.amount }],
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
      formula: 'amount_token*2',
      inputs: [result.plan.refs.amount],
      labels: [{ name: 'amount_token', ref: result.plan.refs.amount }],
    })
    expect(verifyPlan(result.plan).status).toBe('valid')
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected writeFormula command')
    }

    expect(
      verifyPlan({
        ...result.plan,
        commands: [{ ...command, formula: '2' }],
      }).issues,
    ).toEqual([
      expect.objectContaining({
        code: 'formula_label_not_used',
        path: 'commands[0].labels[0].name',
      }),
      expect.objectContaining({
        code: 'missing_concrete_op',
        path: 'commands[0]',
      }),
    ])
  })

  it('keeps raw formula source behavior backward compatible', () => {
    expect(formula.raw('=1+1')).toEqual({
      kind: 'formula',
      source: '1+1',
      inputs: [],
      labels: [],
    })
  })

  it('rejects bare strings as ambiguous formula operands', () => {
    const message = 'Formula operands must be formula expressions, workbook refs, finite numbers, booleans, or formula.text/raw wrappers'

    expect(() => {
      // @ts-expect-error bare strings are intentionally rejected at type and runtime
      formula.source('Sheet1!A1')
    }).toThrowError(message)
    expect(() => {
      // @ts-expect-error bare strings are intentionally rejected at type and runtime
      formula.sum('Sheet1!A1')
    }).toThrowError(message)
    expect(() => {
      // @ts-expect-error bare strings are intentionally rejected at type and runtime
      formula.multiply(2, 'Sheet1!B1')
    }).toThrowError(message)
    expect(formula.source(formula.raw('Sheet1!A1'))).toBe('Sheet1!A1')
    expect(formula.source(formula.text('Sheet1!A1'))).toBe('"Sheet1!A1"')
  })

  it('dedupes raw formula inputs without mutating the caller array', () => {
    const first = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const second = findRange({ sheetName: 'Sheet1', address: 'B1' })
    const inputs = [first, second, first]

    const expression = formula.raw('SUM(Sheet1!A1,Sheet1!B1)', { inputs })

    expect(expression.inputs).toEqual([first, second])
    expect(expression.inputs).not.toBe(inputs)
    expect(inputs).toEqual([first, second, first])
    expect(Object.isFrozen(expression)).toBe(true)
    expect(Object.isFrozen(expression.inputs)).toBe(true)
    expect(Object.isFrozen(expression.labels)).toBe(true)
  })

  it('rejects accessor-backed formula inputs without invoking getters', () => {
    const input = findRange({ sheetName: 'Sheet1', address: 'A1' })

    let optionsGetterInvoked = false
    const rawOptions = {}
    Object.defineProperty(rawOptions, 'inputs', {
      enumerable: true,
      get() {
        optionsGetterInvoked = true
        throw new Error('inputs getter must not run')
      },
    })
    expect(() => formula.raw('Sheet1!A1', rawOptions)).toThrowError('Formula raw options.inputs must be a data property')
    expect(optionsGetterInvoked).toBe(false)

    let labelGetterInvoked = false
    const label = { ref: input }
    Object.defineProperty(label, 'name', {
      enumerable: true,
      get() {
        labelGetterInvoked = true
        throw new Error('label getter must not run')
      },
    })
    expect(() => formula.raw('Sheet1!A1', { labels: [label] })).toThrowError('Formula raw options.labels[0].name must be a data property')
    expect(labelGetterInvoked).toBe(false)

    let operandGetterInvoked = false
    const expressionInputs: unknown[] = []
    Object.defineProperty(expressionInputs, '0', {
      enumerable: true,
      get() {
        operandGetterInvoked = true
        throw new Error('operand getter must not run')
      },
    })
    expressionInputs.length = 1
    expect(() =>
      formula.inputs({
        kind: 'formula',
        source: 'Sheet1!A1',
        inputs: expressionInputs,
        labels: [],
      }),
    ).toThrowError('Formula expression inputs[0] must be a data property')
    expect(operandGetterInvoked).toBe(false)

    const sparseArgs: unknown[] = []
    sparseArgs.length = 1
    expect(() => Reflect.apply(formula.call, undefined, ['SUM', sparseArgs])).toThrowError('Formula arguments[0] must be a data property')
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
            id: 'range_Sheet1_Z9_Z9',
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
})
