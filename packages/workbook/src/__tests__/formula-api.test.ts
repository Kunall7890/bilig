import { describe, expect, it } from 'vitest'
import { defineModel, describePlan, findRange, formula, materializeFormulaLabels, planWorkbookAction, verifyPlan } from '../index.js'

function customPrototypeRecord(fields: Record<string, unknown>): Record<string, unknown> {
  const value: Record<string, unknown> = {}
  Object.setPrototypeOf(value, { inherited: true })
  for (const [key, entry] of Object.entries(fields)) {
    Object.defineProperty(value, key, {
      enumerable: true,
      value: entry,
    })
  }
  return value
}

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

  it('does not treat formula label substrings or quoted text as used refs', () => {
    const model = defineModel({
      name: 'raw-formula-label-token-model',
      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        substring({ refs, workbook }) {
          workbook.writeFormula(
            refs.result,
            formula.raw('DATA+1', {
              labels: [{ name: 'A', ref: refs.amount }],
            }),
          )
        },
        quoted({ refs, workbook }) {
          workbook.writeFormula(
            refs.result,
            formula.raw('"A"', {
              labels: [{ name: 'A', ref: refs.amount }],
            }),
          )
        },
      },
    })

    const substring = planWorkbookAction(model, 'substring')
    const quoted = planWorkbookAction(model, 'quoted')

    expect(substring.status).toBe('planned')
    expect(quoted.status).toBe('planned')
    if (substring.status !== 'planned' || quoted.status !== 'planned') {
      return
    }

    expect(verifyPlan(substring.plan).issues).toEqual([
      expect.objectContaining({
        code: 'formula_label_not_used',
        path: 'commands[0].labels[0].name',
      }),
    ])
    expect(verifyPlan(quoted.plan).issues).toEqual([
      expect.objectContaining({
        code: 'formula_label_not_used',
        path: 'commands[0].labels[0].name',
      }),
    ])
  })

  it('does not treat check formula label substrings or quoted text as used refs', () => {
    const model = defineModel({
      name: 'raw-formula-check-label-token-model',
      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      actions: {
        inspect({ refs, workbook }) {
          workbook.check.formulaEquals(
            refs.result,
            formula.raw('"A"', {
              labels: [{ name: 'A', ref: refs.amount }],
            }),
          )
        },
      },
    })

    const result = planWorkbookAction(model, 'inspect')

    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      return
    }
    expect(verifyPlan(result.plan).issues).toEqual([
      expect.objectContaining({
        code: 'check_expectation_label_not_used',
        path: 'checks[0].expectation.labels[0].name',
      }),
    ])
  })

  it('materializes formula labels by parsed formula tokens', () => {
    expect(
      materializeFormulaLabels('amount_rate+amount+"amount"+LET(amount,1,amount+amount_rate)', [
        { name: 'amount', source: 'Sheet1!A1' },
        { name: 'amount_rate', source: 'Sheet1!B1' },
      ]),
    ).toBe('Sheet1!B1+Sheet1!A1+"amount"+LET(amount,1,amount+Sheet1!B1)')
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
    expect(() =>
      formula.raw(
        'Sheet1!A1',
        customPrototypeRecord({
          inputs: [input],
        }),
      ),
    ).toThrowError('Formula raw options must be an object')

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
    expect(() =>
      formula.raw('Sheet1!A1', {
        labels: [
          customPrototypeRecord({
            name: 'input',
            ref: input,
          }),
        ],
      }),
    ).toThrowError('Formula raw options.labels[0] must be a formula label')

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

    let extraInputGetterInvoked = false
    const expressionInputsWithExtraAccessor = [input]
    Object.defineProperty(expressionInputsWithExtraAccessor, 'hidden', {
      enumerable: true,
      get() {
        extraInputGetterInvoked = true
        throw new Error('extra input getter must not run')
      },
    })
    expect(() =>
      formula.inputs({
        kind: 'formula',
        source: 'Sheet1!A1',
        inputs: expressionInputsWithExtraAccessor,
        labels: [],
      }),
    ).toThrowError('Formula expression inputs.hidden must be a data property')
    expect(extraInputGetterInvoked).toBe(false)

    const expressionInputSubclass = new (class extends Array<unknown> {})()
    expressionInputSubclass.push(input)
    expect(() =>
      formula.inputs({
        kind: 'formula',
        source: 'Sheet1!A1',
        inputs: expressionInputSubclass,
        labels: [],
      }),
    ).toThrowError('Formula expression inputs must be a plain array')

    const sparseArgs: unknown[] = []
    sparseArgs.length = 1
    expect(() => Reflect.apply(formula.call, undefined, ['SUM', sparseArgs])).toThrowError('Formula arguments[0] must be a data property')

    const symbolArgs = [input]
    const secret = Symbol('secret')
    Object.defineProperty(symbolArgs, secret, {
      enumerable: true,
      value: input,
    })
    expect(() => Reflect.apply(formula.call, undefined, ['SUM', symbolArgs])).toThrowError(
      'Formula arguments[Symbol(secret)] must be a data property',
    )

    expect(() =>
      formula.source(
        customPrototypeRecord({
          kind: 'formula',
          source: 'Sheet1!A1',
          inputs: [input],
          labels: [],
        }),
      ),
    ).toThrowError('Formula operands must be formula expressions, workbook refs, finite numbers, booleans, or formula.text/raw wrappers')
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
