import { describe, expect, it } from 'vitest'
import { check, defineModel, describePlanResult, findRange, findTable, formula, planWorkbookAction, verifyPlan } from '../index.js'

describe('@bilig/workbook check api', () => {
  it('exports machine-readable readback checks without runtime dependencies', () => {
    const table = findTable({ name: 'Inputs' })
    const amount = table.column('Amount')
    const output = findRange({ sheetName: 'Model', address: 'C2' })

    expect(check.valueEquals(output, 12)).toEqual({
      status: 'planned',
      kind: 'valueEquals',
      target: output,
      message: 'Model!C2 equals 12',
      expectation: {
        kind: 'valueEquals',
        value: 12,
      },
    })
    expect(check.valuesEqual(output, [[12, 24]])).toEqual({
      status: 'planned',
      kind: 'valuesEqual',
      target: output,
      message: 'Model!C2 values equal [[12,24]]',
      expectation: {
        kind: 'valuesEqual',
        values: [[12, 24]],
      },
    })
    expect(
      check.formulaEquals(output, formula.multiply(amount, 2), {
        message: 'Output formula matches the declared model formula',
      }),
    ).toEqual({
      status: 'planned',
      kind: 'formulaEquals',
      target: output,
      message: 'Output formula matches the declared model formula',
      expectation: {
        kind: 'formulaEquals',
        formula: '(Inputs[Amount])*(2)',
        inputs: [amount],
      },
    })
    expect(check.formulasEqual(output, [['A2+B2', null]])).toEqual({
      status: 'planned',
      kind: 'formulasEqual',
      target: output,
      message: 'Model!C2 formulas equal [["A2+B2",null]]',
      expectation: {
        kind: 'formulasEqual',
        formulas: [['A2+B2', null]],
      },
    })
    expect(() => check.valueEquals(output, Number.NaN)).toThrowError('Workbook readback value must be a finite JSON literal')
    expect(() => Reflect.apply(check.exists, undefined, [{ kind: 'range' }])).toThrowError('Workbook check target must be a WorkbookRef')
    expect(() => check.valuesEqual(output, [[1], [1, 2]])).toThrowError('Workbook readback values must be rectangular')
    expect(() => check.formulasEqual(output, [['A1'], [null, 'B1']])).toThrowError('Workbook readback formulas must be rectangular')
    expect(() =>
      Reflect.apply(check.custom, undefined, [
        {
          kind: 'badRefs',
          refs: [{ kind: 'range' }],
          message: 'Bad refs are rejected',
        },
      ]),
    ).toThrowError('Workbook check ref must be a WorkbookRef')
  })

  it('exports custom planned checks for consumer-defined invariants', () => {
    const inputs = findTable({ name: 'Inputs' })
    const output = findRange({ sheetName: 'Model', address: 'C2' })

    expect(
      check.custom({
        kind: 'balanced',
        target: inputs,
        refs: [output, output],
        message: 'Inputs stay balanced',
      }),
    ).toEqual({
      status: 'planned',
      kind: 'balanced',
      target: inputs,
      refs: [output],
      message: 'Inputs stay balanced',
    })
    expect(
      check.custom({
        kind: 'modelReady',
        message: 'Model has all required inputs',
      }),
    ).toEqual({
      status: 'planned',
      kind: 'modelReady',
      message: 'Model has all required inputs',
    })
    const prePassedOptions = {
      kind: 'prePassed',
      message: 'Consumer code cannot self-prove a check',
      status: 'passed',
    }
    expect(check.custom(prePassedOptions)).toEqual({
      status: 'planned',
      kind: 'prePassed',
      message: 'Consumer code cannot self-prove a check',
    })
    expect(() => check.custom({ kind: ' ', message: 'valid message' })).toThrowError('Workbook check kind cannot be empty')
    expect(() => check.custom({ kind: 'validKind', message: ' ' })).toThrowError('Workbook check message cannot be empty')
  })

  it('plans and describes readback checks through model actions', () => {
    const model = defineModel({
      name: 'readback-check-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs' })
        return {
          amount: table.column('Amount'),
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
          formulaOutput: workbook.findRange({ sheetName: 'Model', address: 'D2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const expectedFormula = formula.add(refs.amount, 1)
          workbook.writeValue(refs.output, 12)
          workbook.check.valueEquals(refs.output, 12)
          workbook.writeFormula(refs.formulaOutput, expectedFormula)
          workbook.check.formulaEquals(refs.formulaOutput, expectedFormula)
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      throw new Error('expected planned result')
    }

    expect(result.plan.checks).toEqual([
      {
        status: 'planned',
        kind: 'valueEquals',
        target: result.plan.refs.output,
        message: 'Model!C2 equals 12',
        expectation: {
          kind: 'valueEquals',
          value: 12,
        },
      },
      {
        status: 'planned',
        kind: 'formulaEquals',
        target: result.plan.refs.formulaOutput,
        message: 'Model!D2 formula equals (Inputs[Amount])+(1)',
        expectation: {
          kind: 'formulaEquals',
          formula: '(Inputs[Amount])+(1)',
          inputs: [result.plan.refs.amount],
        },
      },
    ])
    expect(verifyPlan(result.plan)).toEqual({
      status: 'valid',
      modelName: 'readback-check-model',
      actionName: 'calculate',
      issues: [],
    })
    expect(describePlanResult(result)).toEqual({
      status: 'planned',
      plan: expect.objectContaining({
        modelName: 'readback-check-model',
        actionName: 'calculate',
        checks: [
          {
            status: 'planned',
            kind: 'valueEquals',
            target: {
              kind: 'range',
              id: 'range_p_Model_p_C2_p_C2',
              label: 'Model!C2',
              range: {
                sheetName: 'Model',
                startAddress: 'C2',
                endAddress: 'C2',
              },
            },
            message: 'Model!C2 equals 12',
            expectation: {
              kind: 'valueEquals',
              value: 12,
            },
          },
          {
            status: 'planned',
            kind: 'formulaEquals',
            target: {
              kind: 'range',
              id: 'range_p_Model_p_D2_p_D2',
              label: 'Model!D2',
              range: {
                sheetName: 'Model',
                startAddress: 'D2',
                endAddress: 'D2',
              },
            },
            message: 'Model!D2 formula equals (Inputs[Amount])+(1)',
            expectation: {
              kind: 'formulaEquals',
              formula: '(Inputs[Amount])+(1)',
              inputs: [
                {
                  kind: 'column',
                  id: 'table_p_Inputs_p_Amount',
                  label: 'Inputs.Amount',
                  table: {
                    kind: 'table',
                    id: 'table_p_Inputs',
                    label: 'Inputs',
                    name: 'Inputs',
                  },
                  name: 'Amount',
                },
              ],
            },
          },
        ],
      }),
    })
  })

  it('plans and describes custom checks without built-in business concepts', () => {
    const model = defineModel({
      name: 'custom-check-model',

      find(workbook) {
        return {
          inputs: workbook.findTable({ name: 'Inputs' }),
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [
          workbook.check.custom({
            kind: 'inputContract',
            target: refs.inputs,
            refs: [refs.output],
            message: 'Inputs satisfy the consumer-defined contract',
          }),
          workbook.check.custom({
            kind: 'modelReady',
            message: 'The consumer-defined model is ready to run',
          }),
        ]
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeValue(refs.output, 1)
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    const described = describePlanResult(result)

    expect(result.status).toBe('planned')
    if (result.status === 'planned') {
      expect(result.plan.checks).toEqual([
        {
          status: 'planned',
          kind: 'inputContract',
          target: result.plan.refs.inputs,
          refs: [result.plan.refs.output],
          message: 'Inputs satisfy the consumer-defined contract',
        },
        {
          status: 'planned',
          kind: 'modelReady',
          message: 'The consumer-defined model is ready to run',
        },
      ])
      expect(verifyPlan(result.plan)).toEqual({
        status: 'valid',
        modelName: 'custom-check-model',
        actionName: 'calculate',
        issues: [],
      })
    }
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
    expect(described).toEqual({
      status: 'planned',
      plan: expect.objectContaining({
        modelName: 'custom-check-model',
        actionName: 'calculate',
        checks: [
          {
            status: 'planned',
            kind: 'inputContract',
            target: {
              kind: 'table',
              id: 'table_p_Inputs',
              label: 'Inputs',
              name: 'Inputs',
            },
            refs: [
              {
                kind: 'range',
                id: 'range_p_Model_p_C2_p_C2',
                label: 'Model!C2',
                range: {
                  sheetName: 'Model',
                  startAddress: 'C2',
                  endAddress: 'C2',
                },
              },
            ],
            message: 'Inputs satisfy the consumer-defined contract',
          },
          {
            status: 'planned',
            kind: 'modelReady',
            message: 'The consumer-defined model is ready to run',
          },
        ],
      }),
    })
  })

  it('verifies custom check refs resolve through the model refs contract', () => {
    const model = defineModel({
      name: 'custom-check-ref-verification',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      checks({ workbook }) {
        const hiddenInput = workbook.findRange({ sheetName: 'Model', address: 'A2' })
        return [
          workbook.check.custom({
            kind: 'inputContract',
            refs: [hiddenInput],
            message: 'Input must be available',
          }),
        ]
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeValue(refs.output, 1)
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      throw new Error('expected planned result')
    }

    expect(verifyPlan(result.plan)).toEqual({
      status: 'invalid',
      modelName: 'custom-check-ref-verification',
      actionName: 'calculate',
      issues: [
        {
          code: 'check_ref_not_resolved',
          path: 'checks[0].refs[0]',
          ref: {
            kind: 'range',
            id: 'range_p_Model_p_A2_p_A2',
            label: 'Model!A2',
            range: {
              sheetName: 'Model',
              startAddress: 'A2',
              endAddress: 'A2',
            },
          },
          message: 'Model!A2 appears in checks but is missing from refsUsed',
        },
      ],
    })
  })

  it('verifies formula expectation inputs resolve through the model refs contract', () => {
    const model = defineModel({
      name: 'readback-check-ref-verification',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const hiddenInput = workbook.findRange({ sheetName: 'Model', address: 'A2' })
          workbook.check.formulaEquals(refs.output, formula.add(hiddenInput, 1))
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      throw new Error('expected planned result')
    }

    expect(verifyPlan(result.plan)).toEqual({
      status: 'invalid',
      modelName: 'readback-check-ref-verification',
      actionName: 'calculate',
      issues: [
        {
          code: 'check_expectation_input_not_resolved',
          path: 'checks[0].expectation.inputs[0]',
          ref: {
            kind: 'range',
            id: 'range_p_Model_p_A2_p_A2',
            label: 'Model!A2',
            range: {
              sheetName: 'Model',
              startAddress: 'A2',
              endAddress: 'A2',
            },
          },
          message: 'Model!A2 appears in a formula expectation but is missing from refsUsed',
        },
      ],
    })
  })

  it('verifies formula expectation text is parseable', () => {
    const model = defineModel({
      name: 'readback-check-formula-verification',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.check.formulaEquals(refs.output, formula.raw('1+1'))
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      throw new Error('expected planned result')
    }

    const [plannedCheck] = result.plan.checks
    if (plannedCheck?.expectation?.kind !== 'formulaEquals') {
      throw new Error('expected formula expectation')
    }
    const brokenPlan = {
      ...result.plan,
      checks: [
        {
          ...plannedCheck,
          expectation: {
            ...plannedCheck.expectation,
            formula: 'SUM(',
          },
        },
      ],
    }

    expect(verifyPlan(brokenPlan)).toEqual({
      status: 'invalid',
      modelName: 'readback-check-formula-verification',
      actionName: 'calculate',
      issues: [
        expect.objectContaining({
          code: 'invalid_check_expectation_formula',
          path: 'checks[0].expectation.formula',
        }),
      ],
    })
  })
})
