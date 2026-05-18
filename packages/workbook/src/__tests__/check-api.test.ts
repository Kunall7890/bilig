import { describe, expect, it } from 'vitest'
import { check, defineModel, describePlanResult, findRange, findTable, planWorkbookAction, verifyPlan } from '../index.js'

describe('@bilig/workbook check api', () => {
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
    expect(() => check.custom({ kind: ' ', message: 'valid message' })).toThrowError('Workbook check kind cannot be empty')
    expect(() => check.custom({ kind: 'validKind', message: ' ' })).toThrowError('Workbook check message cannot be empty')
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
              id: 'table_Inputs',
              label: 'Inputs',
              name: 'Inputs',
            },
            refs: [
              {
                kind: 'range',
                id: 'range_Model_C2_C2',
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
            id: 'range_Model_A2_A2',
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
})
