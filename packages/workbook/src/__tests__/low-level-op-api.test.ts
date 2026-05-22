import { describe, expect, it } from 'vitest'
import { buildWorkbookActionPlan, defineModel, describePlan, planWorkbookAction, verifyPlan } from '../index.js'

describe('@bilig/workbook low-level op api', () => {
  it('lets generic model actions add guarded low-level workbook ops', () => {
    const model = defineModel({
      name: 'generic-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'B2',
              value: 42,
            },
            {
              target: refs.output,
              message: 'Seed output value',
            },
          )
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'seed')
    const op = {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'B2',
      value: 42,
    } as const
    const output = {
      kind: 'range',
      id: 'range_p_Sheet1_p_B2_p_B2',
      label: 'Sheet1!B2',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B2',
      },
    } as const

    expect(plan.commands).toEqual([
      {
        kind: 'op',
        op,
        target: plan.refs.output,
        message: 'Seed output value',
      },
    ])
    expect(plan.ops).toEqual([op])
    expect(plan.changed).toEqual([
      {
        kind: 'op',
        target: plan.refs.output,
        message: 'Seed output value',
      },
    ])
    expect(describePlan(plan).commands).toEqual([
      {
        kind: 'op',
        op,
        target: output,
        message: 'Seed output value',
      },
    ])
    expect(verifyPlan(plan)).toEqual({
      status: 'valid',
      modelName: 'generic-op-model',
      actionName: 'seed',
      issues: [],
    })
  })

  it('rejects invalid low-level workbook ops before runtime handoff', () => {
    const model = defineModel({
      name: 'invalid-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          // @ts-expect-error exercising the runtime guard for plain JS callers
          workbook.addOp(
            {
              kind: 'notAWorkbookOp',
              sheetName: 'Sheet1',
              address: 'B2',
              value: 42,
            },
            { target: refs.output },
          )
        },
      },
    })

    expect(planWorkbookAction(model, 'seed')).toEqual({
      status: 'failed',
      modelName: 'invalid-op-model',
      actionName: 'seed',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook op is not a valid WorkbookOp',
        },
      ],
    })
  })

  it('verifies low-level workbook ops remain present in the action plan', () => {
    const model = defineModel({
      name: 'missing-low-level-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'B2',
              value: 42,
            },
            { target: refs.output },
          )
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'seed')

    expect(verifyPlan({ ...plan, ops: [] })).toEqual({
      status: 'invalid',
      modelName: 'missing-low-level-op-model',
      actionName: 'seed',
      issues: [
        {
          code: 'missing_workbook_op',
          path: 'commands[0].op',
          message: 'Low-level workbook op setCellValue is missing from plan ops',
        },
      ],
    })
  })

  it('keeps guarded low-level ops stable after the caller mutates the original object', () => {
    const model = defineModel({
      name: 'immutable-low-level-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          const op: {
            kind: 'setCellValue'
            sheetName: string
            address: string
            value: number
          } = {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 42,
          }
          workbook.addOp(op, { target: refs.output })
          op.value = 99
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'seed')

    expect(plan.commands).toEqual([
      {
        kind: 'op',
        op: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'B2',
          value: 42,
        },
        target: plan.refs.output,
      },
    ])
    expect(plan.ops).toEqual([
      {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'B2',
        value: 42,
      },
    ])
  })

  it('verifies low-level workbook ops with stable structural equality', () => {
    const model = defineModel({
      name: 'stable-low-level-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'B2',
              value: 42,
            },
            { target: refs.output },
          )
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'seed')

    expect(
      verifyPlan({
        ...plan,
        ops: [
          {
            value: 42,
            address: 'B2',
            kind: 'setCellValue',
            sheetName: 'Sheet1',
          },
        ],
      }),
    ).toEqual({
      status: 'valid',
      modelName: 'stable-low-level-op-model',
      actionName: 'seed',
      issues: [],
    })
  })

  it('verifies low-level workbook ops match their declared target when possible', () => {
    const model = defineModel({
      name: 'mismatched-low-level-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        seed({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'Z99',
              value: 42,
            },
            { target: refs.output },
          )
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'seed')

    expect(verifyPlan(plan)).toEqual({
      status: 'invalid',
      modelName: 'mismatched-low-level-op-model',
      actionName: 'seed',
      issues: [
        {
          code: 'op_target_mismatch',
          path: 'commands[0].target',
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
          message: 'Low-level workbook op setCellValue targets Sheet1!Z99:Z99, but command target is Sheet1!B2',
        },
      ],
    })
  })

  it('verifies low-level workbook op targets from filter and nested payload ranges', () => {
    const model = defineModel({
      name: 'nested-target-low-level-op-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }),
        }
      },

      actions: {
        filter({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setFilter',
              sheetName: 'Sheet1',
              range: {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A5',
              },
            },
            { target: refs.output },
          )
        },
        validate({ refs, workbook }) {
          workbook.addOp(
            {
              kind: 'setDataValidation',
              validation: {
                range: {
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A5',
                },
                rule: {
                  kind: 'any',
                },
              },
            },
            { target: refs.output },
          )
        },
      },
    })

    const filterPlan = buildWorkbookActionPlan(model, 'filter')
    const validationPlan = buildWorkbookActionPlan(model, 'validate')
    const mismatch = {
      code: 'op_target_mismatch',
      path: 'commands[0].target',
      ref: {
        kind: 'range',
        id: 'range_p_Sheet1_p_B2_p_C3',
        label: 'Sheet1!B2:C3',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C3',
        },
      },
      message: expect.stringContaining('targets Sheet1!A1:A5'),
    }

    expect(verifyPlan(filterPlan).issues).toEqual([mismatch])
    expect(verifyPlan(validationPlan).issues).toEqual([mismatch])
  })
})
