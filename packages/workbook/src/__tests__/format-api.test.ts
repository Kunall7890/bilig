import { describe, expect, it } from 'vitest'
import { buildWorkbookActionPlan, defineModel, describePlan, verifyPlan } from '../index.js'

describe('@bilig/workbook format api', () => {
  it('compiles known single-cell number formats to concrete workbook ops', () => {
    const model = defineModel({
      name: 'format-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      actions: {
        applyCurrency({ refs, workbook }) {
          workbook.format(refs.output, { numberFormat: '$0.00' })
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'applyCurrency')

    expect(plan.commands).toEqual([
      {
        kind: 'format',
        target: plan.refs.output,
        numberFormat: '$0.00',
      },
    ])
    expect(plan.ops).toEqual([
      {
        kind: 'setCellFormat',
        sheetName: 'Model',
        address: 'C2',
        format: '$0.00',
      },
    ])
    expect(verifyPlan(plan)).toEqual({
      status: 'valid',
      modelName: 'format-model',
      actionName: 'applyCurrency',
      issues: [],
    })
    expect(describePlan(plan).commands).toEqual([
      {
        kind: 'format',
        target: {
          kind: 'range',
          id: 'range_Model_C2_C2',
          label: 'Model!C2',
          range: {
            sheetName: 'Model',
            startAddress: 'C2',
            endAddress: 'C2',
          },
        },
        numberFormat: '$0.00',
      },
    ])
  })

  it('compiles null number formats as explicit format clears', () => {
    const model = defineModel({
      name: 'clear-format-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      actions: {
        clearFormat({ refs, workbook }) {
          workbook.format(refs.output, { numberFormat: null })
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'clearFormat')

    expect(plan.ops).toEqual([
      {
        kind: 'setCellFormat',
        sheetName: 'Model',
        address: 'C2',
        format: null,
      },
    ])
    expect(verifyPlan(plan).status).toBe('valid')
  })

  it('verifies number format commands have matching concrete ops', () => {
    const model = defineModel({
      name: 'format-verification-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Model', address: 'C2' }),
        }
      },

      actions: {
        applyCurrency({ refs, workbook }) {
          workbook.format(refs.output, { numberFormat: '$0.00' })
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'applyCurrency')

    expect(verifyPlan({ ...plan, ops: [] })).toEqual({
      status: 'invalid',
      modelName: 'format-verification-model',
      actionName: 'applyCurrency',
      issues: [
        {
          code: 'missing_concrete_op',
          path: 'commands[0]',
          ref: {
            kind: 'range',
            id: 'range_Model_C2_C2',
            label: 'Model!C2',
            range: {
              sheetName: 'Model',
              startAddress: 'C2',
              endAddress: 'C2',
            },
          },
          message: 'Model!C2 has no matching concrete workbook op',
        },
      ],
    })
  })
})
