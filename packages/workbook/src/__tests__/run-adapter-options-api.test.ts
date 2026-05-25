import { describe, expect, it } from 'vitest'
import { checkWorkbookRunAdapter, defineModel, prepareWorkbookAction } from '../index.js'

describe('@bilig/workbook run adapter option boundary', () => {
  it('rejects accessor-backed adapter check options without invoking getters', async () => {
    const model = defineModel({
      name: 'adapter-options-boundary-model',
      find(workbook) {
        return {
          target: workbook.findName('target'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.target)]
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.target, 1)
        },
      },
    })

    const prepared = prepareWorkbookAction(model, 'write')
    if (prepared.status !== 'prepared') {
      throw new Error('expected adapter options fixture to prepare')
    }

    let getterInvoked = false
    const options: Record<string, unknown> = {}
    Object.defineProperty(options, 'expectedBaseRevision', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('expectedBaseRevision getter must not run')
      },
    })

    const checked = await Reflect.apply(checkWorkbookRunAdapter, undefined, [prepared.planData, {}, options])

    expect(getterInvoked).toBe(false)
    expect(checked).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run option expectedBaseRevision must be a data property',
          path: 'options.expectedBaseRevision',
          issueCode: 'invalid_run_options',
        },
      ],
      issues: [
        {
          code: 'invalid_run_options',
          path: 'options.expectedBaseRevision',
          message: 'Workbook run option expectedBaseRevision must be a data property',
        },
      ],
    })
  })

  it('rejects non-plain and unknown adapter check options', async () => {
    const model = defineModel({
      name: 'adapter-options-plain-model',
      find(workbook) {
        return {
          target: workbook.findName('target'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.target)]
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.target, 1)
        },
      },
    })

    const prepared = prepareWorkbookAction(model, 'write')
    if (prepared.status !== 'prepared') {
      throw new Error('expected adapter options fixture to prepare')
    }

    class AdapterOptions {
      expectedBaseRevision = 1
    }

    await expect(Reflect.apply(checkWorkbookRunAdapter, undefined, [prepared.planData, {}, new AdapterOptions()])).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run options must be a plain object',
          path: 'options',
          issueCode: 'invalid_run_options',
        },
      ],
      issues: [
        {
          code: 'invalid_run_options',
          path: 'options',
          message: 'Workbook run options must be a plain object',
        },
      ],
    })

    await expect(
      Reflect.apply(checkWorkbookRunAdapter, undefined, [prepared.planData, {}, { expectedBaseRevison: 1 }]),
    ).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run option expectedBaseRevison is unknown',
          path: 'options.expectedBaseRevison',
          issueCode: 'invalid_run_options',
        },
      ],
      issues: [
        {
          code: 'invalid_run_options',
          path: 'options.expectedBaseRevison',
          message: 'Workbook run option expectedBaseRevison is unknown',
        },
      ],
    })
  })
})
