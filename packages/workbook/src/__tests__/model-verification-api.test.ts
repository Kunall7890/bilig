import { describe, expect, it } from 'vitest'

import { verifyModel } from '../index.js'

function verifyUnknownModel(model: unknown) {
  return verifyModel(model)
}

describe('@bilig/workbook model verification api', () => {
  it('returns structured invalid_model results for invalid model manifests without invoking getters', () => {
    let modelNameGetterInvoked = false
    const modelWithAccessorName: Record<string, unknown> = {
      actions: {
        write() {},
      },
    }
    Object.defineProperty(modelWithAccessorName, 'name', {
      enumerable: true,
      get() {
        modelNameGetterInvoked = true
        throw new Error('model name getter must not run')
      },
    })

    expect(verifyUnknownModel(modelWithAccessorName)).toEqual({
      status: 'invalid',
      modelName: 'unknown-model',
      errors: [
        {
          code: 'invalid_model',
          message: 'Workbook model name must be a data property',
        },
      ],
      actions: [],
    })
    expect(modelNameGetterInvoked).toBe(false)

    let actionGetterInvoked = false
    const actionMap: Record<string, unknown> = {}
    Object.defineProperty(actionMap, 'write', {
      enumerable: true,
      get() {
        actionGetterInvoked = true
        throw new Error('action getter must not run')
      },
    })

    expect(
      verifyUnknownModel({
        name: 'verify-accessor-action-model',
        actions: actionMap,
      }),
    ).toEqual({
      status: 'invalid',
      modelName: 'verify-accessor-action-model',
      errors: [
        {
          code: 'invalid_model',
          message: 'Workbook model verify-accessor-action-model action write must be a data property',
        },
      ],
      actions: [],
    })
    expect(actionGetterInvoked).toBe(false)
  })
})
