import { describe, expect, it } from 'vitest'

import { verifyModel } from '../index.js'

function verifyUnknownModel(model: unknown) {
  return verifyModel(model)
}

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

    const accessorNameVerification = verifyUnknownModel(modelWithAccessorName)
    expect(accessorNameVerification).toEqual({
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
    expect(Object.isFrozen(accessorNameVerification)).toBe(true)
    expect(Object.isFrozen(accessorNameVerification.errors)).toBe(true)
    expect(Object.isFrozen(accessorNameVerification.errors?.[0])).toBe(true)
    expect(Object.isFrozen(accessorNameVerification.actions)).toBe(true)
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

  it('returns structured invalid results for accessor-backed verification options without invoking getters', () => {
    const model = {
      name: 'verify-options-model',
      actions: {
        write() {},
      },
    }

    let inputsGetterInvoked = false
    const options: Record<string, unknown> = {}
    Object.defineProperty(options, 'inputs', {
      enumerable: true,
      get() {
        inputsGetterInvoked = true
        throw new Error('inputs getter must not run')
      },
    })

    expect(verifyModel(model, options)).toEqual({
      status: 'invalid',
      modelName: 'verify-options-model',
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Workbook model verification options inputs must be a data property',
          path: 'options.inputs',
          issueCode: 'invalid_action_input',
        },
      ],
      actions: [],
    })
    expect(inputsGetterInvoked).toBe(false)

    expect(
      verifyModel(
        model,
        customPrototypeRecord({
          inputs: {},
        }),
      ),
    ).toEqual({
      status: 'invalid',
      modelName: 'verify-options-model',
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Workbook model verification options must be an object',
          path: 'options',
          issueCode: 'invalid_action_input',
        },
      ],
      actions: [],
    })
    expect(
      verifyModel(model, {
        inputs: customPrototypeRecord({
          write: {},
        }),
      }),
    ).toEqual({
      status: 'invalid',
      modelName: 'verify-options-model',
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Workbook model verification options inputs must be an object',
          path: 'options.inputs',
          issueCode: 'invalid_action_input',
        },
      ],
      actions: [],
    })
  })

  it('returns structured failed planning for accessor-backed per-action inputs without invoking getters', () => {
    const model = {
      name: 'verify-action-input-model',
      actions: {
        write() {},
      },
    }

    let actionInputGetterInvoked = false
    const inputs: Record<string, unknown> = {}
    Object.defineProperty(inputs, 'write', {
      enumerable: true,
      get() {
        actionInputGetterInvoked = true
        throw new Error('action input getter must not run')
      },
    })

    expect(verifyModel(model, { inputs })).toEqual({
      status: 'invalid',
      modelName: 'verify-action-input-model',
      actions: [
        {
          actionName: 'write',
          planning: {
            status: 'failed',
            modelName: 'verify-action-input-model',
            actionName: 'write',
            errors: [
              {
                code: 'invalid_action_input',
                message: 'Workbook model verification input for action write must be a data property',
                path: 'inputs.write',
                issueCode: 'invalid_action_input',
              },
            ],
            checks: [],
          },
        },
      ],
    })
    expect(actionInputGetterInvoked).toBe(false)
  })

  it('prefixes invalid nested per-action input paths without invoking getters', () => {
    const model = {
      name: 'verify-nested-input-model',
      actions: {
        write() {},
      },
    }

    let valueGetterInvoked = false
    const input: Record<string, unknown> = {}
    Object.defineProperty(input, 'value', {
      enumerable: true,
      get() {
        valueGetterInvoked = true
        throw new Error('nested input getter must not run')
      },
    })

    expect(verifyModel(model, { inputs: { write: input } })).toEqual({
      status: 'invalid',
      modelName: 'verify-nested-input-model',
      actions: [
        {
          actionName: 'write',
          planning: {
            status: 'failed',
            modelName: 'verify-nested-input-model',
            actionName: 'write',
            errors: [
              {
                code: 'invalid_action_input',
                message: 'Action input at input.value must be a data property',
                path: 'inputs.write.value',
                issueCode: 'invalid_action_input',
              },
            ],
            checks: [],
          },
        },
      ],
    })
    expect(valueGetterInvoked).toBe(false)
  })
})
