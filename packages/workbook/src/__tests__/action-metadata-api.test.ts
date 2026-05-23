import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describeModel,
  inspectModel,
  normalizeWorkbookActionInputDescription,
  verifyModel,
  type WorkbookActionContext,
  type WorkbookActionInput,
  type WorkbookRangeRef,
} from '../index.js'

function inputObject(input: WorkbookActionInput | undefined): Record<string, WorkbookActionInput> {
  if (typeof input !== 'object' || input === null || Array.isArray(input)) {
    throw new Error('action input must be an object')
  }
  return input
}

function defineUnknownModel(config: unknown): unknown {
  return Reflect.apply(defineModel, undefined, [config])
}

function inspectUnknownModel(model: unknown): unknown {
  return Reflect.apply(inspectModel, undefined, [model])
}

function describeUnknownModel(model: unknown): unknown {
  return Reflect.apply(describeModel, undefined, [model])
}

describe('@bilig/workbook action metadata api', () => {
  it('describes action objects with plain input metadata and still plans their run function', () => {
    const model = defineModel({
      name: 'metadata-model',
      description: ' Consumer-owned writer ',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        reset({ refs, workbook }) {
          workbook.clear(refs.output)
        },

        write: {
          description: ' Write a consumer-provided value ',
          input: {
            kind: 'object',
            description: 'Write request',
            fields: {
              value: {
                kind: 'number',
                required: true,
                description: 'Finite value to write',
              },
              note: {
                kind: 'string',
              },
            },
          },
          run({ refs, workbook, input }) {
            const value = inputObject(input).value
            if (typeof value !== 'number') {
              throw new Error('value input must be numeric')
            }
            workbook.writeValue(refs.output, value)
          },
        },
      },
    })

    const description = describeModel(model)

    expect(description).toEqual({
      name: 'metadata-model',
      description: 'Consumer-owned writer',
      actions: ['reset', 'write'],
      actionDetails: [
        {
          name: 'reset',
        },
        {
          name: 'write',
          description: 'Write a consumer-provided value',
          input: {
            kind: 'object',
            description: 'Write request',
            fields: {
              note: {
                kind: 'string',
              },
              value: {
                kind: 'number',
                description: 'Finite value to write',
                required: true,
              },
            },
          },
        },
      ],
      hasChecks: false,
    })
    expect(JSON.parse(JSON.stringify(description))).toEqual(description)

    const plan = buildWorkbookActionPlan(model, 'write', { value: 7, note: 'agent run' })
    expect(plan.commands).toEqual([
      {
        kind: 'writeValue',
        target: plan.refs.output,
        value: 7,
      },
    ])
    expect(verifyModel(model, { inputs: { write: { value: 7 } } }).status).toBe('valid')
  })

  it('freezes normalized model and action metadata at definition time', () => {
    const model = defineModel({
      name: 'stable-metadata-model',
      description: ' Consumer-owned writer ',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write: {
          description: ' Write a consumer-provided value ',
          input: {
            kind: 'object',
            fields: {
              value: { kind: 'number', required: true },
            },
          },
          run({ refs, workbook }) {
            workbook.writeValue(refs.output, 1)
          },
        },
      },
    })
    const writeAction = model.actions.write

    expect(model.description).toBe('Consumer-owned writer')
    expect(Object.isFrozen(model)).toBe(true)
    expect(Object.isFrozen(model.actions)).toBe(true)
    expect(() => Object.defineProperty(model.actions, 'mutated', { value() {} })).toThrowError(TypeError)
    if (typeof writeAction !== 'object' || writeAction === null) {
      throw new Error('expected action object')
    }
    expect(writeAction.description).toBe('Write a consumer-provided value')
    expect(Object.isFrozen(writeAction)).toBe(true)
    expect(Object.isFrozen(writeAction.input)).toBe(true)
    expect(() => Object.defineProperty(writeAction, 'description', { value: 'mutated' })).toThrowError(TypeError)
  })

  it('does not mutate caller-owned action config while freezing the model manifest', () => {
    type CallerRefs = { output: WorkbookRangeRef }
    const required: boolean = true
    const actionConfig = {
      description: ' Write a consumer-provided value ',
      input: {
        kind: 'object',
        fields: {
          value: { kind: 'number', required },
        },
      },
      run({ refs, workbook }: WorkbookActionContext<CallerRefs>) {
        workbook.writeValue(refs.output, 1)
      },
    } satisfies WorkbookActionConfig<CallerRefs>

    const model = defineModel({
      name: 'caller-owned-config-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write: actionConfig,
      },
    })
    const writeAction = model.actions.write

    expect(Object.isFrozen(actionConfig)).toBe(false)
    expect(Object.isFrozen(actionConfig.input)).toBe(false)
    expect(actionConfig.description).toBe(' Write a consumer-provided value ')
    if (typeof writeAction !== 'object' || writeAction === null) {
      throw new Error('expected action object')
    }
    expect(writeAction).not.toBe(actionConfig)
    expect(writeAction.description).toBe('Write a consumer-provided value')
    expect(Object.isFrozen(writeAction)).toBe(true)
    expect(Object.isFrozen(writeAction.input)).toBe(true)

    actionConfig.description = ' Mutated after defineModel '
    actionConfig.input.fields.value.required = false
    expect(describeModel(model).actionDetails).toEqual([
      {
        name: 'write',
        description: 'Write a consumer-provided value',
        input: {
          kind: 'object',
          fields: {
            value: { kind: 'number', required: true },
          },
        },
      },
    ])
  })

  it('only reads own action-object metadata for stable agent manifests', () => {
    type CallerRefs = { output: WorkbookRangeRef }
    const inheritedMetadata = {
      description: 'Inherited description must not leak',
      input: {
        kind: 'object',
        fields: {
          leaked: { kind: 'string' },
        },
      },
    }
    const actionConfig = Object.setPrototypeOf(
      {
        run({ refs, workbook }: WorkbookActionContext<CallerRefs>) {
          workbook.clear(refs.output)
        },
      },
      inheritedMetadata,
    )

    const model = defineModel({
      name: 'own-action-metadata-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write: actionConfig,
      },
    })

    expect(describeModel(model).actionDetails).toEqual([{ name: 'write' }])
    expect(buildWorkbookActionPlan(model, 'write').commands).toEqual([
      {
        kind: 'clear',
        target: expect.objectContaining({ label: 'Sheet1!B2' }),
      },
    ])
  })

  it('rejects action objects whose run function is inherited', () => {
    const inheritedRun = Object.setPrototypeOf(
      {},
      {
        run() {},
      },
    )

    expect(() =>
      defineModel({
        name: 'inherited-run-model',
        find() {
          return {}
        },
        actions: {
          // @ts-expect-error exercising the runtime guard for plain JS callers
          write: inheritedRun,
        },
      }),
    ).toThrowError('Workbook model inherited-run-model action write must be a function or action object with run')
  })

  it('rejects accessor-backed model config and action maps without invoking getters', () => {
    let actionsGetterInvoked = false
    const accessorConfig = {
      name: 'accessor-config-model',
      find() {
        return {}
      },
    }
    Object.defineProperty(accessorConfig, 'actions', {
      enumerable: true,
      get() {
        actionsGetterInvoked = true
        throw new Error('actions getter must not run')
      },
    })

    expect(() => defineUnknownModel(accessorConfig)).toThrowError('Workbook model config actions must be a data property')
    expect(actionsGetterInvoked).toBe(false)

    let actionGetterInvoked = false
    const actions: Record<string, unknown> = {}
    Object.defineProperty(actions, 'write', {
      enumerable: true,
      get() {
        actionGetterInvoked = true
        throw new Error('action getter must not run')
      },
    })

    expect(() =>
      defineUnknownModel({
        name: 'accessor-action-map-model',
        find() {
          return {}
        },
        actions,
      }),
    ).toThrowError('Workbook model accessor-action-map-model action write must be a data property')
    expect(actionGetterInvoked).toBe(false)
  })

  it('rejects accessor-backed action-object metadata without invoking getters', () => {
    let runGetterInvoked = false
    const actionWithAccessorRun: Record<string, unknown> = {}
    Object.defineProperty(actionWithAccessorRun, 'run', {
      enumerable: true,
      get() {
        runGetterInvoked = true
        throw new Error('run getter must not run')
      },
    })

    expect(() =>
      defineUnknownModel({
        name: 'accessor-action-run-model',
        find() {
          return {}
        },
        actions: {
          write: actionWithAccessorRun,
        },
      }),
    ).toThrowError('Workbook model accessor-action-run-model action write run must be a data property')
    expect(runGetterInvoked).toBe(false)

    let descriptionGetterInvoked = false
    const actionWithAccessorDescription: Record<string, unknown> = {
      run() {},
    }
    Object.defineProperty(actionWithAccessorDescription, 'description', {
      enumerable: true,
      get() {
        descriptionGetterInvoked = true
        throw new Error('description getter must not run')
      },
    })

    expect(() =>
      defineUnknownModel({
        name: 'accessor-action-description-model',
        find() {
          return {}
        },
        actions: {
          write: actionWithAccessorDescription,
        },
      }),
    ).toThrowError('Workbook model accessor-action-description-model action write description must be a data property')
    expect(descriptionGetterInvoked).toBe(false)
  })

  it('rejects accessor-backed inspected model manifests without invoking getters', () => {
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

    expect(() => inspectUnknownModel(modelWithAccessorName)).toThrowError('Workbook model name must be a data property')
    expect(modelNameGetterInvoked).toBe(false)

    let modelDescriptionGetterInvoked = false
    const modelWithAccessorDescription: Record<string, unknown> = {
      name: 'inspect-accessor-description-model',
      actions: {
        write() {},
      },
    }
    Object.defineProperty(modelWithAccessorDescription, 'description', {
      enumerable: true,
      get() {
        modelDescriptionGetterInvoked = true
        throw new Error('model description getter must not run')
      },
    })

    expect(() => describeUnknownModel(modelWithAccessorDescription)).toThrowError(
      'Workbook model inspect-accessor-description-model description must be a data property',
    )
    expect(modelDescriptionGetterInvoked).toBe(false)

    let actionGetterInvoked = false
    const actionMap: Record<string, unknown> = {}
    Object.defineProperty(actionMap, 'write', {
      enumerable: true,
      get() {
        actionGetterInvoked = true
        throw new Error('action getter must not run')
      },
    })

    expect(() =>
      inspectUnknownModel({
        name: 'inspect-accessor-action-model',
        actions: actionMap,
      }),
    ).toThrowError('Workbook model inspect-accessor-action-model action write must be a data property')
    expect(actionGetterInvoked).toBe(false)

    let actionDescriptionGetterInvoked = false
    const actionWithAccessorDescription: Record<string, unknown> = {
      run() {},
    }
    Object.defineProperty(actionWithAccessorDescription, 'description', {
      enumerable: true,
      get() {
        actionDescriptionGetterInvoked = true
        throw new Error('action description getter must not run')
      },
    })

    expect(() =>
      describeUnknownModel({
        name: 'inspect-accessor-action-description-model',
        actions: {
          write: actionWithAccessorDescription,
        },
      }),
    ).toThrowError('Workbook action write description must be a data property')
    expect(actionDescriptionGetterInvoked).toBe(false)
  })

  it('describes action metadata without running find, checks, or actions', () => {
    const model = defineModel({
      name: 'metadata-only-model',
      description: 'Inspectable without workbook access',
      find() {
        throw new Error('find should not run during manifest inspection')
      },
      checks() {
        throw new Error('checks should not run during manifest inspection')
      },
      actions: {
        create: {
          description: 'Create generic workbook intent',
          input: {
            kind: 'object',
            fields: {
              title: { kind: 'string', required: true },
            },
          },
          run() {
            throw new Error('action should not run during manifest inspection')
          },
        },
      },
    })

    expect(inspectModel(model)).toEqual({
      name: 'metadata-only-model',
      description: 'Inspectable without workbook access',
      actions: ['create'],
      actionDetails: [
        {
          name: 'create',
          description: 'Create generic workbook intent',
          input: {
            kind: 'object',
            fields: {
              title: { kind: 'string', required: true },
            },
          },
        },
      ],
      hasChecks: true,
    })
  })

  it('normalizes and freezes standalone input descriptions', () => {
    const description = normalizeWorkbookActionInputDescription({
      kind: 'array',
      description: ' Values ',
      items: {
        kind: 'object',
        fields: {
          amount: { kind: 'number', required: true },
        },
      },
    })

    expect(description).toEqual({
      kind: 'array',
      description: 'Values',
      items: {
        kind: 'object',
        fields: {
          amount: { kind: 'number', required: true },
        },
      },
    })
    expect(Object.isFrozen(description)).toBe(true)
    expect(Object.isFrozen(description.items)).toBe(true)
    expect(Object.isFrozen(description.items?.fields)).toBe(true)
  })

  it('rejects malformed action metadata at model definition time', () => {
    expect(() =>
      defineModel({
        name: 'bad-action-description',
        find() {
          return {}
        },
        actions: {
          write: {
            description: ' ',
            run() {},
          },
        },
      }),
    ).toThrowError('Workbook model bad-action-description action write description cannot be empty')

    expect(() =>
      defineModel({
        name: 'bad-action-input',
        find() {
          return {}
        },
        actions: {
          write: {
            input: {
              kind: 'string',
              fields: {
                value: { kind: 'number' },
              },
            },
            run() {},
          },
        },
      }),
    ).toThrowError('Action input description at input.fields can only be used when kind is object')

    expect(() =>
      normalizeWorkbookActionInputDescription({
        kind: 'object',
        fields: {
          ' ': { kind: 'number' },
        },
      }),
    ).toThrowError('Action input description at input.fields cannot contain an empty field name')
  })
})
