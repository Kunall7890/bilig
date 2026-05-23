import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describeModel,
  inspectModel,
  normalizeWorkbookActionInputDescription,
  verifyModel,
  type WorkbookActionConfig,
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
