import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describePlan,
  describePlanResult,
  checkInput,
  isWorkbookActionInput,
  isWorkbookActionInputDescription,
  isWorkbookActionInputDescriptionKind,
  normalizeWorkbookActionInput,
  normalizeWorkbookActionInputDescription,
  planWorkbookAction,
  verifyModel,
  verifyPlan,
  workbookActionInputDescriptionKinds,
  type WorkbookActionInput,
} from '../index.js'

function inputObject(input: WorkbookActionInput | undefined): Record<string, WorkbookActionInput> {
  if (typeof input !== 'object' || input === null || Array.isArray(input)) {
    throw new Error('action input must be an object')
  }
  return input
}

describe('@bilig/workbook action input api', () => {
  it('exports stable action input kind and value guards for agent inspection', () => {
    expect(Object.isFrozen(workbookActionInputDescriptionKinds)).toBe(true)
    expect(workbookActionInputDescriptionKinds).toEqual(['json', 'object', 'array', 'string', 'number', 'boolean', 'null'])
    expect(new Set(workbookActionInputDescriptionKinds).size).toBe(workbookActionInputDescriptionKinds.length)
    expect(isWorkbookActionInputDescriptionKind('object')).toBe(true)
    expect(isWorkbookActionInputDescriptionKind('date')).toBe(false)
    expect(
      isWorkbookActionInputDescription({
        kind: 'object',
        fields: {
          value: { kind: 'number', required: true },
        },
      }),
    ).toBe(true)
    expect(isWorkbookActionInputDescription({ kind: 'object', fields: { value: { kind: 'date' } } })).toBe(false)
    expect(isWorkbookActionInput({ value: 12, label: 'ready', nested: [true, null] })).toBe(true)
    expect(isWorkbookActionInput({ value: Number.NaN })).toBe(false)
    expect(isWorkbookActionInput({ value: new Date('2026-05-22T00:00:00Z') })).toBe(false)
    const sparseInput: unknown[] = ['a', 'b']
    Reflect.deleteProperty(sparseInput, 1)
    expect(isWorkbookActionInput(sparseInput)).toBe(false)
  })

  it('checks payloads against action input metadata without schema dependencies', () => {
    const description = {
      kind: 'object',
      fields: {
        note: { kind: 'string' },
        tags: { kind: 'array', items: { kind: 'string' } },
        value: { kind: 'number', required: true },
      },
    }

    const valid = checkInput(description, {
      tags: ['ready', 'agent'],
      value: 12,
      extra: true,
    })

    expect(valid).toEqual({
      status: 'valid',
      input: {
        extra: true,
        tags: ['ready', 'agent'],
        value: 12,
      },
      issues: [],
    })
    expect(Object.isFrozen(valid.input)).toBe(true)

    expect(
      checkInput(description, {
        tags: ['ready', 3],
      }),
    ).toEqual({
      status: 'invalid',
      input: {
        tags: ['ready', 3],
      },
      issues: [
        {
          code: 'wrong_input_type',
          path: 'input.tags[1]',
          message: 'Action input at input.tags[1] must be a string',
        },
        {
          code: 'missing_required_input',
          path: 'input.value',
          message: 'Action input at input.value is required',
        },
      ],
    })

    expect(
      checkInput(
        { kind: 'json' },
        {
          tags: ['ready', Number.NaN],
        },
      ),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_action_input',
          path: 'input.tags[1]',
          message: 'Action input at input.tags[1] must be a finite number',
        },
      ],
    })
  })

  it('preserves magic JSON keys in normalized action inputs and descriptions', () => {
    const payload = JSON.parse('{"__proto__":{"polluted":true},"constructor":{"nested":1},"value":12}') as unknown

    const normalized = normalizeWorkbookActionInput(payload)

    expect(Object.getPrototypeOf(normalized)).toBe(Object.prototype)
    expect(Object.hasOwn(normalized, '__proto__')).toBe(true)
    expect(Object.hasOwn(normalized, 'constructor')).toBe(true)
    expect(JSON.parse(JSON.stringify(normalized))).toEqual(
      JSON.parse('{"__proto__":{"polluted":true},"constructor":{"nested":1},"value":12}'),
    )

    const description = normalizeWorkbookActionInputDescription(
      JSON.parse(
        '{"kind":"object","fields":{"__proto__":{"kind":"string","required":true},"constructor":{"kind":"number","required":true}}}',
      ),
    )

    expect(description.fields).toBeDefined()
    expect(Object.hasOwn(description.fields ?? {}, '__proto__')).toBe(true)
    expect(Object.hasOwn(description.fields ?? {}, 'constructor')).toBe(true)
    expect(checkInput(description, JSON.parse('{"__proto__":"safe","constructor":7}'))).toEqual({
      status: 'valid',
      input: JSON.parse('{"__proto__":"safe","constructor":7}'),
      issues: [],
    })
    expect(checkInput(description, JSON.parse('{"constructor":"bad"}'))).toEqual({
      status: 'invalid',
      input: {
        constructor: 'bad',
      },
      issues: [
        {
          code: 'missing_required_input',
          path: 'input.__proto__',
          message: 'Action input at input.__proto__ is required',
        },
        {
          code: 'wrong_input_type',
          path: 'input.constructor',
          message: 'Action input at input.constructor must be a number',
        },
      ],
    })
  })

  it('rejects accessor-backed action inputs without executing getters', () => {
    const payload = {
      value: 12,
    }
    Object.defineProperty(payload, 'hidden', {
      enumerable: true,
      get() {
        throw new Error('input accessor should not run')
      },
    })

    expect(() => normalizeWorkbookActionInput(payload)).toThrowError('Action input at input.hidden must be a data property')
    expect(checkInput({ kind: 'json' }, payload)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_action_input',
          path: 'input.hidden',
          message: 'Action input at input.hidden must be a data property',
        },
      ],
    })

    const arrayPayload: unknown[] = [1]
    Object.defineProperty(arrayPayload, '0', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('array input accessor should not run')
      },
    })
    expect(checkInput({ kind: 'array', items: { kind: 'number' } }, arrayPayload)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_action_input',
          path: 'input[0]',
          message: 'Action input at input[0] must be a data property',
        },
      ],
    })
  })

  it('rejects accessor-backed action input descriptions without executing getters', () => {
    const descriptionWithGetter = {
      kind: 'object',
      fields: {
        value: { kind: 'number' },
      },
    }
    Object.defineProperty(descriptionWithGetter, 'required', {
      enumerable: true,
      get() {
        throw new Error('description accessor should not run')
      },
    })

    expect(() => normalizeWorkbookActionInputDescription(descriptionWithGetter)).toThrowError(
      'Action input description at input.required must be a data property',
    )
    expect(checkInput(descriptionWithGetter, { value: 12 })).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_action_input_description',
          path: 'input.required',
          message: 'Action input description at input.required must be a data property',
        },
      ],
    })

    const fieldsWithGetter = {}
    Object.defineProperty(fieldsWithGetter, 'value', {
      enumerable: true,
      get() {
        throw new Error('field description accessor should not run')
      },
    })
    expect(() => normalizeWorkbookActionInputDescription({ kind: 'object', fields: fieldsWithGetter })).toThrowError(
      'Action input description at input.fields.value must be a data property',
    )
  })

  it('uses top-level required metadata for omitted action inputs', () => {
    expect(
      checkInput(
        {
          kind: 'object',
          fields: {
            value: { kind: 'number' },
          },
        },
        undefined,
      ),
    ).toEqual({
      status: 'valid',
      issues: [],
    })

    expect(
      checkInput(
        {
          kind: 'object',
          required: true,
          fields: {
            value: { kind: 'number' },
          },
        },
        undefined,
      ),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_required_input',
          path: 'input',
          message: 'Action input at input is required',
        },
      ],
    })
  })

  it('plans parameterized actions with cloned JSON-safe input', () => {
    const model = defineModel({
      name: 'parameterized-action-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      checks({ refs, workbook, input }) {
        return [
          workbook.check.custom({
            kind: 'inputReady',
            target: refs.output,
            message: inputObject(input).label === 'seed' ? 'Input is ready' : 'Input is unexpected',
          }),
        ]
      },

      actions: {
        write({ refs, workbook, input }) {
          const value = inputObject(input).value
          if (typeof value !== 'number') {
            throw new Error('value input must be numeric')
          }
          workbook.writeValue(refs.output, value)
        },
      },
    })
    const input = {
      value: 42,
      label: 'seed',
      nested: {
        z: false,
        a: 'first',
      },
    }

    const plan = buildWorkbookActionPlan(model, 'write', input)

    expect(plan.input).toEqual({
      label: 'seed',
      nested: {
        a: 'first',
        z: false,
      },
      value: 42,
    })
    expect(Object.isFrozen(plan.input)).toBe(true)
    expect(Object.isFrozen(inputObject(plan.input).nested)).toBe(true)
    expect(JSON.stringify(plan.input)).toBe('{"label":"seed","nested":{"a":"first","z":false},"value":42}')
    input.value = 99
    input.nested.a = 'mutated'
    expect(plan.input).toEqual({
      label: 'seed',
      nested: {
        a: 'first',
        z: false,
      },
      value: 42,
    })
    expect(plan.commands).toEqual([
      {
        kind: 'writeValue',
        target: plan.refs.output,
        value: 42,
      },
    ])
    expect(describePlan(plan).input).toEqual(plan.input)
    expect(JSON.parse(JSON.stringify(describePlan(plan)))).toEqual(describePlan(plan))
    expect(verifyPlan(plan)).toEqual({
      status: 'valid',
      modelName: 'parameterized-action-model',
      actionName: 'write',
      issues: [],
    })
  })

  it('preserves valid input on structured planning failures and descriptions', () => {
    const model = defineModel({
      name: 'parameterized-failure-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write() {
          throw new Error('cannot write this input')
        },
      },
    })
    const input = {
      requestId: 'agent-run-1',
      value: 7,
    }

    expect(planWorkbookAction(model, 'missing', input)).toEqual({
      status: 'failed',
      modelName: 'parameterized-failure-model',
      actionName: 'missing',
      input,
      checks: [],
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model parameterized-failure-model does not define action missing',
        },
      ],
    })

    const result = planWorkbookAction(model, 'write', input)
    expect(result).toEqual({
      status: 'failed',
      modelName: 'parameterized-failure-model',
      actionName: 'write',
      input,
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'cannot write this input',
        },
      ],
    })
    expect(describePlanResult(result)).toEqual(result)
    expect(JSON.parse(JSON.stringify(describePlanResult(result)))).toEqual(describePlanResult(result))
  })

  it('rejects non-JSON-safe action inputs before model code runs', () => {
    const model = defineModel({
      name: 'invalid-action-input-model',
      find() {
        throw new Error('find should not run for invalid input')
      },
      actions: {
        write() {
          throw new Error('action should not run for invalid input')
        },
      },
    })

    expect(planWorkbookAction(model, 'write', { value: Number.NaN })).toEqual({
      status: 'failed',
      modelName: 'invalid-action-input-model',
      actionName: 'write',
      checks: [],
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Action input at input.value must be a finite number',
          path: 'input.value',
          issueCode: 'invalid_action_input',
        },
      ],
    })
    const sparseInput: WorkbookActionInput[] = [1, 2, 3]
    Reflect.deleteProperty(sparseInput, 1)
    expect(planWorkbookAction(model, 'write', sparseInput)).toEqual({
      status: 'failed',
      modelName: 'invalid-action-input-model',
      actionName: 'write',
      checks: [],
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Action input at input[1] must not be a sparse array hole',
          path: 'input[1]',
          issueCode: 'invalid_action_input',
        },
      ],
    })
  })

  it('rejects action metadata mismatches before workbook model code runs', () => {
    let findRan = false
    const model = defineModel({
      name: 'metadata-input-guard-model',
      find(workbook) {
        findRan = true
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      actions: {
        write: {
          input: {
            kind: 'object',
            fields: {
              value: { kind: 'number', required: true },
            },
          },
          run() {
            throw new Error('action should not run for invalid metadata input')
          },
        },
      },
    })

    const result = planWorkbookAction(model, 'write', { value: '12' })

    expect(result).toEqual({
      status: 'failed',
      modelName: 'metadata-input-guard-model',
      actionName: 'write',
      input: {
        value: '12',
      },
      checks: [],
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Action input at input.value must be a number',
          path: 'input.value',
          issueCode: 'wrong_input_type',
        },
      ],
    })
    expect(describePlanResult(result)).toEqual(result)
    expect(findRan).toBe(false)
  })

  it('allows optional action metadata input to be omitted before model code runs', () => {
    let actionInput: WorkbookActionInput | undefined
    const model = defineModel({
      name: 'optional-metadata-input-model',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      actions: {
        write: {
          input: {
            kind: 'object',
            fields: {
              value: { kind: 'number' },
            },
          },
          run({ refs, workbook, input }) {
            actionInput = input
            workbook.writeValue(refs.output, 0)
          },
        },
      },
    })

    const result = planWorkbookAction(model, 'write')

    expect(result.status).toBe('planned')
    expect(actionInput).toBeUndefined()
    if (result.status === 'planned') {
      expect(result.plan).not.toHaveProperty('input')
    }
  })

  it('rejects missing required action metadata input before model code runs', () => {
    let findRan = false
    const model = defineModel({
      name: 'required-metadata-input-model',
      find(workbook) {
        findRan = true
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      actions: {
        write: {
          input: {
            kind: 'object',
            required: true,
            fields: {
              value: { kind: 'number' },
            },
          },
          run() {
            throw new Error('action should not run for missing required input')
          },
        },
      },
    })

    expect(planWorkbookAction(model, 'write')).toEqual({
      status: 'failed',
      modelName: 'required-metadata-input-model',
      actionName: 'write',
      checks: [],
      errors: [
        {
          code: 'invalid_action_input',
          message: 'Action input at input is required',
          path: 'input',
          issueCode: 'missing_required_input',
        },
      ],
    })
    expect(findRan).toBe(false)
  })

  it('verifies manually constructed invalid plan inputs', () => {
    const model = defineModel({
      name: 'invalid-plan-input-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.output, 1)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'write', { value: 1 })

    const brokenPlan = { ...plan }
    Object.defineProperty(brokenPlan, 'input', {
      enumerable: true,
      value: {
        value: new Date('2026-05-21T00:00:00Z'),
      },
    })

    expect(verifyPlan(brokenPlan)).toEqual({
      status: 'invalid',
      modelName: 'invalid-plan-input-model',
      actionName: 'write',
      issues: [
        {
          code: 'invalid_action_input',
          path: 'input.value',
          message: 'Action input at input.value must be a plain JSON object, not Date',
        },
      ],
    })
  })

  it('verifies parameterized models with per-action inputs', () => {
    const model = defineModel({
      name: 'parameterized-model-verification',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write({ refs, workbook, input }) {
          const value = inputObject(input).value
          if (typeof value !== 'number') {
            throw new Error('value input must be numeric')
          }
          workbook.writeValue(refs.output, value)
        },
      },
    })

    expect(verifyModel(model).actions[0]?.planning).toEqual({
      status: 'failed',
      modelName: 'parameterized-model-verification',
      actionName: 'write',
      errors: [
        {
          code: 'action_failed',
          message: 'action input must be an object',
        },
      ],
      checks: [],
    })

    const verification = verifyModel(model, {
      inputs: {
        write: {
          value: 12,
        },
      },
    })

    expect(verification).toEqual({
      status: 'valid',
      modelName: 'parameterized-model-verification',
      actions: [
        {
          actionName: 'write',
          planning: {
            status: 'planned',
            plan: expect.objectContaining({
              modelName: 'parameterized-model-verification',
              actionName: 'write',
              input: {
                value: 12,
              },
            }),
          },
          requirements: {
            modelName: 'parameterized-model-verification',
            actionName: 'write',
            requirements: [
              {
                kind: 'apply',
                capability: 'writeValue',
                commandIndex: 0,
                target: {
                  kind: 'range',
                  id: 'range_Sheet1_B2_B2',
                  label: 'Sheet1!B2',
                  range: {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                  },
                },
                message: 'Apply value write to Sheet1!B2',
              },
            ],
          },
          verification: {
            status: 'valid',
            modelName: 'parameterized-model-verification',
            actionName: 'write',
            issues: [],
          },
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(verification))).toEqual(verification)
  })
})
