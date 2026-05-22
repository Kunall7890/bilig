import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describePlan,
  describePlanResult,
  isWorkbookActionInput,
  isWorkbookActionInputDescription,
  isWorkbookActionInputDescriptionKind,
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
        },
      ],
    })
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
          path: 'input',
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
