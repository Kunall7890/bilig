import { describe, expect, it } from 'vitest'
import { buildWorkbookActionPlan, defineModel, findRange, verifyPlan, type WorkbookActionPlan } from '../index.js'

describe('@bilig/workbook plan refs api', () => {
  it('freezes refs containers in planned handoff data', () => {
    const model = defineModel({
      name: 'frozen-refs-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs', headers: ['Amount'] })
        const amount = table.column('Amount')
        const result = workbook.findRange({ sheetName: 'Sheet1', address: 'D2' })
        return {
          groups: [
            {
              table,
              amount,
            },
          ],
          result,
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeValue(refs.result, 10)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')

    expect(Object.isFrozen(plan.refs)).toBe(true)
    expect(Object.isFrozen(plan.refs.groups)).toBe(true)
    expect(Object.isFrozen(plan.refs.groups[0])).toBe(true)
    expect(Object.isFrozen(plan.refs.groups[0]?.table)).toBe(true)
    expect(Object.isFrozen(plan.refs.groups[0]?.amount)).toBe(true)
    expect(Object.isFrozen(plan.refs.result)).toBe(true)
    expect(plan.refsUsed).toEqual([plan.refs.groups[0]?.table, plan.refs.groups[0]?.amount, plan.refs.result])
  })

  it('does not execute ref container accessors while planning or verifying', () => {
    type Refs = {
      readonly result: ReturnType<typeof findRange>
      readonly hidden?: unknown
    }

    const model = defineModel({
      name: 'accessor-refs-model',

      find(workbook): Refs {
        const refs: Refs = {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'D2' }),
        }
        Object.defineProperty(refs, 'hidden', {
          enumerable: true,
          get() {
            throw new Error('ref accessor should not run')
          },
        })
        return refs
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 10)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'write')

    expect(plan.refsUsed).toEqual([plan.refs.result])
    expect(verifyPlan(plan)).toEqual({
      status: 'valid',
      modelName: 'accessor-refs-model',
      actionName: 'write',
      issues: [],
    })
  })

  it('returns invalid verification for accessor-backed plan data without invoking getters', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })

    let refsUsedGetterInvoked = false
    const topLevelPlan = {
      modelName: 'verify-accessor-plan',
      actionName: 'write',
      refs: { target },
      refsUsed: [target],
      commands: [],
      ops: [],
      changed: [],
      checks: [],
    } satisfies WorkbookActionPlan<{ readonly target: typeof target }>
    Object.defineProperty(topLevelPlan, 'refsUsed', {
      enumerable: true,
      get() {
        refsUsedGetterInvoked = true
        throw new Error('refsUsed getter should not run')
      },
    })

    expect(verifyPlan(topLevelPlan)).toEqual({
      status: 'invalid',
      modelName: 'verify-accessor-plan',
      actionName: 'write',
      issues: [
        {
          code: 'invalid_plan',
          path: 'plan',
          message: 'Workbook plan is invalid: Workbook verification plan.refsUsed must be a data property',
        },
      ],
    })
    expect(refsUsedGetterInvoked).toBe(false)

    let commandGetterInvoked = false
    const command = {
      kind: 'clear',
      target,
    } satisfies WorkbookActionPlan<{ readonly target: typeof target }>['commands'][number]
    Object.defineProperty(command, 'kind', {
      enumerable: true,
      get() {
        commandGetterInvoked = true
        throw new Error('command getter should not run')
      },
    })
    const nestedPlan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'verify-nested-accessor-plan',
      actionName: 'write',
      refs: { target },
      refsUsed: [target],
      commands: [command],
      ops: [],
      changed: [],
      checks: [],
    }

    expect(verifyPlan(nestedPlan)).toEqual({
      status: 'invalid',
      modelName: 'verify-nested-accessor-plan',
      actionName: 'write',
      issues: [
        {
          code: 'invalid_plan',
          path: 'plan',
          message: 'Workbook plan is invalid: Workbook description command.kind must be a data property',
        },
      ],
    })
    expect(commandGetterInvoked).toBe(false)
  })

  it('rejects refsUsed entries that are not discoverable from refs', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const forged = findRange({ sheetName: 'Sheet1', address: 'Z9' })
    const plan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'forged-refs-plan',
      actionName: 'write',
      refs: { target },
      refsUsed: [target, forged],
      commands: [],
      ops: [],
      changed: [],
      checks: [],
    }

    expect(verifyPlan(plan)).toEqual({
      status: 'invalid',
      modelName: 'forged-refs-plan',
      actionName: 'write',
      issues: [
        {
          code: 'ref_not_in_refs',
          path: 'refsUsed[1]',
          ref: {
            kind: 'range',
            id: 'range_Sheet1_Z9_Z9',
            label: 'Sheet1!Z9',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'Z9',
              endAddress: 'Z9',
            },
          },
          message: 'Sheet1!Z9 appears in refsUsed but is not discoverable from refs',
        },
      ],
    })
  })
})
