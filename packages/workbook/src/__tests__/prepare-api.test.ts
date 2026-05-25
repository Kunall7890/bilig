import { describe, expect, it } from 'vitest'
import {
  checkPlanData,
  defineModel,
  formula,
  normalizeWorkbookActionInput,
  prepareWorkbookAction,
  runWorkbookPlan,
  toWorkbookRefData,
  workbookActionCommandDigest,
  workbookPlanId,
} from '../index.js'

describe('@bilig/workbook prepare api', () => {
  it('prepares the canonical agent handoff in one generic result', async () => {
    const model = defineModel({
      name: 'prepared-formula-model',
      find(workbook) {
        return {
          input: workbook.findName('input'),
          factor: workbook.findName('factor'),
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.formulaEquals(refs.result, formula.multiply(refs.input, refs.factor))]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.input, refs.factor))
        },
      },
    })

    const prepared = prepareWorkbookAction(model, 'calculate')

    expect(prepared.status).toBe('prepared')
    if (prepared.status !== 'prepared') {
      throw new Error('model action did not prepare')
    }
    expect(Object.isFrozen(prepared)).toBe(true)
    expect(prepared.verification.status).toBe('valid')
    expect(prepared.planId).toBe(workbookPlanId(prepared.planData))
    expect(checkPlanData(prepared.planData).status).toBe('valid')
    expect(prepared.requirements.requirements.filter((requirement) => requirement.kind === 'apply')).toHaveLength(1)
    expect(prepared.requirements.requirements.filter((requirement) => requirement.kind === 'read')).toEqual([
      expect.objectContaining({ target: expect.objectContaining({ kind: 'name', name: 'result' }) }),
    ])

    const run = await runWorkbookPlan(
      prepared.planData,
      {
        apply(plan) {
          const ops = plan.commands.map((command) => {
            if (command.kind !== 'writeFormula') {
              throw new Error(`unexpected command ${command.kind}`)
            }
            return {
              kind: 'setCellFormula' as const,
              sheetName: 'Resolved',
              address: 'C1',
              formula: command.formula,
            }
          })
          return {
            status: 'applied',
            planId: workbookPlanId(plan),
            matched: true,
            baseRevision: 1,
            revision: 2,
            previewOps: ops,
            appliedOps: ops,
            commandReceipts: plan.commands.map((command, commandIndex) => {
              const op = ops[commandIndex]
              if (op === undefined) {
                throw new Error(`prepared plan is missing op ${String(commandIndex)}`)
              }
              const resolvedRefs: Record<string, unknown> = {}
              if (command.target !== undefined) {
                resolvedRefs['target'] = toWorkbookRefData(command.target)
              }
              if (command.kind === 'writeFormula') {
                resolvedRefs['inputs'] = command.inputs.map((input) => toWorkbookRefData(input))
              }
              return {
                commandIndex,
                commandKind: command.kind,
                commandDigest: workbookActionCommandDigest(command),
                previewOps: [op],
                appliedOps: [op],
                resolvedRefs: normalizeWorkbookActionInput(resolvedRefs),
                formulaLabels:
                  command.kind === 'writeFormula' ? command.labels.map((label) => ({ name: label.name, source: label.name })) : [],
              }
            }),
          }
        },
        read(targets) {
          return targets.map((target) => ({
            target,
            formula: 'input*factor',
          }))
        },
      },
      { strict: true },
    )

    expect(run.status).toBe('done')
  })

  it('returns planning failures without throwing', () => {
    const model = defineModel({
      name: 'prepare-missing-action-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })

    const prepared = prepareWorkbookAction(model, 'missing')

    expect(prepared).toEqual({
      status: 'failed',
      planning: expect.objectContaining({
        status: 'failed',
        modelName: 'prepare-missing-action-model',
        actionName: 'missing',
      }),
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model prepare-missing-action-model does not define action missing',
        },
      ],
      issues: [],
    })
  })

  it('returns invalid action-name failures without coercing runtime values', () => {
    const model = defineModel({
      name: 'prepare-action-name-boundary-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })

    let coerced = false
    const actionName = {
      [Symbol.toPrimitive]() {
        coerced = true
        throw new Error('prepare action name coercion should not run')
      },
    }

    const prepared = Reflect.apply(prepareWorkbookAction, undefined, [model, actionName])

    expect(coerced).toBe(false)
    expect(prepared).toEqual({
      status: 'failed',
      planning: {
        status: 'failed',
        modelName: 'unknown-model',
        actionName: '<invalid-action-name>',
        checks: [],
        errors: [
          {
            code: 'invalid_action_name',
            message: 'Workbook action name must be a string',
            path: 'actionName',
            issueCode: 'invalid_action_name',
          },
        ],
      },
      errors: [
        {
          code: 'invalid_action_name',
          message: 'Workbook action name must be a string',
          path: 'actionName',
          issueCode: 'invalid_action_name',
        },
      ],
      issues: [],
    })
    expect(Object.isFrozen(prepared)).toBe(true)
  })
})
