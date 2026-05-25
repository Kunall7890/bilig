import { describe, expect, it } from 'vitest'
import {
  defineModel,
  formula,
  prepareWorkbookAction,
  workbookActionCommandDigest,
  workbookPlanId,
  type EngineOp,
  type WorkbookActionPlan,
  type WorkbookRunAdapter,
  type WorkbookRunApplyCommandReceipt,
} from '../index.js'
import { assertWorkbookRunAdapter, checkWorkbookRunAdapter } from '../testing.js'

function rangeRef(label: string, address: string) {
  return {
    kind: 'range' as const,
    id: `range_${label}`,
    label,
    range: {
      sheetName: 'Resolved',
      startAddress: address,
      endAddress: address,
    },
  }
}

function model() {
  return defineModel({
    name: 'testing-adapter-model',
    find(workbook) {
      return {
        input: workbook.findName('input'),
        factor: workbook.findName('factor'),
        result: workbook.findName('result'),
      }
    },
    checks({ refs, workbook }) {
      const expected = formula.multiply(refs.input, refs.factor)
      return [workbook.check.formulaEquals(refs.result, expected)]
    },
    actions: {
      calculate({ refs, workbook }) {
        workbook.writeFormula(refs.result, formula.multiply(refs.input, refs.factor))
      },
    },
  })
}

function receipt<Refs>(plan: WorkbookActionPlan<Refs>, commandIndex: number, op: EngineOp): WorkbookRunApplyCommandReceipt {
  const command = plan.commands[commandIndex]
  if (command?.kind !== 'writeFormula') {
    throw new Error('expected formula command')
  }
  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: [op],
    appliedOps: [op],
    resolvedRefs: {
      target: rangeRef('Resolved!C1', 'C1'),
      inputs: [rangeRef('Resolved!A1', 'A1'), rangeRef('Resolved!B1', 'B1')],
    },
    formulaLabels: command.labels.map((label) => ({ name: label.name, source: label.name })),
  }
}

function passingAdapter(): WorkbookRunAdapter<{ readonly refsUsed: ReturnType<typeof prepare>['plan']['refsUsed'] }> {
  return {
    apply(plan) {
      const op: EngineOp = {
        kind: 'setCellFormula',
        sheetName: 'Resolved',
        address: 'C1',
        formula: plan.commands[0]?.kind === 'writeFormula' ? plan.commands[0].formula : 'input*factor',
      }
      return {
        status: 'applied',
        planId: workbookPlanId(plan),
        baseRevision: 4,
        revision: 5,
        previewOps: [op],
        appliedOps: [op],
        commandReceipts: [receipt(plan, 0, op)],
      }
    },
    read(targets) {
      return targets.map((target) => ({
        target,
        formula: 'input*factor',
      }))
    },
  }
}

function prepare() {
  const prepared = prepareWorkbookAction(model(), 'calculate')
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared fixture')
  }
  return prepared
}

describe('@bilig/workbook testing api', () => {
  it('checks a runtime adapter against a strict transported plan', async () => {
    const prepared = prepare()

    const check = await checkWorkbookRunAdapter(prepared.planData, passingAdapter())

    expect(check.status).toBe('passed')
    if (check.status !== 'passed') {
      throw new Error('adapter did not pass')
    }
    expect(Object.isFrozen(check)).toBe(true)
    expect(check.result.status).toBe('done')
    expect(check.description).toMatchObject({
      status: 'done',
      apply: {
        matched: true,
        baseRevision: 4,
        revision: 5,
        commandReceipts: [expect.objectContaining({ commandKind: 'writeFormula' })],
      },
    })
  })

  it('returns adapter capability issues without mutating', async () => {
    const prepared = prepare()
    const adapter = passingAdapter()

    const check = await checkWorkbookRunAdapter(prepared.planData, { apply: (plan) => adapter.apply?.(plan) })

    expect(check).toEqual({
      status: 'failed',
      errors: [],
      issues: [
        {
          code: 'adapter_missing_capability',
          path: 'adapter.read',
          message: 'Adapter is missing read for read',
        },
      ],
    })
  })

  it('returns transported plan issues without throwing', async () => {
    const invalidPlan = JSON.parse(
      JSON.stringify({
        modelName: 'bad-plan-model',
        actionName: 'calculate',
        refsUsed: 'not-an-array',
        commands: [],
        ops: [],
        changed: [],
        checks: [],
      }),
    )

    const check = await checkWorkbookRunAdapter(invalidPlan, passingAdapter())

    expect(check).toEqual({
      status: 'failed',
      errors: [],
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data is invalid: Workbook plan data refsUsed must be an array',
        },
      ],
    })
  })

  it('returns strict proof failures as boring issues', async () => {
    const prepared = prepare()

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          previewOps: [],
          appliedOps: [],
          commandReceipts: [],
        }
      },
      read() {
        return []
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message: 'Workbook action testing-adapter-model.calculate returned invalid command receipts: expected 1 command receipts, got 0',
      },
    ])
    expect(check.result?.status).toBe('failed')
    expect(check.description?.status).toBe('failed')
  })

  it('rejects symbolic ref receipts that are not bound to concrete ranges', async () => {
    const prepared = prepare()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const op: EngineOp = {
      kind: 'setCellFormula',
      sheetName: 'Resolved',
      address: 'C1',
      formula: command.formula,
    }

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: [op],
          appliedOps: [op],
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: [op],
              appliedOps: [op],
              resolvedRefs: {
                target: {
                  kind: 'name',
                  id: 'name_result',
                  label: 'result',
                  name: 'result',
                },
              },
            },
          ],
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'input*factor',
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-model.calculate returned invalid command receipts: commandReceipts[0].resolvedRefs.target must be concrete range ref data',
      },
    ])
  })

  it('rejects materialized ops that do not match the resolved concrete target', async () => {
    const prepared = prepare()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const op: EngineOp = {
      kind: 'setCellFormula',
      sheetName: 'Resolved',
      address: 'D1',
      formula: command.formula,
    }

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: [op],
          appliedOps: [op],
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: [op],
              appliedOps: [op],
              resolvedRefs: {
                target: rangeRef('Resolved!C1', 'C1'),
                inputs: [rangeRef('Resolved!A1', 'A1'), rangeRef('Resolved!B1', 'B1')],
              },
            },
          ],
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'input*factor',
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-model.calculate returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('throws from the assertion helper with the first issue message', async () => {
    const prepared = prepare()
    const adapter = passingAdapter()
    await expect(assertWorkbookRunAdapter(prepared.planData, { apply: (plan) => adapter.apply?.(plan) })).rejects.toThrow(
      'Adapter is missing read for read',
    )

    await expect(assertWorkbookRunAdapter(prepared.planData, passingAdapter())).resolves.toMatchObject({
      status: 'done',
    })
  })
})
