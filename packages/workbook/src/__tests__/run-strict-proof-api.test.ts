import { describe, expect, it, vi } from 'vitest'
import {
  defineModel,
  runWorkbookAction,
  toWorkbookRefData,
  workbookActionCommandDigest,
  workbookPlanId,
  type findRange,
  type findTable,
  type WorkbookActionPlan,
  type WorkbookCheckResult,
  type WorkbookModel,
  type WorkbookRunApplyCommandReceipt,
} from '../index.js'

function valueModel(): WorkbookModel<{ readonly output: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'strict-value-model',

    find(workbook) {
      return {
        output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
      }
    },

    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.output, 12)
        workbook.check.valueEquals(refs.output, 12)
      },
    },
  })
}

function noCheckModel(): WorkbookModel<{ readonly table: ReturnType<typeof findTable> }> {
  return defineModel({
    name: 'strict-no-check-model',

    find(workbook) {
      return {
        table: workbook.findTable({ name: 'Inputs' }),
      }
    },

    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.table, 'ready')
      },
    },
  })
}

function proofCheckModel(): WorkbookModel<{ readonly result: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'strict-proof-check-model',

    find(workbook) {
      return {
        result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
      }
    },

    checks({ refs, workbook }) {
      return [workbook.check.exists(refs.result)]
    },

    actions: {
      inspect({ refs }) {
        void refs.result
      },
    },
  })
}

function commandReceipt<Refs>(plan: WorkbookActionPlan<Refs>, commandIndex = 0): WorkbookRunApplyCommandReceipt {
  const command = plan.commands[commandIndex]
  if (command === undefined) {
    throw new Error('expected planned command')
  }
  const resolvedRefs: Record<string, unknown> = {}
  if (command.target !== undefined) {
    resolvedRefs['target'] = toWorkbookRefData(command.target)
  }
  if (command.kind === 'writeFormula' && command.inputs.length > 0) {
    resolvedRefs['inputs'] = command.inputs.map((input) => toWorkbookRefData(input))
  }
  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: plan.ops,
    appliedOps: plan.ops,
    ...(Object.keys(resolvedRefs).length > 0 ? { resolvedRefs } : {}),
  }
}

describe('@bilig/workbook strict proof api', () => {
  it('strict mode fails before mutating plans that have no checks', async () => {
    const model = noCheckModel()
    const apply = vi.fn()

    const result = await runWorkbookAction(model, 'write', { apply }, undefined, { strict: true })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Strict workbook runs require at least one check before applying mutating plans',
        },
      ],
      changed: [],
      checks: [],
    })
  })

  it('strict mode fails when apply proof omits workbook revisions', async () => {
    const model = valueModel()
    const read = vi.fn()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => ({
          status: 'applied',
          planId: workbookPlanId(plan),
          previewOps: plan.ops,
          appliedOps: plan.ops,
          commandReceipts: [commandReceipt(plan)],
        }),
        read,
      },
      undefined,
      { strict: true },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'plan_not_verified',
          message: 'Adapter did not bind apply proof to workbook revisions',
        },
      ],
      apply: expect.objectContaining({
        matched: true,
        planId: expect.any(String),
      }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('strict mode fails when resolved ref proof does not match the planned command', async () => {
    const model = valueModel()
    const read = vi.fn()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => ({
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 1,
          revision: 2,
          previewOps: plan.ops,
          appliedOps: plan.ops,
          commandReceipts: [
            {
              ...commandReceipt(plan),
              resolvedRefs: {
                target: {
                  kind: 'range',
                  id: 'wrong',
                  label: 'Sheet1!C3',
                  range: { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'C3' },
                },
              },
            },
          ],
        }),
        read,
      },
      undefined,
      { strict: true },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'apply_not_verified',
          message: 'Adapter resolved ref proof for command 0 must bind planned refs to concrete ranges',
        },
      ],
      apply: expect.objectContaining({
        matched: true,
        planId: expect.any(String),
        baseRevision: 1,
        revision: 2,
      }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('rejects apply proof whose base revision does not match the expected base revision', async () => {
    const model = valueModel()
    const read = vi.fn()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => ({
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 7,
          revision: 8,
          previewOps: plan.ops,
          appliedOps: plan.ops,
          commandReceipts: [commandReceipt(plan)],
        }),
        read,
      },
      undefined,
      { expectedBaseRevision: 9 },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'plan_not_verified',
          message: 'Adapter apply proof base revision 7 did not match expected base revision 9',
        },
      ],
      apply: expect.objectContaining({
        baseRevision: 7,
        revision: 8,
      }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('strict mode fails when a passed check has no proof', async () => {
    const model = proofCheckModel()

    const result = await runWorkbookAction(
      model,
      'inspect',
      {
        verifyChecks: (checks) =>
          checks.map(
            (check): WorkbookCheckResult => ({
              ...check,
              status: 'passed',
            }),
          ),
      },
      undefined,
      { strict: true },
    )

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Sheet1!C2 passed check exists without proof',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'passed', kind: 'exists' })],
    })
  })

  it('can require no unverified apply facts without the strict shortcut', async () => {
    const model = valueModel()
    const read = vi.fn()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: () => ({
          status: 'applied',
        }),
        read,
      },
      undefined,
      { requireNoUnverified: true },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'apply_not_verified',
          message: 'Adapter did not return both previewOps and appliedOps',
        },
      ],
      apply: {
        matched: null,
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
        },
        {
          kind: 'apply',
          message: 'Adapter did not return commandReceipts, so planned commands are not bound to materialized ops',
        },
      ],
    })
  })
})
