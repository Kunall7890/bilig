import { describe, expect, it, vi } from 'vitest'
import {
  buildWorkbookCommandBundle,
  buildWorkbookActionPlan,
  defineModel,
  describeCommandBundle,
  findRange,
  isWorkbookCommandBundleIssueCode,
  planWorkbookCommand,
  runWorkbookCommandBundle,
  verifyWorkbookCommandBundle,
  workbookCommandBundleIssueCodes,
  type WorkbookActionPlan,
  type WorkbookCommandBundle,
  type WorkbookRunAdapter,
} from '../index.js'

function commandModel() {
  return defineModel({
    name: 'command-model',

    find(workbook) {
      return {
        output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
      }
    },

    checks({ refs, workbook }) {
      return [workbook.check.exists(refs.output)]
    },

    actions: {
      write: {
        input: {
          kind: 'object',
          fields: {
            value: { kind: 'number', required: true },
          },
        },
        run({ refs, workbook, input }) {
          if (typeof input !== 'object' || input === null || Array.isArray(input) || typeof input.value !== 'number') {
            throw new Error('numeric value required')
          }
          workbook.writeValue(refs.output, input.value)
          workbook.check.valueEquals(refs.output, input.value)
        },
      },
    },
  })
}

function first<T>(values: readonly T[]): T {
  const [value] = values
  if (value === undefined) {
    throw new Error('expected at least one value')
  }
  return value
}

describe('@bilig/workbook command bundle api', () => {
  it('exports stable inspectable command bundle issue codes', () => {
    expect(Object.isFrozen(workbookCommandBundleIssueCodes)).toBe(true)
    expect(workbookCommandBundleIssueCodes).toContain('invalid_command_id')
    expect(workbookCommandBundleIssueCodes).toContain('plan_invalid')
    expect(new Set(workbookCommandBundleIssueCodes).size).toBe(workbookCommandBundleIssueCodes.length)
    expect(isWorkbookCommandBundleIssueCode('requirements_mismatch')).toBe(true)
    expect(isWorkbookCommandBundleIssueCode('consumer_domain_error')).toBe(false)
  })

  it('plans a frozen command bundle with deterministic handoff proof', () => {
    const result = planWorkbookCommand(
      commandModel(),
      'write',
      { value: 12 },
      {
        baseRevision: 'rev-1',
        idempotencyKey: 'retry-write-12',
      },
    )

    expect(result.status).toBe('planned')
    if (result.status !== 'planned') {
      throw new Error('expected command planning to succeed')
    }

    expect(Object.isFrozen(result.command)).toBe(true)
    expect(result.command).toMatchObject({
      schemaVersion: 1,
      commandId: expect.stringMatching(/^cmd_[0-9a-f]{16}$/u),
      idempotencyKey: 'retry-write-12',
      baseRevision: 'rev-1',
      modelName: 'command-model',
      actionName: 'write',
      input: { value: 12 },
      verification: { status: 'valid', modelName: 'command-model', actionName: 'write', issues: [] },
      requirements: { modelName: 'command-model', actionName: 'write' },
    })
    expect(result.command.requirements.requirements.map((requirement) => requirement.kind)).toEqual(['apply', 'read', 'verify'])
    expect(verifyWorkbookCommandBundle(result.command)).toEqual({
      status: 'valid',
      commandId: result.command.commandId,
      modelName: 'command-model',
      actionName: 'write',
      issues: [],
    })

    const described = describeCommandBundle(result.command)
    expect(described.plan.refsUsed).toEqual([
      {
        kind: 'range',
        id: 'range_p_Sheet1_p_B2_p_B2',
        label: 'Sheet1!B2',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B2',
        },
      },
    ])
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
  })

  it('keeps command ids deterministic and revision-aware', () => {
    const plan = buildWorkbookActionPlan(commandModel(), 'write', { value: 12 })
    const firstCommand = buildWorkbookCommandBundle(plan, { baseRevision: 1 })
    const retryCommand = buildWorkbookCommandBundle(plan, { baseRevision: 1 })
    const nextRevisionCommand = buildWorkbookCommandBundle(plan, { baseRevision: 2 })
    const callerKeyCommand = buildWorkbookCommandBundle(plan, { baseRevision: 1, idempotencyKey: 'caller-key' })

    expect(retryCommand.commandId).toBe(firstCommand.commandId)
    expect(nextRevisionCommand.commandId).not.toBe(firstCommand.commandId)
    expect(callerKeyCommand.commandId).not.toBe(firstCommand.commandId)
  })

  it('detects tampered requirements, verification, input, and command id', () => {
    const command = buildWorkbookCommandBundle(buildWorkbookActionPlan(commandModel(), 'write', { value: 12 }))
    const tampered: WorkbookCommandBundle = {
      ...command,
      commandId: 'cmd_bad',
      input: { value: 13 },
      requirements: {
        ...command.requirements,
        actionName: 'changed',
      },
      verification: {
        ...command.verification,
        actionName: 'changed',
      },
    }

    expect(verifyWorkbookCommandBundle(tampered).issues.map((issue) => issue.code)).toEqual([
      'input_mismatch',
      'requirements_mismatch',
      'verification_mismatch',
      'invalid_command_id',
    ])
  })

  it('marks command bundles invalid when the embedded plan is invalid', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C3' })
    const invalidPlan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'invalid-command-plan',
      actionName: 'inspect',
      refs: { target },
      refsUsed: [target],
      commands: [],
      ops: [],
      changed: [],
      checks: [
        {
          status: 'passed',
          kind: 'exists',
          target,
          message: 'Sheet1!C3 exists',
        },
      ],
    }

    const command = buildWorkbookCommandBundle(invalidPlan)

    expect(verifyWorkbookCommandBundle(command)).toMatchObject({
      status: 'invalid',
      issues: [
        {
          code: 'plan_invalid',
          path: 'verification.checks[0].status',
          message: 'Sheet1!C3 check exists must start planned before runtime proof',
        },
      ],
    })
  })

  it('runs command bundles and passes the command to each adapter phase', async () => {
    const planned = planWorkbookCommand(commandModel(), 'write', { value: 12 }, { idempotencyKey: 'retry-write' })
    if (planned.status !== 'planned') {
      throw new Error('expected command planning to succeed')
    }
    const command = planned.command
    const adapter: WorkbookRunAdapter<typeof command.plan.refs> = {
      preview: vi.fn((plan, receivedCommand) => {
        expect(receivedCommand).toBe(command)
        return {
          modelName: plan.modelName,
          actionName: plan.actionName,
          requirements: command.requirements.requirements,
          materializedOps: plan.ops,
        }
      }),
      apply: vi.fn((_plan, receivedCommand) => {
        expect(receivedCommand).toBe(command)
        return { status: 'applied', undo: { id: 'undo-command' } }
      }),
      read: vi.fn((targets, _plan, receivedCommand) => {
        expect(receivedCommand).toBe(command)
        return [{ target: first(targets), value: 12 }]
      }),
      verifyChecks: vi.fn((checks, _plan, receivedCommand) => {
        expect(receivedCommand).toBe(command)
        return checks.map((check) =>
          check.kind === 'exists'
            ? {
                ...check,
                status: 'passed',
                proof: { kind: 'runtime', message: 'range exists' },
              }
            : check,
        )
      }),
    }

    await expect(runWorkbookCommandBundle(command, adapter)).resolves.toMatchObject({
      status: 'done',
      undo: { id: 'undo-command' },
      applied: {
        opCount: 1,
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 12 }],
      },
    })
    expect(adapter.preview).toHaveBeenCalledTimes(1)
    expect(adapter.apply).toHaveBeenCalledTimes(1)
    expect(adapter.read).toHaveBeenCalledTimes(1)
    expect(adapter.verifyChecks).toHaveBeenCalledTimes(1)
  })

  it('rejects invalid command bundles before adapter apply', async () => {
    const command = buildWorkbookCommandBundle(buildWorkbookActionPlan(commandModel(), 'write', { value: 12 }))
    const tampered: WorkbookCommandBundle<typeof command.plan.refs> = {
      ...command,
      commandId: 'cmd_bad',
    }
    const apply = vi.fn<WorkbookRunAdapter<typeof command.plan.refs>['apply']>(() => ({ status: 'applied' }))

    await expect(runWorkbookCommandBundle(tampered, { apply })).resolves.toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_command_bundle',
          message: expect.stringContaining('does not match'),
          path: 'commandId',
        },
      ],
      checks: command.plan.checks,
    })
    expect(apply).not.toHaveBeenCalled()
  })
})
