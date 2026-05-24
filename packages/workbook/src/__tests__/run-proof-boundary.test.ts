import { describe, expect, it } from 'vitest'
import {
  defineModel,
  describeRunResult,
  findRange,
  runWorkbookAction,
  verifyWorkbookReadbacks,
  type WorkbookActionPlan,
  type WorkbookCheckResult,
  type WorkbookModel,
  type WorkbookRunAdapter,
  type WorkbookRunApplyResult,
  workbookActionCommandDigest,
  workbookPlanId,
} from '../index.js'

function valueModel(): WorkbookModel<{ readonly output: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'run-value-model',

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

function proofModel(): WorkbookModel<{ readonly result: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'run-proof-model',

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

function first<T>(values: readonly T[]): T {
  const [value] = values
  if (value === undefined) {
    throw new Error('expected at least one value')
  }
  return value
}

function applied<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRunApplyResult {
  return {
    status: 'applied',
    previewOps: plan.ops,
    appliedOps: plan.ops,
    commandReceipts: plan.commands.map((command, commandIndex) => ({
      commandIndex,
      commandKind: command.kind,
      commandDigest: workbookActionCommandDigest(command),
      previewOps: plan.ops,
      appliedOps: plan.ops,
    })),
  }
}

function commandReceipt<Refs>(plan: WorkbookActionPlan<Refs>, commandIndex = 0) {
  const command = plan.commands[commandIndex]
  if (command === undefined) {
    throw new Error('expected planned command')
  }
  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: plan.ops,
    appliedOps: plan.ops,
  }
}

function invalidProofCheck(checkResult: WorkbookCheckResult): WorkbookCheckResult {
  const verified = {
    ...checkResult,
    status: 'passed' as const,
  }
  Object.defineProperty(verified, 'proof', {
    enumerable: true,
    value: { when: new Date(0) },
  })
  return verified
}

function withUnsupportedField(checkResult: WorkbookCheckResult): WorkbookCheckResult {
  const verified = {
    ...checkResult,
    status: 'passed' as const,
    proof: { source: 'adapter' },
  }
  Object.defineProperty(verified, 'adapterObject', {
    enumerable: true,
    value: { leaked: true },
  })
  return verified
}

function sparseArray(length = 1): unknown[] {
  const value: unknown[] = []
  value.length = length
  return value
}

function accessorArray(get: () => unknown): unknown[] {
  const value = sparseArray()
  Object.defineProperty(value, '0', {
    enumerable: true,
    get,
  })
  return value
}

describe('@bilig/workbook run proof boundary', () => {
  it('rejects apply proof that is not JSON-safe', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'applied',
        proof: { when: new Date(0) },
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message:
            'Workbook action run-value-model.write returned invalid apply proof: Action input at input.when must be a plain JSON object, not Date',
        },
      ],
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects applied apply results that include errors', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'applied',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected after apply',
          },
        ],
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned applied with errors',
        },
      ],
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('keeps apply proof bound to the exact plan id when requested', async () => {
    const model = valueModel()
    let actualPlanId: string | undefined

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => {
          actualPlanId = workbookPlanId(plan)
          return {
            status: 'applied',
            planId: actualPlanId,
            baseRevision: 7,
            revision: 8,
            previewOps: plan.ops,
            appliedOps: plan.ops,
            commandReceipts: [commandReceipt(plan)],
          }
        },
        read: (targets) => [{ target: first(targets), value: 12 }],
      },
      undefined,
      { requirePlanId: true },
    )

    expect(result.status).toBe('done')
    expect(result.apply).toEqual(
      expect.objectContaining({
        planId: actualPlanId,
        baseRevision: 7,
        revision: 8,
        matched: true,
      }),
    )
  })

  it('rejects apply proof with a stale plan id', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'applied',
        planId: `${workbookPlanId(plan)}-stale`,
        previewOps: plan.ops,
        appliedOps: plan.ops,
        commandReceipts: [commandReceipt(plan)],
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned a plan id that does not match the executed plan',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('can require runtime apply proof to include command receipts and a plan id', async () => {
    const model = valueModel()

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
        read: (targets) => [{ target: first(targets), value: 12 }],
      },
      undefined,
      { requireApplyProof: true, requirePlanId: true },
    )

    expect(result.status).toBe('done')
    expect(result.apply).toEqual(
      expect.objectContaining({
        matched: true,
        planId: expect.any(String),
        commandReceipts: [
          expect.objectContaining({
            commandIndex: 0,
            commandKind: 'writeValue',
            commandDigest: expect.stringMatching(/^bilig-command-v1:/),
          }),
        ],
      }),
    )
  })

  it('uses strict mode as the single agent-safe proof option', async () => {
    const model = valueModel()

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
        read: (targets) => [{ target: first(targets), value: 12 }],
      },
      undefined,
      { strict: true },
    )

    expect(result.status).toBe('done')
    expect(result.apply).toEqual(
      expect.objectContaining({
        matched: true,
        planId: expect.any(String),
        commandReceipts: [expect.objectContaining({ commandKind: 'writeValue' })],
      }),
    )
  })

  it('strict mode fails closed when plan id proof is missing', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => ({
          status: 'applied',
          previewOps: plan.ops,
          appliedOps: plan.ops,
          commandReceipts: [commandReceipt(plan)],
        }),
        read: (targets) => [{ target: first(targets), value: 12 }],
      },
      undefined,
      { strict: true },
    )

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'plan_not_verified',
          message: 'Adapter did not bind apply proof to a plan id',
        },
      ],
      apply: expect.objectContaining({
        matched: true,
        commandReceipts: [expect.objectContaining({ commandKind: 'writeValue' })],
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

  it('requires command receipts when apply proof is required', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: (plan) => ({
          status: 'applied',
          planId: workbookPlanId(plan),
          previewOps: plan.ops,
          appliedOps: plan.ops,
        }),
        read: (targets) => [{ target: first(targets), value: 12 }],
      },
      undefined,
      { requireApplyProof: true },
    )

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'apply_not_verified',
          message: 'Adapter did not bind planned commands to materialized ops',
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
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return commandReceipts, so planned commands are not bound to materialized ops',
        },
      ],
    })
  })

  it('rejects command receipts with stale command digests', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'applied',
        previewOps: plan.ops,
        appliedOps: plan.ops,
        commandReceipts: [
          {
            ...commandReceipt(plan),
            commandDigest: 'stale-command',
          },
        ],
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message:
            'Workbook action run-value-model.write returned invalid command receipts: commandReceipts[0].commandDigest does not match the planned command',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('uses stable command digests for equivalent command data', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'B2' })
    const left = {
      kind: 'writeValue' as const,
      target,
      value: 12,
    }
    const right = {
      value: 12,
      target,
      kind: 'writeValue' as const,
    }

    expect(workbookActionCommandDigest(left)).toBe(workbookActionCommandDigest(right))
  })

  it('rejects command receipts whose ops do not match apply ops', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'applied',
        previewOps: plan.ops,
        appliedOps: plan.ops,
        commandReceipts: [
          {
            ...commandReceipt(plan),
            previewOps: [],
            appliedOps: [],
          },
        ],
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message:
            'Workbook action run-value-model.write returned invalid command receipts: commandReceipts previewOps do not match apply previewOps',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('rejects accessor-backed apply ops without invoking getters', async () => {
    const model = valueModel()
    const previewOp = {
      sheetName: 'Sheet1',
      address: 'B2',
      value: 12,
    }
    let getterInvoked = false
    Object.defineProperty(previewOp, 'kind', {
      enumerable: false,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'applied',
        // @ts-expect-error exercising JS adapters that bypass the op type
        previewOps: [previewOp],
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned invalid preview ops',
        },
      ],
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects sparse apply evidence arrays as uninspectable runtime proof', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'applied',
        // @ts-expect-error exercising JS adapters that bypass the op type
        previewOps: sparseArray(),
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned invalid preview ops',
        },
      ],
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects accessor-backed apply error arrays without invoking getters', async () => {
    const model = valueModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        // @ts-expect-error exercising JS adapters that bypass the error type
        errors: accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        }),
      }),
      read: (targets) => [{ target: first(targets), value: 12 }],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned invalid apply errors',
        },
      ],
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('allows generic check verifiers to add JSON-safe proof', async () => {
    const model = proofModel()
    const proof = {
      scannedCells: 24,
      errorCells: [],
      source: 'adapter',
    }

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) =>
        checks.map((checkResult) => ({
          ...checkResult,
          status: 'passed',
          proof,
        })),
    })

    expect(result).toEqual({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'exists',
          proof,
        }),
      ],
    })
    expect(describeRunResult(result)).toEqual({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'exists',
          proof,
        }),
      ],
    })
  })

  it('rejects verifier proof that is not JSON-safe', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) => checks.map(invalidProofCheck),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned invalid proof at index 0: Action input at input.when must be a plain JSON object, not Date',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('rejects accessor-backed verifier proof without invoking getters', async () => {
    const model = proofModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) =>
        checks.map((checkResult) => {
          const verified = {
            ...checkResult,
            status: 'passed' as const,
          }
          Object.defineProperty(verified, 'proof', {
            enumerable: true,
            get() {
              getterInvoked = true
              throw new Error('getter must not run')
            },
          })
          return verified
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned invalid proof at index 0: Workbook check result proof must be a data property',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed verifier check arrays without invoking getters', async () => {
    const model = proofModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: () =>
        // @ts-expect-error exercising JS adapters that bypass the check array type
        accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned an invalid check at index 0',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
    expect(getterInvoked).toBe(false)
  })

  it('strips unsupported verifier fields from check results', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map(withUnsupportedField),
    })

    expect(result.status).toBe('done')
    const check = first(result.checks)
    expect(check).toEqual(
      expect.objectContaining({
        status: 'passed',
        kind: 'exists',
        proof: { source: 'adapter' },
      }),
    )
    expect(Object.hasOwn(check, 'adapterObject')).toBe(false)
  })

  it('rejects readbacks that were not requested by checks', async () => {
    const model = valueModel()
    const unexpected = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const read: Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read'] = (targets) => [
      {
        target: first(targets),
        value: 12,
      },
      {
        target: unexpected,
        value: 99,
      },
    ]

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read,
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_unexpected',
          message: 'Sheet1!D2 was returned by readback but was not requested',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects duplicate readbacks for the same target', async () => {
    const model = valueModel()
    const read: Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read'] = (targets) => {
      const target = first(targets)
      return [
        {
          target,
          value: 12,
        },
        {
          target,
          value: 12,
        },
      ]
    }

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read,
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_duplicate',
          message: 'Sheet1!B2 was returned by readback more than once',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects accessor-backed readback arrays without invoking getters', async () => {
    const model = valueModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: () =>
        // @ts-expect-error exercising JS adapters that bypass the readback array type
        accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0] is invalid',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed readback values without invoking getters', async () => {
    const model = valueModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: (targets) => {
        const readback = {
          target: first(targets),
        }
        Object.defineProperty(readback, 'value', {
          enumerable: true,
          get() {
            getterInvoked = true
            throw new Error('getter must not run')
          },
        })
        return [readback]
      },
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0].value must be a data property',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed public readback arrays without invoking getters', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'B2' })
    let getterInvoked = false

    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'valueEquals',
          target,
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      // @ts-expect-error exercising JS callers that bypass the readback array type
      accessorArray(() => {
        getterInvoked = true
        throw new Error('getter must not run')
      }),
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target,
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      issues: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0] is invalid',
        },
      ],
    })
    expect(Object.isFrozen(verification)).toBe(true)
    expect(Object.isFrozen(verification.checks)).toBe(true)
    expect(Object.isFrozen(verification.checks[0])).toBe(true)
    expect(Object.isFrozen(verification.issues)).toBe(true)
    expect(Object.isFrozen(verification.issues[0])).toBe(true)
    expect(getterInvoked).toBe(false)
  })
})
