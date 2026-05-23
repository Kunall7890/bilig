import { describe, expect, it, vi } from 'vitest'
import {
  describeRunResult,
  defineModel,
  findRange,
  formula,
  isWorkbookRunErrorCode,
  runWorkbookAction,
  runWorkbookPlan,
  verifyWorkbookReadbacks,
  workbookRunErrorCodes,
  type WorkbookActionPlan,
  type WorkbookModel,
  type WorkbookRunAdapter,
  type WorkbookRunApplyResult,
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
  }
}

describe('@bilig/workbook run api', () => {
  it('exports stable inspectable run error codes', () => {
    expect(Object.isFrozen(workbookRunErrorCodes)).toBe(true)
    expect(workbookRunErrorCodes).toContain('action_not_found')
    expect(workbookRunErrorCodes).toContain('invalid_action_input')
    expect(workbookRunErrorCodes).toContain('ref_not_in_refs')
    expect(workbookRunErrorCodes).toContain('formula_input_not_resolved')
    expect(workbookRunErrorCodes).toContain('apply_not_verified')
    expect(workbookRunErrorCodes).toContain('apply_mismatch')
    expect(workbookRunErrorCodes).toContain('readback_missing')
    expect(workbookRunErrorCodes).toContain('adapter_missing_capability')
    expect(workbookRunErrorCodes).toContain('runtime_rejected')
    expect(new Set(workbookRunErrorCodes).size).toBe(workbookRunErrorCodes.length)
    expect(isWorkbookRunErrorCode('check_not_verified')).toBe(true)
    expect(isWorkbookRunErrorCode('custom_runtime_error')).toBe(false)
  })

  it('plans, verifies, applies, reads back, and returns done for value checks', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>((plan) => ({
      status: 'applied',
      previewOps: plan.ops,
      appliedOps: plan.ops,
      undo: { id: 'undo-1' },
    }))
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      {
        target: first(targets),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(model, 'write', { apply, read })

    expect(apply).toHaveBeenCalledTimes(1)
    expect(read).toHaveBeenCalledTimes(1)
    expect(read.mock.calls[0]?.[0]).toEqual([expect.objectContaining({ label: 'Sheet1!B2' })])
    expect(result).toMatchObject({
      status: 'done',
      apply: {
        matched: true,
        previewOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 12,
          },
        ],
        appliedOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 12,
          },
        ],
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        {
          status: 'passed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
          proof: {
            source: 'readback',
            value: 12,
          },
        },
      ],
      undo: { id: 'undo-1' },
    })
    expect(describeRunResult(result).checks[0]?.proof).toEqual({
      source: 'readback',
      value: 12,
    })
  })

  it('reports unverified apply proof when an adapter does not return preview and applied ops', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'applied' }),
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'done',
      apply: {
        matched: null,
      },
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
        },
      ],
    })
  })

  it('can require apply proof before readback and check verification', async () => {
    const model = valueModel()
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>(() => [
      {
        target: findRange({ sheetName: 'Sheet1', address: 'B2' }),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: () => ({ status: 'applied' }),
        read,
      },
      undefined,
      { requireApplyProof: true },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toMatchObject({
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
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
        },
      ],
    })
  })

  it('fails when an adapter applies ops that do not match its preview', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'applied',
        previewOps: plan.ops,
        appliedOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 13,
          },
        ],
      }),
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'apply_mismatch',
          message: 'Adapter applied ops do not match its preview ops',
        },
      ],
      apply: {
        matched: false,
      },
    })
  })

  it('describes successful run results without leaking ref helper functions', async () => {
    const model = defineModel({
      name: 'run-description-model',

      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs', headers: ['Amount'] }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },

      actions: {
        inspect({ refs }) {
          void refs.table
        },
      },
    })

    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({
      status: 'applied',
      undo: {
        id: 'undo-1',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
      },
    }))

    const result = await runWorkbookAction(model, 'inspect', {
      apply,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'passed' })),
    })
    const described = describeRunResult(result)

    expect(apply).not.toHaveBeenCalled()
    expect(described).toEqual({
      status: 'done',
      changed: [],
      checks: [
        {
          status: 'passed',
          kind: 'exists',
          target: {
            kind: 'table',
            id: 'table_Inputs_Amount',
            label: 'Inputs',
            name: 'Inputs',
            headers: ['Amount'],
          },
          message: 'Inputs exists',
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
  })

  it('runs readback-only plans without requiring an apply adapter', async () => {
    const model = defineModel({
      name: 'run-readback-only-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      actions: {
        inspect({ refs, workbook }) {
          workbook.check.valueEquals(refs.result, 12)
        },
      },
    })
    const read = vi.fn<Required<WorkbookRunAdapter<{ result: ReturnType<typeof findRange> }>>['read']>((targets) => [
      {
        target: first(targets),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(model, 'inspect', { read })

    expect(read).toHaveBeenCalledTimes(1)
    expect(result).toMatchObject({
      status: 'done',
      changed: [],
      checks: [
        {
          status: 'passed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!C2' }),
          message: 'Sheet1!C2 equals 12',
          proof: {
            source: 'readback',
            value: 12,
          },
        },
      ],
    })
    expect(result).not.toHaveProperty('apply')
  })

  it('passes formula readback checks with exact normalized formula text', async () => {
    const model = defineModel({
      name: 'run-formula-model',

      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          rate: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const expected = formula.multiply(refs.amount, refs.rate)
          workbook.writeFormula(refs.result, expected)
          workbook.check.formulaEquals(refs.result, expected)
        },
      },
    })
    const source = '(Sheet1!A2)*(Sheet1!B2)'

    const result = await runWorkbookAction(model, 'calculate', {
      apply: applied,
      read: (targets) => [
        {
          target: first(targets),
          formula: source,
        },
      ],
    })

    expect(result.status).toBe('done')
    expect(result.checks).toEqual([
      {
        status: 'passed',
        kind: 'formulaEquals',
        target: expect.objectContaining({ label: 'Sheet1!C2' }),
        message: `Sheet1!C2 formula equals ${source}`,
        expectation: {
          kind: 'formulaEquals',
          formula: source,
          inputs: [expect.objectContaining({ label: 'Sheet1!A2' }), expect.objectContaining({ label: 'Sheet1!B2' })],
          labels: [
            { name: 'Sheet1!A2', ref: expect.objectContaining({ label: 'Sheet1!A2' }) },
            { name: 'Sheet1!B2', ref: expect.objectContaining({ label: 'Sheet1!B2' }) },
          ],
        },
        proof: {
          source: 'readback',
          formula: source,
        },
      },
    ])
  })

  it('fails before apply when non-readback checks require a missing verifier', async () => {
    const model = defineModel({
      name: 'run-legacy-check-model',

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

    const apply = vi.fn<Required<WorkbookRunAdapter<{ result: ReturnType<typeof findRange> }>>['apply']>(applied)

    const result = await runWorkbookAction(model, 'inspect', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'adapter_missing_capability',
          message: 'Adapter is missing verifyChecks for verifyCheck',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('does not let apply results drop planned checks', async () => {
    const model = defineModel({
      name: 'run-apply-proof-boundary-model',

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

    const result = await runWorkbookAction(model, 'inspect', {
      apply: () => ({
        status: 'applied',
        checks: [],
      }),
      verifyChecks: (checks) => checks,
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Sheet1!C2 did not verify check exists: Sheet1!C2 exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('fails when the generic check verifier leaves checks planned', async () => {
    const model = defineModel({
      name: 'run-unverified-check-model',

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

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks,
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Sheet1!C2 did not verify check exists: Sheet1!C2 exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('passes non-readback and custom checks through the generic check verifier', async () => {
    const model = defineModel({
      name: 'run-check-proof-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs', headers: ['Amount', 'Rate'] })
        return {
          table,
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [
          workbook.check.exists(refs.table),
          workbook.check.noFormulaErrors(refs.result),
          workbook.check.custom({
            kind: 'consumerInvariant',
            target: refs.result,
            refs: [refs.table],
            message: 'Consumer invariant holds',
          }),
        ]
      },

      actions: {
        inspect({ refs }) {
          void refs.table
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'passed' })),
    })

    expect(result).toMatchObject({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({ status: 'passed', kind: 'exists', message: 'Inputs exists' }),
        expect.objectContaining({ status: 'passed', kind: 'noFormulaErrors', message: 'Sheet1!C2 has no formula errors' }),
        expect.objectContaining({ status: 'passed', kind: 'consumerInvariant', message: 'Consumer invariant holds' }),
      ],
    })
  })

  it('returns failed when the generic check verifier marks a check failed', async () => {
    const model = defineModel({
      name: 'run-check-failure-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [
          workbook.check.custom({
            kind: 'consumerInvariant',
            target: refs.result,
            message: 'Consumer invariant holds',
          }),
        ]
      },

      actions: {
        inspect({ refs }) {
          void refs.result
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'failed' })),
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_failed',
          message: 'Sheet1!C2 failed check consumerInvariant: Consumer invariant holds',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'failed',
          kind: 'consumerInvariant',
          message: 'Consumer invariant holds',
        }),
      ],
    })
  })

  it('rejects malformed generic check verifier output', async () => {
    const model = defineModel({
      name: 'run-malformed-check-proof-model',

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

    await expect(
      runWorkbookAction(model, 'inspect', {
        apply: applied,
        verifyChecks: () => [],
      }),
    ).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned 0 checks for 1 planned checks',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })

    await expect(
      runWorkbookAction(model, 'inspect', {
        apply: applied,
        verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, message: 'Changed message' })),
      }),
    ).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier changed the check contract at index 0 for exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('rejects in-place generic check verifier contract mutations', async () => {
    const model = defineModel({
      name: 'run-mutating-check-proof-model',

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

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks(checks) {
        Object.defineProperty(first(checks), 'message', { value: 'Changed message' })
        return checks
      },
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier changed the check contract at index 0 for exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('returns failed when the generic check verifier throws', async () => {
    const model = defineModel({
      name: 'run-throwing-check-proof-model',

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

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks() {
        throw new Error('check backend unavailable')
      },
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_verification_failed',
          message: 'check backend unavailable',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('does not apply when action planning fails', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'missing', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model run-value-model does not define action missing',
        },
      ],
      checks: [],
    })
  })

  it('does not apply when static plan verification fails', async () => {
    const hidden = findRange({ sheetName: 'Sheet1', address: 'Z9' })
    const model = defineModel({
      name: 'invalid-run-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.raw('Sheet1!Z9', { inputs: [hidden] }))
        },
      },
    })
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'calculate', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'formula_input_not_resolved',
          message: 'Sheet1!Z9 is used as a formula input but is missing from refsUsed',
        },
      ],
      checks: [],
    })
  })

  it('does not apply when a plan contains already-proved checks', async () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookPlan(
      {
        modelName: 'pre-proved-run-plan',
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
            message: 'Sheet1!C2 exists',
          },
        ],
      },
      { apply },
    )

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_status_not_planned',
          message: 'Sheet1!C2 check exists must start planned before runtime proof',
        },
      ],
      checks: [
        {
          status: 'passed',
          kind: 'exists',
          target,
          message: 'Sheet1!C2 exists',
        },
      ],
    })
  })

  it('returns failed when the adapter apply step fails', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected the plan',
            path: 'adapter.apply',
            issueCode: 'runtime_refusal',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected the plan',
          path: 'adapter.apply',
          issueCode: 'runtime_refusal',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('does not report changed when failed apply proof says no ops were applied', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'failed',
        previewOps: plan.ops,
        appliedOps: [],
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected before writing',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected before writing',
        },
      ],
      apply: {
        matched: false,
        appliedOps: [],
      },
      changed: [],
    })
  })

  it('preserves changed proof when failed apply reports actual applied ops', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'failed',
        previewOps: plan.ops,
        appliedOps: plan.ops,
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime failed after writing',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime failed after writing',
        },
      ],
      apply: {
        matched: true,
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
    })
  })

  it('describes failed run results as JSON-safe errors and checks', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected the plan',
            path: 'adapter.apply',
            issueCode: 'runtime_refusal',
          },
        ],
      }),
      read: () => [],
    })
    const described = describeRunResult(result)

    expect(described).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected the plan',
          path: 'adapter.apply',
          issueCode: 'runtime_refusal',
        },
      ],
      apply: {
        matched: null,
      },
      changed: [],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
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
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
  })

  it('fails before apply when expected readback requires a missing reader', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>(applied)

    const result = await runWorkbookAction(model, 'write', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'adapter_missing_capability',
          message: 'Adapter is missing read for read',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('returns failed when value readback mismatches', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        ...applied(plan),
        undo: { id: 'undo-mismatch' },
      }),
      read: (targets) => [
        {
          target: first(targets),
          value: 13,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'value_mismatch',
          message: 'Sheet1!B2 expected value 12 but read 13',
        },
      ],
      apply: {
        matched: true,
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      undo: { id: 'undo-mismatch' },
    })
  })

  it('returns failed when formula readback mismatches', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
          },
        },
      ],
      [{ target, formula: '=A2+B2' }],
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [
        {
          status: 'failed',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
          },
        },
      ],
      issues: [
        {
          code: 'formula_mismatch',
          check: expect.objectContaining({ kind: 'formulaEquals' }),
          target,
          expected: 'A2+B2',
          actual: '=A2+B2',
          message: 'Sheet1!C2 expected formula A2+B2 but read =A2+B2',
        },
      ],
    })
  })

  it('supports async apply and async read adapters', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: async (plan) => applied(plan),
      read: async (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(planned.status).toBe('done')
  })

  it('runs an already planned action without model access', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'failed' }),
    })

    expect(planned.status).toBe('failed')
    const planResult = await runWorkbookPlan(
      {
        modelName: 'empty-run-plan',
        actionName: 'noop',
        refs: {},
        refsUsed: [],
        commands: [],
        ops: [],
        changed: [],
        checks: [],
      },
      {
        apply: applied,
      },
    )

    expect(planResult).toMatchObject({
      status: 'done',
      changed: [],
      checks: [],
    })
  })
})
