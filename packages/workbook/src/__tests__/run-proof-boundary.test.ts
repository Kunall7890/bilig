import { describe, expect, it } from 'vitest'
import {
  defineModel,
  describeRunResult,
  findRange,
  runWorkbookAction,
  type WorkbookCheckResult,
  type WorkbookModel,
  type WorkbookRunAdapter,
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

describe('@bilig/workbook run proof boundary', () => {
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
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned applied with errors',
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
  })

  it('allows generic check verifiers to add JSON-safe proof', async () => {
    const model = proofModel()
    const proof = {
      scannedCells: 24,
      errorCells: [],
      source: 'adapter',
    }

    const result = await runWorkbookAction(model, 'inspect', {
      apply: () => ({ status: 'applied' }),
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
      apply: () => ({ status: 'applied' }),
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
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('strips unsupported verifier fields from check results', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      apply: () => ({ status: 'applied' }),
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
      apply: () => ({ status: 'applied' }),
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
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })
})
