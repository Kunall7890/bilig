import { describe, expect, it } from 'vitest'
import { defineModel, describeRunResult, runWorkbookAction } from '../index.js'

describe('@bilig/workbook runtime proof api', () => {
  it('allows generic check verifiers to attach runtime proof without changing the check contract', async () => {
    const model = defineModel({
      name: 'run-check-runtime-proof-model',

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
      apply: () => ({ status: 'applied' }),
      verifyChecks: (checks) =>
        checks.map((checkResult) => ({
          ...checkResult,
          status: 'passed',
          proof: {
            kind: 'runtime',
            message: 'Runtime verifier confirmed the invariant',
            data: { checkedBy: 'adapter', rule: checkResult.kind },
          },
        })),
    })

    expect(result).toEqual({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'consumerInvariant',
          message: 'Consumer invariant holds',
          proof: {
            kind: 'runtime',
            message: 'Runtime verifier confirmed the invariant',
            data: { checkedBy: 'adapter', rule: 'consumerInvariant' },
          },
        }),
      ],
    })
    expect(describeRunResult(result)).toMatchObject({
      checks: [
        {
          proof: {
            kind: 'runtime',
            message: 'Runtime verifier confirmed the invariant',
            data: { checkedBy: 'adapter', rule: 'consumerInvariant' },
          },
        },
      ],
    })
  })
})
