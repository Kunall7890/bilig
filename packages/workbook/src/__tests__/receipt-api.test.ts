import { describe, expect, it } from 'vitest'
import {
  describeRunResult,
  defineModel,
  findRange,
  isWorkbookReceiptProofKind,
  planWorkbookCommand,
  runWorkbookAction,
  runWorkbookCommandBundle,
  workbookReceiptProofKinds,
  type WorkbookRunAdapter,
} from '../index.js'

function receiptModel() {
  return defineModel({
    name: 'receipt-model',

    find(workbook) {
      return {
        output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
      }
    },

    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.output, 'done')
        workbook.check.valueEquals(refs.output, 'done')
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

describe('@bilig/workbook receipt api', () => {
  it('exports stable inspectable receipt proof kinds', () => {
    expect(Object.isFrozen(workbookReceiptProofKinds)).toBe(true)
    expect(workbookReceiptProofKinds).toEqual([
      'preview',
      'apply',
      'authoritativeReadback',
      'renderedReadback',
      'semanticReadback',
      'recalculation',
      'undo',
      'check',
      'custom',
    ])
    expect(isWorkbookReceiptProofKind('renderedReadback')).toBe(true)
    expect(isWorkbookReceiptProofKind('domainSpecificProof')).toBe(false)
  })

  it('builds a command-aware receipt with revisions, rendered diffs, proof, and undo', async () => {
    const planned = planWorkbookCommand(receiptModel(), 'write', undefined, {
      baseRevision: 'r1',
      idempotencyKey: 'write-once',
    })
    if (planned.status !== 'planned') {
      throw new Error('expected command planning to succeed')
    }
    const command = planned.command
    const output = findRange({ sheetName: 'Sheet1', address: 'B2' })

    const result = await runWorkbookCommandBundle(command, {
      preview: (plan) => ({
        modelName: plan.modelName,
        actionName: plan.actionName,
        requirements: command.requirements.requirements,
        materializedOps: plan.ops,
      }),
      apply: () => ({
        status: 'applied',
        undo: { id: 'undo-r2' },
        receipt: {
          appliedRevision: 'r2',
          calculatedRevision: 'r2',
          renderedRevision: 'frame-r2',
          rendered: {
            revision: 'frame-r2',
            diffs: [
              {
                kind: 'cell',
                target: output,
                message: 'Rendered Sheet1!B2 changed.',
              },
            ],
            message: 'Rendered frame includes target change.',
          },
          proof: [
            {
              kind: 'renderedReadback',
              status: 'passed',
              revision: 'frame-r2',
              target: output,
              message: 'Visible rendered target matches authoritative value.',
            },
            {
              kind: 'undo',
              status: 'passed',
              revision: 'r2',
              message: 'Undo metadata was captured.',
            },
          ],
        },
      }),
      read: (targets) => [{ target: first(targets), value: 'done' }],
    })

    expect(result).toMatchObject({
      status: 'done',
      undo: { id: 'undo-r2' },
      receipt: {
        commandId: command.commandId,
        idempotencyKey: 'write-once',
        modelName: 'receipt-model',
        actionName: 'write',
        baseRevision: 'r1',
        appliedRevision: 'r2',
        calculatedRevision: 'r2',
        renderedRevision: 'frame-r2',
        previewed: true,
        applied: true,
        verified: true,
        checkCount: 1,
        passedCheckCount: 1,
        failedCheckCount: 0,
        unverifiedCheckCount: 0,
        undo: { id: 'undo-r2' },
      },
    })
    expect(result.status === 'done' ? result.receipt?.proof.map((proof) => proof.kind) : []).toEqual([
      'preview',
      'apply',
      'authoritativeReadback',
      'check',
      'renderedReadback',
      'undo',
    ])
    const described = describeRunResult(result)
    expect(described).toMatchObject({
      status: 'done',
      receipt: {
        rendered: {
          diffs: [
            {
              kind: 'cell',
              target: {
                kind: 'range',
                label: 'Sheet1!B2',
              },
              message: 'Rendered Sheet1!B2 changed.',
            },
          ],
        },
      },
    })
    expect(described.status === 'done' ? described.receipt?.proof : []).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'renderedReadback',
          target: expect.objectContaining({
            kind: 'range',
            label: 'Sheet1!B2',
          }),
        }),
      ]),
    )
  })

  it('rejects malformed runtime receipt proof before reporting success', async () => {
    const adapter = {
      apply: () => ({
        status: 'applied',
        receipt: {
          proof: [
            {
              kind: 'notAProof',
              status: 'passed',
              message: 'bad',
            },
          ],
        },
      }),
    } satisfies WorkbookRunAdapter

    await expect(runWorkbookAction(receiptModel(), 'write', adapter)).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_runtime_result',
          path: 'apply.receipt.proof[0].kind',
        },
      ],
    })
  })
})
