import { describe, expect, it } from 'vitest'
import { createWorkbookAgentCommandBundle } from '@bilig/agent-api'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { applyPendingWorkbookMutationToEngine } from '../worker-runtime-mutation-replay.js'
import type { WorkerEngine } from '../worker-runtime-support.js'

function createMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'doc-1:pending:1',
    localSeq: 1,
    baseRevision: 0,
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
    ...overrides,
  }
}

async function createEngine(): Promise<SpreadsheetEngine & WorkerEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'mutation-replay-test' })
  await engine.ready()
  engine.createSheet('Sheet1')
  return engine
}

describe('applyPendingWorkbookMutationToEngine', () => {
  it('replays valid pending workbook mutations into the projection engine', async () => {
    const engine = await createEngine()

    applyPendingWorkbookMutationToEngine(engine, createMutation())

    expect(engine.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
  })

  it('fails closed for malformed pending workbook mutations instead of silently dropping optimistic state', async () => {
    const engine = await createEngine()
    const mutation = createMutation({
      args: ['Sheet1', 'A1', Number.NaN],
    })

    expect(() => applyPendingWorkbookMutationToEngine(engine, mutation)).toThrow('Invalid pending workbook mutation replay')
    expect(engine.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })
  })

  it('replays agent table command bundles through the projection engine', async () => {
    const engine = await createEngine()
    const bundle = createWorkbookAgentCommandBundle({
      documentId: 'doc-1',
      threadId: 'toolbar',
      turnId: 'turn-1',
      goalText: 'Create table from selection',
      baseRevision: 0,
      context: null,
      commands: [
        {
          kind: 'upsertTable',
          table: {
            name: 'Table1',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Name', 'Amount'],
            columns: [{ name: 'Name' }, { name: 'Amount' }],
            headerRow: true,
            totalsRow: false,
          },
        },
      ],
      now: 100,
    })

    applyPendingWorkbookMutationToEngine(engine, createMutation({ method: 'applyAgentCommandBundle', args: [bundle] }))

    expect(engine.getTable('Table1')).toMatchObject({
      name: 'Table1',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Name', 'Amount'],
    })
  })
})
