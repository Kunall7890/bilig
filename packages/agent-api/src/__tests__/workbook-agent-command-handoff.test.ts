import { checkWorkbookCommandBundle } from '@bilig/workbook'
import { describe, expect, it } from 'vitest'
import { createWorkbookAgentCommandBundle, type WorkbookAgentCommand, type WorkbookAgentCommandBundle } from '../workbook-agent-bundles.js'
import { toWorkbookCommandBundle } from '../workbook-agent-command-handoff.js'

function createBundle(commands: readonly WorkbookAgentCommand[]) {
  return createWorkbookAgentCommandBundle({
    bundleId: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thread-1',
    turnId: 'turn-1',
    goalText: 'Make a safe workbook change',
    baseRevision: 12,
    context: null,
    commands,
    now: 100,
  })
}

function isInputObject(value: unknown): value is { readonly [key: string]: unknown } {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

describe('workbook agent command handoff', () => {
  it('converts agent commands into the generic workbook command bundle contract', () => {
    const bundle = createBundle([
      {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B2',
        values: [
          [1, 2],
          [3, 4],
        ],
      },
    ])

    const handoff = toWorkbookCommandBundle(bundle)
    const check = checkWorkbookCommandBundle(handoff)

    expect(check.status).toBe('valid')
    expect(handoff).toEqual(
      expect.objectContaining({
        id: 'bundle-1',
        targetRevision: 12,
        idempotencyKey: 'bundle-1',
      }),
    )
    expect(handoff.commands).toHaveLength(1)

    const [command] = handoff.commands
    expect(command).toEqual(
      expect.objectContaining({
        id: 'bundle-1:0:writeRange',
        kind: 'request',
        destructive: true,
        touchedRanges: [
          {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'C3',
          },
        ],
      }),
    )
    if (command?.kind !== 'request') {
      throw new Error('Expected request command')
    }
    expect(command.request).toEqual({
      featureId: 'workbook-agent',
      commandId: 'workbookAgent.writeRange',
      category: 'mutation',
      mode: 'applyAndVerify',
      input: {
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B2',
        values: [
          [1, 2],
          [3, 4],
        ],
      },
    })
    if (check.status !== 'valid') {
      throw new Error('Expected valid workbook command bundle')
    }
    expect(check.result.touchedCellCount).toBe(4)
  })

  it('strips undefined optional command fields before public input validation', () => {
    const formatCommand: WorkbookAgentCommand = {
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    Object.defineProperty(formatCommand, 'patch', {
      value: undefined,
      enumerable: true,
      configurable: true,
    })
    Object.defineProperty(formatCommand, 'numberFormat', {
      value: undefined,
      enumerable: true,
      configurable: true,
    })
    const bundle = createBundle([formatCommand])

    const handoff = toWorkbookCommandBundle(bundle)
    const [command] = handoff.commands
    if (command?.kind !== 'request') {
      throw new Error('Expected request command')
    }
    const input = command.request.input
    if (!isInputObject(input)) {
      throw new Error('Expected object input')
    }

    expect(input).toEqual({
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    })
    expect(Object.hasOwn(input, 'patch')).toBe(false)
    expect(Object.hasOwn(input, 'numberFormat')).toBe(false)
  })

  it('rejects agent commands that cannot become canonical workbook ranges', () => {
    const bundle: WorkbookAgentCommandBundle = {
      id: 'bundle-1',
      documentId: 'doc-1',
      threadId: 'thread-1',
      turnId: 'turn-1',
      goalText: 'Clear a range',
      summary: 'Clear a range',
      scope: 'sheet',
      riskClass: 'medium',
      baseRevision: 12,
      createdAtUnixMs: 100,
      context: null,
      commands: [
        {
          kind: 'clearRange',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'not-a-cell',
            endAddress: 'A1',
          },
        },
      ],
      affectedRanges: [],
      estimatedAffectedCells: null,
    }

    expect(() => toWorkbookCommandBundle(bundle)).toThrow()
  })
})
