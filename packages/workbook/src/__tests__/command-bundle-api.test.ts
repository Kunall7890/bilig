import { describe, expect, it } from 'vitest'
import {
  checkWorkbookCommandBundle,
  isWorkbookCommandBundle,
  isWorkbookCommandBundleCommandKind,
  normalizeWorkbookCommandBundle,
  workbookCommandBundleCommandKinds,
  workbookCommandResultFor,
} from '../index.js'

const previewRequest = {
  featureId: 'inspect',
  commandId: 'inspect.selection',
  category: 'command',
  mode: 'preview',
} as const

const mutationRequest = {
  featureId: 'cells',
  commandId: 'cells.setValue',
  category: 'mutation',
  mode: 'applyAndVerify',
} as const

const setCellValueOp = {
  kind: 'setCellValue',
  sheetName: 'Sheet1',
  address: 'A1',
  value: 1,
} as const

describe('@bilig/workbook command bundle api', () => {
  it('exports frozen command bundle vocabulary', () => {
    expect(workbookCommandBundleCommandKinds).toEqual(['request', 'op'])
    expect(Object.isFrozen(workbookCommandBundleCommandKinds)).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('request')).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('op')).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('macro')).toBe(false)
  })

  it('checks and normalizes a portable command bundle result', () => {
    const check = checkWorkbookCommandBundle({
      id: 'bundle-1',
      targetRevision: 42,
      idempotencyKey: 'agent-run-1',
      scope: {
        maxTouchedCells: 8,
      },
      commands: [
        {
          id: 'preview-first',
          kind: 'request',
          request: previewRequest,
          touchedRanges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'b2',
              endAddress: 'C3',
            },
          ],
        },
        {
          id: 'write-second',
          kind: 'op',
          destructive: true,
          op: setCellValueOp,
          touchedRanges: [
            {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'A1',
            },
          ],
        },
      ],
    })

    expect(check.status).toBe('valid')
    if (check.status !== 'valid') {
      throw new Error('expected valid command bundle')
    }

    expect(Object.isFrozen(check.bundle)).toBe(true)
    expect(Object.isFrozen(check.bundle.commands)).toBe(true)
    expect(isWorkbookCommandBundle(check.bundle)).toBe(true)
    expect(check.bundle.commands.map((command) => command.id)).toEqual(['preview-first', 'write-second'])
    expect(check.bundle.commands).toMatchObject([
      {
        kind: 'request',
        request: {
          featureId: 'inspect',
          commandId: 'inspect.selection',
        },
        touchedRanges: [
          {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'C3',
          },
        ],
      },
      {
        kind: 'op',
        destructive: true,
        touchedRanges: [
          {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
        ],
      },
    ])
    expect(check.result).toEqual({
      status: 'accepted',
      bundleId: 'bundle-1',
      targetRevision: 42,
      idempotencyKey: 'agent-run-1',
      commandCount: 2,
      touchedRanges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C3',
        },
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
      ],
      touchedCellCount: 5,
    })
    expect(workbookCommandResultFor(check.bundle)).toEqual(check.result)
  })

  it('rejects bundles without revision, idempotency, or commands', () => {
    expect(checkWorkbookCommandBundle({})).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_target_revision',
          path: 'targetRevision',
          message: 'Workbook command bundle targetRevision is required',
        },
        {
          code: 'missing_idempotency_key',
          path: 'idempotencyKey',
          message: 'Workbook command bundle idempotencyKey is required',
        },
        {
          code: 'missing_commands',
          path: 'commands',
          message: 'Workbook command bundle commands must be a non-empty array',
        },
      ],
    })
  })

  it('rejects unknown command kinds before runtime handoff', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        commands: [
          {
            kind: 'macro',
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'unknown_command_kind',
          path: 'commands[0].kind',
          message: 'Workbook command bundle command kind is unknown',
        },
      ],
    })
  })

  it('rejects invalid touched ranges', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        commands: [
          {
            kind: 'request',
            request: previewRequest,
            touchedRanges: [
              {
                sheetName: 'Sheet1',
                startAddress: 'C2',
                endAddress: 'A1',
              },
            ],
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_range',
          path: 'commands[0].touchedRanges[0]',
          message: 'Workbook command bundle commands[0].touchedRanges[0] endAddress must not be before startAddress',
        },
      ],
    })
  })

  it('requires explicit destructive confirmation for mutations and ops', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        commands: [
          {
            kind: 'request',
            request: mutationRequest,
            touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          },
          {
            kind: 'op',
            op: setCellValueOp,
            touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'destructive_not_confirmed',
          path: 'commands[0].destructive',
          message: 'Workbook command bundle request mutates the workbook and must set destructive: true',
        },
        {
          code: 'destructive_not_confirmed',
          path: 'commands[1].destructive',
          message: 'Workbook command bundle op mutates the workbook and must set destructive: true',
        },
      ],
    })
  })

  it('rejects command bundles that exceed maxTouchedCells', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        scope: {
          maxTouchedCells: 3,
        },
        commands: [
          {
            kind: 'request',
            request: previewRequest,
            touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }],
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'too_many_touched_cells',
          path: 'scope.maxTouchedCells',
          message: 'Workbook command bundle touches 4 cells, exceeding scope.maxTouchedCells 3',
        },
      ],
    })
  })

  it('normalizes without requiring core runtime execution', () => {
    const bundle = normalizeWorkbookCommandBundle({
      targetRevision: 7,
      idempotencyKey: 'agent-run-2',
      commands: [
        {
          kind: 'request',
          request: previewRequest,
        },
      ],
    })

    expect(bundle).toEqual({
      targetRevision: 7,
      idempotencyKey: 'agent-run-2',
      commands: [
        {
          kind: 'request',
          request: previewRequest,
        },
      ],
    })
    expect(workbookCommandResultFor(bundle)).toEqual({
      status: 'accepted',
      targetRevision: 7,
      idempotencyKey: 'agent-run-2',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
    })
    expect(() => normalizeWorkbookCommandBundle({ commands: [] })).toThrowError(
      'Workbook command bundle is invalid: Workbook command bundle targetRevision is required',
    )
  })
})
