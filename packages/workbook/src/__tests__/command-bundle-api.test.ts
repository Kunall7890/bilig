import { describe, expect, it } from 'vitest'
import {
  checkWorkbookCommandBundle,
  checkWorkbookCommandResult,
  checkWorkbookCommandResultForBundle,
  isWorkbookCommandBundle,
  isWorkbookCommandBundleCommandKind,
  isWorkbookCommandResultForBundle,
  isWorkbookCommandResultStatus,
  normalizeWorkbookCommandBundle,
  normalizeWorkbookCommandResult,
  workbookCommandBundleCommandKinds,
  workbookOpCommandReceipt,
  workbookOpCommandReceiptIdentity,
  workbookCommandResultFor,
  workbookCommandResultForReceipts,
  workbookCommandResultStatuses,
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
    expect(workbookCommandResultStatuses).toEqual(['accepted', 'previewed', 'applied', 'rejected', 'noop'])
    expect(Object.isFrozen(workbookCommandBundleCommandKinds)).toBe(true)
    expect(Object.isFrozen(workbookCommandResultStatuses)).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('request')).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('op')).toBe(true)
    expect(isWorkbookCommandBundleCommandKind('macro')).toBe(false)
    expect(isWorkbookCommandResultStatus('applied')).toBe(true)
    expect(isWorkbookCommandResultStatus('done')).toBe(false)
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

    expect(Object.isFrozen(check)).toBe(true)
    expect(Object.isFrozen(check.issues)).toBe(true)
    expect(Object.isFrozen(check.bundle)).toBe(true)
    expect(Object.isFrozen(check.bundle.commands)).toBe(true)
    expect(Object.isFrozen(check.result)).toBe(true)
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

  it('ignores command bundle envelope scratch fields without invoking getters', () => {
    let getterInvoked = false
    const clean = {
      id: 'bundle-scratch',
      targetRevision: 42,
      idempotencyKey: 'agent-run-scratch',
      commands: [
        {
          id: 'preview-first',
          kind: 'request',
          request: previewRequest,
        },
      ],
    }
    const noisy = { ...clean, agentScratchpad: { ignored: true } }
    Object.defineProperty(noisy, 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandBundle(noisy)).toEqual(checkWorkbookCommandBundle(clean))
    expect(getterInvoked).toBe(false)
  })

  it('ignores command-level scratch fields without invoking getters', () => {
    let getterInvoked = false
    const clean = {
      id: 'bundle-command-scratch',
      targetRevision: 42,
      idempotencyKey: 'agent-run-command-scratch',
      commands: [
        {
          id: 'preview-first',
          kind: 'request',
          request: {
            ...previewRequest,
          },
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    }
    const noisy = {
      ...clean,
      commands: [
        {
          ...clean.commands[0],
          request: {
            ...clean.commands[0].request,
          },
          touchedRanges: [{ ...clean.commands[0].touchedRanges[0] }],
        },
      ],
    }
    Object.defineProperty(noisy.commands[0], 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })
    Object.defineProperty(noisy.commands[0].request, 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })
    Object.defineProperty(noisy.commands[0].touchedRanges[0], 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandBundle(noisy)).toEqual(checkWorkbookCommandBundle(clean))
    expect(getterInvoked).toBe(false)
  })

  it('ignores command result envelope scratch fields without invoking getters', () => {
    let getterInvoked = false
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-result-scratch',
      targetRevision: 7,
      idempotencyKey: 'bundle-result-scratch',
      commands: [
        {
          id: 'bundle-result-scratch:0:inspect',
          kind: 'request',
          request: previewRequest,
        },
      ],
    })
    const clean = workbookCommandResultFor(bundle)
    const noisy = { ...clean, agentScratchpad: { ignored: true } }
    Object.defineProperty(noisy, 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandResult(noisy)).toEqual(checkWorkbookCommandResult(clean))
    expect(normalizeWorkbookCommandResult(noisy)).toEqual(clean)
    expect(getterInvoked).toBe(false)
  })

  it('ignores receipt-level scratch fields without invoking getters', () => {
    let getterInvoked = false
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-receipt-scratch',
      targetRevision: 7,
      idempotencyKey: 'bundle-receipt-scratch',
      commands: [
        {
          id: 'bundle-receipt-scratch:0:inspect',
          kind: 'request',
          request: previewRequest,
        },
      ],
    })
    const clean = workbookCommandResultForReceipts(
      bundle,
      [
        {
          status: 'applied',
          featureId: 'inspect',
          commandId: 'inspect.selection',
          category: 'command',
          proof: {
            inspected: true,
          },
        },
      ],
      { revision: 8 },
    )
    if (clean.status === 'accepted') {
      throw new Error('expected settled result')
    }
    const noisy = {
      ...clean,
      receipts: [{ ...clean.receipts[0] }],
    }
    Object.defineProperty(noisy.receipts[0], 'hiddenScratchpad', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandResult(noisy)).toEqual(checkWorkbookCommandResult(clean))
    expect(normalizeWorkbookCommandResult(noisy)).toEqual(clean)
    expect(getterInvoked).toBe(false)
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

  it('rejects duplicate command ids before runtime handoff', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        commands: [
          {
            id: 'inspect-selection',
            kind: 'request',
            request: previewRequest,
          },
          {
            id: 'inspect-selection',
            kind: 'request',
            request: previewRequest,
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'duplicate_command_id',
          path: 'commands[1].id',
          message: 'Workbook command bundle command id inspect-selection already used by commands[0].id',
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

  it('rejects accessor-backed op proof without invoking getters', () => {
    let getterInvoked = false
    const op = { ...setCellValueOp }
    Object.defineProperty(op, 'extra', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        commands: [
          {
            kind: 'op',
            destructive: true,
            op,
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command',
          path: 'commands[0].op.extra',
          message: 'Workbook command bundle op must contain only data properties',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
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

  it('requires touched ranges for scoped destructive commands', () => {
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'agent-run-1',
        scope: {
          maxTouchedCells: 10,
        },
        commands: [
          {
            kind: 'request',
            destructive: true,
            request: mutationRequest,
          },
          {
            kind: 'op',
            destructive: true,
            op: setCellValueOp,
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_touched_ranges',
          path: 'commands[0].touchedRanges',
          message: 'Scoped destructive workbook command must declare touchedRanges so scope.maxTouchedCells is enforceable',
        },
        {
          code: 'missing_touched_ranges',
          path: 'commands[1].touchedRanges',
          message: 'Scoped destructive workbook command must declare touchedRanges so scope.maxTouchedCells is enforceable',
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

  it('builds a validated applied command result from runtime receipts', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-1',
      targetRevision: 7,
      idempotencyKey: 'bundle-1',
      commands: [
        {
          id: 'bundle-1:0:write',
          kind: 'request',
          destructive: true,
          request: mutationRequest,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })

    const result = workbookCommandResultForReceipts(
      bundle,
      [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          previewOps: [setCellValueOp],
          appliedOps: [
            {
              value: 1,
              address: 'A1',
              sheetName: 'Sheet1',
              kind: 'setCellValue',
            },
          ],
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          proof: {
            bundleCommandId: 'bundle-1:0:write',
          },
        },
      ],
      {
        revision: 8,
        undo: {
          id: 'bundle-1:undo',
          ops: [
            {
              kind: 'clearCell',
              sheetName: 'Sheet1',
              address: 'A1',
            },
          ],
        },
      },
    )

    expect(result).toEqual({
      status: 'applied',
      bundleId: 'bundle-1',
      targetRevision: 7,
      idempotencyKey: 'bundle-1',
      commandCount: 1,
      touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
      touchedCellCount: 1,
      revision: 8,
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          previewOps: [setCellValueOp],
          appliedOps: [setCellValueOp],
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          proof: {
            bundleCommandId: 'bundle-1:0:write',
          },
        },
      ],
      matched: true,
      changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
      undo: {
        id: 'bundle-1:undo',
        ops: [
          {
            kind: 'clearCell',
            sheetName: 'Sheet1',
            address: 'A1',
          },
        ],
      },
    })
    expect(Object.isFrozen(result)).toBe(true)
    const resultCheck = checkWorkbookCommandResult(result)
    expect(resultCheck).toEqual({
      status: 'valid',
      result,
      issues: [],
    })
    expect(Object.isFrozen(resultCheck)).toBe(true)
    expect(Object.isFrozen(resultCheck.issues)).toBe(true)

    const bundleResultCheck = checkWorkbookCommandResultForBundle(bundle, result)
    expect(bundleResultCheck).toEqual({
      status: 'valid',
      result,
      issues: [],
    })
    expect(Object.isFrozen(bundleResultCheck)).toBe(true)
    expect(Object.isFrozen(bundleResultCheck.issues)).toBe(true)
    expect(isWorkbookCommandResultForBundle(bundle, result)).toBe(true)
    expect(normalizeWorkbookCommandResult(result)).toEqual(result)
  })

  it('rejects accepted command results that carry settled proof fields', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-accepted',
      targetRevision: 7,
      idempotencyKey: 'bundle-accepted',
      commands: [
        {
          id: 'bundle-accepted:0:inspect',
          kind: 'request',
          request: previewRequest,
        },
      ],
    })
    const accepted = workbookCommandResultFor(bundle)

    expect(
      checkWorkbookCommandResult({
        ...accepted,
        receipts: [],
        matched: null,
        changedRanges: [],
        revision: 8,
        undo: {
          id: 'bundle-accepted:undo',
        },
        errors: [],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command_result',
          path: 'receipts',
          message: 'Accepted workbook command result must not include receipts',
        },
        {
          code: 'invalid_command_result',
          path: 'matched',
          message: 'Accepted workbook command result must not include matched',
        },
        {
          code: 'invalid_command_result',
          path: 'changedRanges',
          message: 'Accepted workbook command result must not include changedRanges',
        },
        {
          code: 'invalid_command_result',
          path: 'revision',
          message: 'Accepted workbook command result must not include revision',
        },
        {
          code: 'invalid_command_result',
          path: 'undo',
          message: 'Accepted workbook command result must not include undo',
        },
        {
          code: 'invalid_command_result',
          path: 'errors',
          message: 'Accepted workbook command result must not include errors',
        },
      ],
    })
    expect(() =>
      normalizeWorkbookCommandResult({
        ...accepted,
        receipts: [],
      }),
    ).toThrow('Workbook command result is invalid: Accepted workbook command result must not include receipts')
  })

  it('rejects command results that are not bound to the bundle', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-1',
      targetRevision: 7,
      idempotencyKey: 'bundle-1',
      commands: [
        {
          id: 'bundle-1:0:write',
          kind: 'request',
          destructive: true,
          request: mutationRequest,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })
    const result = workbookCommandResultForReceipts(
      bundle,
      [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
      { revision: 8 },
    )

    expect(
      checkWorkbookCommandResultForBundle(bundle, {
        ...result,
        idempotencyKey: 'other-run',
        receipts: [
          {
            status: 'applied',
            featureId: 'cells',
            commandId: 'cells.other',
            category: 'mutation',
            changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'bundle_result_mismatch',
          path: 'idempotencyKey',
          message: 'Workbook command result idempotencyKey does not match bundle',
        },
        {
          code: 'receipt_command_mismatch',
          path: 'receipts[0]',
          message: 'Workbook command result receipt 0 does not match command request',
        },
      ],
    })
  })

  it('rejects command results whose summary fields do not match receipts', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-1',
      targetRevision: 7,
      idempotencyKey: 'bundle-1',
      commands: [
        {
          id: 'bundle-1:0:write',
          kind: 'request',
          destructive: true,
          request: mutationRequest,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })
    const result = workbookCommandResultForReceipts(
      bundle,
      [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          previewOps: [setCellValueOp],
          appliedOps: [setCellValueOp],
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
      { revision: 8 },
    )

    expect(
      checkWorkbookCommandResultForBundle(bundle, {
        ...result,
        status: 'rejected',
        matched: false,
        changedRanges: [],
        errors: ['manually supplied error'],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'receipt_result_mismatch',
          path: 'status',
          message: 'Workbook command result status does not match receipts',
        },
        {
          code: 'receipt_result_mismatch',
          path: 'matched',
          message: 'Workbook command result matched does not match receipts',
        },
        {
          code: 'receipt_result_mismatch',
          path: 'changedRanges',
          message: 'Workbook command result changedRanges do not match receipts',
        },
        {
          code: 'receipt_result_mismatch',
          path: 'errors',
          message: 'Workbook command result errors do not match receipts',
        },
      ],
    })
  })

  it('rejects command result changed ranges outside declared command scope', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-scope',
      targetRevision: 7,
      idempotencyKey: 'bundle-scope',
      commands: [
        {
          id: 'bundle-scope:0:write',
          kind: 'request',
          destructive: true,
          request: mutationRequest,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })
    const receipt = {
      status: 'applied' as const,
      featureId: 'cells',
      commandId: 'cells.setValue',
      category: 'mutation' as const,
      changedRanges: [{ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' }],
    }

    expect(() => workbookCommandResultForReceipts(bundle, [receipt], { revision: 8 })).toThrow(
      'Workbook command result is invalid: Workbook command result receipt 0 changed range Sheet1!B1 is outside commands[0].touchedRanges',
    )

    expect(
      checkWorkbookCommandResultForBundle(bundle, {
        ...workbookCommandResultFor(bundle),
        status: 'applied',
        revision: 8,
        receipts: [receipt],
        matched: null,
        changedRanges: [{ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' }],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'changed_range_out_of_scope',
          path: 'receipts[0].changedRanges[0]',
          message: 'Workbook command result receipt 0 changed range Sheet1!B1 is outside commands[0].touchedRanges',
        },
      ],
    })
  })

  it('requires applied command results to carry a final revision when checked against a bundle', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-1',
      targetRevision: 7,
      idempotencyKey: 'bundle-1',
      commands: [
        {
          id: 'bundle-1:0:write',
          kind: 'request',
          destructive: true,
          request: mutationRequest,
        },
      ],
    })
    const result = workbookCommandResultForReceipts(bundle, [
      {
        status: 'applied',
        featureId: 'cells',
        commandId: 'cells.setValue',
        category: 'mutation',
        changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
      },
    ])

    expect(checkWorkbookCommandResultForBundle(bundle, result)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'revision_mismatch',
          path: 'revision',
          message: 'Applied workbook command result must include a revision',
        },
      ],
    })
  })

  it('builds a rejected command result with inspectable errors', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-2',
      targetRevision: 7,
      idempotencyKey: 'bundle-2',
      commands: [
        {
          kind: 'request',
          destructive: true,
          request: mutationRequest,
        },
      ],
    })

    expect(
      workbookCommandResultForReceipts(bundle, [
        {
          status: 'rejected',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          message: 'Range is protected',
        },
      ]),
    ).toEqual({
      status: 'rejected',
      bundleId: 'bundle-2',
      targetRevision: 7,
      idempotencyKey: 'bundle-2',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
      receipts: [
        {
          status: 'rejected',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          message: 'Range is protected',
        },
      ],
      matched: null,
      changedRanges: [],
      errors: ['Range is protected'],
    })
  })

  it('builds deterministic receipts for low-level op commands', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-op',
      targetRevision: 7,
      idempotencyKey: 'bundle-op',
      commands: [
        {
          id: 'op-write-a1',
          kind: 'op',
          destructive: true,
          op: setCellValueOp,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })
    const [command] = bundle.commands
    if (command === undefined) {
      throw new Error('expected command')
    }

    expect(workbookOpCommandReceiptIdentity(command, 0)).toEqual({
      featureId: 'workbook-op',
      commandId: 'op-write-a1',
      category: 'operation',
    })

    const receipt = workbookOpCommandReceipt(command, 0, {
      status: 'applied',
      previewOps: [setCellValueOp],
      appliedOps: [setCellValueOp],
      changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
    })
    const result = workbookCommandResultForReceipts(bundle, [receipt], { revision: 8 })

    expect(checkWorkbookCommandResultForBundle(bundle, result)).toEqual({
      status: 'valid',
      result,
      issues: [],
    })
  })

  it('rejects command result receipts that do not match request commands', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-3',
      targetRevision: 7,
      idempotencyKey: 'bundle-3',
      commands: [
        {
          kind: 'request',
          destructive: true,
          request: mutationRequest,
        },
      ],
    })

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        {
          status: 'applied',
          featureId: 'other',
          commandId: 'cells.setValue',
          category: 'mutation',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ]),
    ).toThrow('Workbook command result is invalid: receipts[0] does not match commands[0].request')
  })

  it('rejects command result receipts that do not match op commands', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-op',
      targetRevision: 7,
      idempotencyKey: 'bundle-op',
      commands: [
        {
          id: 'op-write-a1',
          kind: 'op',
          destructive: true,
          op: setCellValueOp,
        },
      ],
    })

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        {
          status: 'applied',
          featureId: 'other',
          commandId: 'op-write-a1',
          category: 'operation',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ]),
    ).toThrow('Workbook command result is invalid: receipts[0] does not match commands[0].op')
  })

  it('rejects uninspectable command result data without invoking getters', () => {
    let getterInvoked = false
    const result = {
      status: 'applied',
      targetRevision: 7,
      idempotencyKey: 'bundle-4',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
        },
      ],
      matched: null,
      changedRanges: [],
    }
    Object.defineProperty(result.receipts[0], 'proof', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandResult(result)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_receipt',
          path: 'receipts[0].proof',
          message: 'Workbook command receipt proof must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })
})
