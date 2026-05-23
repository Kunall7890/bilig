import { describe, expect, it } from 'vitest'
import {
  checkWorkbookCommandRequest,
  checkWorkbookCommandReceipt,
  defineWorkbookFeaturePlugin,
  isWorkbookCommandCategory,
  isWorkbookCommandExecutionMode,
  isWorkbookCommandReceiptStatus,
  isWorkbookCommandRequest,
  isWorkbookCommandReceipt,
  isWorkbookProjectionInterceptorPoint,
  isWorkbookUiContributionSlot,
  normalizeWorkbookCommandRequest,
  normalizeWorkbookCommandReceipt,
  workbookCommandReceiptOpsMatch,
  workbookCommandCategories,
  workbookCommandExecutionModes,
  workbookCommandReceiptStatuses,
  workbookProjectionInterceptorPoints,
  workbookUiContributionSlots,
} from '../index.js'

describe('@bilig/workbook feature api', () => {
  it('exports frozen command vocabulary for agent tools', () => {
    expect(workbookCommandCategories).toEqual(['command', 'operation', 'mutation'])
    expect(workbookCommandExecutionModes).toEqual(['preview', 'apply', 'applyAndVerify'])
    expect(workbookCommandReceiptStatuses).toEqual(['previewed', 'applied', 'rejected', 'noop'])
    expect(workbookProjectionInterceptorPoints).toEqual([
      'cellDisplay',
      'cellStyle',
      'rangeChrome',
      'rowVisibility',
      'beforeCommand',
      'commandMetadata',
    ])
    expect(workbookUiContributionSlots).toEqual(['toolbar', 'sidePanel', 'floatingOverlay', 'status'])
    expect(Object.isFrozen(workbookCommandCategories)).toBe(true)
    expect(Object.isFrozen(workbookCommandExecutionModes)).toBe(true)
    expect(Object.isFrozen(workbookCommandReceiptStatuses)).toBe(true)
    expect(Object.isFrozen(workbookProjectionInterceptorPoints)).toBe(true)
    expect(Object.isFrozen(workbookUiContributionSlots)).toBe(true)

    expect(isWorkbookCommandCategory('command')).toBe(true)
    expect(isWorkbookCommandCategory('other')).toBe(false)
    expect(isWorkbookCommandExecutionMode('applyAndVerify')).toBe(true)
    expect(isWorkbookCommandExecutionMode('later')).toBe(false)
    expect(isWorkbookCommandReceiptStatus('applied')).toBe(true)
    expect(isWorkbookCommandReceiptStatus('done')).toBe(false)
    expect(isWorkbookProjectionInterceptorPoint('rangeChrome')).toBe(true)
    expect(isWorkbookProjectionInterceptorPoint('gridChrome')).toBe(false)
    expect(isWorkbookUiContributionSlot('toolbar')).toBe(true)
    expect(isWorkbookUiContributionSlot('menu')).toBe(false)
  })

  it('checks and normalizes command requests before runtime handoff', () => {
    const request = normalizeWorkbookCommandRequest({
      featureId: 'tables',
      commandId: 'tables.createFromSelection',
      category: 'command',
      mode: 'applyAndVerify',
      input: {
        tableName: 'Sales',
        headerRow: true,
      },
    })

    expect(Object.isFrozen(request)).toBe(true)
    expect(isWorkbookCommandRequest(request)).toBe(true)
    expect(checkWorkbookCommandRequest(request)).toEqual({
      status: 'valid',
      request,
      issues: [],
    })
    expect(request).toEqual({
      featureId: 'tables',
      commandId: 'tables.createFromSelection',
      category: 'command',
      mode: 'applyAndVerify',
      input: {
        headerRow: true,
        tableName: 'Sales',
      },
    })

    expect(checkWorkbookCommandRequest('not-a-request')).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command_request',
          path: 'request',
          message: 'Workbook command request must be an object',
        },
      ],
    })

    expect(
      checkWorkbookCommandRequest({
        featureId: ' tables ',
        commandId: '',
        category: 'bad',
        mode: 'later',
        input: () => undefined,
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command_request',
          path: 'featureId',
          message: 'Workbook command request feature id must not have leading or trailing whitespace',
        },
        {
          code: 'invalid_command_request',
          path: 'commandId',
          message: 'Workbook command request command id cannot be empty',
        },
        {
          code: 'invalid_command_request',
          path: 'category',
          message: 'Workbook command request category is invalid',
        },
        {
          code: 'invalid_command_request',
          path: 'mode',
          message: 'Workbook command request mode is invalid',
        },
        {
          code: 'invalid_command_request',
          path: 'input',
          message: 'Workbook command request input must be JSON-safe',
        },
      ],
    })
    expect(() => normalizeWorkbookCommandRequest({ featureId: 'tables' })).toThrowError(
      'Workbook command request is invalid: Workbook command request command id must be a string',
    )
  })

  it('defines immutable feature plugins with command, projection, and UI contribution metadata', () => {
    const plugin = defineWorkbookFeaturePlugin({
      id: 'tables',
      version: '1.0.0',
      dependsOn: ['core'],
      commands: [
        {
          id: 'tables.createFromSelection',
          featureId: 'tables',
          category: 'command',
          label: 'Create table',
          description: 'Create a table from the selected range',
        },
      ],
      projectionInterceptors: [
        {
          id: 'tables.rangeChrome',
          featureId: 'tables',
          point: 'rangeChrome',
          priority: 20,
        },
      ],
      uiContributions: [
        {
          id: 'tables.toolbar.create',
          featureId: 'tables',
          slot: 'toolbar',
          label: 'Create table',
          order: 10,
        },
      ],
    })

    expect(Object.isFrozen(plugin)).toBe(true)
    expect(Object.isFrozen(plugin.commands)).toBe(true)
    expect(plugin).toMatchObject({
      id: 'tables',
      version: '1.0.0',
      dependsOn: ['core'],
      commands: [
        {
          id: 'tables.createFromSelection',
          featureId: 'tables',
          category: 'command',
          label: 'Create table',
        },
      ],
      projectionInterceptors: [
        {
          id: 'tables.rangeChrome',
          featureId: 'tables',
          point: 'rangeChrome',
          priority: 20,
        },
      ],
      uiContributions: [
        {
          id: 'tables.toolbar.create',
          featureId: 'tables',
          slot: 'toolbar',
          label: 'Create table',
          order: 10,
        },
      ],
    })
  })

  it('rejects feature metadata that would make plugin ownership ambiguous', () => {
    expect(() =>
      defineWorkbookFeaturePlugin({
        id: 'tables',
        version: '1.0.0',
        commands: [
          {
            id: 'tables.createFromSelection',
            featureId: 'filters',
            category: 'command',
            label: 'Create table',
          },
        ],
        projectionInterceptors: [],
        uiContributions: [],
      }),
    ).toThrowError('does not match plugin tables')

    expect(() =>
      defineWorkbookFeaturePlugin({
        id: ' tables ',
        version: '1.0.0',
        commands: [],
        projectionInterceptors: [],
        uiContributions: [],
      }),
    ).toThrowError('must not have leading or trailing whitespace')
  })

  it('normalizes and validates command receipts with preview/apply parity', () => {
    const receipt = normalizeWorkbookCommandReceipt({
      status: 'applied',
      featureId: 'tables',
      commandId: 'tables.createFromSelection',
      category: 'command',
      previewOps: [
        {
          kind: 'upsertTable',
          table: {
            name: 'Sales',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        },
      ],
      appliedOps: [
        {
          kind: 'upsertTable',
          table: {
            name: 'Sales',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        },
      ],
      undo: { id: 'undo-1' },
      changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }],
      proof: { tableName: 'Sales' },
    })

    expect(isWorkbookCommandReceipt(receipt)).toBe(true)
    expect(checkWorkbookCommandReceipt(receipt)).toEqual({
      status: 'valid',
      receipt,
      issues: [],
    })
    expect(workbookCommandReceiptOpsMatch(receipt)).toBe(true)
    expect(Object.isFrozen(receipt.previewOps)).toBe(true)
    expect(Object.isFrozen(receipt.undo)).toBe(true)
    expect(receipt).toMatchObject({
      status: 'applied',
      featureId: 'tables',
      commandId: 'tables.createFromSelection',
      category: 'command',
      undo: { id: 'undo-1' },
      changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }],
      proof: { tableName: 'Sales' },
    })
  })

  it('rejects command receipts that contain invalid ops or ranges', () => {
    expect(checkWorkbookCommandReceipt('not-a-receipt')).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command_receipt',
          path: 'receipt',
          message: 'Workbook command receipt must be an object',
        },
      ],
    })

    expect(
      checkWorkbookCommandReceipt({
        status: 'done',
        featureId: ' tables ',
        commandId: '',
        category: 'bad',
        previewOps: [
          {
            kind: 'notARealOp',
          },
        ],
        appliedOps: 'not-an-array',
        undo: {
          id: ' undo ',
          ops: [
            {
              kind: 'notARealOp',
            },
          ],
        },
        changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1' }],
        proof: () => undefined,
        message: '  ',
        metadata: Number.NaN,
        errors: [42, ' '],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command_receipt',
          path: 'status',
          message: 'Workbook command receipt status is invalid',
        },
        {
          code: 'invalid_command_receipt',
          path: 'featureId',
          message: 'Workbook command receipt feature id must not have leading or trailing whitespace',
        },
        {
          code: 'invalid_command_receipt',
          path: 'commandId',
          message: 'Workbook command receipt command id cannot be empty',
        },
        {
          code: 'invalid_command_receipt',
          path: 'category',
          message: 'Workbook command receipt category is invalid',
        },
        {
          code: 'invalid_command_receipt',
          path: 'previewOps[0]',
          message: 'Workbook command receipt preview op is invalid',
        },
        {
          code: 'invalid_command_receipt',
          path: 'appliedOps',
          message: 'Workbook command receipt applied ops must be an array',
        },
        {
          code: 'invalid_command_receipt',
          path: 'undo.id',
          message: 'Workbook command receipt undo id must not have leading or trailing whitespace',
        },
        {
          code: 'invalid_command_receipt',
          path: 'undo.ops[0]',
          message: 'Workbook command receipt undo op is invalid',
        },
        {
          code: 'invalid_command_receipt',
          path: 'changedRanges[0]',
          message: 'Workbook command receipt changed range is invalid',
        },
        {
          code: 'invalid_command_receipt',
          path: 'proof',
          message: 'Workbook command receipt proof must be JSON-safe',
        },
        {
          code: 'invalid_command_receipt',
          path: 'message',
          message: 'Workbook command receipt message cannot be empty',
        },
        {
          code: 'invalid_command_receipt',
          path: 'metadata',
          message: 'Workbook command receipt metadata must be JSON-safe',
        },
        {
          code: 'invalid_command_receipt',
          path: 'errors[0]',
          message: 'Workbook command receipt error must be a string',
        },
        {
          code: 'invalid_command_receipt',
          path: 'errors[1]',
          message: 'Workbook command receipt error cannot be empty',
        },
      ],
    })

    expect(() =>
      normalizeWorkbookCommandReceipt({
        status: 'applied',
        featureId: 'tables',
        commandId: 'tables.createFromSelection',
        category: 'command',
        previewOps: [
          {
            // @ts-expect-error exercising runtime receipt validation for JS callers
            kind: 'notARealOp',
          },
        ],
      }),
    ).toThrowError('Workbook command receipt is invalid: Workbook command receipt preview op is invalid')

    expect(
      isWorkbookCommandReceipt({
        status: 'applied',
        featureId: 'tables',
        commandId: 'tables.createFromSelection',
        category: 'command',
        changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1' }],
      }),
    ).toBe(false)
  })
})
