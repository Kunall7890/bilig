import { describe, expect, it } from 'vitest'
import {
  defineWorkbookFeaturePlugin,
  isWorkbookCommandReceipt,
  normalizeWorkbookCommandReceipt,
  workbookCommandReceiptOpsMatch,
} from '../index.js'

describe('@bilig/workbook feature api', () => {
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
    expect(workbookCommandReceiptOpsMatch(receipt)).toBe(true)
    expect(Object.isFrozen(receipt.previewOps)).toBe(true)
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
    ).toThrowError('preview op is invalid')

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
