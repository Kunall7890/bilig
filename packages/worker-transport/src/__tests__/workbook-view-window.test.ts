import { describe, expect, it } from 'vitest'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import {
  WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION,
  buildUnavailableWorkbookViewWindow,
  buildWorkbookViewWindowFromViewportPatch,
  type WorkbookViewWindowSubscription,
} from '../workbook-view-window.js'

const METRICS: RecalcMetrics = {
  batchId: 12,
  changedInputCount: 1,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

const REQUEST: WorkbookViewWindowSubscription = {
  sheetId: 7,
  sheetName: 'Sheet1',
  sheetOrdinal: 2,
  rowStart: 0,
  rowEnd: 3,
  colStart: 0,
  colEnd: 2,
}

describe('WorkbookViewWindow contract', () => {
  it('promotes a viewport patch into a sheet-id keyed authoritative window', () => {
    const window = buildWorkbookViewWindowFromViewportPatch({
      request: REQUEST,
      sheet: {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 2,
      },
      renderAck: {
        status: 'presented',
        batchId: 12,
        renderRevision: 4,
        proofSignature: 'scene:4',
      },
      patch: {
        version: 3,
        authoritativeRevision: 9,
        full: true,
        freezeRows: 1,
        freezeCols: 1,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 0,
          rowEnd: 3,
          colStart: 0,
          colEnd: 2,
        },
        metrics: METRICS,
        styles: [{ id: 'style-live', fill: { backgroundColor: '#34a853' } }],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              value: { tag: ValueTag.String, stringId: 1, value: 'ready' },
              flags: 0,
              version: 5,
            },
            displayText: 'ready',
            copyText: 'ready',
            editorText: 'ready',
            formatId: 0,
            styleId: 'style-live',
          },
        ],
        columns: [{ index: 1, size: 104, hidden: false }],
        rows: [{ index: 1, size: 22, hidden: false }],
        merges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }],
      },
    })

    expect(window).toMatchObject({
      schemaVersion: WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION,
      status: 'ready',
      request: { sheetId: 7, sheetName: 'Sheet1' },
      sheet: { sheetId: 7, sheetName: 'Sheet1', sheetOrdinal: 2 },
      authoritativeRevision: 9,
      patchVersion: 3,
      renderAck: {
        status: 'presented',
        batchId: 12,
        renderRevision: 4,
        proofSignature: 'scene:4',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 3,
        colStart: 0,
        colEnd: 2,
      },
      cells: [
        {
          row: 1,
          col: 1,
          address: 'B2',
          displayText: 'ready',
          editorText: 'ready',
          styleId: 'style-live',
        },
      ],
      merges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }],
    })
  })

  it('returns a typed unavailable window instead of a blank grid success', () => {
    const window = buildUnavailableWorkbookViewWindow({
      status: 'missing-sheet',
      reason: 'sheet-id-not-found',
      request: { ...REQUEST, sheetId: 99, sheetName: 'Deleted' },
      authoritativeRevision: 11,
    })

    expect(window).toMatchObject({
      schemaVersion: WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION,
      status: 'missing-sheet',
      reason: 'sheet-id-not-found',
      sheet: null,
      authoritativeRevision: 11,
      patchVersion: null,
      cells: [],
      styles: [],
      renderAck: {
        status: 'rejected',
        reason: 'sheet-id-not-found',
      },
    })
  })
})
