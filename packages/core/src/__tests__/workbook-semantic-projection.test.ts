import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  normalizeWorkbookSnapshotForSemanticComparison,
  diffWorkbookSemanticSnapshots,
  projectWorkbookSemanticSnapshot,
  workbookSemanticSnapshotsEqual,
} from '../semantics/index.js'

const baseSnapshot: WorkbookSnapshot = {
  version: 1,
  workbook: {
    name: 'semantic-fixture',
  },
  sheets: [
    {
      name: 'Sheet1',
      order: 0,
      cells: [],
    },
  ],
}

describe('workbook semantic projection', () => {
  it('normalizes metadata ordering and equivalent range coverage for engine comparisons', () => {
    const snapshot: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [
            { id: 'style-b', font: { bold: true } },
            { id: 'style-a', fill: { backgroundColor: '#ffffff' } },
          ],
          definedNames: [
            { name: 'Totals', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'C2' } },
            { name: 'Inputs', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            rows: [
              { id: 'row-2', index: 2, size: 24 },
              { id: 'row-1', index: 1, size: 18 },
            ],
            styleRanges: [
              { range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' }, styleId: 'style-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, styleId: 'style-a' },
            ],
          },
          cells: [
            { address: 'B1', value: 2 },
            { address: 'A1', value: 1 },
          ],
        },
      ],
    }

    const normalized = normalizeWorkbookSnapshotForSemanticComparison(snapshot)

    expect(normalized.workbook.metadata?.styles?.map((style) => style.id)).toEqual(['style-a', 'style-b'])
    expect(normalized.workbook.metadata?.definedNames?.map((definedName) => definedName.name)).toEqual(['Inputs', 'Totals'])
    expect(normalized.sheets[0]?.metadata?.rows).toEqual([
      { index: 1, size: 18 },
      { index: 2, size: 24 },
    ])
    expect(normalized.sheets[0]?.metadata?.styleRanges).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' }, styleId: 'style-a' },
    ])
  })

  it('normalizes rich workbook and sheet metadata deterministically', () => {
    const snapshot: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          properties: [
            { key: 'subject', value: 'model' },
            { key: 'author', value: 'bilig' },
          ],
          styles: [
            { id: 'style-z', font: { italic: true } },
            { id: 'style-a', fill: { backgroundColor: '#fef3c7' } },
          ],
          formats: [
            { id: 'fmt-z', code: '0.000', kind: 'number' },
            { id: 'fmt-a', code: '$#,##0.00', kind: 'currency' },
          ],
          tables: [
            {
              name: 'TableZ',
              sheetName: 'Sheet1',
              startAddress: 'D1',
              endAddress: 'E3',
              columnNames: ['D', 'E'],
              headerRow: true,
              totalsRow: false,
            },
            {
              name: 'TableA',
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'B3',
              columnNames: ['A', 'B'],
              headerRow: true,
              totalsRow: true,
            },
          ],
          pivots: [
            {
              name: 'PivotB',
              sheetName: 'Sheet1',
              address: 'H1',
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Amount', summarizeBy: 'sum' }],
              rows: 5,
              cols: 3,
            },
            {
              name: 'PivotA',
              sheetName: 'Sheet1',
              address: 'G1',
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Amount', summarizeBy: 'count' }],
              rows: 4,
              cols: 2,
            },
          ],
          charts: [
            {
              id: 'chart-z',
              sheetName: 'Sheet1',
              address: 'K1',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
              chartType: 'line',
              rows: 8,
              cols: 5,
            },
            {
              id: 'chart-a',
              sheetName: 'Sheet1',
              address: 'J1',
              source: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'D4' },
              chartType: 'bar',
              rows: 6,
              cols: 4,
            },
          ],
          images: [
            { id: 'image-z', sheetName: 'Sheet1', address: 'M1', sourceUrl: 'https://example.test/z.png', rows: 2, cols: 2 },
            { id: 'image-a', sheetName: 'Sheet1', address: 'L1', sourceUrl: 'https://example.test/a.png', rows: 2, cols: 2 },
          ],
          shapes: [
            { id: 'shape-z', sheetName: 'Sheet1', address: 'O1', shapeType: 'rectangle', rows: 2, cols: 2 },
            { id: 'shape-a', sheetName: 'Sheet1', address: 'N1', shapeType: 'ellipse', rows: 2, cols: 2 },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            rows: [
              { id: 'row-4', index: 4, size: null },
              { id: 'row-1', index: 1, size: 18 },
            ],
            columns: [
              { id: 'col-4', index: 4, size: 128 },
              { id: 'col-1', index: 1, size: null },
            ],
            rowMetadata: [
              { start: 5, count: 2, hidden: true },
              { start: 2, count: 1, size: 24 },
            ],
            columnMetadata: [
              { start: 3, count: 2, hidden: true },
              { start: 1, count: 1, size: 96 },
            ],
            styleRanges: [
              { range: { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D2' }, styleId: 'style-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, styleId: 'style-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' }, styleId: 'style-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' }, styleId: 'style-z' },
            ],
            formatRanges: [
              { range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'B1' }, formatId: 'fmt-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A3' }, formatId: 'fmt-z' },
            ],
            filters: [
              { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'D5' },
              { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' },
            ],
            sorts: [
              { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' }, keys: [{ keyAddress: 'A1', direction: 'desc' }] },
              { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' }, keys: [{ keyAddress: 'A1', direction: 'asc' }] },
            ],
            validations: [
              {
                range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C4' },
                rule: { kind: 'whole', operator: 'between', values: [1, 10] },
              },
              {
                range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B4' },
                rule: { kind: 'list', values: ['open', 'closed'] },
              },
            ],
            conditionalFormats: [
              {
                id: 'cf-z',
                range: { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D5' },
                rule: { kind: 'formula', formula: 'D2>0' },
                style: { font: { bold: true } },
              },
              {
                id: 'cf-a',
                range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B5' },
                rule: { kind: 'textContains', text: 'late' },
                style: { fill: { backgroundColor: '#fee2e2' } },
              },
            ],
            protectedRanges: [
              { id: 'lock-z', range: { sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'E5' }, hideFormulas: true },
              { id: 'lock-a', range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' } },
            ],
            commentThreads: [
              {
                threadId: 'thread-z',
                sheetName: 'Sheet1',
                address: 'C3',
                comments: [{ id: 'comment-2', body: 'second', authorDisplayName: 'Beta' }],
              },
              {
                threadId: 'thread-a',
                sheetName: 'Sheet1',
                address: 'A1',
                comments: [{ id: 'comment-1', body: 'first', authorDisplayName: 'Alpha' }],
              },
            ],
            notes: [
              { sheetName: 'Sheet1', address: 'D4', text: 'later' },
              { sheetName: 'Sheet1', address: 'A1', text: 'first' },
            ],
          },
          cells: [{ address: 'A1', value: 1 }],
        },
      ],
    }

    const normalized = normalizeWorkbookSnapshotForSemanticComparison(snapshot)

    expect(normalized.workbook.metadata?.properties?.map((property) => property.key)).toEqual(['author', 'subject'])
    expect(normalized.workbook.metadata?.formats?.map((format) => format.id)).toEqual(['fmt-a', 'fmt-z'])
    expect(normalized.workbook.metadata?.tables?.map((table) => table.name)).toEqual(['TableA', 'TableZ'])
    expect(normalized.workbook.metadata?.pivots?.map((pivot) => pivot.name)).toEqual(['PivotA', 'PivotB'])
    expect(normalized.workbook.metadata?.charts?.map((chart) => chart.id)).toEqual(['chart-a', 'chart-z'])
    expect(normalized.workbook.metadata?.images?.map((image) => image.id)).toEqual(['image-a', 'image-z'])
    expect(normalized.workbook.metadata?.shapes?.map((shape) => shape.id)).toEqual(['shape-a', 'shape-z'])
    expect(normalized.sheets[0]?.metadata?.columns).toEqual([
      { index: 1, size: null },
      { index: 4, size: 128 },
    ])
    expect(normalized.sheets[0]?.metadata?.rowMetadata?.map((metadata) => metadata.start)).toEqual([2, 5])
    expect(normalized.sheets[0]?.metadata?.columnMetadata?.map((metadata) => metadata.start)).toEqual([1, 3])
    expect(normalized.sheets[0]?.metadata?.styleRanges).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'D2' }, styleId: 'style-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' }, styleId: 'style-z' },
    ])
    expect(normalized.sheets[0]?.metadata?.formatRanges).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1' }, formatId: 'fmt-a' },
      { range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A3' }, formatId: 'fmt-z' },
    ])
    expect(normalized.sheets[0]?.metadata?.filters?.map((filter) => filter.startAddress)).toEqual(['A1', 'C1'])
    expect(normalized.sheets[0]?.metadata?.sorts?.map((sort) => sort.keys[0]?.direction)).toEqual(['asc', 'desc'])
    expect(normalized.sheets[0]?.metadata?.validations?.map((validation) => validation.range.startAddress)).toEqual(['B2', 'C2'])
    expect(normalized.sheets[0]?.metadata?.conditionalFormats?.map((format) => format.id)).toEqual(['cf-a', 'cf-z'])
    expect(normalized.sheets[0]?.metadata?.protectedRanges?.map((range) => range.id)).toEqual(['lock-a', 'lock-z'])
    expect(normalized.sheets[0]?.metadata?.commentThreads?.map((thread) => thread.threadId)).toEqual(['thread-a', 'thread-z'])
    expect(normalized.sheets[0]?.metadata?.notes?.map((note) => note.address)).toEqual(['A1', 'D4'])
  })

  it('projects stable workbook semantics independent of generated style ids and defaults', () => {
    const left: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [{ id: 'left-style', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } }],
          charts: [
            {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'E1',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
              chartType: 'column',
              rows: 8,
              cols: 4,
            },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, styleId: 'left-style' }],
            freezePane: { rows: 1, cols: 1, topLeftCell: 'B2', activePane: 'bottomRight' },
          },
          cells: [
            { address: 'B1', formula: 'A1*2' },
            { address: 'A1', value: 12, format: '$0.00' },
          ],
        },
      ],
    }
    const right: WorkbookSnapshot = {
      ...left,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [{ id: 'right-style', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } }],
          charts: [
            {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'E1',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
              chartType: 'column',
              seriesOrientation: 'columns',
              rows: 8,
              cols: 4,
            },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, styleId: 'right-style' }],
            freezePane: { rows: 1, cols: 1 },
          },
          cells: [
            { address: 'A1', value: 12, format: '$0.00' },
            { address: 'B1', formula: 'A1*2' },
          ],
        },
      ],
    }

    expect(projectWorkbookSemanticSnapshot(left)).toEqual(projectWorkbookSemanticSnapshot(right))
    expect(workbookSemanticSnapshotsEqual(left, right)).toBe(true)
  })

  it('projects rich workbook metadata into semantic digests', () => {
    const snapshot: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          properties: [
            { key: 'zeta', value: false },
            { key: 'alpha', value: 'stable' },
          ],
          calculationSettings: {
            mode: 'manual',
            compatibilityMode: 'excel-modern',
            dateSystem: '1904',
            iterate: true,
            iterateCount: 7,
            iterateDelta: '0.001',
            fullCalcOnLoad: true,
          },
          definedNames: [
            { name: 'ZetaName', value: { kind: 'formula', formula: 'SUM(Sheet1!A1:B2)' } },
            { name: 'AlphaName', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
          ],
          styles: [
            {
              id: 'portable',
              fill: { backgroundColor: '#dbeafe' },
              font: { family: 'Inter', size: 12, bold: true, color: '#1e3a8a' },
              alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
              borders: { bottom: { style: 'solid', weight: 'thin', color: '#1f2937' } },
              protection: { locked: true, hidden: false },
            },
          ],
          tables: [
            {
              name: 'Sales',
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'D5',
              columnNames: ['Region', 'Rep', 'Amount', 'Status'],
              columns: [{ name: 'Amount', totalsRowFunction: 'sum' }],
              headerRow: true,
              totalsRow: true,
              style: { name: 'TableStyleMedium2', showRowStripes: true },
              sortState: '<sortState />',
            },
          ],
          charts: [
            {
              id: 'chart-full',
              sheetName: 'Sheet1',
              address: 'G2',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D5' },
              chartType: 'column',
              seriesOrientation: 'rows',
              firstRowAsHeaders: true,
              firstColumnAsLabels: true,
              title: 'Sales by Region',
              legendPosition: 'bottom',
              rows: 12,
              cols: 7,
            },
            {
              id: 'chart-partial',
              sheetName: 'Sheet1',
              address: 'P2',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' },
              chartType: 'pie',
              title: 'Share',
              legendPosition: 'right',
              rows: 8,
              cols: 5,
            },
          ],
          pivots: [
            {
              name: 'SalesPivot',
              sheetName: 'Sheet1',
              address: 'J10',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' },
              groupBy: ['Region'],
              columnFields: ['Status'],
              pageFields: [{ sourceColumn: 'Rep', selectedValue: 'Ana' }],
              filters: [{ sourceColumn: 'Status', includedValues: ['Open'] }],
              hiddenItems: [{ sourceColumn: 'Region', values: ['West'] }],
              calculatedFields: [{ name: 'Net', formula: '=Amount-Discount', clause: '18.10' }],
              calculatedItems: [{ name: 'EastPlus', formula: '=East+South', clause: '3.2.3.1' }],
              values: [
                { sourceColumn: 'Amount', summarizeBy: 'sum', outputLabel: 'Total Amount' },
                { sourceColumn: 'Rep', summarizeBy: 'count' },
              ],
              rows: 8,
              cols: 5,
            },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            rows: [
              { id: 'row-1', index: 1, size: 24 },
              { id: 'row-4', index: 4 },
            ],
            columns: [
              { id: 'col-2', index: 2, size: 96 },
              { id: 'col-5', index: 5 },
            ],
            merges: [
              { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D3' },
              { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
            ],
            styleRanges: [
              { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, styleId: 'portable' },
              { range: { sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'E1' }, styleId: 'missing-style' },
            ],
            freezePane: { rows: 1, cols: 0, topLeftCell: 'A4', activePane: 'topLeft' },
            filters: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' }],
            sorts: [
              { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' }, keys: [{ keyAddress: 'C2', direction: 'desc' }] },
            ],
            validations: [
              {
                range: { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D20' },
                rule: { kind: 'list', values: ['Open', 'Closed'] },
                allowBlank: true,
                promptTitle: 'Status',
                errorStyle: 'warning',
              },
            ],
            conditionalFormats: [
              {
                id: 'cf-priority',
                range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C20' },
                rule: { kind: 'cellIs', operator: 'greaterThan', values: [1000] },
                style: { fill: { backgroundColor: '#dcfce7' } },
                stopIfTrue: true,
                priority: 1,
              },
            ],
            sheetProtection: {
              sheetName: 'Sheet1',
              hideFormulas: true,
              xmlAttributes: [{ name: 'sheet', value: '1' }],
            },
            protectedRanges: [{ id: 'locked-formulas', range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B20' } }],
            commentThreads: [
              {
                threadId: 'thread-1',
                sheetName: 'Sheet1',
                address: 'C2',
                comments: [
                  { id: 'comment-1', body: 'check this', authorDisplayName: 'Analyst' },
                  { id: 'comment-2', body: 'resolved' },
                ],
              },
            ],
          },
          cells: [
            { address: 'B2', formula: 'A2*2' },
            { address: 'A2', value: 42, format: '$0.00' },
            { address: 'C2', value: 'Open' },
          ],
        },
      ],
    }

    const projection = projectWorkbookSemanticSnapshot(snapshot)

    expect(projection.properties).toEqual([
      { key: 'alpha', value: 'stable' },
      { key: 'zeta', value: false },
    ])
    expect(projection.calculationSettings).toMatchObject({ mode: 'manual', dateSystem: '1904', iterateCount: 7 })
    expect(projection.definedNames.map((definedName) => definedName.name)).toEqual(['AlphaName', 'ZetaName'])
    expect(projection.commentThreads).toEqual([
      {
        sheetName: 'Sheet1',
        address: 'C2',
        comments: [{ body: 'check this', authorDisplayName: 'Analyst' }, { body: 'resolved' }],
      },
    ])
    expect(projection.styleRanges).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        style: {
          fill: { backgroundColor: '#dbeafe' },
          font: { family: 'Inter', size: 12, bold: true, color: '#1e3a8a' },
          alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
          borders: { bottom: { style: 'solid', weight: 'thin', color: '#1f2937' } },
          protection: { locked: true, hidden: false },
        },
      },
      { range: { sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'E1' }, style: undefined },
    ])
    expect(projection.tables).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'D5',
        columnNames: ['Region', 'Rep', 'Amount', 'Status'],
        headerRow: true,
        totalsRow: true,
      },
    ])
    expect(projection.charts).toEqual([
      {
        id: 'chart-full',
        sheetName: 'Sheet1',
        address: 'G2',
        source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D5' },
        chartType: 'column',
        seriesOrientation: 'rows',
        firstRowAsHeaders: true,
        firstColumnAsLabels: true,
        title: 'Sales by Region',
        legendPosition: 'bottom',
        rows: 12,
        cols: 7,
      },
      {
        id: 'chart-partial',
        sheetName: 'Sheet1',
        address: 'P2',
        source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' },
        chartType: 'pie',
        seriesOrientation: 'columns',
        title: 'Share',
        legendPosition: 'right',
        rows: 8,
        cols: 5,
      },
    ])
    expect(projection.pivots).toEqual([
      {
        name: 'SalesPivot',
        sheetName: 'Sheet1',
        address: 'J10',
        source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' },
        groupBy: ['Region'],
        columnFields: ['Status'],
        pageFields: [{ sourceColumn: 'Rep', selectedValue: 'Ana' }],
        filters: [{ sourceColumn: 'Status', includedValues: ['Open'] }],
        hiddenItems: [{ sourceColumn: 'Region', values: ['West'] }],
        calculatedFields: [{ name: 'Net', formula: '=Amount-Discount', clause: '18.10' }],
        calculatedItems: [{ name: 'EastPlus', formula: '=East+South', clause: '3.2.3.1' }],
        values: [
          { sourceColumn: 'Amount', summarizeBy: 'sum', outputLabel: 'Total Amount' },
          { sourceColumn: 'Rep', summarizeBy: 'count' },
        ],
        rows: 8,
        cols: 5,
      },
    ])
    expect(projection.validations).toHaveLength(1)
    expect(projection.conditionalFormats).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C20' },
        rule: { kind: 'cellIs', operator: 'greaterThan', values: [1000] },
        style: { fill: { backgroundColor: '#dcfce7' } },
        stopIfTrue: true,
        priority: 1,
      },
    ])
    expect(projection.freezePanes).toEqual([
      {
        sheetName: 'Sheet1',
        freezePane: { rows: 1, cols: 0, topLeftCell: 'A4', activePane: 'topLeft' },
      },
    ])
    expect(projection.filters).toEqual([{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' }])
    expect(projection.sorts).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D20' }, keys: [{ keyAddress: 'C2', direction: 'desc' }] },
    ])
    expect(projection.sheetProtections).toEqual([
      { sheetName: 'Sheet1', hideFormulas: true, xmlAttributes: [{ name: 'sheet', value: '1' }] },
    ])
    expect(projection.protectedRanges).toEqual([
      { id: 'locked-formulas', range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B20' } },
    ])
    expect(projection.valueFormulaFormatSheets[0]?.cells).toEqual([
      { address: 'A2', value: 42, format: '$0.00' },
      { address: 'B2', formula: 'A2*2' },
      { address: 'C2', value: 'Open' },
    ])
    expect(projection.dimensionSheets).toEqual([
      {
        name: 'Sheet1',
        columns: [{ index: 2, size: 96 }, { index: 5 }],
        rows: [{ index: 1, size: 24 }, { index: 4 }],
        merges: [
          { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
          { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D3' },
        ],
      },
    ])
  })

  it('reports stable projection paths for semantic differences', () => {
    const left: WorkbookSnapshot = {
      ...baseSnapshot,
      sheets: [{ name: 'Sheet1', order: 0, cells: [{ address: 'A1', value: 12 }] }],
    }
    const right: WorkbookSnapshot = {
      ...baseSnapshot,
      sheets: [{ name: 'Sheet1', order: 0, cells: [{ address: 'A1', value: 13 }] }],
    }

    expect(diffWorkbookSemanticSnapshots(left, right)).toContainEqual({
      path: '$.valueFormulaFormatSheets.0.cells.0.value',
      left: 12,
      right: 13,
    })
  })
})
