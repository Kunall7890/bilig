import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('MS-OI29500 pivot semantic import', () => {
  it('imports row, column, page, hidden item, and aggregate variant semantics from worksheet pivots', () => {
    const imported = importXlsx(buildPivotSemanticsWorkbookBytes(), 'ms-oi29500-pivot-semantics.xlsx')

    expect(imported.snapshot.workbook.metadata?.pivots).toEqual([
      expect.objectContaining({
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'A1',
        cacheId: 1,
        sourceKind: 'worksheet',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'E6' },
        groupBy: ['Region'],
        columnFields: ['Quarter'],
        pageFields: [{ sourceColumn: 'Status', selectedValue: 'Closed' }],
        hiddenItems: [{ sourceColumn: 'Status', values: ['Open'] }],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' },
          { sourceColumn: 'Sales', summarizeBy: 'average', outputLabel: 'Average Sales' },
          { sourceColumn: 'Units', summarizeBy: 'countNums', outputLabel: 'Numeric Units' },
          { sourceColumn: 'Sales', summarizeBy: 'min', outputLabel: 'Min Sales' },
          { sourceColumn: 'Sales', summarizeBy: 'max', outputLabel: 'Max Sales' },
          { sourceColumn: 'Units', summarizeBy: 'product', outputLabel: 'Product Units' },
        ],
      }),
    ])
  })

  it('imports external cache-only pivots semantically when cache records are present', () => {
    const imported = importXlsx(buildExternalCacheOnlyPivotWorkbookBytes(), 'ms-oi29500-external-pivot-cache.xlsx')
    const pivot = imported.snapshot.workbook.metadata?.pivots?.[0]

    expect(pivot).toEqual(
      expect.objectContaining({
        name: 'ExternalSales',
        sheetName: 'Pivot',
        address: 'A1',
        cacheId: 1,
        sourceKind: 'external-cache-only',
        cacheOnly: true,
        cacheFields: ['Region', 'Sales'],
        cachedRecords: [
          ['East', 10],
          ['West', 7],
        ],
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
      }),
    )
    expect(pivot?.source).toBeUndefined()
  })
})

function buildPivotSemanticsWorkbookBytes(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildPivotWorkbookSnapshot()))
  const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml'
  const cacheXml = strFromU8(zip[cachePath] ?? new Uint8Array())
  zip[cachePath] = strToU8(
    cacheXml.replace(
      '<cacheField name="Status" numFmtId="0"><sharedItems/></cacheField>',
      '<cacheField name="Status" numFmtId="0"><sharedItems count="2"><s v="Closed"/><s v="Open"/></sharedItems></cacheField>',
    ),
  )
  const pivotPath = 'xl/pivotTables/pivotTable1.xml'
  const pivotXml = strFromU8(zip[pivotPath] ?? new Uint8Array())
  zip[pivotPath] = strToU8(
    pivotXml
      .replace(
        '<pivotField showAll="0"/><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/>',
        [
          '<pivotField axis="axisCol" showAll="0"><items count="1"><item t="default"/></items></pivotField>',
          '<pivotField axis="axisPage" showAll="0"><items count="2"><item x="0"/><item x="1" h="1"/></items></pivotField>',
          '<pivotField dataField="1" showAll="0"/>',
        ].join(''),
      )
      .replace(
        '<rowFields count="1"><field x="0"/></rowFields>',
        '<rowFields count="1"><field x="0"/></rowFields><colFields count="1"><field x="1"/></colFields><pageFields count="1"><pageField fld="2" item="0"/></pageFields>',
      )
      .replace(
        '<dataFields count="1"><dataField name="Sales Total" fld="3" subtotal="sum"/></dataFields>',
        [
          '<dataFields count="6">',
          '<dataField name="Sales Total" fld="3" subtotal="sum"/>',
          '<dataField name="Average Sales" fld="3" subtotal="average"/>',
          '<dataField name="Numeric Units" fld="4" subtotal="countNums"/>',
          '<dataField name="Min Sales" fld="3" subtotal="min"/>',
          '<dataField name="Max Sales" fld="3" subtotal="max"/>',
          '<dataField name="Product Units" fld="4" subtotal="product"/>',
          '</dataFields>',
        ].join(''),
      ),
  )
  return zipSync(zip)
}

function buildExternalCacheOnlyPivotWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), 'Pivot')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  zip['xl/workbook.xml'] = strToU8(
    workbookXml.replace('</sheets>', '</sheets><pivotCaches><pivotCache cacheId="1" r:id="rIdExternalPivotCache"/></pivotCaches>'),
  )
  const workbookRelsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    workbookRelsXml.replace(
      '</Relationships>',
      '<Relationship Id="rIdExternalPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>',
    ),
  )
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml
      .replace(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      )
      .replace('</worksheet>', '<pivotTableDefinition r:id="rIdExternalPivot"/></worksheet>'),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rIdExternalPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>',
      '</Relationships>',
    ].join(''),
  )
  zip['xl/pivotTables/pivotTable1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="ExternalSales" cacheId="1">',
      '<location ref="A1:B3" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>',
      '<pivotFields count="2">',
      '<pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>',
      '<pivotField dataField="1" showAll="0"/>',
      '</pivotFields>',
      '<rowFields count="1"><field x="0"/></rowFields>',
      '<dataFields count="1"><dataField name="Sales Total" fld="1" subtotal="sum"/></dataFields>',
      '</pivotTableDefinition>',
    ].join(''),
  )
  zip['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdRecords" recordCount="2">',
      '<cacheSource type="external"/>',
      '<cacheFields count="2">',
      '<cacheField name="Region"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>',
      '<cacheField name="Sales"><sharedItems count="2"><n v="10"/><n v="7"/></sharedItems></cacheField>',
      '</cacheFields>',
      '</pivotCacheDefinition>',
    ].join(''),
  )
  zip['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>',
      '</Relationships>',
    ].join(''),
  )
  zip['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">',
      '<r><x v="0"/><x v="0"/></r>',
      '<r><x v="1"/><x v="1"/></r>',
      '</pivotCacheRecords>',
    ].join(''),
  )
  return zipSync(zip)
}

function buildPivotWorkbookSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'pivot-semantics',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'A1',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'E6' },
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
            rows: 4,
            cols: 2,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Quarter' },
          { address: 'C1', value: 'Status' },
          { address: 'D1', value: 'Sales' },
          { address: 'E1', value: 'Units' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'Q1' },
          { address: 'C2', value: 'Closed' },
          { address: 'D2', value: 10 },
          { address: 'E2', value: 2 },
          { address: 'A3', value: 'East' },
          { address: 'B3', value: 'Q1' },
          { address: 'C3', value: 'Open' },
          { address: 'D3', value: 100 },
          { address: 'E3', value: 9 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'Q2' },
          { address: 'C4', value: 'Closed' },
          { address: 'D4', value: 5 },
          { address: 'E4', value: 4 },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 'Q2' },
          { address: 'C5', value: 'Closed' },
          { address: 'D5', value: 7 },
          { address: 'E5', value: 5 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [],
      },
    ],
  }
}
