import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'
import { createWorkbookFacade } from '../workbook-facade.js'
import { WorkbookFeatureRegistry } from '../workbook-feature-registry.js'
import { WorkbookProjectionInterceptorService } from '../workbook-projection-interceptors.js'
import { WORKBOOK_TABLE_COMMAND_IDS, WORKBOOK_TABLES_FEATURE_ID } from '../workbook-tables-feature.js'

describe('workbook feature kernel', () => {
  it('activates feature dependencies first and disposes in reverse order', () => {
    const events: string[] = []
    const registry = new WorkbookFeatureRegistry()

    registry.register({
      id: 'tables',
      version: '1.0.0',
      dependsOn: ['core'],
      commands: [],
      projectionInterceptors: [],
      uiContributions: [],
      activate() {
        events.push('activate:tables')
      },
      dispose() {
        events.push('dispose:tables')
      },
    })
    registry.register({
      id: 'core',
      version: '1.0.0',
      commands: [],
      projectionInterceptors: [],
      uiContributions: [],
      activate() {
        events.push('activate:core')
      },
      dispose() {
        events.push('dispose:core')
      },
    })

    expect(registry.activateAll()).toEqual(['core', 'tables'])
    registry.disposeAll()
    expect(events).toEqual(['activate:core', 'activate:tables', 'dispose:tables', 'dispose:core'])
  })

  it('rejects duplicate features and missing dependencies', () => {
    const duplicateRegistry = new WorkbookFeatureRegistry()
    const plugin = {
      id: 'tables',
      version: '1.0.0',
      commands: [],
      projectionInterceptors: [],
      uiContributions: [],
    }
    duplicateRegistry.register(plugin)
    expect(() => duplicateRegistry.register(plugin)).toThrowError('already registered')

    const missingDependencyRegistry = new WorkbookFeatureRegistry()
    missingDependencyRegistry.register({ ...plugin, dependsOn: ['missing'] })
    expect(() => missingDependencyRegistry.activateAll()).toThrowError('depends on missing feature missing')
  })

  it('orders range chrome projection interceptors by priority and registration order', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'feature-kernel-projection-order' })
    await engine.ready()
    const projection = new WorkbookProjectionInterceptorService(engine)

    projection.register({
      id: 'low',
      featureId: 'low-feature',
      point: 'rangeChrome',
      priority: 1,
      projectRangeChrome(input) {
        return [{ id: 'low', featureId: 'low-feature', source: 'runtime', range: input.range, role: 'low' }]
      },
    })
    projection.register({
      id: 'high',
      featureId: 'high-feature',
      point: 'rangeChrome',
      priority: 10,
      projectRangeChrome(input) {
        return [{ id: 'high', featureId: 'high-feature', source: 'runtime', range: input.range, role: 'high' }]
      },
    })
    projection.register({
      id: 'high-later',
      featureId: 'high-feature',
      point: 'rangeChrome',
      priority: 10,
      projectRangeChrome(input) {
        return [{ id: 'high-later', featureId: 'high-feature', source: 'runtime', range: input.range, role: 'high-later' }]
      },
    })

    expect(projection.rangeChrome({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }).map((chrome) => chrome.id)).toEqual([
      'high',
      'high-later',
      'low',
    ])
  })

  it('creates tables through the facade with preview/apply proof, undo, and metadata-derived projection chrome', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'feature-kernel-create-table' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 20],
    ])
    const facade = createWorkbookFacade(engine)

    const createTable = facade
      .selection({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' })
      .createTable({ name: 'Sales', hasHeaders: true })
    const preview = await createTable.preview()

    expect(preview).toMatchObject({
      status: 'previewed',
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      commandId: WORKBOOK_TABLE_COMMAND_IDS.createFromSelection,
      previewOps: [
        {
          kind: 'upsertTable',
          table: {
            name: 'Sales',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        },
      ],
    })
    expect(engine.getTable('Sales')).toBeUndefined()

    const applied = await createTable.applyAndVerify()

    expect(applied.status).toBe('applied')
    expect(applied.undo?.ops).toEqual([{ kind: 'deleteTable', name: 'Sales' }])
    expect(engine.getTable('Sales')).toMatchObject({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Amount'],
    })
    expect(facade.table('Sales').column('Amount')).toEqual({ tableName: 'Sales', columnName: 'Amount', columnIndex: 1 })
    expect(facade.projection().rangeChrome({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' })).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'tables:Sales:table',
          featureId: WORKBOOK_TABLES_FEATURE_ID,
          source: 'workbook-metadata',
          role: 'table',
        }),
        expect.objectContaining({
          id: 'tables:Sales:header',
          role: 'header',
        }),
        expect.objectContaining({
          id: 'tables:Sales:dataBody',
          role: 'dataBody',
        }),
      ]),
    )

    expect(engine.undo()).toBe(true)
    expect(engine.getTable('Sales')).toBeUndefined()
    facade.dispose()
  })

  it('upserts tables without dropping style, filter, or column metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'feature-kernel-upsert-table' })
    await engine.ready()
    engine.createSheet('Data')
    const facade = createWorkbookFacade(engine)

    await facade
      .command({
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        commandId: WORKBOOK_TABLE_COMMAND_IDS.upsert,
        category: 'operation',
        input: {
          table: {
            name: 'Sales',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'B4',
            columnNames: ['Region', 'Amount'],
            columns: [
              { name: 'Region', totalsRowLabel: 'Total' },
              { name: 'Amount', calculatedColumnFormula: '=[@Units]*[@Price]', totalsRowFunction: 'sum' },
            ],
            headerRow: true,
            totalsRow: true,
            style: { name: 'TableStyleMedium2', showRowStripes: true },
            autoFilter: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4', criteria: [] },
            sortState: 'Amount desc',
          },
        },
      })
      .applyAndVerify()

    expect(engine.getTable('Sales')).toMatchObject({
      columns: [
        { name: 'Region', totalsRowLabel: 'Total' },
        { name: 'Amount', calculatedColumnFormula: '=[@Units]*[@Price]', totalsRowFunction: 'sum' },
      ],
      style: { name: 'TableStyleMedium2', showRowStripes: true },
      autoFilter: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4', criteria: [] },
      sortState: 'Amount desc',
    })
    facade.dispose()
  })

  it('renames table headers through the command service without losing structured-reference rewrites', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'feature-kernel-table-header-rename' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 20],
    ])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Data', 'D1', 'SUM(Sales[Amount])')
    const facade = createWorkbookFacade(engine)

    const receipt = await facade
      .command({
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        commandId: WORKBOOK_TABLE_COMMAND_IDS.renameHeader,
        category: 'command',
        input: {
          tableName: 'Sales',
          columnName: 'Amount',
          name: 'Revenue',
        },
      })
      .applyAndVerify()

    expect(receipt.status).toBe('applied')
    expect(engine.getTable('Sales')).toMatchObject({
      columnNames: ['Region', 'Revenue'],
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Revenue])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Amount])')
    facade.dispose()
  })
})
