import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'
import { translateImportedFormulaStructuredReferences } from '../xlsx-formula-translation.js'

describe('structured reference XLSX import', () => {
  it('translates Excel table sections and this-row references into executable formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Financial Model',
        metadata: {
          definedNames: [
            { name: 'Currency', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'F7' } },
            { name: 'Start_Year', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'B10' } },
          ],
          tables: [
            {
              name: 'tblPFS',
              sheetName: 'Constants',
              startAddress: 'H6',
              endAddress: 'O7',
              columnNames: ['Period', 'Revenue', 'CoGS', 'Gross Profit', 'Opex', 'EBITDA', 'Tax', 'Cash'],
              headerRow: true,
              totalsRow: false,
            },
            {
              name: 'tblActuals',
              sheetName: 'Imports',
              startAddress: 'A6',
              endAddress: 'D8',
              columnNames: ['Account', 'Value', 'Year', 'Period'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [
            { address: 'B10', value: 2012 },
            { address: 'F7', value: 'USD' },
            { address: 'F9', formula: 'Currency & "  000s"' },
            { address: 'H6', value: 'Period' },
            { address: 'I6', value: 'Revenue' },
            { address: 'J6', value: 'CoGS' },
            { address: 'K6', value: 'Gross Profit' },
            { address: 'L6', value: 'Opex' },
            { address: 'M6', value: 'EBITDA' },
            { address: 'N6', value: 'Tax' },
            { address: 'O6', value: 'Cash' },
            { address: 'H7', formula: 'ROW()-ROW(tblPFS[#Headers])' },
          ],
        },
        {
          id: 2,
          name: 'Imports',
          order: 1,
          cells: [
            { address: 'A6', value: 'Account' },
            { address: 'B6', value: 'Value' },
            { address: 'C6', value: 'Year' },
            { address: 'D6', value: 'Period' },
            { address: 'A7', value: 'Revenue' },
            { address: 'B7', value: 100 },
            { address: 'C7', value: 2011 },
            { address: 'D7', formula: 'tblActuals[[#This Row],[Year]]-Start_Year+1' },
            { address: 'A8', value: 'Revenue' },
            { address: 'B8', value: 125 },
            { address: 'C8', value: 2012 },
            { address: 'D8', formula: 'tblActuals[[#This Row],[Year]]-Start_Year+1' },
            { address: 'F10', formula: 'SUM(tblActuals[Value])' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-financial-model.xlsx')
    const constants = imported.snapshot.sheets.find((sheet) => sheet.name === 'Constants')
    const imports = imported.snapshot.sheets.find((sheet) => sheet.name === 'Imports')

    expect(constants?.cells.find((cell) => cell.address === 'H7')?.formula).not.toContain('tblPFS[#Headers]')
    expect(imports?.cells.find((cell) => cell.address === 'D7')?.formula).toBe('tblActuals[[#This Row],[Year]]-Start_Year+1')
    expect(imports?.cells.find((cell) => cell.address === 'F10')?.formula).toBe('SUM(tblActuals[Value])')

    const engine = new SpreadsheetEngine({ workbookName: 'structured-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Constants', 'F9')).toMatchObject({ tag: ValueTag.String, value: 'USD  000s' })
    expect(engine.getCellValue('Constants', 'H7')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Imports', 'D7')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Imports', 'D8')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Imports', 'F10')).toEqual({ tag: ValueTag.Number, value: 225 })
  })

  it('translates whole-table structured references into data-body ranges', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Lookup Panel',
        metadata: {
          tables: [
            {
              name: 'Data_Origin_Table',
              sheetName: 'Panel',
              startAddress: 'A10',
              endAddress: 'B12',
              columnNames: ['Origin', 'Code'],
              headerRow: true,
              totalsRow: false,
            },
            {
              name: 'Data_Quality_Table',
              sheetName: 'Panel',
              startAddress: 'A15',
              endAddress: 'B17',
              columnNames: ['Quality', 'Code'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Panel',
          order: 0,
          cells: [
            { address: 'B2', formula: 'CONCATENATE(C8,C13," ",)' },
            { address: 'B8', value: 'Public' },
            { address: 'C8', formula: 'IFERROR(VLOOKUP(B8,Data_Origin_Table[],2,FALSE),"X")' },
            { address: 'B13', value: 'Not calculated / not applicable' },
            { address: 'C13', formula: 'IFERROR(VLOOKUP(B13,Data_Quality_Table[],2,FALSE),"X")' },
            { address: 'A10', value: 'Origin' },
            { address: 'B10', value: 'Code' },
            { address: 'A11', value: 'Public' },
            { address: 'B11', value: 'PUB-' },
            { address: 'A12', value: 'Private' },
            { address: 'B12', value: 'PRI-' },
            { address: 'A15', value: 'Quality' },
            { address: 'B15', value: 'Code' },
            { address: 'A16', value: 'Not calculated / not applicable' },
            { address: 'A17', value: 'High' },
            { address: 'B17', value: 'H-' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-lookup-panel.xlsx')
    const panel = imported.snapshot.sheets.find((sheet) => sheet.name === 'Panel')

    expect(panel?.cells.find((cell) => cell.address === 'C8')?.formula).toBe('IFERROR(VLOOKUP(B8,\'Panel\'!A11:B12,2,FALSE),"X")')
    expect(panel?.cells.find((cell) => cell.address === 'C13')?.formula).toBe('IFERROR(VLOOKUP(B13,\'Panel\'!A16:B17,2,FALSE),"X")')

    const engine = new SpreadsheetEngine({ workbookName: 'structured-lookup-panel' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Panel', 'C8')).toMatchObject({ tag: ValueTag.String, value: 'PUB-' })
    expect(engine.getCellValue('Panel', 'C13')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Panel', 'B2')).toMatchObject({ tag: ValueTag.String, value: 'PUB-0 ' })
  })

  it('translates cross-sheet current-row structured references by row position', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Cross Sheet Rows',
        metadata: {
          tables: [
            {
              name: 'RevenueTable',
              sheetName: 'Data',
              startAddress: 'A1',
              endAddress: 'C4',
              columnNames: ['Segment', '2024', '2025'],
              headerRow: true,
              totalsRow: false,
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
            { address: 'A1', value: 'Segment' },
            { address: 'B1', value: '2024' },
            { address: 'C1', value: '2025' },
            { address: 'A2', value: 'A' },
            { address: 'B2', value: 100 },
            { address: 'C2', value: 125 },
            { address: 'A3', value: 'B' },
            { address: 'B3', value: 200 },
            { address: 'C3', value: 250 },
            { address: 'A4', value: 'C' },
            { address: 'B4', value: 300 },
            { address: 'C4', value: 390 },
          ],
        },
        {
          id: 2,
          name: 'Ratios',
          order: 1,
          cells: [{ address: 'B3', formula: 'RevenueTable[[#This Row],[2025]]/RevenueTable[[#This Row],[2024]]' }],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-cross-sheet-rows.xlsx')
    const ratios = imported.snapshot.sheets.find((sheet) => sheet.name === 'Ratios')

    expect(ratios?.cells.find((cell) => cell.address === 'B3')?.formula).toBe(
      'RevenueTable[[#This Row],[2025]]/RevenueTable[[#This Row],[2024]]',
    )

    const engine = new SpreadsheetEngine({ workbookName: 'structured-cross-sheet-rows' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Ratios', 'B3')).toEqual({ tag: ValueTag.Number, value: 1.25 })
  })

  it('normalizes XML line endings in structured reference column names', () => {
    const formula = 'Budget[[#This Row],[Projected\r\ncost]]-Budget[[#This Row],[Actual \r\ncost]]'

    expect(
      translateImportedFormulaStructuredReferences({
        formula,
        ownerSheetName: 'Summary',
        ownerAddress: 'E16',
        tables: [
          {
            name: 'Budget',
            sheetName: 'Summary',
            startAddress: 'B15',
            endAddress: 'E26',
            columnNames: ['Category', 'Projected\ncost', 'Actual \ncost', 'Difference'],
            headerRow: true,
            totalsRow: true,
          },
        ],
      }),
    ).toBe("'Summary'!C16-'Summary'!D16")
  })

  it('decodes Excel-escaped table column names before rewriting structured references', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Escaped Table Headers',
        metadata: {
          tables: [
            {
              name: 'Table41',
              sheetName: 'Local and Regional Initiatives',
              startAddress: 'A1',
              endAddress: 'C4',
              columnNames: ['Year', 'Number of\n people inducted', 'Road show \nvisits'],
              columns: [
                { name: 'Year', totalsRowLabel: 'Total' },
                { name: 'Number of\n people inducted', totalsRowFunction: 'sum' },
                { name: 'Road show \nvisits', totalsRowFunction: 'sum' },
              ],
              headerRow: true,
              totalsRow: true,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Local and Regional Initiatives',
          order: 0,
          cells: [
            { address: 'A1', value: 'Year' },
            { address: 'B1', value: 'Number of\n people inducted' },
            { address: 'C1', value: 'Road show \nvisits' },
            { address: 'A2', value: 2013 },
            { address: 'B2', value: 64 },
            { address: 'C2', value: 10 },
            { address: 'A3', value: 2014 },
            { address: 'B3', value: 15 },
            { address: 'C3', value: 5 },
            { address: 'A4', value: 'Total' },
            { address: 'B4', formula: 'SUBTOTAL(109,Table41[Number of\n people inducted])' },
            { address: 'C4', formula: 'SUBTOTAL(109,Table41[Road show \nvisits])' },
          ],
        },
      ],
    }

    const imported = importXlsx(withExcelEscapedTableColumnNames(exportXlsx(snapshot)), 'escaped-table-headers.xlsx')
    const table = imported.snapshot.workbook.metadata?.tables?.find((candidate) => candidate.name === 'Table41')
    const sheet = imported.snapshot.sheets.find((candidate) => candidate.name === 'Local and Regional Initiatives')

    expect(table?.columnNames).toEqual(['Year', 'Number of\n people inducted', 'Road show \nvisits'])
    expect(sheet?.cells.find((cell) => cell.address === 'B4')?.formula).toBe('SUBTOTAL(109,Table41[Number of\n people inducted])')
    expect(sheet?.cells.find((cell) => cell.address === 'C4')?.formula).toBe('SUBTOTAL(109,Table41[Road show \nvisits])')

    const engine = new SpreadsheetEngine({ workbookName: 'escaped-table-headers' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Local and Regional Initiatives', 'B4')).toEqual({ tag: ValueTag.Number, value: 79 })
    expect(engine.getCellValue('Local and Regional Initiatives', 'C4')).toEqual({ tag: ValueTag.Number, value: 15 })

    const reexportedTableXml = readOnlyTableXml(exportXlsx(imported.snapshot))
    expect(reexportedTableXml).toMatch(/name="Number of_x000a_ people inducted"/iu)
    expect(reexportedTableXml).toMatch(/name="Road show _x000a_visits"/iu)
  })

  it('rewrites imported structured references with special-character table headers', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Special Structured Headers',
        metadata: {
          tables: [
            {
              name: 'Sales',
              sheetName: 'Data',
              startAddress: 'A1',
              endAddress: 'E3',
              columnNames: ['Revenue, Net', 'A:B', '# Units', "Owner's Share", 'A[B]'],
              headerRow: true,
              totalsRow: false,
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
            { address: 'A1', value: 'Revenue, Net' },
            { address: 'B1', value: 'A:B' },
            { address: 'C1', value: '# Units' },
            { address: 'D1', value: "Owner's Share" },
            { address: 'E1', value: 'A[B]' },
            { address: 'A2', value: 10 },
            { address: 'B2', value: 2 },
            { address: 'C2', value: 4 },
            { address: 'D2', value: 0.25 },
            { address: 'E2', value: 100 },
            { address: 'A3', value: 20 },
            { address: 'B3', value: 3 },
            { address: 'C3', value: 6 },
            { address: 'D3', value: 0.75 },
            { address: 'E3', value: 200 },
            { address: 'G1', formula: 'SUM(Sales[Revenue, Net])' },
            { address: 'G2', formula: 'SUM(Sales[A:B])' },
            { address: 'G3', formula: "SUM(Sales['# Units])" },
            { address: 'G4', formula: "SUM(Sales[Owner''s Share])" },
            { address: 'G5', formula: "SUM(Sales[A'[B']])" },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'special-structured-headers.xlsx')
    const sheet = imported.snapshot.sheets.find((candidate) => candidate.name === 'Data')

    expect(sheet?.cells.find((cell) => cell.address === 'G1')?.formula).toBe('SUM(Sales[Revenue, Net])')
    expect(sheet?.cells.find((cell) => cell.address === 'G2')?.formula).toBe('SUM(Sales[A:B])')
    expect(sheet?.cells.find((cell) => cell.address === 'G3')?.formula).toBe("SUM(Sales['# Units])")
    expect(sheet?.cells.find((cell) => cell.address === 'G4')?.formula).toBe("SUM(Sales[Owner''s Share])")
    expect(sheet?.cells.find((cell) => cell.address === 'G5')?.formula).toBe("SUM(Sales[A'[B']])")

    const engine = new SpreadsheetEngine({ workbookName: 'special-structured-headers' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'G2')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Data', 'G3')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Data', 'G4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Data', 'G5')).toEqual({ tag: ValueTag.Number, value: 300 })
  })

  it('rewrites Excel total-row structured-reference aliases', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Totals',
        metadata: {
          tables: [
            {
              name: 'SalesTable',
              sheetName: 'Sales',
              startAddress: 'A1',
              endAddress: 'B4',
              columnNames: ['Region', 'Amount'],
              headerRow: true,
              totalsRow: true,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sales',
          order: 0,
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Amount' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'Total' },
            { address: 'B4', value: 17 },
            { address: 'D2', formula: 'SalesTable[[#Total Row],[Amount]]' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-total-row.xlsx')
    const sheet = imported.snapshot.sheets.find((candidate) => candidate.name === 'Sales')

    expect(sheet?.cells.find((cell) => cell.address === 'D2')?.formula).toBe('SalesTable[[#Total Row],[Amount]]')
  })

  it('infers omitted table totals row flags from totals-row column formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Omitted Totals Flag',
        metadata: {
          tables: [
            {
              name: 'PorosityTable',
              sheetName: 'Porosity',
              startAddress: 'A1',
              endAddress: 'B4',
              columnNames: ['Sample', 'Average Total Porosity'],
              columns: [{ name: 'Sample' }, { name: 'Average Total Porosity', totalsRowFunction: 'custom' }],
              headerRow: true,
              totalsRow: true,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Porosity',
          order: 0,
          cells: [
            { address: 'A1', value: 'Sample' },
            { address: 'B1', value: 'Average Total Porosity' },
            { address: 'A2', value: 'A' },
            { address: 'B2', value: 0.25 },
            { address: 'A3', value: 'B' },
            { address: 'B3', value: 0.75 },
            { address: 'A4', value: 'Average' },
            { address: 'B4', formula: 'AVERAGE(PorosityTable[Average Total Porosity])' },
          ],
        },
      ],
    }

    const imported = importXlsx(withOmittedTotalsRowShownAndCustomTotalFormula(exportXlsx(snapshot)), 'omitted-totals-flag.xlsx')
    const table = imported.snapshot.workbook.metadata?.tables?.find((candidate) => candidate.name === 'PorosityTable')
    const sheet = imported.snapshot.sheets.find((candidate) => candidate.name === 'Porosity')

    expect(table?.totalsRow).toBe(true)
    expect(sheet?.cells.find((cell) => cell.address === 'B4')?.formula).toBe('AVERAGE(PorosityTable[Average Total Porosity])')

    const engine = new SpreadsheetEngine({ workbookName: 'omitted-totals-flag' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Porosity', 'B4')).toEqual({ tag: ValueTag.Number, value: 0.5 })
  })

  it('preserves lowercase defined names and named ranges used by simulation formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Simulation Model',
        metadata: {
          definedNames: [
            { name: 'distribution', value: { kind: 'range-ref', sheetName: 'newsvendor', startAddress: 'H35', endAddress: 'K36' } },
            { name: 'mean', value: { kind: 'cell-ref', sheetName: 'newsvendor', address: 'D9' } },
            { name: 'std', value: { kind: 'cell-ref', sheetName: 'newsvendor', address: 'D10' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'newsvendor',
          order: 0,
          cells: [
            { address: 'D9', value: 5000 },
            { address: 'D10', value: 250 },
            { address: 'H35', value: 0.25 },
            { address: 'I35', value: 125 },
            { address: 'J35', value: 'low' },
            { address: 'K35', value: true },
            { address: 'H36', value: 0.5 },
            { address: 'I36', value: 225 },
            { address: 'J36', value: 'high' },
            { address: 'K36', value: false },
            { address: 'I17', value: 0.975 },
            { address: 'I18', formula: 'mean+(NORM.S.INV(I17)*std)' },
            { address: 'I19', formula: 'VLOOKUP(0.5,distribution,2,FALSE)' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'newsvendor-simulation.xlsm')
    const definedNames = imported.snapshot.workbook.metadata?.definedNames ?? []

    expect(definedNames.find((name) => name.name === 'mean')?.value).toEqual({
      kind: 'cell-ref',
      sheetName: 'newsvendor',
      address: 'D9',
    })
    expect(definedNames.find((name) => name.name === 'distribution')?.value).toEqual({
      kind: 'range-ref',
      sheetName: 'newsvendor',
      startAddress: 'H35',
      endAddress: 'K36',
    })

    const engine = new SpreadsheetEngine({ workbookName: 'newsvendor-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    const i18 = engine.getCellValue('newsvendor', 'I18')
    expect(i18.tag).toBe(ValueTag.Number)
    expect(i18.tag === ValueTag.Number ? i18.value : Number.NaN).toBeCloseTo(5489.990996, 5)
    expect(engine.getCellValue('newsvendor', 'I19')).toEqual({ tag: ValueTag.Number, value: 225 })
  })

  it('preserves native current-row shorthand and multi-column structured references from imported formulas', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Span Model',
        metadata: {
          tables: [
            {
              name: 'Metrics',
              sheetName: 'PlayerData',
              startAddress: 'A1',
              endAddress: 'D4',
              columnNames: ['Feet', 'Inches', 'Height', 'RowTotal'],
              headerRow: true,
              totalsRow: true,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'PlayerData',
          order: 0,
          cells: [
            { address: 'A1', value: 'Feet' },
            { address: 'B1', value: 'Inches' },
            { address: 'C1', value: 'Height' },
            { address: 'D1', value: 'RowTotal' },
            { address: 'A2', value: 5 },
            { address: 'B2', value: 7 },
            { address: 'C2', formula: '[@Feet]+([@Inches]/12)' },
            { address: 'D2', formula: 'SUM([@[Feet]:[Inches]])' },
            { address: 'A3', value: 5 },
            { address: 'B3', value: 9 },
            { address: 'C3', formula: '[@Feet]+([@Inches]/12)' },
            { address: 'D3', formula: 'SUM([@[Feet]:[Inches]])' },
            { address: 'C4', formula: 'SUM(C2:C3)' },
            { address: 'D4', formula: 'SUM(D2:D3)' },
            { address: 'F1', formula: 'SUM(Metrics[[Feet]:[Inches]])' },
            { address: 'F2', formula: 'COUNTA(Metrics[[#Headers],[Feet]:[Inches]])' },
            { address: 'F3', formula: 'Metrics[[#Totals],[Height]]' },
            { address: 'F4', formula: 'SUM(Metrics[[#Totals],[Height]:[RowTotal]])' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'structured-span-model.xlsx')
    const playerData = imported.snapshot.sheets.find((sheet) => sheet.name === 'PlayerData')

    expect(playerData?.cells.find((cell) => cell.address === 'C2')?.formula).toBe('[@Feet]+([@Inches]/12)')
    expect(playerData?.cells.find((cell) => cell.address === 'D2')?.formula).toBe('SUM([@[Feet]:[Inches]])')
    expect(playerData?.cells.find((cell) => cell.address === 'F1')?.formula).toBe('SUM(Metrics[[Feet]:[Inches]])')
    expect(playerData?.cells.find((cell) => cell.address === 'F2')?.formula).toBe('COUNTA(Metrics[[#Headers],[Feet]:[Inches]])')
    expect(playerData?.cells.find((cell) => cell.address === 'F4')?.formula).toBe('SUM(Metrics[[#Totals],[Height]:[RowTotal]])')

    const engine = new SpreadsheetEngine({ workbookName: 'structured-span-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    const c2 = engine.getCellValue('PlayerData', 'C2')
    expect(c2.tag).toBe(ValueTag.Number)
    expect(c2.tag === ValueTag.Number ? c2.value : Number.NaN).toBeCloseTo(5.583333333333333, 10)

    const c3 = engine.getCellValue('PlayerData', 'C3')
    expect(c3.tag).toBe(ValueTag.Number)
    expect(c3.tag === ValueTag.Number ? c3.value : Number.NaN).toBeCloseTo(5.75, 10)

    expect(engine.getCellValue('PlayerData', 'D2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('PlayerData', 'D3')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('PlayerData', 'F1')).toEqual({ tag: ValueTag.Number, value: 26 })
    expect(engine.getCellValue('PlayerData', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })

    const f3 = engine.getCellValue('PlayerData', 'F3')
    expect(f3.tag).toBe(ValueTag.Number)
    expect(f3.tag === ValueTag.Number ? f3.value : Number.NaN).toBeCloseTo(11.333333333333332, 10)

    const f4 = engine.getCellValue('PlayerData', 'F4')
    expect(f4.tag).toBe(ValueTag.Number)
    expect(f4.tag === ValueTag.Number ? f4.value : Number.NaN).toBeCloseTo(37.33333333333333, 10)
  })
})

function withExcelEscapedTableColumnNames(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const tablePath = onlyTablePath(zip)
  const tableXml = strFromU8(zip[tablePath] ?? new Uint8Array())
  zip[tablePath] = strToU8(
    tableXml
      .replace(/Number of(?:\r?\n|_x000a_) people inducted/giu, 'Number of_x000a_ people inducted')
      .replace(/Road show (?:\r?\n|_x000a_)visits/giu, 'Road show _x000a_visits'),
  )
  return zipSync(zip)
}

function withOmittedTotalsRowShownAndCustomTotalFormula(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const tablePath = onlyTablePath(zip)
  const tableXml = strFromU8(zip[tablePath] ?? new Uint8Array())
  zip[tablePath] = strToU8(
    tableXml
      .replace(/\s+totalsRowShown="1"/u, '')
      .replace(
        '<tableColumn id="2" name="Average Total Porosity" totalsRowFunction="custom"/>',
        '<tableColumn id="2" name="Average Total Porosity" totalsRowFunction="custom"><totalsRowFormula>AVERAGE(PorosityTable[Average Total Porosity])</totalsRowFormula></tableColumn>',
      ),
  )
  return zipSync(zip)
}

function readOnlyTableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  return strFromU8(zip[onlyTablePath(zip)] ?? new Uint8Array())
}

function onlyTablePath(zip: Record<string, Uint8Array>): string {
  const tablePaths = Object.keys(zip).filter((path) => /^xl\/tables\/table[0-9]+\.xml$/u.test(path))
  expect(tablePaths).toHaveLength(1)
  return tablePaths[0]
}
