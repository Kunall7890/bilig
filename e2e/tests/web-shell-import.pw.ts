import { basename } from 'node:path'
import { readFileSync } from 'node:fs'
import { readFile, writeFile } from 'node:fs/promises'
import { expect, type Page, test } from '@playwright/test'
import {
  decodeCellAddress,
  encodeCellAddress,
  readXlsxWorkbookCells,
  writeSimpleXlsxWorkbook,
  type SimpleXlsxAxisEntry,
  type SimpleXlsxCell,
  type SimpleXlsxMergeRange,
  type SimpleXlsxSheet,
} from '@bilig/xlsx'
import { getProductColumnWidth, gotoWorkbookShell, waitForWorkbookReady } from './web-shell-helpers.js'

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
type WorkbookFixtureValue = boolean | number | string | null

interface WorkbookImportExpectation {
  readonly workbookName: string
  readonly uploadFileName?: string
  readonly sheetNames: readonly string[]
  readonly activeSheetName: string
  readonly expectedColumnWidth?: number
  readonly cells: readonly {
    readonly sheetName: string
    readonly address: string
    readonly value: string
  }[]
}

function fixtureCells(rows: ReadonlyArray<ReadonlyArray<WorkbookFixtureValue>>): SimpleXlsxCell[] {
  const cells: SimpleXlsxCell[] = []
  for (let row = 0; row < rows.length; row += 1) {
    const rowValues = rows[row] ?? []
    for (let col = 0; col < rowValues.length; col += 1) {
      const value = rowValues[col]
      if (value === null || value === undefined) {
        continue
      }
      cells.push({ address: encodeCellAddress({ r: row, c: col }), row, col, value })
    }
  }
  return cells
}

function formulaCell(address: string, formula: string): SimpleXlsxCell {
  const { r, c } = decodeCellAddress(address)
  return { address, row: r, col: c, formula }
}

function columnSizes(sizes: readonly number[]): SimpleXlsxAxisEntry[] {
  return sizes.map((size, index) => ({ index, size }))
}

function rowSizes(entries: readonly (number | null)[]): SimpleXlsxAxisEntry[] {
  return entries.flatMap((size, index) => (size === null ? [] : [{ index, size }]))
}

function merge(startAddress: string, endAddress: string): SimpleXlsxMergeRange {
  return { startAddress, endAddress }
}

function writeWorkbook(sheets: readonly SimpleXlsxSheet[]): Uint8Array {
  return writeSimpleXlsxWorkbook({ sheets })
}

function buildMultiSheetOperationsWorkbook(): Uint8Array {
  const dashboardCells = fixtureCells([
    ['OPERATIONS DASHBOARD', null, null, null],
    [],
    ['Metric', 'Value'],
    ['Total budget'],
    ['Open balance'],
    ['Completion rate'],
  ])
  dashboardCells.push(
    formulaCell('B4', 'SUM(Ledger!F:F)'),
    formulaCell('B5', 'SUMIF(Ledger!H:H,"Open",Ledger!G:G)'),
    formulaCell('B6', 'IF(B4>0,1-B5/B4,0)'),
  )

  const ledgerCells = fixtureCells([
    ['OPERATIONS LEDGER', null, null, null, null, null, null, null],
    [],
    ['ID', 'Date', 'Owner', 'Workstream', 'Category', 'Budget', 'Open Balance', 'Status'],
    ['OP001', 45292, 'Facilities', 'Office refresh', 'Capital', 12000, null, 'Open'],
    ['OP002', 45323, 'Engineering', 'Data migration', 'Platform', 18000, null, 'Open'],
  ])
  ledgerCells.push(
    formulaCell('G4', 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)'),
    formulaCell('G5', 'F5-SUMIF(Rollforward!$B:$B,A5,Rollforward!$E:$E)'),
  )

  const rollforwardCells = fixtureCells([
    ['ROLLFORWARD', null, null, null, null],
    [],
    ['Period', 'Item ID', 'Description', 'Monthly Change', 'Cumulative Change'],
    ['Jan 2024', 'OP001', 'Office refresh'],
    ['Feb 2024', 'OP001', 'Office refresh'],
    ['Mar 2024', 'OP002', 'Data migration'],
  ])
  rollforwardCells.push(
    formulaCell('D4', 'VLOOKUP(B4,Ledger!A:F,6,FALSE())/12'),
    formulaCell('E4', 'D4'),
    formulaCell('D5', 'VLOOKUP(B5,Ledger!A:F,6,FALSE())/12'),
    formulaCell('E5', 'IF(B5=B4,E4+D5,D5)'),
    formulaCell('D6', 'VLOOKUP(B6,Ledger!A:F,6,FALSE())/12'),
    formulaCell('E6', 'IF(B6=B5,E5+D6,D6)'),
  )

  return writeWorkbook([
    {
      name: 'Dashboard',
      cells: dashboardCells,
      columns: columnSizes([180, 118, 96, 96]),
      rows: rowSizes([30, null, 24]),
      merges: [merge('A1', 'D1')],
    },
    {
      name: 'Ledger',
      cells: ledgerCells,
      columns: columnSizes([132, 96, 142, 210, 138, 118, 138, 92]),
      rows: rowSizes([30, null, 24]),
      merges: [merge('A1', 'H1')],
    },
    {
      name: 'Rollforward',
      cells: rollforwardCells,
      columns: columnSizes([112, 96, 210, 126, 148]),
      rows: rowSizes([30, null, 24]),
      merges: [merge('A1', 'E1')],
    },
    {
      name: 'Lookups',
      cells: fixtureCells([['Category'], ['Capital'], ['Platform']]),
    },
  ])
}

function buildSingleSheetPlanningWorkbook(): Uint8Array {
  const cells = fixtureCells([
    ['Monthly Planning Schedule', null, null, null, null, null, null, null, null],
    ['Owner', 'Workstream', 'Start Date', 'End Date', 'Budget', 'Jan 2026', 'Feb 2026', 'Planned', 'Remaining'],
    ['TenantWorks', 'Facilities platform', 46054, 46234, 6600],
    ['Blue Harbor', 'Insurance binder', 46023, 46388, 12000],
  ])
  cells.push(
    formulaCell('F3', 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)'),
    formulaCell('G3', 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,2,1),0))-MAX($C3,DATE(2026,2,1))+1)/($D3-$C3+1),0),2)'),
    formulaCell('H3', 'ROUND(SUM(F3:G3),2)'),
    formulaCell('I3', 'ROUND(E3-H3,2)'),
    formulaCell('F4', 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,1,1),0))-MAX($C4,DATE(2026,1,1))+1)/($D4-$C4+1),0),2)'),
    formulaCell('G4', 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,2,1),0))-MAX($C4,DATE(2026,2,1))+1)/($D4-$C4+1),0),2)'),
    formulaCell('H4', 'ROUND(SUM(F4:G4),2)'),
    formulaCell('I4', 'ROUND(E4-H4,2)'),
  )
  return writeWorkbook([
    {
      name: 'Monthly Plan',
      cells,
      columns: columnSizes([168, 190, 104, 104, 118, 96, 96, 134, 138]),
      rows: rowSizes([30, 24]),
      merges: [merge('A1', 'I1')],
    },
  ])
}

async function writeFixture(testInfo: { outputPath: (pathSegment: string) => string }, name: string, bytes: Uint8Array): Promise<string> {
  const path = testInfo.outputPath(`${name}.xlsx`)
  await writeFile(path, bytes)
  return path
}

function normalizeWorkbookName(fileName: string): string {
  return basename(fileName).replace(/\.(xlsx|csv)$/i, '') || 'Imported workbook'
}

function readExternalWorkbookExpectation(path: string): WorkbookImportExpectation {
  const workbook = readXlsxWorkbookCells(readFileSync(path))
  const cells: WorkbookImportExpectation['cells'] = []
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells) {
      if (cells.length >= 6) {
        break
      }
      if (typeof cell.formula === 'string' && cell.formula.trim().length > 0) {
        cells.push({ sheetName: sheet.name, address: cell.address, value: `=${cell.formula}` })
        continue
      }
      if (typeof cell.value === 'string' && cell.value.trim().length > 0) {
        cells.push({ sheetName: sheet.name, address: cell.address, value: cell.value })
      }
    }
  }
  const sheetNames = workbook.sheets.map((sheet) => sheet.name)
  if (sheetNames.length === 0 || cells.length === 0) {
    throw new Error('External workbook verifier needs at least one sheet and one string or formula cell')
  }
  return {
    workbookName: normalizeWorkbookName(path),
    uploadFileName: basename(path),
    sheetNames,
    activeSheetName: sheetNames[0] ?? 'Sheet1',
    cells,
  }
}

async function importWorkbookThroughUi(page: Page, path: string, expectation: WorkbookImportExpectation): Promise<void> {
  await page.getByTestId('workbook-import-toggle').click()
  await page.getByTestId('workbook-import-file').setInputFiles({
    name: expectation.uploadFileName ?? `${expectation.workbookName}.xlsx`,
    mimeType: XLSX_MIME_TYPE,
    buffer: await readFile(path),
  })
  await expect(page.getByTestId('workbook-import-preview-list')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByText(expectation.workbookName, { exact: true })).toBeVisible()
  await Promise.all(
    expectation.sheetNames.map(async (sheetName) => expect(page.getByText(sheetName, { exact: true }).first()).toBeVisible()),
  )

  await page.getByTestId('workbook-import-create').click()
  await page.waitForURL(/document=xlsx%3A/, { timeout: 15_000 })
  await waitForWorkbookReady(page)
  await Promise.all(expectation.sheetNames.map(async (sheetName) => expect(page.getByRole('tab', { name: sheetName })).toBeVisible()))
}

async function expectImportedCell(page: Page, sheetName: string, address: string, value: string): Promise<void> {
  await page.getByRole('tab', { name: sheetName }).click()
  const nameBox = page.getByTestId('name-box')
  await nameBox.fill(address)
  await nameBox.press('Enter')
  await expect(nameBox).toHaveValue(address)
  await expect(page.getByTestId('formula-input')).toHaveValue(value)
}

const generatedFixtures: readonly {
  readonly name: string
  readonly bytes: () => Uint8Array
  readonly expectation: WorkbookImportExpectation
}[] = [
  {
    name: 'multi-sheet-operations',
    bytes: buildMultiSheetOperationsWorkbook,
    expectation: {
      workbookName: 'multi-sheet-operations',
      sheetNames: ['Dashboard', 'Ledger', 'Rollforward', 'Lookups'],
      activeSheetName: 'Ledger',
      expectedColumnWidth: 132,
      cells: [
        { sheetName: 'Dashboard', address: 'A1', value: 'OPERATIONS DASHBOARD' },
        { sheetName: 'Dashboard', address: 'B4', value: '=SUM(Ledger!F:F)' },
        { sheetName: 'Ledger', address: 'A4', value: 'OP001' },
        { sheetName: 'Ledger', address: 'B4', value: '45292' },
        { sheetName: 'Ledger', address: 'G4', value: '=F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)' },
        { sheetName: 'Rollforward', address: 'E5', value: '=IF(B5=B4,E4+D5,D5)' },
      ],
    },
  },
  {
    name: 'single-sheet-planning',
    bytes: buildSingleSheetPlanningWorkbook,
    expectation: {
      workbookName: 'single-sheet-planning',
      sheetNames: ['Monthly Plan'],
      activeSheetName: 'Monthly Plan',
      expectedColumnWidth: 168,
      cells: [
        { sheetName: 'Monthly Plan', address: 'A1', value: 'Monthly Planning Schedule' },
        { sheetName: 'Monthly Plan', address: 'A3', value: 'TenantWorks' },
        {
          sheetName: 'Monthly Plan',
          address: 'F3',
          value: '=ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
        },
        { sheetName: 'Monthly Plan', address: 'I4', value: '=ROUND(E4-H4,2)' },
      ],
    },
  },
]

for (const fixture of generatedFixtures) {
  test(`web app imports generated workbook fixture: ${fixture.name}`, async ({ page }, testInfo) => {
    await gotoWorkbookShell(page)
    await waitForWorkbookReady(page)

    const path = await writeFixture(testInfo, fixture.name, fixture.bytes())
    await importWorkbookThroughUi(page, path, fixture.expectation)

    await page.getByRole('tab', { name: fixture.expectation.activeSheetName }).click()
    await expect.poll(async () => await getProductColumnWidth(page, 0), { timeout: 15_000 }).toBe(fixture.expectation.expectedColumnWidth)
    for (const cell of fixture.expectation.cells) {
      // oxlint-disable-next-line eslint(no-await-in-loop)
      await expectImportedCell(page, cell.sheetName, cell.address, cell.value)
    }
  })
}

test('web app imports an external workbook when BILIG_REFERENCE_WORKBOOK_XLSX is set', async ({ page }) => {
  const referencePath = process.env['BILIG_REFERENCE_WORKBOOK_XLSX']
  test.skip(!referencePath, 'Set BILIG_REFERENCE_WORKBOOK_XLSX to verify a local workbook through the normal import UI.')

  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)

  const expectation = readExternalWorkbookExpectation(referencePath)
  await importWorkbookThroughUi(page, referencePath, expectation)

  for (const cell of expectation.cells.slice(0, 4)) {
    // oxlint-disable-next-line eslint(no-await-in-loop)
    await expectImportedCell(page, cell.sheetName, cell.address, cell.value)
  }
})
