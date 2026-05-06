import { basename } from 'node:path'
import { readFile, writeFile } from 'node:fs/promises'
import { expect, type Page, test } from '@playwright/test'
import * as XLSX from 'xlsx'
import { getProductColumnWidth, gotoWorkbookShell, waitForWorkbookReady } from './web-shell-helpers.js'

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

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

function writeWorkbook(workbook: XLSX.WorkBook): Uint8Array {
  const bytes: unknown = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'buffer',
  })
  if (!(bytes instanceof Uint8Array)) {
    throw new Error('Expected XLSX writer to return workbook bytes')
  }
  return bytes
}

function buildMultiSheetOperationsWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()

  const dashboard = XLSX.utils.aoa_to_sheet([
    ['OPERATIONS DASHBOARD', null, null, null],
    [],
    ['Metric', 'Value'],
    ['Total budget'],
    ['Open balance'],
    ['Completion rate'],
  ])
  dashboard.B4 = { t: 'n', f: 'SUM(Ledger!F:F)' }
  dashboard.B5 = { t: 'n', f: 'SUMIF(Ledger!H:H,"Open",Ledger!G:G)' }
  dashboard.B6 = { t: 'n', f: 'IF(B4>0,1-B5/B4,0)' }
  dashboard['!ref'] = 'A1:D6'
  dashboard['!cols'] = [{ wpx: 180 }, { wpx: 118 }, { wpx: 96 }, { wpx: 96 }]
  dashboard['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  dashboard['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }]

  const ledger = XLSX.utils.aoa_to_sheet([
    ['OPERATIONS LEDGER', null, null, null, null, null, null, null],
    [],
    ['ID', 'Date', 'Owner', 'Workstream', 'Category', 'Budget', 'Open Balance', 'Status'],
    ['OP001', 45292, 'Facilities', 'Office refresh', 'Capital', 12000, null, 'Open'],
    ['OP002', 45323, 'Engineering', 'Data migration', 'Platform', 18000, null, 'Open'],
  ])
  ledger.G4 = { t: 'n', f: 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)' }
  ledger.G5 = { t: 'n', f: 'F5-SUMIF(Rollforward!$B:$B,A5,Rollforward!$E:$E)' }
  ledger['!ref'] = 'A1:H5'
  ledger['!cols'] = [{ wpx: 132 }, { wpx: 96 }, { wpx: 142 }, { wpx: 210 }, { wpx: 138 }, { wpx: 118 }, { wpx: 138 }, { wpx: 92 }]
  ledger['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  ledger['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }]

  const rollforward = XLSX.utils.aoa_to_sheet([
    ['ROLLFORWARD', null, null, null, null],
    [],
    ['Period', 'Item ID', 'Description', 'Monthly Change', 'Cumulative Change'],
    ['Jan 2024', 'OP001', 'Office refresh'],
    ['Feb 2024', 'OP001', 'Office refresh'],
    ['Mar 2024', 'OP002', 'Data migration'],
  ])
  rollforward.D4 = { t: 'n', f: 'VLOOKUP(B4,Ledger!A:F,6,FALSE())/12' }
  rollforward.E4 = { t: 'n', f: 'D4' }
  rollforward.D5 = { t: 'n', f: 'VLOOKUP(B5,Ledger!A:F,6,FALSE())/12' }
  rollforward.E5 = { t: 'n', f: 'IF(B5=B4,E4+D5,D5)' }
  rollforward.D6 = { t: 'n', f: 'VLOOKUP(B6,Ledger!A:F,6,FALSE())/12' }
  rollforward.E6 = { t: 'n', f: 'IF(B6=B5,E5+D6,D6)' }
  rollforward['!ref'] = 'A1:E6'
  rollforward['!cols'] = [{ wpx: 112 }, { wpx: 96 }, { wpx: 210 }, { wpx: 126 }, { wpx: 148 }]
  rollforward['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  rollforward['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }]

  XLSX.utils.book_append_sheet(workbook, dashboard, 'Dashboard')
  XLSX.utils.book_append_sheet(workbook, ledger, 'Ledger')
  XLSX.utils.book_append_sheet(workbook, rollforward, 'Rollforward')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Category'], ['Capital'], ['Platform']]), 'Lookups')

  return writeWorkbook(workbook)
}

function buildSingleSheetPlanningWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Monthly Planning Schedule', null, null, null, null, null, null, null, null],
    ['Owner', 'Workstream', 'Start Date', 'End Date', 'Budget', 'Jan 2026', 'Feb 2026', 'Planned', 'Remaining'],
    ['TenantWorks', 'Facilities platform', 46054, 46234, 6600],
    ['Blue Harbor', 'Insurance binder', 46023, 46388, 12000],
  ])
  sheet.F3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)' }
  sheet.G3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,2,1),0))-MAX($C3,DATE(2026,2,1))+1)/($D3-$C3+1),0),2)' }
  sheet.H3 = { t: 'n', f: 'ROUND(SUM(F3:G3),2)' }
  sheet.I3 = { t: 'n', f: 'ROUND(E3-H3,2)' }
  sheet.F4 = { t: 'n', f: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,1,1),0))-MAX($C4,DATE(2026,1,1))+1)/($D4-$C4+1),0),2)' }
  sheet.G4 = { t: 'n', f: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,2,1),0))-MAX($C4,DATE(2026,2,1))+1)/($D4-$C4+1),0),2)' }
  sheet.H4 = { t: 'n', f: 'ROUND(SUM(F4:G4),2)' }
  sheet.I4 = { t: 'n', f: 'ROUND(E4-H4,2)' }
  sheet['!ref'] = 'A1:I4'
  sheet['!cols'] = [
    { wpx: 168 },
    { wpx: 190 },
    { wpx: 104 },
    { wpx: 104 },
    { wpx: 118 },
    { wpx: 96 },
    { wpx: 96 },
    { wpx: 134 },
    { wpx: 138 },
  ]
  sheet['!rows'] = [{ hpx: 30 }, { hpx: 24 }]
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }]
  XLSX.utils.book_append_sheet(workbook, sheet, 'Monthly Plan')
  return writeWorkbook(workbook)
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
  const workbook = XLSX.readFile(path, {
    cellFormula: true,
    cellText: false,
  })
  const cells: WorkbookImportExpectation['cells'] = []
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet?.['!ref']) {
      continue
    }
    const range = XLSX.utils.decode_range(sheet['!ref'])
    for (let row = range.s.r; row <= range.e.r && cells.length < 6; row += 1) {
      for (let col = range.s.c; col <= range.e.c && cells.length < 6; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        const cell = sheet[address]
        if (!cell) {
          continue
        }
        if (typeof cell.f === 'string' && cell.f.trim().length > 0) {
          cells.push({ sheetName, address, value: `=${cell.f}` })
          continue
        }
        if (typeof cell.v === 'string' && cell.v.trim().length > 0) {
          cells.push({ sheetName, address, value: cell.v })
        }
      }
    }
  }
  if (workbook.SheetNames.length === 0 || cells.length === 0) {
    throw new Error('External workbook verifier needs at least one sheet and one string or formula cell')
  }
  return {
    workbookName: normalizeWorkbookName(path),
    uploadFileName: basename(path),
    sheetNames: workbook.SheetNames,
    activeSheetName: workbook.SheetNames[0] ?? 'Sheet1',
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
