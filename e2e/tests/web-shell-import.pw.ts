import { readFile, writeFile } from 'node:fs/promises'
import { expect, type Page, test } from '@playwright/test'
import * as XLSX from 'xlsx'
import { getProductColumnWidth, gotoWorkbookShell, waitForWorkbookReady } from './web-shell-helpers.js'

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

interface PrepaidImportExpectation {
  readonly workbookName: string
  readonly sheetNames: readonly string[]
  readonly activeSheetName: string
  readonly expectedColumnWidth: number
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

function buildReferenceStylePrepaidWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()

  const summary = XLSX.utils.aoa_to_sheet([
    ['PREPAID EXPENSE DASHBOARD', null, null, null],
    [],
    ['KEY METRICS'],
    ['Metric', 'Value'],
    ['Total prepaids'],
    ['Active balance'],
    ['Average monthly amortization'],
  ])
  summary.B5 = { t: 'n', f: "SUM('Prepaid Tracking'!F:F)" }
  summary.B6 = { t: 'n', f: "SUMIF('Prepaid Tracking'!L:L,\"Active\",'Prepaid Tracking'!K:K)" }
  summary.B7 = { t: 'n', f: 'IF(B5>0,B6/B5,0)' }
  summary['!ref'] = 'A1:D7'
  summary['!cols'] = [{ wpx: 180 }, { wpx: 118 }, { wpx: 96 }, { wpx: 96 }]
  summary['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  summary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }]

  const tracking = XLSX.utils.aoa_to_sheet([
    ['PREPAID EXPENSE TRACKING', null, null, null, null, null, null, null, null, null, null, null],
    [],
    [
      'ID',
      'Date Paid',
      'Vendor',
      'Description',
      'Category',
      'Total Amount',
      'Start Date',
      'End Date',
      'Life Months',
      'Monthly Amount',
      'Remaining Balance',
      'Status',
    ],
    ['PE001', 45292, 'Acme Insurance', 'Annual insurance premium', 'Insurance', 12000, 45292, 45657, null, null, null, 'Active'],
    ['PE002', 45323, 'Northstar SaaS', 'Platform subscription', 'Software Licenses', 18000, 45323, 45687, null, null, null, 'Active'],
  ])
  tracking.I4 = { t: 'n', f: 'DATEDIF(G4,H4,"M")+1' }
  tracking.J4 = { t: 'n', f: 'F4/I4' }
  tracking.K4 = { t: 'n', f: "F4-SUMIF('Amortization Schedule'!$B:$B,A4,'Amortization Schedule'!$E:$E)" }
  tracking.I5 = { t: 'n', f: 'DATEDIF(G5,H5,"M")+1' }
  tracking.J5 = { t: 'n', f: 'F5/I5' }
  tracking.K5 = { t: 'n', f: "F5-SUMIF('Amortization Schedule'!$B:$B,A5,'Amortization Schedule'!$E:$E)" }
  tracking['!ref'] = 'A1:L5'
  tracking['!cols'] = [
    { wpx: 132 },
    { wpx: 96 },
    { wpx: 142 },
    { wpx: 210 },
    { wpx: 138 },
    { wpx: 118 },
    { wpx: 96 },
    { wpx: 96 },
    { wpx: 92 },
    { wpx: 118 },
    { wpx: 138 },
    { wpx: 92 },
  ]
  tracking['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  tracking['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }]

  const amortization = XLSX.utils.aoa_to_sheet([
    ['AMORTIZATION SCHEDULE', null, null, null, null, null],
    [],
    ['Month', 'Prepaid ID', 'Description', 'Monthly Amount', 'Cumulative Amortized', 'Remaining Balance'],
    ['Jan 2024', 'PE001', 'Annual insurance premium'],
    ['Feb 2024', 'PE001', 'Annual insurance premium'],
    ['Mar 2024', 'PE002', 'Platform subscription'],
  ])
  amortization.D4 = { t: 'n', f: "VLOOKUP(B4,'Prepaid Tracking'!A:J,10,FALSE())" }
  amortization.E4 = { t: 'n', f: 'D4' }
  amortization.F4 = { t: 'n', f: "VLOOKUP(B4,'Prepaid Tracking'!A:F,6,FALSE())-E4" }
  amortization.D5 = { t: 'n', f: "VLOOKUP(B5,'Prepaid Tracking'!A:J,10,FALSE())" }
  amortization.E5 = { t: 'n', f: 'IF(B5=B4,E4+D5,D5)' }
  amortization.F5 = { t: 'n', f: "VLOOKUP(B5,'Prepaid Tracking'!A:F,6,FALSE())-E5" }
  amortization.D6 = { t: 'n', f: "VLOOKUP(B6,'Prepaid Tracking'!A:J,10,FALSE())" }
  amortization.E6 = { t: 'n', f: 'IF(B6=B5,E5+D6,D6)' }
  amortization.F6 = { t: 'n', f: "VLOOKUP(B6,'Prepaid Tracking'!A:F,6,FALSE())-E6" }
  amortization['!ref'] = 'A1:F6'
  amortization['!cols'] = [{ wpx: 112 }, { wpx: 96 }, { wpx: 210 }, { wpx: 126 }, { wpx: 148 }, { wpx: 138 }]
  amortization['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
  amortization['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }]

  XLSX.utils.book_append_sheet(workbook, summary, 'Summary Dashboard')
  XLSX.utils.book_append_sheet(workbook, tracking, 'Prepaid Tracking')
  XLSX.utils.book_append_sheet(workbook, amortization, 'Amortization Schedule')
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([['Expense Categories'], ['Insurance'], ['Software Licenses']]),
    'Categories',
  )

  return writeWorkbook(workbook)
}

function buildSingleSheetDailyPrepaidWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Daily Prepaid Schedule', null, null, null, null, null, null, null, null],
    ['Vendor', 'Description', 'Start Date', 'End Date', 'Total Amount', 'Jan 2026', 'Feb 2026', '2026 Amortized', 'Remaining Balance'],
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
  XLSX.utils.book_append_sheet(workbook, sheet, 'Daily Prepaids')
  return writeWorkbook(workbook)
}

async function writeFixture(testInfo: { outputPath: (pathSegment: string) => string }, name: string, bytes: Uint8Array): Promise<string> {
  const path = testInfo.outputPath(`${name}.xlsx`)
  await writeFile(path, bytes)
  return path
}

async function importWorkbookThroughUi(page: Page, path: string, expectation: PrepaidImportExpectation): Promise<void> {
  await page.getByTestId('workbook-import-toggle').click()
  await page.getByTestId('workbook-import-file').setInputFiles({
    name: `${expectation.workbookName}.xlsx`,
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
  readonly expectation: PrepaidImportExpectation
}[] = [
  {
    name: 'reference-style-prepaid',
    bytes: buildReferenceStylePrepaidWorkbook,
    expectation: {
      workbookName: 'reference-style-prepaid',
      sheetNames: ['Summary Dashboard', 'Prepaid Tracking', 'Amortization Schedule', 'Categories'],
      activeSheetName: 'Prepaid Tracking',
      expectedColumnWidth: 132,
      cells: [
        { sheetName: 'Summary Dashboard', address: 'A1', value: 'PREPAID EXPENSE DASHBOARD' },
        { sheetName: 'Summary Dashboard', address: 'B5', value: "=SUM('Prepaid Tracking'!F:F)" },
        { sheetName: 'Prepaid Tracking', address: 'A4', value: 'PE001' },
        { sheetName: 'Prepaid Tracking', address: 'B4', value: '45292' },
        {
          sheetName: 'Prepaid Tracking',
          address: 'K4',
          value: "=F4-SUMIF('Amortization Schedule'!$B:$B,A4,'Amortization Schedule'!$E:$E)",
        },
        { sheetName: 'Amortization Schedule', address: 'E5', value: '=IF(B5=B4,E4+D5,D5)' },
      ],
    },
  },
  {
    name: 'single-sheet-daily-prepaid',
    bytes: buildSingleSheetDailyPrepaidWorkbook,
    expectation: {
      workbookName: 'single-sheet-daily-prepaid',
      sheetNames: ['Daily Prepaids'],
      activeSheetName: 'Daily Prepaids',
      expectedColumnWidth: 168,
      cells: [
        { sheetName: 'Daily Prepaids', address: 'A1', value: 'Daily Prepaid Schedule' },
        { sheetName: 'Daily Prepaids', address: 'A3', value: 'TenantWorks' },
        {
          sheetName: 'Daily Prepaids',
          address: 'F3',
          value: '=ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
        },
        { sheetName: 'Daily Prepaids', address: 'I4', value: '=ROUND(E4-H4,2)' },
      ],
    },
  },
]

for (const fixture of generatedFixtures) {
  test(`web app imports generated prepaid workbook fixture: ${fixture.name}`, async ({ page }, testInfo) => {
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

test('web app imports the external referenced prepaid workbook when BILIG_REFERENCE_PREPAID_XLSX is set', async ({ page }) => {
  const referencePath = process.env['BILIG_REFERENCE_PREPAID_XLSX']
  test.skip(!referencePath, 'Set BILIG_REFERENCE_PREPAID_XLSX to verify the local referenced prepaid workbook.')

  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)

  await importWorkbookThroughUi(page, referencePath, {
    workbookName: 'Prepaid Expense Template',
    sheetNames: ['Summary Dashboard', 'Prepaid Tracking', 'Amortization Schedule', 'Categories'],
    activeSheetName: 'Prepaid Tracking',
    expectedColumnWidth: 11,
    cells: [
      { sheetName: 'Summary Dashboard', address: 'A1', value: 'PREPAID EXPENSE DASHBOARD' },
      { sheetName: 'Summary Dashboard', address: 'B5', value: "=SUM('Prepaid Tracking'!F:F)" },
      { sheetName: 'Prepaid Tracking', address: 'A4', value: 'PE001' },
      { sheetName: 'Prepaid Tracking', address: 'B4', value: '01/01/2024' },
      { sheetName: 'Prepaid Tracking', address: 'K4', value: "=F4-SUMIF('Amortization Schedule'!$B:$B,A4,'Amortization Schedule'!$E:$E)" },
      { sheetName: 'Amortization Schedule', address: 'E5', value: '=IF(B5=B4,E4+D5,D5)' },
    ],
  })

  await page.getByRole('tab', { name: 'Prepaid Tracking' }).click()
  expect(await getProductColumnWidth(page, 0)).toBeGreaterThan(10)
})
