import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { WorkPaper, exportXlsx } from '../../../packages/xlsx-formula-recalc/dist/index.js'

const exampleDir = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const outputPath = resolve(exampleDir, 'fixtures/stale-pricing.xlsx')

const rows: Array<[number | string, number | string]> = [['Input', 'Output']]
for (let row = 2; row <= 61; row += 1) {
  rows.push([row - 1, `=A${row}*10`])
}

const workbook = WorkPaper.buildFromSheets({
  Sheet1: rows,
})

mkdirSync(dirname(outputPath), { recursive: true })
try {
  writeFileSync(
    outputPath,
    Buffer.from(
      replaceWorksheetCellXml(
        exportXlsx(workbook.exportSnapshot()),
        'xl/worksheets/sheet1.xml',
        'B61',
        '<c r="B61"><f>A61*10</f><v>999</v></c>',
      ),
    ),
  )
} finally {
  workbook.dispose()
}

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  zip[path] = strToU8(xml.replace(new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'), replacement))
  return zipSync(zip)
}
