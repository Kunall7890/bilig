import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

const workpaperPackageVersion = '0.164.4'
const nextWinRate = Number(process.argv[2] ?? '0.4')
if (!Number.isFinite(nextWinRate) || nextWinRate < 0 || nextWinRate > 1) {
  throw new Error('win rate must be a number between 0 and 1')
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
    ['Expansion multiplier', 1.1],
    ['ARR target', 100000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ['Expected ARR', '=B2*Inputs!B4'],
    ['Expansion ARR', '=B3*Inputs!B5'],
    ['Target gap', '=B4-Inputs!B6'],
  ],
})

try {
  const before = readSummary(workbook)
  const inputBefore = readNumber(workbook, 'Inputs!B3')
  const editedCell = requireAddress(workbook, 'Inputs!B3')

  workbook.setCellContents(editedCell, nextWinRate)

  const after = readSummary(workbook)
  const persistedDocument = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(persistedDocument))
  const afterRestore = readSummary(restored)
  const restoredInput = readNumber(restored, 'Inputs!B3')

  const verified =
    nearlyEqual(after.expectedCustomers, 20 * nextWinRate) &&
    nearlyEqual(after.expectedArr, after.expectedCustomers * 12000) &&
    nearlyEqual(after.expansionArr, after.expectedArr * 1.1) &&
    nearlyEqual(after.targetGap, after.expansionArr - 100000) &&
    nearlyEqual(restoredInput, nextWinRate) &&
    sameJson(after, afterRestore) &&
    persistedDocument.includes('=Inputs!B2*Inputs!B3')

  const payload = {
    schemaVersion: 'bilig.huggingface.workpaper-readback.v1',
    package: '@bilig/workpaper',
    packageVersion: workpaperPackageVersion,
    editedCell: 'Inputs!B3',
    input: {
      previousWinRate: inputBefore,
      nextWinRate,
    },
    formulas: {
      expectedCustomers: '=Inputs!B2*Inputs!B3',
      expectedArr: '=B2*Inputs!B4',
      expansionArr: '=B3*Inputs!B5',
      targetGap: '=B4-Inputs!B6',
    },
    before,
    after,
    afterRestore,
    persistedDocumentBytes: Buffer.byteLength(persistedDocument, 'utf8'),
    checks: {
      formulaOutputChanged: !sameJson(before, after),
      restoredMatchesAfter: sameJson(after, afterRestore),
      inputPersisted: nearlyEqual(restoredInput, nextWinRate),
      formulasPersisted: persistedDocument.includes('=Inputs!B2*Inputs!B3'),
    },
    limitations: [
      'This public Space is a no-key readback fixture.',
      'Use the file-backed WorkPaper MCP server for private persisted workbook files.',
    ],
    verified,
  }

  if (!verified) {
    console.error(JSON.stringify(payload, null, 2))
    process.exitCode = 1
  } else {
    console.log(JSON.stringify(payload, null, 2))
  }

  restored.dispose()
} finally {
  workbook.dispose()
}

function readSummary(target) {
  return {
    expectedCustomers: readNumber(target, 'Summary!B2'),
    expectedArr: readNumber(target, 'Summary!B3'),
    expansionArr: readNumber(target, 'Summary!B4'),
    targetGap: readNumber(target, 'Summary!B5'),
  }
}

function readNumber(target, address) {
  const parsed = requireAddress(target, address)
  const cell = target.getCellValue(parsed)
  if (typeof cell !== 'object' || cell === null || typeof cell.value !== 'number') {
    throw new Error(`expected numeric cell at ${address}, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 1000000) / 1000000
}

function requireAddress(target, address) {
  const parsed = target.simpleCellAddressFromString(address)
  if (parsed === undefined) {
    throw new Error(`unknown cell: ${address}`)
  }
  return parsed
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right)
}

function nearlyEqual(left, right) {
  return Math.abs(left - right) < 0.000001
}
