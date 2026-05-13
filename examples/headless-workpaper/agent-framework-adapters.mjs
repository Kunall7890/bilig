import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const aiSdkWorkbook = buildWorkbook()
const langChainWorkbook = buildWorkbook()
const aiSdkTools = createAiSdkTools(createWorkPaperTools(aiSdkWorkbook))
const langChainTools = createLangChainTools(createWorkPaperTools(langChainWorkbook))

const output = {
  aiSdk: {
    toolNames: Object.keys(aiSdkTools),
    readResult: aiSdkTools.readWorkPaperSummary.execute({
      range: 'Summary!A1:B5',
    }),
    writeResult: aiSdkTools.setWorkPaperInputCell.execute({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
  langChain: {
    toolNames: langChainTools.map((tool) => tool.name),
    readResult: requireTool(langChainTools, 'read_workpaper_summary').invoke({
      range: 'Summary!A1:B5',
    }),
    writeResult: requireTool(langChainTools, 'set_workpaper_input_cell').invoke({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  },
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function buildWorkbook() {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Qualified opportunities', 20],
      ['Win rate', 0.25],
      ['Average ARR', 12000],
      ['Expansion multiplier', 1.1],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ['Expected ARR', '=B2*Inputs!B4'],
      ['Expansion ARR', '=B3*Inputs!B5'],
      ['Target gap', '=B4-100000'],
    ],
  })
}

function createAiSdkTools(workPaperTools) {
  return {
    readWorkPaperSummary: {
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            default: 'Summary!A1:B5',
          },
        },
      },
      execute({ range = 'Summary!A1:B5' } = {}) {
        return workPaperTools.readWorkPaperSummary(range)
      },
    },

    setWorkPaperInputCell: {
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
          },
          address: {
            type: 'string',
          },
          value: {
            oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
          },
        },
      },
      execute(args) {
        return workPaperTools.setWorkPaperInputCell(args)
      },
    },
  }
}

function createLangChainTools(workPaperTools) {
  return [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      schema: {
        range: 'string',
      },
      invoke({ range = 'Summary!A1:B5' } = {}) {
        return workPaperTools.readWorkPaperSummary(range)
      },
    },
    {
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      schema: {
        sheetName: 'string',
        address: 'string',
        value: 'string | number | boolean | null',
      },
      invoke(args) {
        return workPaperTools.setWorkPaperInputCell(args)
      },
    },
  ]
}

function createWorkPaperTools(workbook) {
  const summarySheet = requireSheet(workbook, 'Summary')

  return {
    readWorkPaperSummary(range = 'Summary!A1:B5') {
      const parsedRange = workbook.simpleCellRangeFromString(range, summarySheet)
      if (parsedRange === undefined) {
        throw new Error(`Invalid readable range: ${range}`)
      }

      return {
        range,
        values: workbook.getRangeValues(parsedRange),
        serialized: workbook.getRangeSerialized(parsedRange),
      }
    },

    setWorkPaperInputCell({ sheetName, address, value }) {
      if (sheetName !== 'Inputs') {
        throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
      }

      const target = requireCellAddress(workbook, sheetName, address)
      const before = readSummary(workbook, summarySheet)
      const formulaContracts = readFormulaContracts(workbook, summarySheet)
      const previousValue = workbook.getCellSerialized(target)

      workbook.setCellContents(target, value)

      const after = readSummary(workbook, summarySheet)
      const serialized = serializeWorkbook(workbook)
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const restoredSummary = readSummary(restored, restoredSummarySheet)
      const restoredFormulaContracts = readFormulaContracts(restored, restoredSummarySheet)

      return {
        editedCell: workbook.simpleCellAddressToString(target, {
          includeSheetName: true,
        }),
        before,
        after,
        restored: restoredSummary,
        formulaContracts,
        checks: {
          previousValue,
          newValue: workbook.getCellSerialized(target),
          formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
          restoredMatchesAfter: sameJson(after, restoredSummary),
          expectedArrChanged: after.expectedArr > before.expectedArr,
          serializedBytes: Buffer.byteLength(serialized, 'utf8'),
        },
      }
    },
  }
}

function requireTool(tools, name) {
  const tool = tools.find((candidate) => candidate.name === name)
  if (tool === undefined) {
    throw new Error(`Missing framework tool: ${name}`)
  }
  return tool
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireCellAddress(workpaper, sheetName, a1Address) {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readSummary(workpaper, summary) {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper, summary) {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper, sheet, row, col, label) {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function serializeWorkbook(workpaper) {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual) {
  const expectedBefore = {
    expectedCustomers: 5,
    expectedArr: 60000,
    expansionArr: 66000,
    targetGap: -34000,
  }
  const expectedAfter = {
    expectedCustomers: 8,
    expectedArr: 96000,
    expansionArr: 105600,
    targetGap: 5600,
  }
  const expectedFormulaContracts = {
    expectedCustomers: '=Inputs!B2*Inputs!B3',
    expectedArr: '=B2*Inputs!B4',
    expansionArr: '=B3*Inputs!B5',
    targetGap: '=B4-100000',
  }

  for (const [framework, result] of Object.entries({
    aiSdk: actual.aiSdk.writeResult,
    langChain: actual.langChain.writeResult,
  })) {
    if (
      result.editedCell !== 'Inputs!B3' ||
      !sameJson(result.before, expectedBefore) ||
      !sameJson(result.after, expectedAfter) ||
      !sameJson(result.restored, expectedAfter) ||
      !sameJson(result.formulaContracts, expectedFormulaContracts) ||
      result.checks.previousValue !== 0.25 ||
      result.checks.newValue !== 0.4 ||
      result.checks.formulasPersisted !== true ||
      result.checks.restoredMatchesAfter !== true ||
      result.checks.expectedArrChanged !== true ||
      result.checks.serializedBytes <= 0
    ) {
      throw new Error(`Unexpected ${framework} adapter result: ${JSON.stringify(result)}`)
    }
  }
}
