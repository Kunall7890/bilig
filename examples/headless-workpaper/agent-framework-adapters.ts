import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type SetInputCellArgs = {
  sheetName: string
  address: string
  value: string | number | boolean | null
}
type ReadSummaryArgs = {
  range?: string
}
type WorkPaperToolSet = ReturnType<typeof createWorkPaperTools>
type WorkPaperWriteResult = ReturnType<WorkPaperToolSet['setWorkPaperInputCell']>
type LangChainReadTool = {
  name: 'read_workpaper_summary'
  description: string
  schema: Record<string, string>
  invoke(args?: ReadSummaryArgs): ReturnType<WorkPaperToolSet['readWorkPaperSummary']>
}
type LangChainWriteTool = {
  name: 'set_workpaper_input_cell'
  description: string
  schema: Record<string, string>
  invoke(args: SetInputCellArgs): WorkPaperWriteResult
}
type LangChainTool = LangChainReadTool | LangChainWriteTool

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

function createAiSdkTools(workPaperTools: WorkPaperToolSet) {
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
      execute({ range = 'Summary!A1:B5' }: ReadSummaryArgs = {}) {
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
      execute(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(args)
      },
    },
  }
}

function createLangChainTools(workPaperTools: WorkPaperToolSet): LangChainTool[] {
  return [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      schema: {
        range: 'string',
      },
      invoke({ range = 'Summary!A1:B5' }: ReadSummaryArgs = {}) {
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
      invoke(args: SetInputCellArgs) {
        return workPaperTools.setWorkPaperInputCell(args)
      },
    },
  ]
}

function createWorkPaperTools(workbook: WorkPaperInstance) {
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

    setWorkPaperInputCell({ sheetName, address, value }: SetInputCellArgs) {
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

function requireTool(tools: LangChainTool[], name: 'read_workpaper_summary'): LangChainReadTool
function requireTool(tools: LangChainTool[], name: 'set_workpaper_input_cell'): LangChainWriteTool
function requireTool(tools: LangChainTool[], name: LangChainTool['name']): LangChainTool {
  const tool = tools.find((candidate) => candidate.name === name)
  if (tool === undefined) {
    throw new Error(`Missing framework tool: ${name}`)
  }
  return tool
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireCellAddress(workpaper: WorkPaperInstance, sheetName: string, a1Address: string): CellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readSummary(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function serializeWorkbook(workpaper: WorkPaperInstance): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual: typeof output): void {
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

  const writeResults: [string, WorkPaperWriteResult][] = [
    ['aiSdk', actual.aiSdk.writeResult],
    ['langChain', actual.langChain.writeResult],
  ]

  for (const [framework, result] of writeResults) {
    if (
      result.editedCell !== 'Inputs!B3' ||
      !sameJson(result.before, expectedBefore) ||
      !sameJson(result.after, expectedAfter) ||
      !sameJson(result.restored, expectedAfter) ||
      !sameJson(result.formulaContracts, expectedFormulaContracts) ||
      result.checks.previousValue !== 0.25 ||
      result.checks.newValue !== 0.4 ||
      !result.checks.formulasPersisted ||
      !result.checks.restoredMatchesAfter ||
      !result.checks.expectedArrChanged ||
      result.checks.serializedBytes <= 0
    ) {
      throw new Error(`Unexpected ${framework} adapter result: ${JSON.stringify(result)}`)
    }
  }
}
