import { createTool, noopObserve } from '@mastra/core/tools'
import type { ToolExecutionContext } from '@mastra/core/tools'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'
import { z } from 'zod'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>

type SummaryReadback = {
  expectedCustomers: number
  expectedArr: number
  expansionArr: number
  targetGap: number
}

type FormulaContracts = {
  expectedCustomers: string
  expectedArr: string
  expansionArr: string
  targetGap: string
}

const readSummaryInputSchema = z.object({
  range: z.string().default('Summary!A1:B5'),
})

const readSummaryOutputSchema = z.object({
  range: z.string(),
  values: z.array(z.array(z.unknown())),
  serialized: z.array(z.array(z.unknown())),
})

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

const summaryReadbackSchema = z.object({
  expectedCustomers: z.number(),
  expectedArr: z.number(),
  expansionArr: z.number(),
  targetGap: z.number(),
})

const formulaContractsSchema = z.object({
  expectedCustomers: z.string(),
  expectedArr: z.string(),
  expansionArr: z.string(),
  targetGap: z.string(),
})

const writeOutputSchema = z.object({
  editedCell: z.string(),
  before: summaryReadbackSchema,
  after: summaryReadbackSchema,
  restored: summaryReadbackSchema,
  beforeContracts: formulaContractsSchema,
  afterContracts: formulaContractsSchema,
  checks: z.object({
    previousValue: z.unknown(),
    newValue: z.unknown(),
    formulasPersisted: z.boolean(),
    restoredMatchesAfter: z.boolean(),
    expectedArrChanged: z.boolean(),
    serializedBytes: z.number(),
  }),
})

const workpaper = buildWorkbook()
const localToolContext = {
  observe: noopObserve,
} satisfies ToolExecutionContext
const tools = {
  readWorkPaperSummary: createTool({
    id: 'read-workpaper-summary',
    description: 'Read computed Bilig WorkPaper summary values for a small range.',
    inputSchema: readSummaryInputSchema,
    outputSchema: readSummaryOutputSchema,
    execute: async ({ range = 'Summary!A1:B5' }) => readSummaryRange(workpaper, range),
  }),
  setWorkPaperInputCell: createTool({
    id: 'set-workpaper-input-cell',
    description: 'Set one Bilig WorkPaper input cell and return formula readback.',
    inputSchema: setInputCellInputSchema,
    outputSchema: writeOutputSchema,
    execute: async (input) => setInputCell(workpaper, input),
  }),
}

const readResult = readSummaryOutputSchema.parse(
  await tools.readWorkPaperSummary.execute?.(
    {
      range: 'Summary!A1:B5',
    },
    localToolContext,
  ),
)
const writeResult = writeOutputSchema.parse(
  await tools.setWorkPaperInputCell.execute?.(
    {
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    },
    localToolContext,
  ),
)

const proof = {
  apiShape: 'Mastra createTool -> execute -> WorkPaper readback',
  toolIds: [tools.readWorkPaperSummary.id, tools.setWorkPaperInputCell.id],
  readResult,
  writeResult,
}

assertProof(proof)
console.log(JSON.stringify(proof, null, 2))

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

function readSummaryRange(target: WorkPaperInstance, range: string) {
  const summarySheet = requireSheet(target, 'Summary')
  const parsedRange = target.simpleCellRangeFromString(range, summarySheet)
  if (parsedRange === undefined) {
    throw new Error(`Invalid readable WorkPaper range: ${range}`)
  }

  return {
    range,
    values: target.getRangeValues(parsedRange),
    serialized: target.getRangeSerialized(parsedRange),
  }
}

function setInputCell(target: WorkPaperInstance, args: z.infer<typeof setInputCellInputSchema>) {
  const parsedArgs = setInputCellInputSchema.parse(args)
  const summarySheet = requireSheet(target, 'Summary')
  const address = requireCellAddress(target, parsedArgs.sheetName, parsedArgs.address)
  const before = readSummary(target, summarySheet)
  const beforeContracts = readFormulaContracts(target, summarySheet)
  const previousValue = target.getCellSerialized(address)

  target.setCellContents(address, parsedArgs.value)

  const after = readSummary(target, summarySheet)
  const afterContracts = readFormulaContracts(target, summarySheet)
  const saved = serializeWorkPaperDocument(
    exportWorkPaperDocument(target, {
      includeConfig: true,
    }),
  )
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
  const restoredSummary = readSummary(restored, requireSheet(restored, 'Summary'))

  return {
    editedCell: target.simpleCellAddressToString(address, {
      includeSheetName: true,
    }),
    before,
    after,
    restored: restoredSummary,
    beforeContracts,
    afterContracts,
    checks: {
      previousValue,
      newValue: target.getCellSerialized(address),
      formulasPersisted: sameJson(afterContracts, readFormulaContracts(restored, requireSheet(restored, 'Summary'))),
      restoredMatchesAfter: sameJson(after, restoredSummary),
      expectedArrChanged: before.expectedArr !== after.expectedArr,
      serializedBytes: Buffer.byteLength(saved, 'utf8'),
    },
  }
}

function readSummary(target: WorkPaperInstance, summarySheet: number): SummaryReadback {
  return {
    expectedCustomers: readNumber(target, summarySheet, 'B2'),
    expectedArr: readNumber(target, summarySheet, 'B3'),
    expansionArr: readNumber(target, summarySheet, 'B4'),
    targetGap: readNumber(target, summarySheet, 'B5'),
  }
}

function readFormulaContracts(target: WorkPaperInstance, summarySheet: number): FormulaContracts {
  return {
    expectedCustomers: readFormula(target, summarySheet, 'B2'),
    expectedArr: readFormula(target, summarySheet, 'B3'),
    expansionArr: readFormula(target, summarySheet, 'B4'),
    targetGap: readFormula(target, summarySheet, 'B5'),
  }
}

function readNumber(target: WorkPaperInstance, sheetId: number, address: string): number {
  const value = target.getCellValue(requireCellAddressBySheetId(target, sheetId, address))
  if (!value || typeof value !== 'object' || !('value' in value) || typeof value.value !== 'number') {
    throw new Error(`Expected ${address} to contain a number, received ${JSON.stringify(value)}`)
  }
  return Math.round(value.value * 100) / 100
}

function readFormula(target: WorkPaperInstance, sheetId: number, address: string): string {
  const value = target.getCellFormula(requireCellAddressBySheetId(target, sheetId, address))
  if (value === undefined) {
    throw new Error(`Expected ${address} to contain a formula`)
  }
  return value
}

function requireCellAddress(target: WorkPaperInstance, sheetName: string, address: string): CellAddress {
  return requireCellAddressBySheetId(target, requireSheet(target, sheetName), address)
}

function requireCellAddressBySheetId(target: WorkPaperInstance, sheetId: number, address: string): CellAddress {
  const parsed = target.simpleCellAddressFromString(address, sheetId)
  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid WorkPaper address: ${address}`)
  }
  return parsed
}

function requireSheet(target: WorkPaperInstance, sheetName: string): number {
  const sheetId = target.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected WorkPaper sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertProof(value: typeof proof): void {
  if (value.apiShape !== 'Mastra createTool -> execute -> WorkPaper readback') {
    throw new Error(`Unexpected API shape: ${value.apiShape}`)
  }
  if (!sameJson(value.toolIds, ['read-workpaper-summary', 'set-workpaper-input-cell'])) {
    throw new Error(`Unexpected tool ids: ${JSON.stringify(value.toolIds)}`)
  }
  if (value.writeResult.editedCell !== 'Inputs!B3') {
    throw new Error(`Unexpected edited cell: ${value.writeResult.editedCell}`)
  }
  if (value.writeResult.before.expectedArr !== 60000 || value.writeResult.after.expectedArr !== 96000) {
    throw new Error(`Unexpected expected ARR readback: ${JSON.stringify(value.writeResult)}`)
  }
  if (
    !value.writeResult.checks.formulasPersisted ||
    !value.writeResult.checks.restoredMatchesAfter ||
    !value.writeResult.checks.expectedArrChanged
  ) {
    throw new Error(`Mastra WorkPaper checks failed: ${JSON.stringify(value.writeResult.checks)}`)
  }
}
