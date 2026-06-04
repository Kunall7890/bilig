import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import { WorkPaper, createWorkPaperFromDocument, parseWorkPaperDocument } from '@bilig/workpaper'
import { exportXlsx } from '@bilig/workpaper/xlsx'

type JsonObject = Record<string, unknown>

interface XlsxRiskPreflightOutput {
  schemaVersion: 'bilig-agent-xlsx-risk-preflight.v1'
  transport: 'stdio'
  command: string
  requestMethods: string[]
  source: {
    xlsxFile: string
    persistedWorkPaperFile: string
  }
  toolNames: string[]
  risk: {
    schemaVersion: string
    verified: boolean
    fileName: string
    formulaCellCount: number
    level: string
    excelParity: 'not_proven'
  }
  readback: {
    editedCell: 'Inputs!B3'
    beforeExpectedArr: 60000
    afterExpectedArr: 96000
    restoredExpectedArr: 96000
    persisted: boolean
    restoredReadbackMatchesAfter: boolean
  }
  export: {
    serializedBytes: number
  }
  verified: true
  limitations: string[]
}

const workspace = mkdtempSync(join(tmpdir(), 'bilig-xlsx-risk-preflight-'))
const xlsxPath = join(workspace, 'pricing-risk-preflight.xlsx')
const workpaperPath = join(workspace, 'pricing-risk-preflight.workpaper.json')

try {
  writeFileSync(xlsxPath, buildPricingXlsx())

  const requestLines = [
    jsonLine({ jsonrpc: '2.0', id: 1, method: 'initialize' }),
    jsonLine({ jsonrpc: '2.0', method: 'notifications/initialized' }),
    jsonLine({ jsonrpc: '2.0', id: 2, method: 'tools/list' }),
    jsonLine({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'analyze_workbook_risk',
        arguments: {
          inspectLimit: 'all',
        },
      },
    }),
    jsonLine({
      jsonrpc: '2.0',
      id: 4,
      method: 'tools/call',
      params: {
        name: 'set_cell_contents_and_readback',
        arguments: {
          sheetName: 'Inputs',
          address: 'B3',
          value: 0.4,
          readbackRange: 'Summary!A1:B4',
        },
      },
    }),
    jsonLine({
      jsonrpc: '2.0',
      id: 5,
      method: 'tools/call',
      params: {
        name: 'export_workpaper_document',
        arguments: {
          includeConfig: true,
        },
      },
    }),
  ]

  const responseLines = await runPublishedMcpBinary(xlsxPath, workpaperPath, requestLines)
  const output = createPreflightOutput(xlsxPath, workpaperPath, requestLines, responseLines)
  assertPreflightOutput(output)

  console.log(JSON.stringify(output, null, 2))
} finally {
  rmSync(workspace, { recursive: true, force: true })
}

function buildPricingXlsx(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Opportunities', 20],
      ['Win rate', 0.25],
      ['ARR per customer', 12000],
      ['Target ARR', 100000],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ['Expected ARR', '=B2*Inputs!B4'],
      ['Target gap', '=B3-Inputs!B5'],
    ],
  })

  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

function jsonLine(value: unknown): string {
  return JSON.stringify(value)
}

async function runPublishedMcpBinary(xlsxFile: string, workpaperFile: string, lines: string[]): Promise<string[]> {
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm'
  const child = spawn(
    npmCommand,
    [
      'exec',
      '--loglevel=error',
      '--package',
      '@bilig/workpaper@latest',
      '--',
      'bilig-workpaper-mcp',
      '--from-xlsx',
      xlsxFile,
      '--workpaper',
      workpaperFile,
      '--writable',
    ],
    {
      cwd: process.cwd(),
      env: {
        ...process.env,
        NODE_NO_WARNINGS: '1',
      },
      stdio: ['pipe', 'pipe', 'pipe'],
    },
  )

  let stdout = ''
  let stderr = ''
  child.stdout.setEncoding('utf8')
  child.stderr.setEncoding('utf8')
  child.stdout.on('data', (chunk: string) => {
    stdout += chunk
  })
  child.stderr.on('data', (chunk: string) => {
    stderr += chunk
  })

  child.stdin.end(`${lines.join('\n')}\n`)

  const exitCode = await new Promise<number | null>((resolve, reject) => {
    child.on('error', reject)
    child.on('close', resolve)
  })

  if (exitCode !== 0) {
    throw new Error(`XLSX risk preflight MCP transcript failed with exit ${String(exitCode)}: ${stderr}`)
  }
  if (stderr.trim().length > 0) {
    throw new Error(`XLSX risk preflight MCP transcript wrote stderr: ${stderr}`)
  }

  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
}

function createPreflightOutput(xlsxFile: string, workpaperFile: string, requests: string[], responses: string[]): XlsxRiskPreflightOutput {
  const toolsResponse = requireResponse(responses, 2)
  const riskResponse = requireResponse(responses, 3)
  const writeResponse = requireResponse(responses, 4)
  const exportResponse = requireResponse(responses, 5)

  const toolsResult = readRecord(toolsResponse.result, 'tools/list result')
  const risk = readRecord(readRecord(riskResponse.result, 'risk result').structuredContent, 'risk structured content')
  const input = readRecord(risk.input, 'risk input')
  const workbook = readRecord(risk.workbook, 'risk workbook')
  const riskLevel = readRecord(risk.risk, 'risk level')
  const write = readRecord(readRecord(writeResponse.result, 'write result').structuredContent, 'write structured content')
  const exportResult = readRecord(readRecord(exportResponse.result, 'export result').structuredContent, 'export structured content')
  const checks = readRecord(write.checks, 'write checks')
  const persistence = readRecord(write.persistence, 'write persistence')
  const beforeReadback = readRecord(write.beforeReadback, 'before readback')
  const afterReadback = readRecord(write.afterReadback, 'after readback')
  const restoredReadback = readRecord(write.restoredReadback, 'restored readback')

  return {
    schemaVersion: 'bilig-agent-xlsx-risk-preflight.v1',
    transport: 'stdio',
    command:
      'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx pricing-risk-preflight.xlsx --workpaper pricing-risk-preflight.workpaper.json --writable',
    requestMethods: requests.map((line) => readString(readRecord(JSON.parse(line), 'request line').method, 'request method')),
    source: {
      xlsxFile: basename(xlsxFile),
      persistedWorkPaperFile: basename(workpaperFile),
    },
    toolNames: readToolNames(toolsResult.tools),
    risk: {
      schemaVersion: readString(risk.schemaVersion, 'risk schemaVersion'),
      verified: readBoolean(risk.verified, 'risk verified'),
      fileName: readString(input.fileName, 'risk input fileName'),
      formulaCellCount: readNumber(workbook.formulaCellCount, 'risk formulaCellCount'),
      level: readString(riskLevel.level, 'risk level'),
      excelParity: readExcelParity(risk.excelParity),
    },
    readback: {
      editedCell: readEditedCell(write.editedCell),
      beforeExpectedArr: readExpectedArr(beforeReadback, 'before readback', 60000),
      afterExpectedArr: readExpectedArr(afterReadback, 'after readback', 96000),
      restoredExpectedArr: readExpectedArr(restoredReadback, 'restored readback', 96000),
      persisted: readBoolean(persistence.persisted, 'persistence persisted'),
      restoredReadbackMatchesAfter: readBoolean(checks.restoredReadbackMatchesAfter, 'restoredReadbackMatchesAfter'),
    },
    export: {
      serializedBytes: readNumber(exportResult.serializedBytes, 'export serializedBytes'),
    },
    verified: true,
    limitations: [
      'The risk tool is a local preflight diagnostic, not an Excel compatibility certification.',
      'The edit proof is WorkPaper formula readback and JSON persistence, not desktop Excel UI proof.',
    ],
  }
}

function requireResponse(lines: string[], id: number): JsonObject {
  for (const line of lines) {
    const parsed = readRecord(JSON.parse(line), `response ${id.toString()}`)
    if (parsed.id === id && parsed.jsonrpc === '2.0') {
      return parsed
    }
  }
  throw new Error(`Missing JSON-RPC response id ${id.toString()}: ${lines.join('\n')}`)
}

function readToolNames(value: unknown): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected tools/list result to be an array, got ${JSON.stringify(value)}`)
  }

  return value.map((entry, index) => readString(readRecord(entry, `tool ${index.toString()}`).name, `tool ${index.toString()} name`))
}

function readExpectedArr<const Expected extends 60000 | 96000>(readback: JsonObject, label: string, expected: Expected): Expected {
  const values = read2dArray(readback.values, `${label} values`)
  const row = values[2]
  if (row === undefined) {
    throw new Error(`Expected ${label} to include Summary row 3, got ${JSON.stringify(values)}`)
  }
  const value = readCellValue(row[1], `${label} Summary!B3`)
  if (value !== expected) {
    throw new Error(`Expected ${label} Summary!B3 to be ${expected.toString()}, got ${JSON.stringify(value)}`)
  }
  return expected
}

function readCellValue(value: unknown, _label: string): unknown {
  if (isJsonObject(value) && 'value' in value) {
    return value.value
  }
  return value
}

function read2dArray(value: unknown, label: string): unknown[][] {
  if (!Array.isArray(value) || !value.every((row) => Array.isArray(row))) {
    throw new Error(`Expected ${label} to be a two-dimensional array, got ${JSON.stringify(value)}`)
  }
  return value
}

function readEditedCell(value: unknown): 'Inputs!B3' {
  if (value !== 'Inputs!B3') {
    throw new Error(`Expected edited cell to be Inputs!B3, got ${JSON.stringify(value)}`)
  }
  return value
}

function readExcelParity(value: unknown): 'not_proven' {
  if (value !== 'not_proven') {
    throw new Error(`Expected excelParity to be not_proven, got ${JSON.stringify(value)}`)
  }
  return value
}

function assertPreflightOutput(output: XlsxRiskPreflightOutput): void {
  const requiredTools = ['analyze_workbook_risk', 'set_cell_contents_and_readback', 'export_workpaper_document']
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))
  const summarySheet = restored.getSheetId('Summary')
  if (summarySheet === undefined) {
    throw new Error('Expected restored WorkPaper to include Summary sheet')
  }
  const restoredSummaryArr = readCellValue(restored.getCellValue({ sheet: summarySheet, row: 2, col: 1 }), 'restored Summary!B3')

  if (
    !requiredTools.every((toolName) => output.toolNames.includes(toolName)) ||
    output.risk.schemaVersion !== 'bilig-workbook-compatibility-report.v1' ||
    !output.risk.verified ||
    output.risk.fileName !== 'pricing-risk-preflight.xlsx' ||
    output.risk.formulaCellCount !== 3 ||
    output.risk.excelParity !== 'not_proven' ||
    output.readback.editedCell !== 'Inputs!B3' ||
    output.readback.beforeExpectedArr !== 60000 ||
    output.readback.afterExpectedArr !== 96000 ||
    output.readback.restoredExpectedArr !== 96000 ||
    !output.readback.persisted ||
    !output.readback.restoredReadbackMatchesAfter ||
    output.export.serializedBytes <= 0 ||
    restoredSummaryArr !== 96000
  ) {
    throw new Error(`Unexpected XLSX risk preflight output: ${JSON.stringify(output)}`)
  }
}

function readRecord(value: unknown, label: string): JsonObject {
  if (!isJsonObject(value)) {
    throw new Error(`Expected ${label} to be an object, got ${JSON.stringify(value)}`)
  }
  return value
}

function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string, got ${JSON.stringify(value)}`)
  }
  return value
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`Expected ${label} to be a number, got ${JSON.stringify(value)}`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean, got ${JSON.stringify(value)}`)
  }
  return value
}
