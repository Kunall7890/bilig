import { writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { parseJsonRecord, parseRecordValue } from './workpaper-external-smoke-parser-helpers.ts'

export type PackageMcpXlsxRiskSummary = {
  excelParity: 'not_proven'
  inputFileName: string
  schemaVersion: string
  toolNames: string[]
  verified: boolean
}

export function writeMcpXlsxRiskFixtureScript(projectDir: string): void {
  writeFileSync(
    join(projectDir, 'mcp-xlsx-risk-fixture.ts'),
    [
      'import { mkdirSync, writeFileSync } from "node:fs";',
      'import { WorkPaper } from "@bilig/workpaper";',
      'import { exportXlsx } from "@bilig/workpaper/xlsx";',
      '',
      'mkdirSync("fixtures", { recursive: true });',
      'const workbook = WorkPaper.buildFromSheets({',
      '  Inputs: [["Metric", "Value"], ["Units", 40], ["Price", 1200]],',
      '  Summary: [["Metric", "Value"], ["Revenue", "=Inputs!B2*Inputs!B3"]],',
      '});',
      'writeFileSync("fixtures/pricing-risk.xlsx", exportXlsx(workbook.exportSnapshot()));',
      'workbook.dispose();',
      '',
    ].join('\n'),
  )
}

export function parsePackageMcpXlsxRiskOutput(output: string): PackageMcpXlsxRiskSummary {
  const responses = output
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .map((line, index) => parseJsonRecord(line, `package MCP XLSX risk response ${index + 1}`))
  const listResult = parseRecordValue(
    requireJsonRpcResponse(responses, 2, 'package MCP XLSX risk tools/list response').result,
    'package MCP XLSX risk tools/list result',
  )
  const riskResult = parseRecordValue(
    requireJsonRpcResponse(responses, 3, 'package MCP XLSX risk tools/call response').result,
    'package MCP XLSX risk tools/call result',
  )
  const toolNames = parseMcpToolNames(listResult.tools, 'package MCP XLSX risk tool')
  const structuredContent = parseRecordValue(riskResult.structuredContent, 'package MCP XLSX risk structured content')
  const input = parseRecordValue(structuredContent.input, 'package MCP XLSX risk input')
  const schemaVersion = structuredContent.schemaVersion
  const inputFileName = input.fileName

  if (
    !toolNames.includes('analyze_workbook_risk') ||
    schemaVersion !== 'bilig-workbook-compatibility-report.v1' ||
    structuredContent.verified !== true ||
    inputFileName !== 'pricing-risk.xlsx' ||
    structuredContent.excelParity !== 'not_proven'
  ) {
    throw new Error(`Unexpected package MCP XLSX risk output: ${output}`)
  }

  return {
    excelParity: 'not_proven',
    inputFileName,
    schemaVersion,
    toolNames,
    verified: true,
  }
}

function requireJsonRpcResponse(responses: Record<string, unknown>[], id: number, context: string): Record<string, unknown> {
  const response = responses.find((entry) => entry.id === id)
  if (response === undefined || response.jsonrpc !== '2.0') {
    throw new Error(`Missing ${context}: ${JSON.stringify(responses)}`)
  }
  return response
}

function parseMcpToolNames(value: unknown, context: string): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected ${context} list to be an array: ${JSON.stringify(value)}`)
  }
  return value.map((entry, index) => {
    const tool = parseRecordValue(entry, `${context} ${index + 1}`)
    if (typeof tool.name !== 'string') {
      throw new Error(`Unexpected ${context}: ${JSON.stringify(entry)}`)
    }
    return tool.name
  })
}
