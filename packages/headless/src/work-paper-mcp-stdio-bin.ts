#!/usr/bin/env node
import { createFileBackedWorkPaperMcpToolServer, createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import { parseWorkPaperMcpStdioCliArgs, workPaperMcpStdioHelpText } from './work-paper-mcp-stdio-cli.js'
import { runDemoWorkPaperMcpStdioServer } from './work-paper-mcp-stdio-server.js'
import { buildDemoWorkPaper, type WorkPaperMcpToolServer } from './work-paper-mcp-server.js'

type XlsxWorkbookRiskToolWrapper = (
  server: WorkPaperMcpToolServer,
  input: {
    readonly xlsxPath: string
  },
) => WorkPaperMcpToolServer

const cliOptions = parseWorkPaperMcpStdioCliArgs(process.argv.slice(2))
if (cliOptions.help) {
  process.stdout.write(workPaperMcpStdioHelpText())
  process.exit(0)
}

if (cliOptions.demoWorkPaperTools) {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServer({
      workbook: buildDemoWorkPaper(),
      sourcePath: 'demo://bilig-workpaper',
      writable: false,
    }),
  })
} else if (cliOptions.fromXlsxPath !== undefined) {
  const { createFileBackedWorkPaperMcpToolServerFromXlsxFile, createWorkPaperMcpToolServerFromXlsxFile } =
    await import('./work-paper-mcp-xlsx-file.js')
  const server =
    cliOptions.workpaperPath === undefined
      ? createWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: cliOptions.fromXlsxPath,
        })
      : createFileBackedWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: cliOptions.fromXlsxPath,
          overwriteWorkPaper: cliOptions.overwriteWorkPaper,
          workpaperPath: cliOptions.workpaperPath,
          writable: cliOptions.writable,
        })
  runDemoWorkPaperMcpStdioServer({
    server: await withOptionalXlsxWorkbookRiskTool(server, cliOptions.fromXlsxPath),
  })
} else if (cliOptions.workpaperPath === undefined) {
  runDemoWorkPaperMcpStdioServer()
} else {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServerFromFile({
      initDemoWorkPaper: cliOptions.initDemoWorkPaper,
      workpaperPath: cliOptions.workpaperPath,
      writable: cliOptions.writable,
    }),
  })
}

async function withOptionalXlsxWorkbookRiskTool(server: WorkPaperMcpToolServer, xlsxPath: string): Promise<WorkPaperMcpToolServer> {
  try {
    const riskModuleSpecifier = ['bilig-workpaper', 'mcp'].join('/')
    const riskModule: unknown = await import(riskModuleSpecifier)
    if (hasXlsxWorkbookRiskToolWrapper(riskModule)) {
      return riskModule.withXlsxWorkbookRiskTool(server, { xlsxPath })
    }
  } catch (error) {
    if (!isMissingOptionalBiligWorkpaperMcpModule(error)) {
      throw error
    }
  }
  return server
}

function hasXlsxWorkbookRiskToolWrapper(value: unknown): value is {
  readonly withXlsxWorkbookRiskTool: XlsxWorkbookRiskToolWrapper
} {
  return isRecord(value) && typeof value['withXlsxWorkbookRiskTool'] === 'function'
}

function isMissingOptionalBiligWorkpaperMcpModule(error: unknown): boolean {
  if (!isRecord(error)) {
    return false
  }
  const code = error['code']
  return code === 'ERR_MODULE_NOT_FOUND' || code === 'MODULE_NOT_FOUND'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
