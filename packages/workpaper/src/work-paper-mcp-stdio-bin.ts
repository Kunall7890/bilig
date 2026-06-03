#!/usr/bin/env node
import {
  buildDemoWorkPaper,
  createFileBackedWorkPaperMcpToolServer,
  createFileBackedWorkPaperMcpToolServerFromFile,
  createFileBackedWorkPaperMcpToolServerFromXlsxFile,
  parseWorkPaperMcpStdioCliArgs,
  runDemoWorkPaperMcpStdioServer,
  withXlsxWorkbookRiskTool,
  workPaperMcpStdioHelpText,
} from 'bilig-workpaper/mcp'

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
} else if (cliOptions.workpaperPath === undefined) {
  runDemoWorkPaperMcpStdioServer()
} else if (cliOptions.fromXlsxPath !== undefined) {
  runDemoWorkPaperMcpStdioServer({
    server: withXlsxWorkbookRiskTool(
      createFileBackedWorkPaperMcpToolServerFromXlsxFile({
        fromXlsxPath: cliOptions.fromXlsxPath,
        overwriteWorkPaper: cliOptions.overwriteWorkPaper,
        workpaperPath: cliOptions.workpaperPath,
        writable: cliOptions.writable,
      }),
      { xlsxPath: cliOptions.fromXlsxPath },
    ),
  })
} else {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServerFromFile({
      initDemoWorkPaper: cliOptions.initDemoWorkPaper,
      workpaperPath: cliOptions.workpaperPath,
      writable: cliOptions.writable,
    }),
  })
}
