#!/usr/bin/env node
import { createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import { parseWorkPaperMcpStdioCliArgs, workPaperMcpStdioHelpText } from './work-paper-mcp-stdio-cli.js'
import { runDemoWorkPaperMcpStdioServer } from './work-paper-mcp-stdio-server.js'

const cliOptions = parseWorkPaperMcpStdioCliArgs(process.argv.slice(2))
if (cliOptions.help) {
  process.stdout.write(workPaperMcpStdioHelpText())
  process.exit(0)
}

if (cliOptions.workpaperPath === undefined) {
  runDemoWorkPaperMcpStdioServer()
} else {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServerFromFile({
      workpaperPath: cliOptions.workpaperPath,
      writable: cliOptions.writable,
    }),
  })
}
