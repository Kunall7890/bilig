#!/usr/bin/env node
import { createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import { runDemoWorkPaperMcpStdioServer } from './work-paper-mcp-stdio-server.js'

interface CliOptions {
  workpaperPath?: string
  writable: boolean
  help: boolean
}

const cliOptions = parseArgs(process.argv.slice(2))
if (cliOptions.help) {
  process.stdout.write(helpText())
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

function parseArgs(args: string[]): CliOptions {
  const options: CliOptions = {
    writable: false,
    help: false,
  }

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === '--help' || arg === '-h') {
      options.help = true
      continue
    }
    if (arg === '--writable') {
      options.writable = true
      continue
    }
    if (arg === '--workpaper') {
      const next = args[index + 1]
      if (next === undefined || next.startsWith('-')) {
        throw new Error('--workpaper requires a path')
      }
      options.workpaperPath = next
      index += 1
      continue
    }
    throw new Error(`Unknown bilig-workpaper-mcp argument: ${arg}`)
  }

  return options
}

function helpText(): string {
  return [
    'Usage: bilig-workpaper-mcp [--workpaper ./model.workpaper.json] [--writable]',
    '',
    'Without --workpaper, starts the built-in demo WorkPaper MCP server.',
    'With --workpaper, loads a persisted WorkPaper JSON document and exposes file-backed tools.',
    '--writable persists set_cell_contents edits back to the same JSON file.',
    '',
  ].join('\n')
}
