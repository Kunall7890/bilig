export interface WorkPaperMcpStdioCliOptions {
  readonly workpaperPath?: string
  readonly writable: boolean
  readonly help: boolean
}

export function parseWorkPaperMcpStdioCliArgs(args: readonly string[]): WorkPaperMcpStdioCliOptions {
  let workpaperPath: string | undefined
  let writable = false
  let help = false

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--writable') {
      writable = true
      continue
    }
    if (arg === '--workpaper') {
      const next = args[index + 1]
      if (next === undefined || next.trim().length === 0 || next.startsWith('-')) {
        throw new Error('--workpaper requires a path')
      }
      workpaperPath = next
      index += 1
      continue
    }
    throw new Error(`Unknown bilig-workpaper-mcp argument: ${arg}`)
  }

  if (workpaperPath === undefined) {
    return { help, writable }
  }
  return { help, writable, workpaperPath }
}

export function workPaperMcpStdioHelpText(): string {
  return [
    'Usage: bilig-workpaper-mcp [--workpaper ./model.workpaper.json] [--writable]',
    '',
    'Without --workpaper, starts the built-in demo WorkPaper MCP server.',
    'With --workpaper, loads a persisted WorkPaper JSON document and exposes file-backed tools.',
    '--writable persists set_cell_contents edits back to the same JSON file.',
    '',
  ].join('\n')
}
