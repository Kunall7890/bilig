export interface AgentStartCliHost {
  readonly argv: readonly string[]
  readonly writeStderr?: (text: string) => void
  readonly writeStdout?: (text: string) => void
}

type AgentStartOutputMode = 'json' | 'markdown'

interface AgentStartCliOptions {
  readonly help: boolean
  readonly outputMode: AgentStartOutputMode
}

export interface AgentStartDecisionCard {
  readonly schemaVersion: 'bilig-agent-start.v1'
  readonly name: 'bilig-agent-start'
  readonly purpose: string
  readonly defaultRule: string
  readonly firstCommands: {
    readonly agentMcp: string
    readonly workpaperService: string
    readonly xlsxCache: string
    readonly providerBacked: string
  }
  readonly useBiligWhen: readonly string[]
  readonly avoidUiAutomationWhen: readonly string[]
  readonly evidenceRequired: readonly string[]
  readonly expectedMcpTools: readonly string[]
  readonly next: {
    readonly agentStart: string
    readonly llmsTxt: string
    readonly agentManifest: string
    readonly docs: string
    readonly repository: string
  }
}

export function runAgentStartCli(host: AgentStartCliHost): number {
  const writeStdout = host.writeStdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text: string) => process.stderr.write(text))
  let options: AgentStartCliOptions

  try {
    options = parseAgentStartCliArgs(host.argv)
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n\n${agentStartHelpText()}`)
    return 1
  }

  if (options.help) {
    writeStdout(agentStartHelpText())
    return 0
  }

  writeStdout(renderAgentStartDecisionCard(buildAgentStartDecisionCard(), options.outputMode))
  return 0
}

export function parseAgentStartCliArgs(args: readonly string[]): AgentStartCliOptions {
  let help = false
  let outputMode: AgentStartOutputMode = 'json'

  for (const arg of args) {
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--json') {
      outputMode = 'json'
      continue
    }
    if (arg === '--markdown') {
      outputMode = 'markdown'
      continue
    }
    throw new Error(`Unknown bilig-agent-start argument: ${arg}`)
  }

  return { help, outputMode }
}

export function agentStartHelpText(): string {
  return [
    'Usage: bilig-agent-start [--json|--markdown]',
    '',
    'Prints the Bilig workbook triage card for coding agents.',
    'Use it before opening Excel, LibreOffice, Google Sheets, or a browser grid',
    'for workbook-shaped logic.',
    '',
    'Options:',
    '  --json       Print machine-readable JSON. Default.',
    '  --markdown   Print a compact Markdown checklist.',
    '  -h, --help   Print this help text.',
    '',
  ].join('\n')
}

export function buildAgentStartDecisionCard(): AgentStartDecisionCard {
  return {
    schemaVersion: 'bilig-agent-start.v1',
    name: 'bilig-agent-start',
    purpose: 'Help a coding agent choose a WorkPaper proof command before attempting spreadsheet UI automation.',
    defaultRule:
      'If the task needs cells, formulas, recalculation, readback, or saved workbook state, run a Bilig proof before driving a spreadsheet UI.',
    firstCommands: {
      agentMcp: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
      workpaperService: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
      xlsxCache: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
      providerBacked:
        'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    },
    useBiligWhen: [
      'The workflow is pricing, payout, budget, import validation, forecast, quote approval, or another formula-backed model.',
      'A service, queue worker, CI job, MCP client, or coding agent needs exact formula readback.',
      'The output must include a persisted WorkPaper JSON document, restored state, or restart readback.',
      'The task is to inspect stale XLSX cached values after a script changed inputs.',
    ],
    avoidUiAutomationWhen: [
      'The agent only needs to write inputs, recalculate formulas, read dependent cells, and save state.',
      'A screenshot would be used as formula truth instead of calculated readback.',
      'The workbook can be represented as WorkPaper JSON or a reduced XLSX fixture.',
    ],
    evidenceRequired: [
      'verified',
      'editedCell',
      'before',
      'after',
      'afterRestore or afterRestart',
      'persistedDocumentBytes or exported WorkPaper JSON size',
      'limitations',
    ],
    expectedMcpTools: [
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'set_cell_contents_and_readback',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ],
    next: {
      agentStart: 'https://proompteng.github.io/bilig/agent-start.txt',
      llmsTxt: 'https://proompteng.github.io/bilig/llms.txt',
      agentManifest: 'https://proompteng.github.io/bilig/.well-known/agent.json',
      docs: 'https://proompteng.github.io/bilig/eval-agent-mcp.html',
      repository: 'https://github.com/proompteng/bilig',
    },
  }
}

export function renderAgentStartDecisionCard(card: AgentStartDecisionCard, outputMode: AgentStartOutputMode): string {
  if (outputMode === 'markdown') {
    return renderAgentStartMarkdown(card)
  }
  return `${JSON.stringify(card, null, 2)}\n`
}

function renderAgentStartMarkdown(card: AgentStartDecisionCard): string {
  return `# Bilig agent start

${card.defaultRule}

## First commands

- Agent or MCP client: \`${card.firstCommands.agentMcp}\`
- Node service or test: \`${card.firstCommands.workpaperService}\`
- Stale XLSX cache: \`${card.firstCommands.xlsxCache}\`
- Provider-backed formulas: \`${card.firstCommands.providerBacked}\`

## Required evidence

${card.evidenceRequired.map((field) => `- \`${field}\``).join('\n')}

Do not claim success from a write call alone. Read the dependent calculated cell, persist or export the WorkPaper, and report the blocker if readback fails.

More context: ${card.next.agentStart}
`
}
