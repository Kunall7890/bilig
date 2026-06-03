import { describe, expect, it } from 'vitest'

import { agentStartHelpText, buildAgentStartDecisionCard, parseAgentStartCliArgs, runAgentStartCli } from '../agent-start-cli.js'

const expectedRuleTargets = [
  ['codex', 'AGENTS.md'],
  ['claude', 'CLAUDE.md'],
  [
    'copilot',
    '.github/copilot-instructions.md, .github/instructions/bilig-workpaper.instructions.md, .github/prompts/bilig-workpaper-proof.prompt.md, .vscode/mcp.json',
  ],
  ['cursor', '.cursor/rules/bilig-workpaper.mdc'],
  ['opencode', 'opencode.jsonc, .opencode/agents/bilig-workpaper.md'],
  ['cline', '.clinerules/bilig-workpaper.md'],
  ['continue', '.continue/rules/bilig-workpaper.md'],
  ['windsurf', '.devin/rules/bilig-workpaper.md or .windsurf/rules/bilig-workpaper.md'],
  ['gemini', 'GEMINI.md, gemini-extension.json, gemini-workpaper-context.md'],
  ['junie', '.junie/mcp/mcp.json'],
  ['vscode-mcp', '.vscode/mcp.json'],
] as const

function runCli(argv: readonly string[]): { readonly exitCode: number; readonly stdout: string; readonly stderr: string } {
  let stdout = ''
  let stderr = ''
  const exitCode = runAgentStartCli({
    argv,
    writeStdout(text) {
      stdout += text
    },
    writeStderr(text) {
      stderr += text
    },
  })
  return { exitCode, stdout, stderr }
}

describe('bilig-agent-start', () => {
  it('builds the agent workbook decision card', () => {
    const card = buildAgentStartDecisionCard()

    expect(card).toMatchObject({
      schemaVersion: 'bilig-agent-start.v1',
      name: 'bilig-agent-start',
      firstCommands: {
        agentMcp: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
        workpaperService: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
        xlsxCache: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
        providerBacked:
          'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
      },
    })
    expect(card.evidenceRequired).toContain('afterRestore or afterRestart')
    expect(card.expectedMcpTools).toContain('set_cell_contents_and_readback')
    for (const [target, file] of expectedRuleTargets) {
      expect(card.agentRuleTargets).toContainEqual({
        target,
        file,
        command: `npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules ${target}`,
      })
    }
  })

  it('prints JSON by default', () => {
    const { exitCode, stdout } = runCli([])

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed).toMatchObject({
      schemaVersion: 'bilig-agent-start.v1',
      name: 'bilig-agent-start',
    })
    expect(parsed.firstCommands.agentMcp).toContain('bilig-evaluate --door agent-mcp --json')
    expect(parsed.avoidUiAutomationWhen.join(' ')).toContain('calculated readback')
    expect(parsed.agentRuleTargets.map((target: { readonly target: string }) => target.target)).toEqual(
      expectedRuleTargets.map(([target]) => target),
    )
  })

  it('prints markdown when requested', () => {
    const { exitCode, stdout } = runCli(['--markdown'])

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig agent start')
    expect(stdout).toContain('bilig-evaluate --door agent-mcp --json')
    expect(stdout).toContain('Do not claim success from a write call alone.')
  })

  it('prints target-specific agent rules', () => {
    const { exitCode, stdout } = runCli(['--rules', 'cursor'])

    expect(exitCode).toBe(0)
    expect(stdout).toContain('Recommended file: `.cursor/rules/bilig-workpaper.mdc`')
    expect(stdout).toContain('bilig-evaluate --door agent-mcp --json')
    expect(stdout).toContain('bilig-workpaper-mcp --from-xlsx ./pricing.xlsx')
    expect(stdout).toContain('Never claim success from a write call alone.')
    expect(stdout).toContain('alwaysApply: false')
  })

  it.each([
    ['cline', '.clinerules/bilig-workpaper.md', 'Cline', 'Cline can read this workspace rule'],
    ['continue', '.continue/rules/bilig-workpaper.md', 'Continue', 'name: Bilig WorkPaper Formula Check'],
    ['opencode', 'opencode.jsonc, .opencode/agents/bilig-workpaper.md', 'OpenCode', 'mode: subagent'],
    ['windsurf', '.devin/rules/bilig-workpaper.md or .windsurf/rules/bilig-workpaper.md', 'Windsurf/Cascade', 'trigger: model_decision'],
    ['gemini', 'GEMINI.md, gemini-extension.json, gemini-workpaper-context.md', 'Gemini CLI', 'gemini-workpaper-context.md'],
    ['junie', '.junie/mcp/mcp.json', 'Junie', 'Junie reads project guidelines'],
    ['vscode-mcp', '.vscode/mcp.json', 'VS Code agent mode', 'biligWorkpaperFile'],
    [
      'copilot',
      '.github/copilot-instructions.md, .github/instructions/bilig-workpaper.instructions.md, .github/prompts/bilig-workpaper-proof.prompt.md, .vscode/mcp.json',
      'GitHub Copilot',
      '.github/prompts/bilig-workpaper-proof.prompt.md',
    ],
  ])('prints %s agent rules', (target, expectedFile, expectedName, targetMarker) => {
    const { exitCode, stdout } = runCli(['--rules', target])

    expect(exitCode).toBe(0)
    expect(stdout).toContain(`# Bilig WorkPaper rule for ${expectedName}`)
    expect(stdout).toContain(`Recommended file: \`${expectedFile}\``)
    expect(stdout).toContain(targetMarker)
    expect(stdout).toContain('bilig-evaluate --door agent-mcp --json')
    expect(stdout).toContain('Never claim success from a write call alone.')
    expect(stdout).not.toContain('alwaysApply: false')
  })

  it('validates arguments and help', () => {
    expect(parseAgentStartCliArgs(['--json'])).toEqual({
      help: false,
      outputMode: 'json',
      ruleTarget: 'codex',
    })
    expect(parseAgentStartCliArgs(['--rules=copilot'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'copilot',
    })
    expect(parseAgentStartCliArgs(['--rules', 'continue'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'continue',
    })
    expect(parseAgentStartCliArgs(['--rules=windsurf'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'windsurf',
    })
    expect(parseAgentStartCliArgs(['--rules=vscode-mcp'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'vscode-mcp',
    })
    expect(parseAgentStartCliArgs(['--rules=opencode'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'opencode',
    })
    expect(parseAgentStartCliArgs(['--rules=junie'])).toEqual({
      help: false,
      outputMode: 'rules',
      ruleTarget: 'junie',
    })
    expect(agentStartHelpText()).toContain('Usage: bilig-agent-start')
    expect(agentStartHelpText()).toContain('cline        .clinerules/bilig-workpaper.md')
    expect(agentStartHelpText()).toContain('opencode     opencode.jsonc, .opencode/agents/bilig-workpaper.md')
    expect(agentStartHelpText()).toContain('windsurf     .devin/rules/bilig-workpaper.md or .windsurf/rules/bilig-workpaper.md')
    expect(agentStartHelpText()).toContain('junie        .junie/mcp/mcp.json')
    expect(agentStartHelpText()).toContain('vscode-mcp   .vscode/mcp.json')
    expect(() => parseAgentStartCliArgs(['--bad'])).toThrow('Unknown bilig-agent-start argument')
    expect(() => parseAgentStartCliArgs(['--rules', 'bad'])).toThrow(
      'Unknown bilig-agent-start rules target: bad. Use one of: codex, claude, copilot, cursor, opencode, cline, continue, windsurf, gemini, junie, vscode-mcp',
    )
    expect(() => parseAgentStartCliArgs(['--rules'])).toThrow(
      'Missing target for --rules. Use one of: codex, claude, copilot, cursor, opencode, cline, continue, windsurf, gemini, junie, vscode-mcp',
    )
  })
})
