import { describe, expect, it } from 'vitest'

import { agentStartHelpText, buildAgentStartDecisionCard, parseAgentStartCliArgs, runAgentStartCli } from '../agent-start-cli.js'

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
  })

  it('prints JSON by default', () => {
    let stdout = ''
    const exitCode = runAgentStartCli({
      argv: [],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed).toMatchObject({
      schemaVersion: 'bilig-agent-start.v1',
      name: 'bilig-agent-start',
    })
    expect(parsed.firstCommands.agentMcp).toContain('bilig-evaluate --door agent-mcp --json')
    expect(parsed.avoidUiAutomationWhen.join(' ')).toContain('calculated readback')
  })

  it('prints markdown when requested', () => {
    let stdout = ''
    const exitCode = runAgentStartCli({
      argv: ['--markdown'],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig agent start')
    expect(stdout).toContain('bilig-evaluate --door agent-mcp --json')
    expect(stdout).toContain('Do not claim success from a write call alone.')
  })

  it('validates arguments and help', () => {
    expect(parseAgentStartCliArgs(['--json'])).toEqual({
      help: false,
      outputMode: 'json',
    })
    expect(agentStartHelpText()).toContain('Usage: bilig-agent-start')
    expect(() => parseAgentStartCliArgs(['--bad'])).toThrow('Unknown bilig-agent-start argument')
  })
})
