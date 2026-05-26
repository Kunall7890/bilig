import { describe, expect, it } from 'vitest'

import {
  agentWorkbookChallengeHelpText,
  buildAgentWorkbookChallengeProof,
  parseAgentWorkbookChallengeCliArgs,
  runAgentWorkbookChallengeCli,
} from '../agent-workbook-challenge-cli.js'

describe('bilig-agent-challenge', () => {
  it('builds the verified WorkPaper proof object', () => {
    expect(buildAgentWorkbookChallengeProof()).toMatchObject({
      editedCell: 'Inputs!B2',
      dependentCell: 'Summary!B2',
      before: 24_000,
      after: 38_400,
      afterRestore: 38_400,
      checks: {
        formulaReadbackChanged: true,
        exportedWorkPaperDocument: true,
        restoredMatchesAfter: true,
      },
      verified: true,
      star: 'https://github.com/proompteng/bilig/stargazers',
      watchReleases: 'https://github.com/proompteng/bilig/subscription',
      adoptionBlocker: 'https://github.com/proompteng/bilig/discussions/new?category=general',
      nextStep: {
        ifUseful: 'If this agent workbook proof matched your workflow, star or bookmark Bilig so you can find it again.',
        star: 'https://github.com/proompteng/bilig/stargazers',
        watchReleases: 'https://github.com/proompteng/bilig/subscription',
        ifBlocked: 'If it almost worked, open the concrete workbook or agent blocker.',
        adoptionBlocker: 'https://github.com/proompteng/bilig/discussions/new?category=general',
      },
    })
  })

  it('prints JSON by default', () => {
    let stdout = ''
    const exitCode = runAgentWorkbookChallengeCli({
      argv: [],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed).toMatchObject({
      editedCell: 'Inputs!B2',
      after: 38_400,
      verified: true,
      star: 'https://github.com/proompteng/bilig/stargazers',
      watchReleases: 'https://github.com/proompteng/bilig/subscription',
      nextStep: {
        star: 'https://github.com/proompteng/bilig/stargazers',
        watchReleases: 'https://github.com/proompteng/bilig/subscription',
      },
    })
  })

  it('prints a markdown report when requested', () => {
    let stdout = ''
    const exitCode = runAgentWorkbookChallengeCli({
      argv: ['--markdown'],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig agent workbook challenge')
    expect(stdout).toContain('"verified": true')
    expect(stdout).toContain('Inputs!B2')
  })

  it('validates arguments and help', () => {
    expect(parseAgentWorkbookChallengeCliArgs(['--json'])).toEqual({
      help: false,
      outputMode: 'json',
    })
    expect(agentWorkbookChallengeHelpText()).toContain('Usage: bilig-agent-challenge')
    expect(() => parseAgentWorkbookChallengeCliArgs(['--bad'])).toThrow('Unknown bilig-agent-challenge argument')
  })
})
