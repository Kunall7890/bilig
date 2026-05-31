import { spawnSync } from 'node:child_process'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseNodeMarkdownReportOutput } from '../workpaper-external-smoke-parsers.ts'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)
const exampleDir = resolve(repoRoot, 'examples', 'headless-workpaper')

describe('headless WorkPaper markdown report example', () => {
  it('prints deterministic report rows for downstream job summaries', () => {
    const result = spawnSync('npm', ['run', '--silent', 'markdown-report'], {
      cwd: exampleDir,
      encoding: 'utf8',
      env: { ...process.env, NODE_NO_WARNINGS: '1' },
    })

    expect(result.stderr).toBe('')
    expect(result.status).toBe(0)

    const output = parseNodeMarkdownReportOutput(result.stdout)
    expect(output).toEqual({
      verified: true,
      report: [
        '| Metric | Value |',
        '| --- | ---: |',
        '| Committed MRR | $39,600 |',
        '| Weighted pipeline MRR | $43,400 |',
        '| Target gap | $10,400 |',
      ].join('\n'),
    })
  })
})
