import { describe, expect, it } from 'vitest'

import { listBiligEvaluatorDoors, runBiligEvaluatorCli } from '../evaluator-cli.js'

describe('bilig-evaluate CLI', () => {
  it('lists only the XLSX evaluator doors', () => {
    expect(listBiligEvaluatorDoors().map((door) => door.door)).toEqual(['xlsx-cache', 'workbook-compatibility'])
  })

  it('prints a verified XLSX stale-cache proof', async () => {
    let stdout = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'xlsx-cache', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const proof = readProof(stdout)
    expect(proof.schemaVersion).toBe('bilig-evaluator.v1')
    expect(proof.door).toBe('xlsx-cache')
    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      target: 'Summary!B2',
      before: 60_000,
      after: 72_000,
      staleCachedFormulaCount: 1,
      suggestedReads: ['Summary!B2'],
      checks: {
        commandSucceeded: true,
        inspectionCompleted: true,
        recalculationCompleted: true,
        staleCachedFormulaFound: true,
        readbackSuggested: true,
      },
    })
  })

  it('prints a verified workbook compatibility proof without a compatibility score', async () => {
    let stdout = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'workbook-compatibility', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const proof = readProof(stdout)
    expect(proof.schemaVersion).toBe('bilig-evaluator.v1')
    expect(proof.door).toBe('workbook-compatibility')
    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      riskLevel: 'high',
      unsupportedFunctions: [{ name: 'CUBEVALUE', count: 1 }],
      volatileFunctions: [{ name: 'NOW', count: 1 }],
      formulaCellCount: 3,
      staleCachedFormulaCount: 1,
      checks: {
        commandSucceeded: true,
        inspectionCompleted: true,
        metadataScanCompleted: true,
        riskReasonsExplainFindings: true,
        noCompatibilityScore: true,
        unsupportedFunctionsReported: true,
      },
    })
    expect(stdout).not.toMatch(/compatibilityScore|excelCompatibilityPercent/u)
  })

  it('rejects WorkPaper doors from the XLSX evaluator package', async () => {
    let stderr = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'agent-mcp'], {
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Unknown bilig-evaluate door: agent-mcp')
  })

  it('rejects unknown doors with a focused error', async () => {
    let stderr = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'screenshots'], {
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Unknown bilig-evaluate door: screenshots')
  })
})

interface EvaluatorProofForTest {
  readonly schemaVersion: string
  readonly door: string
  readonly verified: boolean
  readonly evidence: Record<string, unknown>
}

function readProof(stdout: string): EvaluatorProofForTest {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error('Expected evaluator proof object.')
  }
  const evidence = parsed.evidence
  if (!isRecord(evidence)) {
    throw new Error('Expected evaluator proof evidence object.')
  }
  return {
    schemaVersion: requireString(parsed.schemaVersion),
    door: requireString(parsed.door),
    verified: parsed.verified === true,
    evidence,
  }
}

function requireString(value: unknown): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected string, got ${JSON.stringify(value)}.`)
  }
  return value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
