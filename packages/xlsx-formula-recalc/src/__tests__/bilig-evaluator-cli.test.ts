import { describe, expect, it } from 'vitest'

import { buildBiligEvaluatorProof, listBiligEvaluatorDoors, runBiligEvaluatorCli } from '../evaluator-cli.js'

describe('bilig-evaluate CLI', () => {
  it('lists the four production evaluator doors', () => {
    expect(listBiligEvaluatorDoors().map((door) => door.door)).toEqual([
      'xlsx-cache',
      'workbook-compatibility',
      'workpaper-service',
      'agent-mcp',
    ])
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

  it('prints a verified WorkPaper service proof', async () => {
    const proof = await buildBiligEvaluatorProof('workpaper-service')

    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
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
    })
  })

  it('prints a verified agent MCP proof', async () => {
    const proof = await buildBiligEvaluatorProof('agent-mcp')

    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      editedCell: 'Inputs!B3',
      dependentCell: 'Summary!B3',
      before: 60_000,
      after: 96_000,
      afterRestore: 96_000,
      afterRestart: 96_000,
      checks: {
        listedFileBackedTools: true,
        restartReadbackMatchesAfter: true,
      },
    })
    expect(proof.evidence.tools).toContain('read_cell')
  })

  it('prints a verified revenue-plan agent MCP proof', async () => {
    let stdout = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'agent-mcp', '--scenario', 'revenue-plan', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const proof = readProof(stdout)
    expect(proof.schemaVersion).toBe('bilig-evaluator.v1')
    expect(proof.door).toBe('agent-mcp')
    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      scenario: 'revenue-plan',
      editedCell: 'Deals!C2',
      readbackRange: 'Summary!B2:B8',
      formulaFamilies: ['SUM', 'SUMIF', 'XLOOKUP', 'FILTER', 'named-expression', 'persistence', 'restart'],
      before: {
        totalRevenue: 27_300,
        westCustomers: 30,
        enterpriseArpa: 1_200,
        targetRevenue: 30_576,
        qualifiedCustomerCounts: [30, 18],
      },
      after: {
        totalRevenue: 36_900,
        westCustomers: 38,
        enterpriseArpa: 1_200,
        targetRevenue: 41_328,
        qualifiedCustomerCounts: [20, 30, 18],
      },
      afterRestart: {
        totalRevenue: 36_900,
        westCustomers: 38,
        enterpriseArpa: 1_200,
        targetRevenue: 41_328,
        qualifiedCustomerCounts: [20, 30, 18],
      },
      checks: {
        totalRevenueRecalculated: true,
        sumifReadbackChanged: true,
        xlookupReadbackStable: true,
        filterSpillUpdated: true,
        namedExpressionApplied: true,
        restartReadbackMatchesAfter: true,
      },
    })
  })

  it('prints a verified provider-backed agent MCP proof', async () => {
    let stdout = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'agent-mcp', '--scenario', 'provider-backed', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const proof = readProof(stdout)
    expect(proof.schemaVersion).toBe('bilig-evaluator.v1')
    expect(proof.door).toBe('agent-mcp')
    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      scenario: 'provider-backed',
      providerFunction: 'IMPORTRANGE',
      adapterSurface: 'web',
      target: 'Imports!B2',
      formula: '=IMPORTRANGE("source","Revenue!B2")',
      adapterFormula: '=IMPORTRANGE("source","Revenue!B2")+0',
      before: {
        displayValue: '#BLOCKED!',
      },
      after: {
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
      },
      afterRestart: {
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
      },
      diagnostics: [
        {
          code: 'provider-backed-adapter-missing',
          functionName: 'IMPORTRANGE',
          adapterSurface: 'web',
          errorText: '#BLOCKED!',
        },
      ],
      checks: {
        blockedReadbackIsBlocked: true,
        blockedDiagnosticExplainsAdapter: true,
        adapterBackedReadbackIsFresh: true,
        adapterBackedDiagnosticsCleared: true,
        restartReadbackMatchesAfter: true,
      },
    })
  })

  it('rejects unsupported door and scenario combinations', async () => {
    let stderr = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'workpaper-service', '--scenario', 'provider-backed'], {
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Scenario "provider-backed" is only available for --door agent-mcp.')
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
