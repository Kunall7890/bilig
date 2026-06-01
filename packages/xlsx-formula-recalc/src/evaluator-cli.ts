import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { WorkPaper } from '@bilig/headless'
import { buildAgentWorkbookChallengeProof, buildMcpChallengeProof, buildMcpRevenuePlanChallengeProof } from '@bilig/headless/cli'
import { runXlsxFormulaRecalcCli } from './cli-api.js'

export const biligEvaluatorSchemaVersion = 'bilig-evaluator.v1'

export type BiligEvaluatorDoor = 'xlsx-cache' | 'workpaper-service' | 'agent-mcp'
export type BiligEvaluatorScenario = 'default' | 'revenue-plan'

export interface BiligEvaluatorCliContext {
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

export interface BiligEvaluatorDoorSummary {
  readonly door: BiligEvaluatorDoor
  readonly label: string
  readonly command: string
  readonly docs: string
}

export interface BiligEvaluatorEvidence {
  readonly scenario?: BiligEvaluatorScenario | undefined
  readonly target?: string | undefined
  readonly editedCell?: string | undefined
  readonly dependentCell?: string | undefined
  readonly dependentCells?: readonly string[] | undefined
  readonly readbackRange?: string | undefined
  readonly formulaFamilies?: readonly string[] | undefined
  readonly before?: unknown
  readonly after?: unknown
  readonly afterRestore?: unknown
  readonly afterRestart?: unknown
  readonly persistedDocumentBytes?: number | undefined
  readonly staleCachedFormulaCount?: number | undefined
  readonly suggestedReads?: readonly string[] | undefined
  readonly toolCount?: number | undefined
  readonly tools?: readonly string[] | undefined
  readonly checks: Readonly<Record<string, boolean>>
}

export interface BiligEvaluatorProof {
  readonly schemaVersion: typeof biligEvaluatorSchemaVersion
  readonly door: BiligEvaluatorDoor
  readonly doorName: string
  readonly command: string
  readonly packageVersions: Readonly<Record<string, string>>
  readonly scenario?: BiligEvaluatorScenario | undefined
  readonly evidence: BiligEvaluatorEvidence
  readonly verified: boolean
  readonly durationMs: number
  readonly limitations: readonly string[]
  readonly next: {
    readonly docs: string
    readonly command: string
  }
  readonly sourceProof: unknown
}

interface ParsedEvaluatorArgs {
  readonly door: BiligEvaluatorDoor | undefined
  readonly scenario: BiligEvaluatorScenario
  readonly json: boolean
  readonly markdown: boolean
  readonly list: boolean
  readonly help: boolean
}

interface XlsxCacheDoctorFormulaSummary {
  readonly target?: string | undefined
  readonly formula?: string | undefined
  readonly cachedValue?: unknown
  readonly literalRecalculatedValue?: unknown
  readonly stale?: boolean | undefined
}

interface XlsxCacheDoctorProofSummary {
  readonly commandSucceeded?: boolean | undefined
  readonly inspectionCompleted?: boolean | undefined
  readonly recalculationCompleted?: boolean | undefined
  readonly formulaCellCount?: number | undefined
  readonly staleCachedFormulaCount?: number | undefined
  readonly suggestedReads?: readonly string[] | undefined
  readonly formulas?: readonly XlsxCacheDoctorFormulaSummary[] | undefined
  readonly [key: string]: unknown
}

const xlsxCacheDocs = 'https://proompteng.github.io/bilig/eval-xlsx-cache-doctor.html'
const workpaperDocs = 'https://proompteng.github.io/bilig/eval-workpaper-service.html'
const mcpDocs = 'https://proompteng.github.io/bilig/eval-agent-mcp.html'

const doorSummaries: readonly BiligEvaluatorDoorSummary[] = [
  {
    door: 'xlsx-cache',
    label: 'Detect and recalculate a stale cached XLSX formula without Excel or LibreOffice.',
    command: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
    docs: xlsxCacheDocs,
  },
  {
    door: 'workpaper-service',
    label: 'Edit, recalculate, restore, and persist a WorkPaper service document.',
    command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
    docs: workpaperDocs,
  },
  {
    door: 'agent-mcp',
    label: 'Exercise the WorkPaper MCP tools through a local file-backed server restart.',
    command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    docs: mcpDocs,
  },
]

export function listBiligEvaluatorDoors(): readonly BiligEvaluatorDoorSummary[] {
  return doorSummaries
}

export function buildBiligEvaluatorProof(
  door: BiligEvaluatorDoor,
  options: { readonly scenario?: BiligEvaluatorScenario } = {},
): BiligEvaluatorProof {
  const startedAt = Date.now()
  const proof = buildDoorProof(door, options.scenario ?? 'default')
  return {
    ...proof,
    durationMs: Math.max(0, Date.now() - startedAt),
  }
}

export function runBiligEvaluatorCli(args: readonly string[], context: BiligEvaluatorCliContext = {}): number {
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    const parsed = parseEvaluatorArgs(args)
    if (parsed.help) {
      writeStdout(renderEvaluatorHelp())
      return 0
    }
    if (parsed.list) {
      const summary = {
        schemaVersion: biligEvaluatorSchemaVersion,
        doors: listBiligEvaluatorDoors(),
      }
      writeStdout(parsed.markdown ? renderDoorListMarkdown(summary.doors) : `${JSON.stringify(summary, null, 2)}\n`)
      return 0
    }
    if (parsed.door === undefined) {
      throw new Error('Expected --door xlsx-cache, --door workpaper-service, or --door agent-mcp.')
    }

    const proof = buildBiligEvaluatorProof(parsed.door, { scenario: parsed.scenario })
    writeStdout(parsed.markdown ? renderProofMarkdown(proof) : `${JSON.stringify(proof, null, 2)}\n`)
    return proof.verified ? 0 : 1
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

function buildDoorProof(door: BiligEvaluatorDoor, scenario: BiligEvaluatorScenario): Omit<BiligEvaluatorProof, 'durationMs'> {
  if (scenario !== 'default' && door !== 'agent-mcp') {
    throw new Error(`Scenario "${scenario}" is only available for --door agent-mcp.`)
  }
  switch (door) {
    case 'xlsx-cache':
      return buildXlsxCacheEvaluatorProof()
    case 'workpaper-service':
      return buildWorkpaperServiceEvaluatorProof()
    case 'agent-mcp':
      return buildAgentMcpEvaluatorProof(scenario)
  }
}

function buildXlsxCacheEvaluatorProof(): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof = buildXlsxCacheDoctorSourceProof()
  const firstFormula = sourceProof.formulas?.[0]
  const checks = {
    commandSucceeded: sourceProof.commandSucceeded === true,
    inspectionCompleted: sourceProof.inspectionCompleted === true,
    recalculationCompleted: sourceProof.recalculationCompleted === true,
    staleCachedFormulaFound: (sourceProof.staleCachedFormulaCount ?? 0) > 0,
    readbackSuggested: (sourceProof.suggestedReads?.length ?? 0) > 0,
  }
  const verified = Object.values(checks).every(Boolean)
  const summary = requireDoorSummary('xlsx-cache')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'xlsx-cache',
    doorName: 'XLSX stale-cache proof',
    command: summary.command,
    packageVersions: {
      'xlsx-formula-recalc': readLocalPackageVersion(),
      '@bilig/workpaper-runtime': WorkPaper.version,
    },
    evidence: {
      target: firstFormula?.target,
      before: firstFormula?.cachedValue,
      after: firstFormula?.literalRecalculatedValue,
      staleCachedFormulaCount: sourceProof.staleCachedFormulaCount,
      suggestedReads: sourceProof.suggestedReads,
      checks,
    },
    verified,
    limitations: [
      'This door proves stale cached formula detection and Bilig recalculation readback; it does not prove Excel desktop parity.',
      'Use Excel, LibreOffice, or an oracle harness for macros, pivots, charts, unsupported functions, and UI-specific workbook behavior.',
    ],
    next: {
      docs: summary.docs,
      command: summary.command,
    },
    sourceProof,
  }
}

function buildWorkpaperServiceEvaluatorProof(): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof = buildAgentWorkbookChallengeProof()
  const checks = normalizeBooleanRecord(recordValue(sourceProof, 'checks'))
  const summary = requireDoorSummary('workpaper-service')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'workpaper-service',
    doorName: 'WorkPaper service proof',
    command: summary.command,
    packageVersions: {
      '@bilig/workpaper': WorkPaper.version,
      'xlsx-formula-recalc': readLocalPackageVersion(),
    },
    evidence: {
      editedCell: stringValue(sourceProof, 'editedCell'),
      dependentCell: stringValue(sourceProof, 'dependentCell'),
      before: propertyValue(sourceProof, 'before'),
      after: propertyValue(sourceProof, 'after'),
      afterRestore: propertyValue(sourceProof, 'afterRestore'),
      persistedDocumentBytes: numberValue(sourceProof, 'persistedDocumentBytes'),
      checks,
    },
    verified: booleanValue(sourceProof, 'verified'),
    limitations: stringArrayValue(sourceProof, 'limitations'),
    next: {
      docs: summary.docs,
      command: summary.command,
    },
    sourceProof,
  }
}

function buildAgentMcpEvaluatorProof(scenario: BiligEvaluatorScenario): Omit<BiligEvaluatorProof, 'durationMs'> {
  const sourceProof = scenario === 'revenue-plan' ? buildMcpRevenuePlanChallengeProof() : buildMcpChallengeProof()
  const checks = normalizeBooleanRecord(recordValue(sourceProof, 'checks'))
  const tools = stringArrayValue(sourceProof, 'tools')
  const summary = requireDoorSummary('agent-mcp')
  const command = commandForScenario(summary, scenario)
  const formulaFamilies = stringArrayValue(sourceProof, 'formulaFamilies')

  return {
    schemaVersion: biligEvaluatorSchemaVersion,
    door: 'agent-mcp',
    doorName: scenario === 'revenue-plan' ? 'Agent MCP revenue-plan proof' : 'Agent MCP proof',
    command,
    packageVersions: {
      '@bilig/workpaper': WorkPaper.version,
      'xlsx-formula-recalc': readLocalPackageVersion(),
    },
    ...(scenario === 'default' ? {} : { scenario }),
    evidence: {
      ...(scenario === 'default' ? {} : { scenario }),
      editedCell: stringValue(sourceProof, 'editedCell'),
      dependentCell: stringValue(sourceProof, 'dependentCell'),
      ...(scenario === 'revenue-plan'
        ? {
            dependentCells: ['Summary!B2', 'Summary!B3', 'Summary!B4', 'Summary!B5', 'Summary!B6:B8'],
            readbackRange: stringValue(sourceProof, 'readbackRange'),
            formulaFamilies,
          }
        : {}),
      before: propertyValue(sourceProof, 'before'),
      after: propertyValue(sourceProof, 'after'),
      afterRestore: propertyValue(sourceProof, 'afterRestore'),
      afterRestart: propertyValue(sourceProof, 'afterRestart'),
      persistedDocumentBytes: numberValue(sourceProof, 'persistedDocumentBytes'),
      toolCount: tools.length,
      tools,
      checks,
    },
    verified: booleanValue(sourceProof, 'verified'),
    limitations: stringArrayValue(sourceProof, 'limitations'),
    next: {
      docs: summary.docs,
      command,
    },
    sourceProof,
  }
}

function buildXlsxCacheDoctorSourceProof(): XlsxCacheDoctorProofSummary {
  let stdout = ''
  const exitCode = runXlsxFormulaRecalcCli(['--demo', '--json'], {
    commandName: 'xlsx-cache-doctor',
    stdout: (text) => {
      stdout += text
    },
  })
  if (exitCode !== 0) {
    throw new Error('xlsx-cache-doctor evaluator door failed.')
  }
  return parseXlsxCacheDoctorProof(stdout)
}

function parseXlsxCacheDoctorProof(text: string): XlsxCacheDoctorProofSummary {
  const parsed: unknown = JSON.parse(text)
  const source = requireRecord(parsed, 'xlsx-cache-doctor JSON output')
  const formulasRaw = arrayValue(source, 'formulas')
  const formulas = formulasRaw.map((formula) => {
    const record = requireRecord(formula, 'xlsx-cache-doctor formula summary')
    return {
      target: optionalString(record['target']),
      formula: optionalString(record['formula']),
      cachedValue: record['cachedValue'],
      literalRecalculatedValue: record['literalRecalculatedValue'],
      stale: optionalBoolean(record['stale']),
    }
  })

  return {
    ...source,
    commandSucceeded: optionalBoolean(source['commandSucceeded']),
    inspectionCompleted: optionalBoolean(source['inspectionCompleted']),
    recalculationCompleted: optionalBoolean(source['recalculationCompleted']),
    formulaCellCount: optionalNumber(source['formulaCellCount']),
    staleCachedFormulaCount: optionalNumber(source['staleCachedFormulaCount']),
    suggestedReads: stringArray(source['suggestedReads']),
    formulas,
  }
}

function parseEvaluatorArgs(args: readonly string[]): ParsedEvaluatorArgs {
  let door: BiligEvaluatorDoor | undefined
  let scenario: BiligEvaluatorScenario = 'default'
  let json = true
  let markdown = false
  let list = false
  let help = false

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error('Unexpected missing bilig-evaluate argument.')
    }
    if (arg === '--door') {
      door = parseDoor(requireNextArg(args, index, '--door'))
      index += 1
      continue
    }
    if (arg.startsWith('--door=')) {
      door = parseDoor(arg.slice('--door='.length))
      continue
    }
    if (arg === '--scenario') {
      scenario = parseScenario(requireNextArg(args, index, '--scenario'))
      index += 1
      continue
    }
    if (arg.startsWith('--scenario=')) {
      scenario = parseScenario(arg.slice('--scenario='.length))
      continue
    }
    switch (arg) {
      case '--json':
        json = true
        markdown = false
        break
      case '--markdown':
        markdown = true
        json = false
        break
      case '--list':
        list = true
        break
      case '--help':
      case '-h':
        help = true
        break
      default:
        throw new Error(`Unknown bilig-evaluate option: ${arg}`)
    }
  }

  return { door, scenario, json, markdown, list, help }
}

function parseDoor(value: string): BiligEvaluatorDoor {
  switch (value) {
    case 'xlsx':
    case 'xlsx-cache':
    case 'xlsx-cache-doctor':
      return 'xlsx-cache'
    case 'workpaper':
    case 'workpaper-service':
    case 'service':
      return 'workpaper-service'
    case 'agent':
    case 'agent-mcp':
    case 'mcp':
      return 'agent-mcp'
    default:
      throw new Error(`Unknown bilig-evaluate door: ${value}`)
  }
}

function parseScenario(value: string): BiligEvaluatorScenario {
  switch (value) {
    case 'default':
    case 'basic':
      return 'default'
    case 'revenue':
    case 'revenue-plan':
      return 'revenue-plan'
    default:
      throw new Error(`Unknown bilig-evaluate scenario: ${value}`)
  }
}

function requireNextArg(args: readonly string[], index: number, option: string): string {
  const value = args[index + 1]
  if (value === undefined || value.startsWith('-')) {
    throw new Error(`Expected value after ${option}.`)
  }
  return value
}

function renderEvaluatorHelp(): string {
  return `Usage: bilig-evaluate --door <door> [--scenario <scenario>] [--json|--markdown]

Doors:
  xlsx-cache          Detect stale cached XLSX formulas and prove recalculated readback.
  workpaper-service   Prove WorkPaper edit, recalc, restore, and persistence.
  agent-mcp           Prove MCP tools, restart persistence, and readback.

Options:
  --door <door>        Door to run.
  --scenario <name>    Optional scenario. Use revenue-plan with agent-mcp for a multi-formula workbook proof.
  --list              Print available doors.
  --json              Print JSON proof (default).
  --markdown          Print compact Markdown proof.
  -h, --help          Show this help.
`
}

function renderDoorListMarkdown(doors: readonly BiligEvaluatorDoorSummary[]): string {
  return ['# Bilig Evaluator Doors', '', ...doors.map((door) => `- \`${door.door}\`: ${door.label}\n  \`${door.command}\``), ''].join('\n')
}

function renderProofMarkdown(proof: BiligEvaluatorProof): string {
  return [
    `# ${proof.doorName}`,
    '',
    `- Door: \`${proof.door}\``,
    `- Verified: \`${String(proof.verified)}\``,
    `- Command: \`${proof.command}\``,
    `- Docs: ${proof.next.docs}`,
    `- Evidence: \`${JSON.stringify(proof.evidence)}\``,
    '',
  ].join('\n')
}

function requireDoorSummary(door: BiligEvaluatorDoor): BiligEvaluatorDoorSummary {
  const summary = doorSummaries.find((entry) => entry.door === door)
  if (summary === undefined) {
    throw new Error(`Missing evaluator door summary for ${door}.`)
  }
  return summary
}

function commandForScenario(summary: BiligEvaluatorDoorSummary, scenario: BiligEvaluatorScenario): string {
  if (summary.door !== 'agent-mcp' || scenario === 'default') {
    return summary.command
  }
  return 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json'
}

function readLocalPackageVersion(): string {
  const packageJsonPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json')
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  const record = requireRecord(parsed, packageJsonPath)
  const version = record['version']
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error(`Expected package version in ${packageJsonPath}.`)
  }
  return version
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object.`)
  }
  return value
}

function propertyValue(source: unknown, key: string): unknown {
  return requireRecord(source, 'proof')[key]
}

function recordValue(source: unknown, key: string): Record<string, unknown> {
  const value = propertyValue(source, key)
  if (value === undefined) {
    return {}
  }
  return requireRecord(value, key)
}

function normalizeBooleanRecord(source: Record<string, unknown>): Readonly<Record<string, boolean>> {
  const normalized: Record<string, boolean> = {}
  for (const [key, value] of Object.entries(source)) {
    if (typeof value === 'boolean') {
      normalized[key] = value
    }
  }
  return normalized
}

function stringValue(source: unknown, key: string): string | undefined {
  return optionalString(propertyValue(source, key))
}

function numberValue(source: unknown, key: string): number | undefined {
  return optionalNumber(propertyValue(source, key))
}

function booleanValue(source: unknown, key: string): boolean {
  return propertyValue(source, key) === true
}

function arrayValue(source: Record<string, unknown>, key: string): readonly unknown[] {
  const value = source[key]
  return Array.isArray(value) ? value : []
}

function stringArrayValue(source: unknown, key: string): readonly string[] {
  return stringArray(propertyValue(source, key))
}

function stringArray(value: unknown): readonly string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : []
}

function optionalString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function optionalNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function optionalBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
