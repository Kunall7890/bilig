import { resolve } from 'node:path'

import { isWorkbookBenchmarkCorpusId, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'

export interface CaptureArgs {
  readonly allowIncompleteEvidence: boolean
  readonly biligProductionHost: string
  readonly biligProductionPort: number
  readonly biligUrl: string
  readonly biligUrlSource: 'default-dev' | 'explicit' | 'served-production'
  readonly biligStorageStatePath: string | null
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly deltaX: number
  readonly deltaY: number
  readonly googleSheetsUrl: string
  readonly googleSheetsStorageStatePath: string | null
  readonly headless: boolean
  readonly microsoftExcelWebUrl: string | null
  readonly microsoftExcelWebStorageStatePath: string | null
  readonly outputPath: string
  readonly readyTimeoutMs: number
  readonly sampleCount: number
  readonly storageStatePath: string | null
}

export interface EmitXlsxArgs {
  readonly check: boolean
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly targetDirectory: string
}

export interface SaveStorageStateArgs {
  readonly authUrl: string
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly headless: boolean
  readonly product: UiResponsivenessSameCorpusProduct
  readonly readyTimeoutMs: number
  readonly targetPath: string
}

export interface PreflightArgs {
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly googleSheetsUrl: string | null
  readonly googleSheetsStorageStatePath: string | null
  readonly headless: boolean
  readonly microsoftExcelWebUrl: string | null
  readonly microsoftExcelWebStorageStatePath: string | null
  readonly outputPath: string | null
  readonly readyTimeoutMs: number
  readonly storageStatePath: string | null
}

export const defaultCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
export const defaultViewport = { width: 1440, height: 900 } as const
export const defaultBiligProductionPreviewHost = '127.0.0.1'
export const defaultBiligProductionPreviewPort = 4180

export function defaultBiligSameCorpusUrl(corpusId: WorkbookBenchmarkCorpusId): string {
  return `http://localhost:5173/?benchmarkCorpus=${encodeURIComponent(corpusId)}`
}

export function productionBiligSameCorpusUrl(host: string, port: number, corpusId: WorkbookBenchmarkCorpusId): string {
  const browserHost = host === '0.0.0.0' ? '127.0.0.1' : host
  return `http://${browserHost}:${String(port)}/?benchmarkCorpus=${encodeURIComponent(corpusId)}&persist=0`
}

export function parsePreflightArgs(argv: readonly string[]): PreflightArgs | null {
  if (!argv.includes('--preflight')) {
    return null
  }
  const googleSheetsUrl = argumentValue(argv, '--google-sheets-url')
  const microsoftExcelWebUrl = argumentValue(argv, '--microsoft-excel-web-url')
  if (!googleSheetsUrl && !microsoftExcelWebUrl) {
    throw new Error('Same-corpus preflight requires --google-sheets-url, --microsoft-excel-web-url, or both.')
  }
  return {
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    googleSheetsUrl,
    googleSheetsStorageStatePath: resolveOptionalPath(argumentValue(argv, '--google-sheets-storage-state')),
    headless: !argv.includes('--headed'),
    microsoftExcelWebUrl,
    microsoftExcelWebStorageStatePath: resolveOptionalPath(argumentValue(argv, '--microsoft-excel-web-storage-state')),
    outputPath: resolveOptionalPath(argumentValue(argv, '--output')),
    readyTimeoutMs: parsePositiveInteger(argumentValue(argv, '--ready-timeout-ms') ?? '60000', '--ready-timeout-ms'),
    storageStatePath: resolveOptionalPath(argumentValue(argv, '--storage-state')),
  }
}

export function parseEmitXlsxArgs(argv: readonly string[]): EmitXlsxArgs | null {
  const emitIndex = argv.indexOf('--emit-xlsx')
  if (emitIndex === -1) {
    return null
  }
  const targetDirectory = requiredArgumentValue(argv, emitIndex, 'Missing directory after --emit-xlsx')
  return {
    check: argv.includes('--check'),
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    targetDirectory: resolve(targetDirectory),
  }
}

export function parseSaveStorageStateArgs(argv: readonly string[]): SaveStorageStateArgs | null {
  const saveIndex = argv.indexOf('--save-storage-state')
  if (saveIndex === -1) {
    return null
  }
  const targetPath = requiredArgumentValue(argv, saveIndex, 'Missing file path after --save-storage-state')
  const product = parseSameCorpusProduct(argumentValue(argv, '--auth-product') ?? 'google-sheets')
  const authUrl = argumentValue(argv, '--auth-url') ?? authUrlFromProductArgs(argv, product)
  if (!authUrl) {
    throw new Error('Missing auth URL. Pass --auth-url <url> or the product-specific URL flag.')
  }
  return {
    authUrl,
    corpusId: parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId),
    headless: argv.includes('--headless'),
    product,
    readyTimeoutMs: parsePositiveInteger(argumentValue(argv, '--ready-timeout-ms') ?? '300000', '--ready-timeout-ms'),
    targetPath: resolve(targetPath),
  }
}

export function parseCaptureArgs(argv: readonly string[]): CaptureArgs {
  const corpusId = parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId)
  const outputPath = argumentValue(argv, '--output')
  const googleSheetsUrl = argumentValue(argv, '--google-sheets-url')
  const microsoftExcelWebUrl = argumentValue(argv, '--microsoft-excel-web-url')
  if (!outputPath || !googleSheetsUrl) {
    throw new Error(
      [
        'Missing required arguments.',
        'Usage: bun scripts/capture-ui-responsiveness-same-corpus.ts',
        '  --output <capture.json>',
        '  --google-sheets-url <same-corpus-google-sheets-url>',
        '  [--microsoft-excel-web-url <same-corpus-excel-web-url>]',
        '  or: --emit-xlsx <directory>',
        '  [--bilig-url <local-bilig-url>] [--corpus wide-mixed-250k] [--samples 3] [--delta-x 0] [--delta-y 720] [--headed]',
        '  [--serve-bilig-production] [--bilig-production-port 4180] [--bilig-production-host 127.0.0.1]',
        '  [--storage-state <state.json>]',
        '  [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>] [--bilig-storage-state <state.json>]',
        '  [--allow-incomplete-evidence]',
        '  [--ready-timeout-ms 60000]',
      ].join('\n'),
    )
  }
  const sampleCount = parsePositiveInteger(argumentValue(argv, '--samples') ?? '3', '--samples')
  const readyTimeoutMs = parsePositiveInteger(argumentValue(argv, '--ready-timeout-ms') ?? '60000', '--ready-timeout-ms')
  const serveBiligProduction = argv.includes('--serve-bilig-production')
  const explicitBiligUrl = argumentValue(argv, '--bilig-url')
  if (serveBiligProduction && explicitBiligUrl) {
    throw new Error('Use either --serve-bilig-production or --bilig-url, not both.')
  }
  const biligProductionHost = parseHost(argumentValue(argv, '--bilig-production-host') ?? defaultBiligProductionPreviewHost)
  const biligProductionPort = parsePort(argumentValue(argv, '--bilig-production-port') ?? String(defaultBiligProductionPreviewPort))
  return {
    allowIncompleteEvidence: argv.includes('--allow-incomplete-evidence'),
    biligProductionHost,
    biligProductionPort,
    biligUrl: serveBiligProduction
      ? productionBiligSameCorpusUrl(biligProductionHost, biligProductionPort, corpusId)
      : (explicitBiligUrl ?? defaultBiligSameCorpusUrl(corpusId)),
    biligUrlSource: serveBiligProduction ? 'served-production' : explicitBiligUrl ? 'explicit' : 'default-dev',
    biligStorageStatePath: resolveOptionalPath(argumentValue(argv, '--bilig-storage-state')),
    corpusId,
    deltaX: parseNonNegativeNumber(argumentValue(argv, '--delta-x') ?? '0', '--delta-x'),
    deltaY: parseNonNegativeNumber(argumentValue(argv, '--delta-y') ?? '720', '--delta-y'),
    googleSheetsUrl,
    googleSheetsStorageStatePath: resolveOptionalPath(argumentValue(argv, '--google-sheets-storage-state')),
    headless: !argv.includes('--headed'),
    microsoftExcelWebUrl,
    microsoftExcelWebStorageStatePath: resolveOptionalPath(argumentValue(argv, '--microsoft-excel-web-storage-state')),
    outputPath: resolve(outputPath),
    readyTimeoutMs,
    sampleCount,
    storageStatePath: resolveOptionalPath(argumentValue(argv, '--storage-state')),
  }
}

export function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected same-corpus product: ${value}`)
}

function authUrlFromProductArgs(argv: readonly string[], product: UiResponsivenessSameCorpusProduct): string | null {
  if (product === 'bilig') {
    return argumentValue(argv, '--bilig-url')
  }
  if (product === 'google-sheets') {
    return argumentValue(argv, '--google-sheets-url')
  }
  return argumentValue(argv, '--microsoft-excel-web-url')
}

function parseCorpusId(value: string): WorkbookBenchmarkCorpusId {
  if (!isWorkbookBenchmarkCorpusId(value)) {
    throw new Error(`Unexpected workbook benchmark corpus id: ${value}`)
  }
  return value
}

function parsePositiveInteger(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error(`${flag} must be a positive integer`)
  }
  return parsed
}

function parsePort(value: string): number {
  const parsed = parsePositiveInteger(value, '--bilig-production-port')
  if (parsed > 65_535) {
    throw new Error('--bilig-production-port must be between 1 and 65535')
  }
  return parsed
}

function parseHost(value: string): string {
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed.includes('/') || trimmed.includes('?') || trimmed.includes('#')) {
    throw new Error('--bilig-production-host must be a host name or IP address')
  }
  return trimmed
}

function parseNonNegativeNumber(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(`${flag} must be a non-negative number`)
  }
  return parsed
}

function resolveOptionalPath(value: string | null): string | null {
  return value ? resolve(value) : null
}

function argumentValue(argv: readonly string[], name: string): string | null {
  const index = argv.indexOf(name)
  if (index === -1) {
    return null
  }
  return requiredArgumentValue(argv, index, `Missing value after ${name}`)
}

function requiredArgumentValue(argv: readonly string[], index: number, message: string): string {
  const value = argv[index + 1]
  if (value === undefined || value.trim().length === 0 || value.startsWith('-')) {
    throw new Error(message)
  }
  return value
}
