import { createHash } from 'node:crypto'

import * as XLSX from 'xlsx'

import type { WorkbookBenchmarkCorpusCase } from '../packages/benchmarks/src/workbook-corpus.js'
import type { SameCorpusCaptureCorpusVerification } from './gen-ui-responsiveness-live-browser-scorecard.ts'

export const sameCorpusFingerprintVersion = 'same-corpus-fingerprint-v1'

export interface SameCorpusCaptureCorpusFingerprint {
  readonly version: typeof sameCorpusFingerprintVersion
  readonly corpusCaseId: string
  readonly workbookName: string
  readonly sheetCount: number
  readonly materializedCells: number
  readonly primaryViewport: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
  }
  readonly snapshotSha256: string
}

export interface SameCorpusFingerprint {
  readonly corpusFingerprint: SameCorpusCaptureCorpusFingerprint
  readonly materializedCells: number
  readonly sheetName: string
  readonly checkedCells: readonly {
    readonly address: string
    readonly expected: string
  }[]
}

export function verifyXlsxCorpusFingerprint(
  bytes: Uint8Array,
  corpus: WorkbookBenchmarkCorpusCase,
  method: SameCorpusCaptureCorpusVerification['method'],
): SameCorpusCaptureCorpusVerification {
  const fingerprint = buildSameCorpusFingerprint(corpus)
  const workbook = XLSX.read(Buffer.from(bytes), { type: 'buffer' })
  const worksheet = workbook.Sheets[fingerprint.sheetName]
  if (!worksheet) {
    throw new Error(`Same-corpus XLSX is missing sheet: ${fingerprint.sheetName}`)
  }
  const checkedCells = fingerprint.checkedCells.map((cell) => {
    const actual = normalizeSpreadsheetValue(worksheet[cell.address]?.v)
    if (actual !== cell.expected) {
      throw new Error(
        `Same-corpus XLSX cell mismatch at ${fingerprint.sheetName}!${cell.address}: expected ${cell.expected}, got ${actual}`,
      )
    }
    return {
      address: cell.address,
      expected: cell.expected,
      actual,
    }
  })
  return {
    verified: true,
    method,
    sheetName: fingerprint.sheetName,
    materializedCells: fingerprint.materializedCells,
    corpusFingerprint: fingerprint.corpusFingerprint,
    sourceWorkbookSha256: sha256Hex(bytes),
    checkedCells,
  }
}

export function buildSameCorpusFingerprint(corpus: WorkbookBenchmarkCorpusCase): SameCorpusFingerprint {
  const sheet = corpus.snapshot.sheets.find((candidate) => candidate.name === corpus.primaryViewport.sheetName)
  if (!sheet) {
    throw new Error(`Same-corpus snapshot is missing primary sheet: ${corpus.primaryViewport.sheetName}`)
  }
  const literalCells = sheet.cells
    .filter((cell) => cell.value !== undefined && cell.value !== null)
    .map((cell) => ({ address: cell.address, expected: normalizeSpreadsheetValue(cell.value) }))
  const checkedCells = selectFingerprintCells(literalCells)
  if (checkedCells.length < 3) {
    throw new Error(`Same-corpus fingerprint needs at least 3 literal cells for ${corpus.id}`)
  }
  return {
    corpusFingerprint: {
      version: sameCorpusFingerprintVersion,
      corpusCaseId: corpus.id,
      workbookName: corpus.snapshot.workbook.name,
      sheetCount: corpus.snapshot.sheets.length,
      materializedCells: corpus.materializedCellCount,
      primaryViewport: corpus.primaryViewport,
      snapshotSha256: sha256Hex(stableJsonBytes(corpus.snapshot)),
    },
    sheetName: sheet.name,
    materializedCells: corpus.materializedCellCount,
    checkedCells,
  }
}

function selectFingerprintCells(
  cells: readonly {
    readonly address: string
    readonly expected: string
  }[],
): readonly {
  readonly address: string
  readonly expected: string
}[] {
  const selected = new Map<string, { address: string; expected: string }>()
  for (const index of [0, 1, Math.floor(cells.length / 2), cells.length - 2, cells.length - 1]) {
    const cell = cells[index]
    if (cell) {
      selected.set(cell.address, cell)
    }
  }
  return [...selected.values()]
}

function stableJsonBytes(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(stableJsonValue(value)))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entryValue]) => [key, stableJsonValue(entryValue)]),
    )
  }
  return value
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function normalizeSpreadsheetValue(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value)
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  const serialized = JSON.stringify(value)
  if (serialized === undefined) {
    throw new Error(`Unable to normalize spreadsheet value of type ${typeof value}`)
  }
  return serialized
}
