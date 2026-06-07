import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import { recalculateXlsxFileToFile } from './file-recalc.js'
import type { XlsxFormulaRecalcOptions, XlsxFormulaRecalcResult } from './types.js'

export const xlsxFormulaRecalcBytesApiLimit = 1_000_000

export async function recalculateXlsx(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxFormulaRecalcOptions = {},
): Promise<XlsxFormulaRecalcResult> {
  const bytes = toUint8Array(input)
  assertXlsxFormulaRecalcBytesApiWithinLimit(bytes)
  const tempDir = await mkdtemp(join(tmpdir(), 'xlsx-formula-recalc-'))
  try {
    const inputPath = join(tempDir, tempWorkbookFileName(options.fileName ?? 'workbook.xlsx'))
    const outputPath = join(tempDir, 'workbook.recalculated.xlsx')
    await writeFile(inputPath, bytes)
    const result = await recalculateXlsxFileToFile(inputPath, {
      ...options,
      outputPath,
    })
    return {
      xlsx: new Uint8Array(await readFile(outputPath)),
      warnings: result.warnings,
      sheetNames: result.sheetNames,
      reads: result.reads,
      changes: result.changes,
      ...(result.diagnostics === undefined ? {} : { diagnostics: result.diagnostics }),
    }
  } finally {
    await rm(tempDir, { recursive: true, force: true })
  }
}

function tempWorkbookFileName(fileName: string): string {
  const name = basename(fileName.trim()) || 'workbook.xlsx'
  return /\.[a-z0-9]+$/iu.test(name) ? name : `${name}.xlsx`
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input)
}

function assertXlsxFormulaRecalcBytesApiWithinLimit(bytes: Uint8Array): void {
  if (bytes.byteLength <= xlsxFormulaRecalcBytesApiLimit) {
    return
  }
  throw new Error(
    [
      `recalculateXlsx byte input is small-workbook only: source is ${bytes.byteLength.toString()} bytes`,
      `limit is ${xlsxFormulaRecalcBytesApiLimit.toString()} bytes`,
      'Use recalculateXlsxFileToFile for file-backed native XLSX jobs.',
    ].join('; '),
  )
}
