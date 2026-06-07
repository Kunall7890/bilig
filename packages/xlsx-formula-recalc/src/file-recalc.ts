import type { LiteralInput } from '@bilig/protocol'
import { recalculateXlsxFileToFileStreamingNative } from '@bilig/xlsx'

import type { XlsxFormulaRecalcEdit, XlsxFormulaRecalcFileOptions, XlsxFormulaRecalcFileResult } from './types.js'

export async function recalculateXlsxFileToFile(
  inputPath: string,
  options: XlsxFormulaRecalcFileOptions,
): Promise<XlsxFormulaRecalcFileResult> {
  const engine = options.engine ?? 'auto'
  if (engine === 'workpaper') {
    return await recalculateXlsxToFileWithWorkPaper(inputPath, options)
  }
  try {
    if ((options.externalWorkbooks?.length ?? 0) > 0) {
      throw new Error('streaming-native does not support external workbook companions')
    }
    if (options.config !== undefined) {
      throw new Error('streaming-native does not support WorkPaper config options')
    }
    const edits = nativeLiteralEdits(options.edits)
    return await recalculateXlsxFileToFileStreamingNative(inputPath, {
      outputPath: options.outputPath,
      ...(edits === undefined ? {} : { edits }),
      ...(options.reads === undefined ? {} : { reads: options.reads }),
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
    })
  } catch (error) {
    if (options.fallbackPolicy === 'workpaper') {
      return await recalculateXlsxToFileWithWorkPaper(inputPath, options)
    }
    throw error
  }
}

function nativeLiteralEdits(
  edits: readonly XlsxFormulaRecalcEdit[] | undefined,
): readonly { readonly target: string; readonly value: LiteralInput }[] | undefined {
  if (!edits || edits.length === 0) {
    return undefined
  }
  return edits.map((edit) => {
    const value = edit.value
    if (value === null || typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
      return { target: edit.target, value }
    }
    throw new Error(`streaming-native supports literal edits only: ${edit.target}`)
  })
}

async function recalculateXlsxToFileWithWorkPaper(
  inputPath: string,
  options: XlsxFormulaRecalcFileOptions,
): Promise<XlsxFormulaRecalcFileResult> {
  const [{ readFileSync }, { recalculateXlsxToFile }] = await Promise.all([import('node:fs'), import('./index.js')])
  return recalculateXlsxToFile(readFileSync(inputPath), { ...options, engine: 'workpaper', fileName: options.fileName ?? inputPath })
}
