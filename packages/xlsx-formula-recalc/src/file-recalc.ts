import type { LiteralInput } from '@bilig/protocol'
import { recalculateXlsxFileToFileStreamingNative } from '@bilig/xlsx'

import type { XlsxFormulaRecalcEdit, XlsxFormulaRecalcFileOptions, XlsxFormulaRecalcFileResult } from './types.js'

export async function recalculateXlsxFileToFile(
  inputPath: string,
  options: XlsxFormulaRecalcFileOptions,
): Promise<XlsxFormulaRecalcFileResult> {
  const legacyOptions = options as { readonly config?: unknown; readonly engine?: string; readonly fallbackPolicy?: string }
  const engine = legacyOptions.engine ?? 'auto'
  if (engine === 'workpaper') {
    throw legacyWorkPaperPathError('engine')
  }
  if (legacyOptions.fallbackPolicy === 'workpaper') {
    throw legacyWorkPaperPathError('fallbackPolicy')
  }
  if (legacyOptions.config !== undefined) {
    throw new Error('streaming-native does not support WorkPaper config options')
  }
  const edits = nativeLiteralEdits(options.edits)
  return await recalculateXlsxFileToFileStreamingNative(inputPath, {
    outputPath: options.outputPath,
    ...(edits === undefined ? {} : { edits }),
    ...(options.reads === undefined ? {} : { reads: options.reads }),
    ...(options.externalWorkbooks === undefined ? {} : { externalWorkbooks: options.externalWorkbooks }),
    ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
  })
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

function legacyWorkPaperPathError(optionName: 'engine' | 'fallbackPolicy'): Error {
  return new Error(
    `The primary @bilig/xlsx-formula-recalc file API no longer loads or exports WorkPaper through ${optionName}; use @bilig/workpaper for WorkPaper workflows.`,
  )
}
