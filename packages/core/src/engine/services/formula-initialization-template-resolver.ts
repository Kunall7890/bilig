import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { translateSimpleDirectScalarFormula } from '../../formula/simple-direct-scalar-compile.js'
import {
  tryBuildInitialSimpleRowRelativeBinaryTemplateKey,
  type InitialTemplateFormulaCacheEntry,
} from './formula-initialization-template-keys.js'

export function createInitialTemplateFormulaResolver(
  compileTemplateFormula: (source: string, row: number, col: number) => FormulaTemplateResolution,
): (source: string, row: number, col: number) => FormulaTemplateResolution {
  const simpleTemplateCache = new Map<string, InitialTemplateFormulaCacheEntry>()
  return (source, row, col) => {
    const templateKey = tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source, row, col)
    const cached = templateKey === undefined ? undefined : simpleTemplateCache.get(templateKey)
    if (cached) {
      const anchorRowDelta = row - cached.anchorRow
      const anchorColDelta = col - cached.anchorCol
      const compiled = translateSimpleDirectScalarFormula(cached.anchorCompiled, anchorRowDelta, anchorColDelta, source)
      if (compiled) {
        return {
          ...cached.resolution,
          compiled,
          translated: cached.resolution.translated || anchorRowDelta !== 0 || anchorColDelta !== 0,
          rowDelta: cached.resolution.rowDelta + anchorRowDelta,
          colDelta: cached.resolution.colDelta + anchorColDelta,
        }
      }
    }
    const resolution = compileTemplateFormula(source, row, col)
    if (templateKey !== undefined) {
      simpleTemplateCache.set(templateKey, {
        resolution,
        anchorRow: row,
        anchorCol: col,
        anchorCompiled: resolution.compiled,
      })
    }
    return resolution
  }
}
