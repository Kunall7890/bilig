import type { LiteralInput } from '@bilig/protocol'
import type {
  WorkbookAgentPreviewCellDiff,
  WorkbookAgentPreviewChangeKind,
  WorkbookAgentPreviewEffectSummary,
  WorkbookAgentPreviewRange,
  WorkbookAgentPreviewSemanticTarget,
  WorkbookAgentPreviewSemanticTargetKind,
  WorkbookAgentPreviewSummary,
} from './workbook-agent-bundle-types.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isLiteralInputValue(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function isPreviewChangeKind(value: unknown): value is WorkbookAgentPreviewChangeKind {
  return value === 'input' || value === 'formula' || value === 'style' || value === 'numberFormat'
}

function isPreviewSemanticTargetKind(value: unknown): value is WorkbookAgentPreviewSemanticTargetKind {
  return (
    value === 'table' || value === 'tableColumn' || value === 'tableHeaderRow' || value === 'tableDataBody' || value === 'tableTotalsRow'
  )
}

export function isWorkbookAgentPreviewRange(value: unknown): value is WorkbookAgentPreviewRange {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string' &&
    (value['role'] === 'target' || value['role'] === 'source')
  )
}

export function isWorkbookAgentPreviewSemanticTarget(value: unknown): value is WorkbookAgentPreviewSemanticTarget {
  return (
    isRecord(value) &&
    isPreviewSemanticTargetKind(value['kind']) &&
    typeof value['tableName'] === 'string' &&
    typeof value['label'] === 'string' &&
    (value['range'] === undefined || isWorkbookAgentPreviewRange(value['range'])) &&
    (value['columnName'] === undefined || typeof value['columnName'] === 'string') &&
    (value['columnIndex'] === undefined || (typeof value['columnIndex'] === 'number' && Number.isSafeInteger(value['columnIndex'])))
  )
}

export function isWorkbookAgentPreviewCellDiff(value: unknown): value is WorkbookAgentPreviewCellDiff {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    (value['beforeInput'] === null || isLiteralInputValue(value['beforeInput'])) &&
    (value['beforeFormula'] === null || typeof value['beforeFormula'] === 'string') &&
    (value['afterInput'] === null || isLiteralInputValue(value['afterInput'])) &&
    (value['afterFormula'] === null || typeof value['afterFormula'] === 'string') &&
    Array.isArray(value['changeKinds']) &&
    value['changeKinds'].every((entry) => isPreviewChangeKind(entry))
  )
}

export function isWorkbookAgentPreviewEffectSummary(value: unknown): value is WorkbookAgentPreviewEffectSummary {
  return (
    isRecord(value) &&
    typeof value['displayedCellDiffCount'] === 'number' &&
    Number.isFinite(value['displayedCellDiffCount']) &&
    typeof value['truncatedCellDiffs'] === 'boolean' &&
    typeof value['inputChangeCount'] === 'number' &&
    Number.isFinite(value['inputChangeCount']) &&
    typeof value['formulaChangeCount'] === 'number' &&
    Number.isFinite(value['formulaChangeCount']) &&
    typeof value['styleChangeCount'] === 'number' &&
    Number.isFinite(value['styleChangeCount']) &&
    typeof value['numberFormatChangeCount'] === 'number' &&
    Number.isFinite(value['numberFormatChangeCount']) &&
    typeof value['structuralChangeCount'] === 'number' &&
    Number.isFinite(value['structuralChangeCount'])
  )
}

export function isWorkbookAgentPreviewSummary(value: unknown): value is WorkbookAgentPreviewSummary {
  return (
    isRecord(value) &&
    Array.isArray(value['ranges']) &&
    value['ranges'].every((entry) => isWorkbookAgentPreviewRange(entry)) &&
    Array.isArray(value['structuralChanges']) &&
    value['structuralChanges'].every((entry) => typeof entry === 'string') &&
    Array.isArray(value['cellDiffs']) &&
    value['cellDiffs'].every((entry) => isWorkbookAgentPreviewCellDiff(entry)) &&
    (value['semanticTargets'] === undefined ||
      (Array.isArray(value['semanticTargets']) &&
        value['semanticTargets'].every((entry) => isWorkbookAgentPreviewSemanticTarget(entry)))) &&
    isWorkbookAgentPreviewEffectSummary(value['effectSummary'])
  )
}

function derivePreviewEffectSummary(input: {
  cellDiffs: readonly WorkbookAgentPreviewCellDiff[]
  structuralChanges: readonly string[]
  truncatedCellDiffs: boolean
}): WorkbookAgentPreviewEffectSummary {
  return {
    displayedCellDiffCount: input.cellDiffs.length,
    truncatedCellDiffs: input.truncatedCellDiffs,
    inputChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('input')).length,
    formulaChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('formula')).length,
    styleChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('style')).length,
    numberFormatChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('numberFormat')).length,
    structuralChangeCount: input.structuralChanges.length,
  }
}

function decodeWorkbookAgentPreviewCellDiff(value: unknown): WorkbookAgentPreviewCellDiff | null {
  if (!isRecord(value)) {
    return null
  }
  if (
    typeof value['sheetName'] !== 'string' ||
    typeof value['address'] !== 'string' ||
    (value['beforeInput'] !== null && !isLiteralInputValue(value['beforeInput'])) ||
    (value['beforeFormula'] !== null && typeof value['beforeFormula'] !== 'string') ||
    (value['afterInput'] !== null && !isLiteralInputValue(value['afterInput'])) ||
    (value['afterFormula'] !== null && typeof value['afterFormula'] !== 'string')
  ) {
    return null
  }
  const explicitChangeKinds = Array.isArray(value['changeKinds'])
    ? value['changeKinds'].flatMap((entry) => (isPreviewChangeKind(entry) ? [entry] : []))
    : []
  const derivedChangeKinds = explicitChangeKinds.length
    ? explicitChangeKinds
    : [
        ...(value['beforeFormula'] !== value['afterFormula'] ? (['formula'] as const) : []),
        ...(value['beforeInput'] !== value['afterInput'] ? (['input'] as const) : []),
      ]
  return {
    sheetName: value['sheetName'],
    address: value['address'],
    beforeInput: (value['beforeInput'] as LiteralInput | null | undefined) ?? null,
    beforeFormula: (value['beforeFormula'] as string | null | undefined) ?? null,
    afterInput: (value['afterInput'] as LiteralInput | null | undefined) ?? null,
    afterFormula: (value['afterFormula'] as string | null | undefined) ?? null,
    changeKinds: [...new Set(derivedChangeKinds)],
  }
}

export function decodeWorkbookAgentPreviewSummary(value: unknown): WorkbookAgentPreviewSummary | null {
  if (!isRecord(value)) {
    return null
  }
  if (
    !Array.isArray(value['ranges']) ||
    !value['ranges'].every((entry) => isWorkbookAgentPreviewRange(entry)) ||
    !Array.isArray(value['structuralChanges']) ||
    !value['structuralChanges'].every((entry) => typeof entry === 'string') ||
    !Array.isArray(value['cellDiffs'])
  ) {
    return null
  }
  const cellDiffs = value['cellDiffs'].flatMap((entry) => {
    const decoded = decodeWorkbookAgentPreviewCellDiff(entry)
    return decoded ? [decoded] : []
  })
  if (cellDiffs.length !== value['cellDiffs'].length) {
    return null
  }
  const semanticTargets = Array.isArray(value['semanticTargets'])
    ? value['semanticTargets'].flatMap((entry) => (isWorkbookAgentPreviewSemanticTarget(entry) ? [{ ...entry }] : []))
    : []
  if (Array.isArray(value['semanticTargets']) && semanticTargets.length !== value['semanticTargets'].length) {
    return null
  }
  const truncatedCellDiffs = isWorkbookAgentPreviewEffectSummary(value['effectSummary']) ? value['effectSummary'].truncatedCellDiffs : false
  return {
    ranges: value['ranges'].map((range) => ({ ...range })),
    structuralChanges: [...value['structuralChanges']],
    cellDiffs,
    semanticTargets,
    effectSummary: isWorkbookAgentPreviewEffectSummary(value['effectSummary'])
      ? { ...value['effectSummary'] }
      : derivePreviewEffectSummary({
          cellDiffs,
          structuralChanges: value['structuralChanges'],
          truncatedCellDiffs,
        }),
  }
}

export function sameWorkbookAgentPreviewRange(left: WorkbookAgentPreviewRange, right: WorkbookAgentPreviewRange): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.startAddress === right.startAddress &&
    left.endAddress === right.endAddress &&
    left.role === right.role
  )
}

function samePreviewCellDiff(left: WorkbookAgentPreviewCellDiff, right: WorkbookAgentPreviewCellDiff): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.beforeInput === right.beforeInput &&
    left.beforeFormula === right.beforeFormula &&
    left.afterInput === right.afterInput &&
    left.afterFormula === right.afterFormula &&
    left.changeKinds.length === right.changeKinds.length &&
    left.changeKinds.every((kind, index) => kind === right.changeKinds[index])
  )
}

function samePreviewSemanticTarget(left: WorkbookAgentPreviewSemanticTarget, right: WorkbookAgentPreviewSemanticTarget): boolean {
  return (
    left.kind === right.kind &&
    left.tableName === right.tableName &&
    left.label === right.label &&
    left.columnName === right.columnName &&
    left.columnIndex === right.columnIndex &&
    ((left.range === undefined && right.range === undefined) ||
      (left.range !== undefined && right.range !== undefined && sameWorkbookAgentPreviewRange(left.range, right.range)))
  )
}

export function areWorkbookAgentPreviewSummariesEqual(left: WorkbookAgentPreviewSummary, right: WorkbookAgentPreviewSummary): boolean {
  return (
    left.ranges.length === right.ranges.length &&
    left.ranges.every((range, index) => {
      const other = right.ranges[index]
      return other ? sameWorkbookAgentPreviewRange(range, other) : false
    }) &&
    left.structuralChanges.length === right.structuralChanges.length &&
    left.structuralChanges.every((change, index) => change === right.structuralChanges[index]) &&
    left.cellDiffs.length === right.cellDiffs.length &&
    left.cellDiffs.every((diff, index) => {
      const other = right.cellDiffs[index]
      return other ? samePreviewCellDiff(diff, other) : false
    }) &&
    (left.semanticTargets ?? []).length === (right.semanticTargets ?? []).length &&
    (left.semanticTargets ?? []).every((target, index) => {
      const other = (right.semanticTargets ?? [])[index]
      return other ? samePreviewSemanticTarget(target, other) : false
    }) &&
    left.effectSummary.displayedCellDiffCount === right.effectSummary.displayedCellDiffCount &&
    left.effectSummary.truncatedCellDiffs === right.effectSummary.truncatedCellDiffs &&
    left.effectSummary.inputChangeCount === right.effectSummary.inputChangeCount &&
    left.effectSummary.formulaChangeCount === right.effectSummary.formulaChangeCount &&
    left.effectSummary.styleChangeCount === right.effectSummary.styleChangeCount &&
    left.effectSummary.numberFormatChangeCount === right.effectSummary.numberFormatChangeCount &&
    left.effectSummary.structuralChangeCount === right.effectSummary.structuralChangeCount
  )
}
