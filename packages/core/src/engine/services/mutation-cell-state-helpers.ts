import { ValueTag, type CellSnapshot, type LiteralInput } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { translateMutationFormulaForTarget } from './mutation-cell-content-helpers.js'

export interface ComparableMutationCellState {
  readonly formula?: string
  readonly value: LiteralInput | null
  readonly format: string | null
  readonly styleId?: string
  readonly authoredBlank?: boolean
}

export function readStoredMutationCellState(
  snapshot: CellSnapshot,
  format: string | null,
  cellFlags: number | undefined,
  styleId?: string,
): ComparableMutationCellState {
  if (snapshot.formula !== undefined) {
    return { formula: snapshot.formula, value: null, format, ...(styleId !== undefined ? { styleId } : {}) }
  }
  const authoredBlank = ((cellFlags ?? 0) & CellFlags.AuthoredBlank) !== 0
  switch (snapshot.value.tag) {
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return { value: snapshot.value.value, format, ...(styleId !== undefined ? { styleId } : {}) }
    case ValueTag.Empty:
    case ValueTag.Error:
      return { value: null, format, authoredBlank, ...(styleId !== undefined ? { styleId } : {}) }
  }
}

export function readDesiredMutationCellState(args: {
  readonly targetSheetName: string
  readonly targetAddress: string
  readonly snapshot: CellSnapshot
  readonly sourceSheetName?: string
  readonly sourceAddress?: string
  readonly formatOverride?: string | null
  readonly styleIdOverride?: string
}): ComparableMutationCellState {
  const format = args.formatOverride ?? args.snapshot.format ?? null
  const styleId = args.styleIdOverride ?? args.snapshot.styleId
  if (args.snapshot.formula !== undefined) {
    return {
      formula:
        args.sourceSheetName && args.sourceAddress
          ? translateMutationFormulaForTarget(
              args.snapshot.formula,
              args.sourceSheetName,
              args.sourceAddress,
              args.targetSheetName,
              args.targetAddress,
            )
          : args.snapshot.formula,
      value: null,
      format,
      ...(styleId !== undefined ? { styleId } : {}),
    }
  }
  switch (args.snapshot.value.tag) {
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return { value: args.snapshot.value.value, format, ...(styleId !== undefined ? { styleId } : {}) }
    case ValueTag.Empty:
    case ValueTag.Error:
      return {
        value: null,
        format,
        authoredBlank: (args.snapshot.flags & CellFlags.AuthoredBlank) !== 0,
        ...(styleId !== undefined ? { styleId } : {}),
      }
  }
}

export function hasMutationCellContent(snapshot: CellSnapshot): boolean {
  if (snapshot.formula !== undefined) {
    return true
  }
  return snapshot.value.tag !== ValueTag.Empty || (snapshot.flags & CellFlags.AuthoredBlank) !== 0
}

export function hasStoredMutationCellState(snapshot: CellSnapshot, format: string | undefined, cellFlags: number | undefined): boolean {
  return (
    snapshot.formula !== undefined ||
    snapshot.value.tag !== ValueTag.Empty ||
    format !== undefined ||
    ((cellFlags ?? 0) & CellFlags.AuthoredBlank) !== 0
  )
}

export function shouldApplyMutationCellState(current: ComparableMutationCellState, desired: ComparableMutationCellState): boolean {
  if (
    desired.formula === undefined &&
    desired.value === null &&
    desired.format === null &&
    desired.authoredBlank === true &&
    current.formula === undefined &&
    current.value === null &&
    current.format === null &&
    !(current.authoredBlank ?? false)
  ) {
    return false
  }
  return (
    current.formula !== desired.formula ||
    current.value !== desired.value ||
    current.format !== desired.format ||
    current.styleId !== desired.styleId ||
    (current.authoredBlank ?? false) !== (desired.authoredBlank ?? false)
  )
}
