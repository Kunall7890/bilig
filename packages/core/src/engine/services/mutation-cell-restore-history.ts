import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot, type LiteralInput } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import { CellFlags } from '../../cell-store.js'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../../cell-mutations-at.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { RuntimeFormula, RuntimeStructuralFormulaSourceTransform, TransactionRecord } from '../runtime-state.js'
import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import {
  CapturedCellMutationKind,
  materializeCapturedCellMutationRestores,
  type CapturedCellMutationRestores,
} from './mutation-captured-cell-restores.js'
import type { FastMutationHistoryResult } from './mutation-history-fast-path.js'
import {
  createExistingNumericCellMutationsTransactionRecord,
  createLazyCellMutationTransactionRecord,
  createLazyMaterializedCellMutationTransactionRecord,
  createLazySingleOpTransactionRecord,
} from './mutation-transaction-records.js'

interface MutationFormulaStore {
  get(cellIndex: number): RuntimeFormula | undefined
}

export interface MutationCellRestoreHistoryRuntime {
  readonly workbook: WorkbookStore
  readonly formulas: MutationFormulaStore
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly getFormulaFamilyStructuralSourceTransform?: (cellIndex: number) => RuntimeStructuralFormulaSourceTransform | undefined
  readonly hasFormulaFamilyStructuralSourceTransforms?: () => boolean
}

export interface MutationCellRestoreHistoryHelpers {
  readonly restoreCellOpFromRef: (ref: EngineCellMutationRef) => EngineOp
  readonly tryRestoreSimpleCellOpFromStore: (
    sheetName: string,
    address: string,
  ) => Extract<EngineOp, { kind: 'setCellValue' | 'setCellFormula' | 'clearCell' }> | null
  readonly createLazyInverseCellMutationRecord: (refs: readonly EngineCellMutationRef[]) => TransactionRecord
  readonly tryCreateSingleExistingNumericInverseCellMutationRecord: (refs: readonly EngineCellMutationRef[]) => TransactionRecord | null
  readonly tryCreateExistingNumericInverseCellMutationRecord: (refs: readonly EngineCellMutationRef[]) => TransactionRecord | null
  readonly buildFastMutationHistoryFromRefs: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells: number,
    options?: {
      includeUndoOps?: boolean
    },
  ) => FastMutationHistoryResult
}

export function createMutationCellRestoreHistoryHelpers(args: MutationCellRestoreHistoryRuntime): MutationCellRestoreHistoryHelpers {
  const captureRuntimeFormulaSource = (cellIndex: number, formula: RuntimeFormula): string => {
    if (formula.structuralSourceTransform !== undefined) {
      return getRuntimeFormulaSource(formula)
    }
    const inheritedTransform =
      args.hasFormulaFamilyStructuralSourceTransforms?.() === true ? args.getFormulaFamilyStructuralSourceTransform?.(cellIndex) : undefined
    return getRuntimeFormulaSource(formula, inheritedTransform)
  }

  const restoreCellOpFromRef = (ref: EngineCellMutationRef): EngineOp => {
    const sheet = args.workbook.getSheetById(ref.sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${ref.sheetId}`)
    }
    const address = formatAddress(ref.mutation.row, ref.mutation.col)
    const existingCellIndex =
      sheet.structureVersion === 1
        ? sheet.grid.getPhysical(ref.mutation.row, ref.mutation.col)
        : sheet.grid.get(ref.mutation.row, ref.mutation.col)
    const cellIndex = existingCellIndex === -1 ? undefined : existingCellIndex
    if (cellIndex === undefined) {
      return { kind: 'clearCell', sheetName: sheet.name, address }
    }
    const cellStore = args.workbook.cellStore
    const formulaId = cellStore.formulaIds[cellIndex] ?? 0
    if (formulaId === 0) {
      const tag = cellStore.tags[cellIndex]
      if (tag === ValueTag.Number) {
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: cellStore.numbers[cellIndex] ?? 0,
        }
      }
      if (tag === ValueTag.Boolean) {
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: (cellStore.numbers[cellIndex] ?? 0) !== 0,
        }
      }
      if (tag === ValueTag.Empty || tag === ValueTag.Error || tag === undefined) {
        return { kind: 'clearCell', sheetName: sheet.name, address }
      }
    }
    const runtimeFormula = args.formulas.get(cellIndex)
    if (runtimeFormula?.source !== undefined) {
      return {
        kind: 'setCellFormula',
        sheetName: sheet.name,
        address,
        formula: captureRuntimeFormulaSource(cellIndex, runtimeFormula),
      }
    }
    const snapshot = args.getCellByIndex(cellIndex)
    if (snapshot.formula !== undefined) {
      return {
        kind: 'setCellFormula',
        sheetName: sheet.name,
        address,
        formula: snapshot.formula,
      }
    }
    switch (snapshot.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return (snapshot.flags & CellFlags.AuthoredBlank) !== 0
          ? {
              kind: 'setCellValue',
              sheetName: sheet.name,
              address,
              value: null,
            }
          : { kind: 'clearCell', sheetName: sheet.name, address }
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: snapshot.value.value,
        }
    }
  }

  const tryRestoreSimpleCellOpFromStore = (
    sheetName: string,
    address: string,
  ): Extract<EngineOp, { kind: 'setCellValue' | 'setCellFormula' | 'clearCell' }> | null => {
    const cellIndex = args.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return { kind: 'clearCell', sheetName, address }
    }
    const cellStore = args.workbook.cellStore
    const formulaId = cellStore.formulaIds[cellIndex] ?? 0
    if (formulaId === 0) {
      const tag = cellStore.tags[cellIndex]
      if (tag === ValueTag.Number) {
        return { kind: 'setCellValue', sheetName, address, value: cellStore.numbers[cellIndex] ?? 0 }
      }
      if (tag === ValueTag.Boolean) {
        return { kind: 'setCellValue', sheetName, address, value: (cellStore.numbers[cellIndex] ?? 0) !== 0 }
      }
      if (tag === ValueTag.Empty || tag === ValueTag.Error || tag === undefined) {
        return { kind: 'clearCell', sheetName, address }
      }
      return null
    }
    const runtimeFormula = args.formulas.get(cellIndex)
    if (runtimeFormula?.source === undefined) {
      return null
    }
    return {
      kind: 'setCellFormula',
      sheetName,
      address,
      formula: captureRuntimeFormulaSource(cellIndex, runtimeFormula),
    }
  }

  const captureInverseCellMutationRestores = (refs: readonly EngineCellMutationRef[]): CapturedCellMutationRestores => {
    const count = refs.length
    const sheetIds = new Uint32Array(count)
    const cellIndexPlusOnes = new Uint32Array(count)
    const rows = new Uint32Array(count)
    const cols = new Uint32Array(count)
    const kinds = new Uint8Array(count)
    const numbers = new Float64Array(count)
    let values: Array<LiteralInput | undefined> | undefined
    let formulas: Array<string | undefined> | undefined
    const cellStore = args.workbook.cellStore
    let potentialNewCells = 0
    let cachedSheetId = -1
    let cachedSheet: ReturnType<WorkbookStore['getSheetById']> | undefined

    for (let index = 0; index < count; index += 1) {
      const ref = refs[index]!
      const targetIndex = count - 1 - index
      const { sheetId, mutation } = ref
      sheetIds[targetIndex] = sheetId
      rows[targetIndex] = mutation.row
      cols[targetIndex] = mutation.col
      if (sheetId !== cachedSheetId) {
        cachedSheet = args.workbook.getSheetById(sheetId)
        cachedSheetId = sheetId
      }
      if (!cachedSheet) {
        throw new Error(`Unknown sheet id: ${sheetId}`)
      }
      const candidate = ref.cellIndex
      const existingCellIndex =
        candidate !== undefined &&
        cellStore.sheetIds[candidate] === sheetId &&
        cellStore.rows[candidate] === mutation.row &&
        cellStore.cols[candidate] === mutation.col
          ? candidate
          : cachedSheet.structureVersion === 1
            ? cachedSheet.grid.getPhysical(mutation.row, mutation.col)
            : cachedSheet.grid.get(mutation.row, mutation.col)
      const cellIndex = existingCellIndex === -1 ? undefined : existingCellIndex
      if (cellIndex === undefined) {
        kinds[targetIndex] = CapturedCellMutationKind.Clear
        continue
      }
      cellIndexPlusOnes[targetIndex] = cellIndex + 1
      if ((cellStore.formulaIds[cellIndex] ?? 0) === 0) {
        const tag = cellStore.tags[cellIndex]
        if (tag === ValueTag.Number) {
          kinds[targetIndex] = CapturedCellMutationKind.NumberValue
          numbers[targetIndex] = cellStore.numbers[cellIndex] ?? 0
          potentialNewCells += 1
          continue
        }
        if (tag === ValueTag.Boolean) {
          kinds[targetIndex] = CapturedCellMutationKind.BooleanValue
          numbers[targetIndex] = cellStore.numbers[cellIndex] ?? 0
          potentialNewCells += 1
          continue
        }
        if (tag === ValueTag.String) {
          const snapshot = args.getCellByIndex(cellIndex)
          kinds[targetIndex] = CapturedCellMutationKind.LiteralValue
          values ??= Array<LiteralInput | undefined>(count)
          values[targetIndex] = snapshot.value.tag === ValueTag.String ? snapshot.value.value : null
          potentialNewCells += 1
          continue
        }
        kinds[targetIndex] = CapturedCellMutationKind.Clear
        continue
      }
      const runtimeFormula = args.formulas.get(cellIndex)
      const formula =
        runtimeFormula?.source !== undefined
          ? captureRuntimeFormulaSource(cellIndex, runtimeFormula)
          : args.getCellByIndex(cellIndex).formula
      if (formula !== undefined) {
        kinds[targetIndex] = CapturedCellMutationKind.Formula
        formulas ??= Array<string | undefined>(count)
        formulas[targetIndex] = formula
        potentialNewCells += 1
        continue
      }
      const snapshot = args.getCellByIndex(cellIndex)
      if (
        (snapshot.flags & CellFlags.AuthoredBlank) !== 0 &&
        (snapshot.value.tag === ValueTag.Empty || snapshot.value.tag === ValueTag.Error)
      ) {
        kinds[targetIndex] = CapturedCellMutationKind.NullValue
        potentialNewCells += 1
        continue
      }
      kinds[targetIndex] = CapturedCellMutationKind.Clear
    }

    return {
      sheetIds,
      cellIndexPlusOnes,
      rows,
      cols,
      kinds,
      numbers,
      ...(values === undefined ? {} : { values }),
      ...(formulas === undefined ? {} : { formulas }),
      potentialNewCells,
    }
  }

  const createLazyInverseCellMutationRecord = (refs: readonly EngineCellMutationRef[]): TransactionRecord => {
    const captured = captureInverseCellMutationRestores(refs)
    return createLazyMaterializedCellMutationTransactionRecord(
      () => materializeCapturedCellMutationRestores(captured),
      captured.potentialNewCells,
    )
  }

  const tryCreateSingleExistingNumericInverseCellMutationRecord = (refs: readonly EngineCellMutationRef[]): TransactionRecord | null => {
    if (refs.length !== 1) {
      return null
    }
    const ref = refs[0]!
    const { mutation, sheetId } = ref
    if (mutation.kind !== 'setCellValue' || typeof mutation.value !== 'number') {
      return null
    }

    const cellStore = args.workbook.cellStore
    const candidate = ref.cellIndex
    let existingCellIndex: number | undefined
    if (
      candidate !== undefined &&
      cellStore.sheetIds[candidate] === sheetId &&
      cellStore.rows[candidate] === mutation.row &&
      cellStore.cols[candidate] === mutation.col
    ) {
      existingCellIndex = candidate
    } else {
      const sheet = args.workbook.getSheetById(sheetId)
      if (!sheet) {
        return null
      }
      existingCellIndex = sheet.grid.get(mutation.row, mutation.col)
    }
    if (existingCellIndex === -1 || existingCellIndex === undefined) {
      return null
    }
    if ((cellStore.formulaIds[existingCellIndex] ?? 0) !== 0 || cellStore.tags[existingCellIndex] !== ValueTag.Number) {
      return null
    }

    return createLazyCellMutationTransactionRecord(
      [
        {
          sheetId,
          cellIndex: existingCellIndex,
          mutation: {
            kind: 'setCellValue',
            row: mutation.row,
            col: mutation.col,
            value: cellStore.numbers[existingCellIndex] ?? 0,
          },
        },
      ],
      1,
    )
  }

  const tryCreateExistingNumericInverseCellMutationRecord = (refs: readonly EngineCellMutationRef[]): TransactionRecord | null => {
    const count = refs.length
    if (count === 0) {
      return null
    }
    const sheetIds = new Uint32Array(count)
    const cellIndexPlusOnes = new Uint32Array(count)
    const rows = new Uint32Array(count)
    const cols = new Uint32Array(count)
    const numbers = new Float64Array(count)
    const cellStore = args.workbook.cellStore
    let cachedSheetId = -1
    let cachedSheet: ReturnType<WorkbookStore['getSheetById']> | undefined

    for (let index = 0; index < count; index += 1) {
      const ref = refs[index]!
      const { mutation, sheetId } = ref
      if (mutation.kind !== 'setCellValue' || typeof mutation.value !== 'number') {
        return null
      }
      if (sheetId !== cachedSheetId) {
        cachedSheet = args.workbook.getSheetById(sheetId)
        cachedSheetId = sheetId
      }
      if (!cachedSheet) {
        return null
      }
      const candidate = ref.cellIndex
      const existingCellIndex =
        candidate !== undefined &&
        cellStore.sheetIds[candidate] === sheetId &&
        cellStore.rows[candidate] === mutation.row &&
        cellStore.cols[candidate] === mutation.col
          ? candidate
          : cachedSheet.structureVersion === 1
            ? cachedSheet.grid.getPhysical(mutation.row, mutation.col)
            : cachedSheet.grid.get(mutation.row, mutation.col)
      if (
        existingCellIndex === -1 ||
        (cellStore.formulaIds[existingCellIndex] ?? 0) !== 0 ||
        cellStore.tags[existingCellIndex] !== ValueTag.Number
      ) {
        return null
      }
      const targetIndex = count - 1 - index
      sheetIds[targetIndex] = sheetId
      cellIndexPlusOnes[targetIndex] = existingCellIndex + 1
      rows[targetIndex] = mutation.row
      cols[targetIndex] = mutation.col
      numbers[targetIndex] = cellStore.numbers[existingCellIndex] ?? 0
    }

    return createExistingNumericCellMutationsTransactionRecord(
      {
        sheetIds,
        cellIndexPlusOnes,
        rows,
        cols,
        numbers,
      },
      0,
    )
  }

  const buildFastMutationHistoryFromRefs = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells: number,
    options: {
      includeUndoOps?: boolean
    } = {},
  ): FastMutationHistoryResult => {
    if (refs.length === 1) {
      const ref = refs[0]!
      const forwardOp = cellMutationRefToEngineOp(args.workbook, ref)
      const inverseOp = restoreCellOpFromRef(ref)
      return {
        forward: createLazySingleOpTransactionRecord(forwardOp, potentialNewCells),
        inverse: createLazySingleOpTransactionRecord(inverseOp, 1),
        undoOps: options.includeUndoOps === false ? null : [structuredClone(inverseOp)],
      }
    }
    const forwardOps: EngineOp[] = Array.from({ length: refs.length })
    for (let index = 0; index < refs.length; index += 1) {
      const ref = refs[index]!
      forwardOps[index] = cellMutationRefToEngineOp(args.workbook, ref)
    }

    const inverseOps: EngineOp[] = []
    for (let index = refs.length - 1; index >= 0; index -= 1) {
      inverseOps.push(restoreCellOpFromRef(refs[index]!))
    }

    return {
      forward: { kind: 'ops', ops: forwardOps, potentialNewCells },
      inverse: { kind: 'ops', ops: inverseOps, potentialNewCells: refs.length },
      undoOps: options.includeUndoOps === false ? null : structuredClone(inverseOps),
    }
  }

  return {
    restoreCellOpFromRef,
    tryRestoreSimpleCellOpFromStore,
    createLazyInverseCellMutationRecord,
    tryCreateSingleExistingNumericInverseCellMutationRecord,
    tryCreateExistingNumericInverseCellMutationRecord,
    buildFastMutationHistoryFromRefs,
  }
}

export function tryMutationCellRefsFromOps(workbook: WorkbookStore, ops: readonly EngineOp[]): EngineCellMutationRef[] | null {
  if (ops.length === 0) {
    return []
  }
  const refs = Array<EngineCellMutationRef>(ops.length)
  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index]!
    if (op.kind !== 'setCellValue' && op.kind !== 'setCellFormula' && op.kind !== 'clearCell') {
      return null
    }
    const sheet = workbook.getSheet(op.sheetName)
    if (!sheet) {
      return null
    }
    const parsed = parseCellAddress(op.address, op.sheetName)
    const cellIndex = workbook.getCellIndex(op.sheetName, op.address)
    refs[index] = {
      sheetId: sheet.id,
      ...(cellIndex === undefined ? {} : { cellIndex }),
      mutation:
        op.kind === 'setCellValue'
          ? { kind: 'setCellValue', row: parsed.row, col: parsed.col, value: op.value }
          : op.kind === 'setCellFormula'
            ? { kind: 'setCellFormula', row: parsed.row, col: parsed.col, formula: op.formula }
            : { kind: 'clearCell', row: parsed.row, col: parsed.col },
    }
  }
  return refs
}
