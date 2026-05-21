import { Effect } from 'effect'
import { ValueTag, type CellRangeRef, type CellSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import { CellFlags } from '../../cell-store.js'
import { normalizeRange } from '../../engine-range-utils.js'
import { WorkbookStore } from '../../workbook-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { EngineCellStateError } from '../errors.js'

export interface EngineCellStateService {
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly restoreCellOps: (sheetName: string, address: string) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly readRangeCells: (range: CellRangeRef) => Effect.Effect<CellSnapshot[][], EngineCellStateError>
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride?: string | null,
    styleIdOverride?: string,
    options?: CellStateRestoreOptions,
  ) => Effect.Effect<EngineOp[], EngineCellStateError>
  readonly captureStoredCellOpsNow: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly restoreCellOpsNow: (sheetName: string, address: string) => EngineOp[]
  readonly readRangeCellsNow: (range: CellRangeRef) => CellSnapshot[][]
  readonly toCellStateOpsNow: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride?: string | null,
    styleIdOverride?: string,
    options?: CellStateRestoreOptions,
  ) => EngineOp[]
}

export interface CellStateRestoreOptions {
  readonly clearExistingFormat?: boolean
  readonly forceFormatWrite?: boolean
  readonly skipTableHeaderRename?: boolean
}

function cellStateErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function translateFormulaForTarget(
  formula: string,
  sourceSheetName: string,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): string {
  const source = parseCellAddress(sourceAddress, sourceSheetName)
  const target = parseCellAddress(targetAddress, targetSheetName)
  return translateFormulaReferences(formula, target.row - source.row, target.col - source.col)
}

export function createEngineCellStateService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly getCell: (sheetName: string, address: string) => CellSnapshot
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
}): EngineCellStateService {
  const getStoredStyleId = (sheetName: string, address: string): string => {
    const parsed = parseCellAddress(address, sheetName)
    return args.state.workbook.getStyleId(sheetName, parsed.row, parsed.col)
  }

  const toCellStateOpsNow = (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride: string | null = snapshot.format ?? null,
    styleIdOverride = snapshot.styleId ?? WorkbookStore.defaultStyleId,
    options: CellStateRestoreOptions = { clearExistingFormat: true },
  ): EngineOp[] => {
    const resolvedOptions: CellStateRestoreOptions = { clearExistingFormat: true, ...options }
    const ops: EngineOp[] = []
    const targetCellIndex = args.state.workbook.getCellIndex(sheetName, address)
    const explicitCurrentFormat = targetCellIndex === undefined ? undefined : args.state.workbook.getCellFormat(targetCellIndex)
    const parsedTarget = parseCellAddress(address, sheetName)
    const currentRangeFormatId = args.state.workbook.getRangeFormatId(sheetName, parsedTarget.row, parsedTarget.col)
    const currentRangeFormat =
      currentRangeFormatId === WorkbookStore.defaultFormatId
        ? null
        : (args.state.workbook.getCellNumberFormat(currentRangeFormatId)?.code ?? null)
    const currentFormat = explicitCurrentFormat ?? currentRangeFormat
    const currentStyleId = args.state.workbook.getStyleId(sheetName, parsedTarget.row, parsedTarget.col)
    const nextFormat = formatOverride === '' ? null : formatOverride
    const isAuthoredBlank = (snapshot.flags & CellFlags.AuthoredBlank) !== 0
    const explicitBlankOp = { kind: 'setCellValue' as const, sheetName, address, value: null }
    const shouldRestoreExplicitBlank = (snapshot.version ?? 0) !== 0 || isAuthoredBlank
    if (snapshot.formula !== undefined) {
      const translatedFormula =
        sourceSheetName && sourceAddress
          ? translateFormulaForTarget(snapshot.formula, sourceSheetName, sourceAddress, sheetName, address)
          : snapshot.formula
      ops.push({ kind: 'setCellFormula', sheetName, address, formula: translatedFormula })
    } else {
      switch (snapshot.value.tag) {
        case ValueTag.Empty:
          ops.push(
            shouldRestoreExplicitBlank
              ? explicitBlankOp
              : {
                  kind: 'clearCell',
                  sheetName,
                  address,
                  ...(resolvedOptions.skipTableHeaderRename === true ? { skipTableHeaderRename: true } : {}),
                },
          )
          break
        case ValueTag.Number:
        case ValueTag.Boolean:
        case ValueTag.String:
          ops.push({ kind: 'setCellValue', sheetName, address, value: snapshot.value.value })
          break
        case ValueTag.Error:
          ops.push(
            shouldRestoreExplicitBlank
              ? explicitBlankOp
              : {
                  kind: 'clearCell',
                  sheetName,
                  address,
                  ...(resolvedOptions.skipTableHeaderRename === true ? { skipTableHeaderRename: true } : {}),
                },
          )
          break
      }
    }
    if (nextFormat !== null && (resolvedOptions.forceFormatWrite === true || nextFormat !== currentFormat)) {
      ops.push({
        kind: 'setCellFormat',
        sheetName,
        address,
        format: nextFormat,
      })
    } else if (nextFormat === null && resolvedOptions.clearExistingFormat === true) {
      if (explicitCurrentFormat !== undefined) {
        ops.push({
          kind: 'setCellFormat',
          sheetName,
          address,
          format: nextFormat,
        })
      }
      if (currentRangeFormatId !== WorkbookStore.defaultFormatId) {
        ops.push({
          kind: 'setFormatRange',
          range: {
            sheetName,
            startAddress: address,
            endAddress: address,
          },
          formatId: WorkbookStore.defaultFormatId,
        })
      }
    }
    if (styleIdOverride !== currentStyleId) {
      ops.push({
        kind: 'setStyleRange',
        range: {
          sheetName,
          startAddress: address,
          endAddress: address,
        },
        styleId: styleIdOverride,
      })
    }
    return ops
  }

  const captureStoredCellOpsNow = (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] =>
    toCellStateOpsNow(
      sheetName,
      address,
      args.getCellByIndex(cellIndex),
      sourceSheetName,
      sourceAddress,
      args.state.workbook.getCellFormat(cellIndex) ?? null,
      getStoredStyleId(sheetName, address),
      { forceFormatWrite: true },
    )

  const restoreCellOpsNow = (sheetName: string, address: string): EngineOp[] => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return [{ kind: 'clearCell', sheetName, address }]
    }
    const snapshot = args.getCellByIndex(cellIndex)
    const explicitFormat = args.state.workbook.getCellFormat(cellIndex) ?? null
    const explicitStyleId = getStoredStyleId(sheetName, address)
    const restoreOptions = { clearExistingFormat: false }
    if (snapshot.formula !== undefined) {
      return toCellStateOpsNow(sheetName, address, snapshot, undefined, undefined, explicitFormat, explicitStyleId, restoreOptions)
    }
    switch (snapshot.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return toCellStateOpsNow(sheetName, address, snapshot, undefined, undefined, explicitFormat, explicitStyleId, restoreOptions)
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return toCellStateOpsNow(sheetName, address, snapshot, undefined, undefined, explicitFormat, explicitStyleId, restoreOptions)
    }
  }

  const readRangeCellsNow = (range: CellRangeRef): CellSnapshot[][] => {
    const bounds = normalizeRange(range)
    const rows: CellSnapshot[][] = []
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const cells: CellSnapshot[] = []
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        cells.push(args.getCell(range.sheetName, formatAddress(row, col)))
      }
      rows.push(cells)
    }
    return rows
  }

  return {
    captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress) {
      return Effect.try({
        try: () => captureStoredCellOpsNow(cellIndex, sheetName, address, sourceSheetName, sourceAddress),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to capture stored cell ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    restoreCellOps(sheetName, address) {
      return Effect.try({
        try: () => restoreCellOpsNow(sheetName, address),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to restore cell ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    readRangeCells(range) {
      return Effect.try({
        try: () => readRangeCellsNow(range),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to read range ${range.sheetName}!${range.startAddress}:${range.endAddress}`, cause),
            cause,
          }),
      })
    },
    toCellStateOps(sheetName, address, snapshot, sourceSheetName, sourceAddress, formatOverride, styleIdOverride, options) {
      return Effect.try({
        try: () =>
          toCellStateOpsNow(sheetName, address, snapshot, sourceSheetName, sourceAddress, formatOverride, styleIdOverride, options),
        catch: (cause) =>
          new EngineCellStateError({
            message: cellStateErrorMessage(`Failed to materialize cell state ops for ${sheetName}!${address}`, cause),
            cause,
          }),
      })
    },
    captureStoredCellOpsNow,
    restoreCellOpsNow,
    readRangeCellsNow,
    toCellStateOpsNow,
  }
}
