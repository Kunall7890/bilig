import { useCallback, useSyncExternalStore } from 'react'
import type { CellSnapshot } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import {
  readViewportColumnWidths,
  readViewportHiddenColumns,
  readViewportHiddenRows,
  readViewportRowHeights,
} from './worker-workbook-view-state.js'

type SheetViewportChannel = 'columnWidths' | 'rowHeights' | 'hiddenColumns' | 'hiddenRows' | 'freeze'

type ViewportStoreLike = {
  subscribeSheetChannel(sheetName: string, channel: SheetViewportChannel, listener: () => void): () => void
  subscribeCell(sheetName: string, address: string, listener: () => void): () => void
  peekCell(sheetName: string, address: string): CellSnapshot | undefined
  getFreezeRows(sheetName: string): number
  getFreezeCols(sheetName: string): number
  getColumnWidths(sheetName: string): Readonly<Record<number, number>>
  getRowHeights(sheetName: string): Readonly<Record<number, number>>
  getHiddenColumns(sheetName: string): Readonly<Record<number, true>>
  getHiddenRows(sheetName: string): Readonly<Record<number, true>>
}

type WorkerHandleLike = {
  readonly viewportStore: ViewportStoreLike
}

type StructuralMutationMethod = 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns' | 'setFreezePane'

export function useWorkerWorkbookGridState(input: {
  workerHandle: WorkerHandleLike | null | undefined
  selection: WorkerRuntimeSelection
  emptySelectedCell: CellSnapshot
  invokeMutation: (method: StructuralMutationMethod, ...args: unknown[]) => Promise<void>
}) {
  const { workerHandle, selection, emptySelectedCell, invokeMutation } = input

  const columnWidths = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'columnWidths', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
  )

  const rowHeights = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'rowHeights', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
  )

  const hiddenColumns = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenColumns', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
  )

  const hiddenRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenRows', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
  )

  const freezeRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
  )

  const freezeCols = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
  )

  const selectedCell = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeCell(selection.sheetName, selection.address, listener) ?? (() => {}),
      [selection.address, selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.peekCell(selection.sheetName, selection.address) ?? emptySelectedCell,
    () => emptySelectedCell,
  )

  const invokeInsertRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> => invokeMutation('insertRows', sheetName, startRow, count),
    [invokeMutation],
  )

  const invokeDeleteRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> => invokeMutation('deleteRows', sheetName, startRow, count),
    [invokeMutation],
  )

  const invokeInsertColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> => invokeMutation('insertColumns', sheetName, startCol, count),
    [invokeMutation],
  )

  const invokeDeleteColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> => invokeMutation('deleteColumns', sheetName, startCol, count),
    [invokeMutation],
  )

  const invokeSetFreezePaneMutation = useCallback(
    (sheetName: string, rows: number, cols: number): Promise<void> => invokeMutation('setFreezePane', sheetName, rows, cols),
    [invokeMutation],
  )

  return {
    columnWidths,
    rowHeights,
    hiddenColumns,
    hiddenRows,
    freezeRows,
    freezeCols,
    selectedCell,
    invokeInsertRowsMutation,
    invokeDeleteRowsMutation,
    invokeInsertColumnsMutation,
    invokeDeleteColumnsMutation,
    invokeSetFreezePaneMutation,
  }
}
