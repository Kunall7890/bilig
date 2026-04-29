import { useWorkbookInteractionOverlayState, type WorkbookInteractionOverlayState } from './useWorkbookInteractionOverlayState.js'
import type { Rectangle } from './gridTypes.js'

export function useWorkbookGridInteractionRuntime(input: {
  readonly activeResizeColumn: number | null
  readonly activeResizeRow: number | null
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly hasColumnResizePreview: boolean
  readonly hasRowResizePreview: boolean
  readonly isEditingCell: boolean
  readonly selectedCol: number
  readonly selectedRow: number
  readonly visibleRange: Rectangle
}): WorkbookInteractionOverlayState {
  const {
    activeResizeColumn,
    activeResizeRow,
    getCellLocalBounds,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol,
    selectedRow,
    visibleRange,
  } = input

  return useWorkbookInteractionOverlayState({
    activeResizeColumn,
    activeResizeRow,
    getCellLocalBounds,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol,
    selectedRow,
    visibleRange,
  })
}
