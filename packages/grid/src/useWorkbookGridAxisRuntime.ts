import { useWorkbookAxisResizeState, type WorkbookAxisResizeState } from './useWorkbookAxisResizeState.js'

type AxisSizeOverrides = Readonly<Record<number, number>>
type HiddenAxisOverrides = Readonly<Record<number, true>>

export function useWorkbookGridAxisRuntime(input: {
  readonly controlledColumnWidths?: AxisSizeOverrides | undefined
  readonly controlledHiddenColumns?: HiddenAxisOverrides | undefined
  readonly controlledHiddenRows?: HiddenAxisOverrides | undefined
  readonly controlledRowHeights?: AxisSizeOverrides | undefined
  readonly onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined
  readonly onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined
  readonly sheetName: string
}): WorkbookAxisResizeState {
  return useWorkbookAxisResizeState(input)
}
