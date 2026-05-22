export function resolveRequiresLiveViewportState(input: {
  readonly isEditingCell: boolean
  readonly fillPreviewActive: boolean
  readonly isFillHandleDragging: boolean
  readonly hasActiveHeaderDrag: boolean
  readonly hasActiveResizeColumn: boolean
  readonly hasActiveResizeRow: boolean
  readonly hasColumnResizePreview: boolean
  readonly hasRowResizePreview: boolean
}): boolean {
  void input.isEditingCell
  void input.hasActiveResizeColumn
  void input.hasActiveResizeRow
  return (
    input.hasActiveHeaderDrag ||
    input.hasColumnResizePreview ||
    input.hasRowResizePreview ||
    input.fillPreviewActive ||
    input.isFillHandleDragging
  )
}
