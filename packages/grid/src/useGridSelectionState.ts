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
  void input.fillPreviewActive
  void input.hasActiveResizeColumn
  void input.hasActiveResizeRow
  void input.isFillHandleDragging
  return input.hasActiveHeaderDrag || input.hasColumnResizePreview || input.hasRowResizePreview
}
