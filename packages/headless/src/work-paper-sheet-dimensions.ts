import type { WorkPaperSheetDimensions } from './work-paper-types.js'

export function applyCachedSheetDimensionInsertion(args: {
  readonly axis: 'row' | 'column'
  readonly cache: Map<number, WorkPaperSheetDimensions>
  readonly count: number
  readonly sheetId: number
  readonly start: number
}): boolean {
  const cached = args.cache.get(args.sheetId)
  if (!cached) {
    return false
  }
  if (args.count <= 0) {
    return true
  }
  if (args.axis === 'row') {
    if (args.start < cached.height) {
      cached.height += args.count
    }
    return true
  }
  if (args.start < cached.width) {
    cached.width += args.count
  }
  return true
}
