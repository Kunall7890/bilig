import { ValueTag, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'

export function resolveProjectedStyleDirtyMask(input: {
  readonly snapshot: CellSnapshot
  readonly baseStyle: CellStyleRecord
  readonly nextStyle: CellStyleRecord
}): number {
  let mask = DirtyMaskV3.Style
  if (styleFillChanged(input.baseStyle, input.nextStyle)) {
    mask |= DirtyMaskV3.Rect
  }
  if (styleBordersChanged(input.baseStyle, input.nextStyle)) {
    mask |= DirtyMaskV3.Rect | DirtyMaskV3.Border
  }
  if (snapshotHasVisibleText(input.snapshot) && styleTextLayoutChanged(input.baseStyle, input.nextStyle)) {
    mask |= DirtyMaskV3.Text
  }
  return mask
}

function styleFillChanged(baseStyle: CellStyleRecord, nextStyle: CellStyleRecord): boolean {
  return (baseStyle.fill?.backgroundColor ?? null) !== (nextStyle.fill?.backgroundColor ?? null)
}

function styleBordersChanged(baseStyle: CellStyleRecord, nextStyle: CellStyleRecord): boolean {
  return stableSectionKey(baseStyle.borders) !== stableSectionKey(nextStyle.borders)
}

function styleTextLayoutChanged(baseStyle: CellStyleRecord, nextStyle: CellStyleRecord): boolean {
  return (
    stableSectionKey(baseStyle.font) !== stableSectionKey(nextStyle.font) ||
    stableSectionKey(baseStyle.alignment) !== stableSectionKey(nextStyle.alignment)
  )
}

function snapshotHasVisibleText(snapshot: CellSnapshot): boolean {
  switch (snapshot.value.tag) {
    case ValueTag.Boolean:
    case ValueTag.Error:
    case ValueTag.Number:
      return true
    case ValueTag.String:
      return snapshot.value.value.length > 0
    case ValueTag.Empty:
      return false
  }
}

function stableSectionKey(value: unknown): string {
  return JSON.stringify(value ?? null)
}
