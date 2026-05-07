import type { CellStyleRecord } from '@bilig/protocol'
import type { Rectangle } from './gridTypes.js'
import { parseGpuColor, type GridGpuColor, type GridGpuRect } from './gridGpuPrimitives.js'

export function createBorderRects(
  rect: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  side: 'top' | 'right' | 'bottom' | 'left',
  border: NonNullable<NonNullable<CellStyleRecord['borders']>['top']>,
): GridGpuRect[] {
  const thickness = border.weight === 'thick' ? 3 : border.weight === 'medium' ? 2 : 1
  const isHorizontal = side === 'top' || side === 'bottom'
  const edgeX = side === 'left' ? rect.x : side === 'right' ? rect.x + rect.width - 1 : rect.x
  const edgeY = side === 'top' ? rect.y : side === 'bottom' ? rect.y + rect.height - 1 : rect.y
  const length = isHorizontal ? rect.width : rect.height
  const color = parseGpuColor(border.color)

  if (length <= 0) {
    return []
  }

  switch (border.style) {
    case 'dashed':
      return createPatternBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal, 6, 4)
    case 'dotted':
      return createPatternBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal, 1, 3)
    case 'double':
      return createDoubleBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal)
    case 'solid':
    default:
      return [
        {
          x: isHorizontal ? edgeX : edgeX - thickness / 2,
          y: isHorizontal ? edgeY - thickness / 2 : edgeY,
          width: isHorizontal ? length : thickness,
          height: isHorizontal ? thickness : length,
          color,
        },
      ]
  }
}

function createPatternBorderRects(
  edgeX: number,
  edgeY: number,
  length: number,
  thickness: number,
  color: GridGpuColor,
  isHorizontal: boolean,
  segmentLength: number,
  gapLength: number,
): GridGpuRect[] {
  const rects: GridGpuRect[] = []
  for (let cursor = 0; cursor < length; cursor += segmentLength + gapLength) {
    const currentLength = Math.min(segmentLength, length - cursor)
    rects.push({
      x: isHorizontal ? edgeX + cursor : edgeX - thickness / 2,
      y: isHorizontal ? edgeY - thickness / 2 : edgeY + cursor,
      width: isHorizontal ? currentLength : thickness,
      height: isHorizontal ? thickness : currentLength,
      color,
    })
  }
  return rects
}

function createDoubleBorderRects(
  edgeX: number,
  edgeY: number,
  length: number,
  thickness: number,
  color: GridGpuColor,
  isHorizontal: boolean,
): GridGpuRect[] {
  const span = Math.max(3, thickness + 2)
  const offset = span / 2
  if (isHorizontal) {
    return [
      {
        x: edgeX,
        y: edgeY - offset,
        width: length,
        height: 1,
        color,
      },
      {
        x: edgeX,
        y: edgeY - offset + span - 1,
        width: length,
        height: 1,
        color,
      },
    ]
  }
  return [
    {
      x: edgeX - offset,
      y: edgeY,
      width: 1,
      height: length,
      color,
    },
    {
      x: edgeX - offset + span - 1,
      y: edgeY,
      width: 1,
      height: length,
      color,
    },
  ]
}
