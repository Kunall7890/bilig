import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { Rectangle } from './gridTypes.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

interface GridFillPreviewOverlayProps {
  readonly fillPreviewRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly getGeometrySnapshot: () => GridGeometrySnapshot | null
  readonly scrollTransformStore: WorkbookGridScrollStore
}

export function GridFillPreviewOverlay(props: GridFillPreviewOverlayProps) {
  const { fillPreviewRange, getGeometrySnapshot, scrollTransformStore } = props
  const [scrollVersion, setScrollVersion] = useState(0)

  useEffect(() => {
    return scrollTransformStore.subscribe(() => {
      setScrollVersion((version) => version + 1)
    })
  }, [scrollTransformStore])

  const rects = useMemo(() => {
    void scrollVersion
    if (!fillPreviewRange) {
      return []
    }
    return getGeometrySnapshot()?.rangeScreenRects(fillPreviewRange) ?? []
  }, [fillPreviewRange, getGeometrySnapshot, scrollVersion])

  if (rects.length === 0) {
    return null
  }

  return (
    <div aria-hidden="true" className="pointer-events-none absolute inset-0 z-20 overflow-hidden">
      {rects.map((bounds) => (
        <div
          className="pointer-events-none absolute box-border"
          data-grid-fill-preview="true"
          key={keyForFillPreviewRect(fillPreviewRange, bounds)}
          style={styleForFillPreviewRect(bounds)}
        />
      ))}
    </div>
  )
}

function keyForFillPreviewRect(range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null, bounds: Rectangle): string {
  return [range?.x ?? 0, range?.y ?? 0, range?.width ?? 0, range?.height ?? 0, bounds.x, bounds.y, bounds.width, bounds.height].join(':')
}

function styleForFillPreviewRect(bounds: Rectangle): CSSProperties {
  return {
    backgroundColor: 'var(--wb-fill-preview-fill)',
    border: '1px solid var(--wb-selection-accent)',
    height: bounds.height,
    left: bounds.x,
    top: bounds.y,
    width: bounds.width,
  }
}
