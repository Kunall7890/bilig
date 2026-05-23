import { memo, useCallback, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { Rectangle } from '../gridTypes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import { CELL_TEXT_PADDING_X } from '../text/gridTextPacket.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import {
  WORKBOOK_DEFAULT_FONT_SIZE,
  WORKBOOK_FONT_SANS,
  workbookDisplayFontCssPx,
  workbookDisplayLineHeightCssPx,
  workbookFontPointSizeToCssPx,
  workbookSnapCssPixel,
} from '../workbookTheme.js'
import { workbookNativeTextQualityStyle } from '../workbookTextQuality.js'
import { getWorkbookDevicePixelRatio, subscribeWorkbookDevicePixelRatioChange } from '../workbookDevicePixelRatio.js'
import type { TextQuadRun } from './line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'

type TextLayerPane = GridHeaderPaneState | WorkbookRenderTilePaneState

interface NativeTextRunVisibleClipV3 {
  readonly outerHeight: number
  readonly outerLeft: number
  readonly outerTop: number
  readonly outerWidth: number
  readonly innerHeight: number
  readonly innerLeft: number
  readonly innerTop: number
  readonly innerWidth: number
}

interface NativeTextViewportOffsetV3 {
  readonly x: number
  readonly y: number
}

export interface NativeTextRunFontStyleV3 {
  readonly fontFamily: string
  readonly fontSize: number
  readonly fontStyle: 'italic' | 'normal'
  readonly fontWeight: number | 'normal' | 'bold'
}

export interface SuppressedNativeTextCellV3 {
  readonly col: number
  readonly row: number
}

export interface WorkbookPaneNativeTextLayerV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly presentedScrollSnapshot?: WorkbookGridScrollSnapshot | null | undefined
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly selectionOcclusionRanges?: readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] | null | undefined
  readonly suppressedTextCell?: SuppressedNativeTextCellV3 | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

const EMPTY_SCROLL_SNAPSHOT: WorkbookGridScrollSnapshot = Object.freeze({ tx: 0, ty: 0 })
const EMPTY_VIEWPORT_OFFSET: NativeTextViewportOffsetV3 = Object.freeze({ x: 0, y: 0 })

function subscribeNoop(): () => void {
  return () => {}
}

function getNullSnapshot(): GridGeometrySnapshot | null {
  return null
}

function resolvePaneRenderOffset(
  pane: {
    readonly contentOffset: { readonly x: number; readonly y: number }
    readonly scrollAxes: { readonly x: boolean; readonly y: boolean }
  },
  scrollSnapshot: {
    readonly tx: number
    readonly ty: number
    readonly renderTx?: number | undefined
    readonly renderTy?: number | undefined
  },
): { readonly x: number; readonly y: number } {
  const renderTx = scrollSnapshot.renderTx ?? scrollSnapshot.tx
  const renderTy = scrollSnapshot.renderTy ?? scrollSnapshot.ty
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? renderTx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? renderTy : 0),
  }
}

function resolveLatestGeometry(
  propGeometry: GridGeometrySnapshot | null,
  liveGeometry: GridGeometrySnapshot | null,
): GridGeometrySnapshot | null {
  if (!propGeometry) {
    return liveGeometry
  }
  if (!liveGeometry) {
    return propGeometry
  }
  return liveGeometry.camera.seq > propGeometry.camera.seq ? liveGeometry : propGeometry
}

export function resolveNativeTextLayerDrawScrollSnapshotV3(input: {
  readonly geometry: GridGeometrySnapshot | null
  readonly liveScrollSnapshot: WorkbookGridScrollSnapshot
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly presentedScrollSnapshot?: WorkbookGridScrollSnapshot | null | undefined
}): WorkbookGridScrollSnapshot {
  if (input.presentedScrollSnapshot) {
    return input.presentedScrollSnapshot
  }
  return resolveTypeGpuV3DrawScrollSnapshot({
    fallback: input.liveScrollSnapshot,
    geometry: input.geometry,
    panes: input.panes,
  })
}

function isPaneDrawVisible(pane: TextLayerPane): boolean {
  return pane.drawVisible !== false
}

function getPaneTextRuns(pane: TextLayerPane): readonly TextQuadRun[] {
  return 'tile' in pane ? pane.tile.textRuns : pane.textRuns
}

function isNativeTextRunSuppressed(
  pane: TextLayerPane,
  run: TextQuadRun,
  suppressedTextCell: SuppressedNativeTextCellV3 | null | undefined,
): boolean {
  return Boolean(suppressedTextCell && 'tile' in pane && run.row === suppressedTextCell.row && run.col === suppressedTextCell.col)
}

function resolveNativeTextRunKey(pane: TextLayerPane, run: TextQuadRun): string {
  const paneIdentity =
    'tile' in pane
      ? [pane.paneId, pane.tile.tileId, pane.tile.coord.sheetOrdinal, pane.tile.coord.rowTile, pane.tile.coord.colTile].join(':')
      : pane.paneId
  const runIdentity =
    run.row !== undefined && run.col !== undefined ? ['cell', run.row, run.col].join(':') : ['header', run.text, run.x, run.y].join(':')
  return [
    paneIdentity,
    runIdentity,
    run.x,
    run.y,
    run.width ?? '',
    run.height ?? '',
    run.clipX ?? '',
    run.clipY ?? '',
    run.clipWidth ?? '',
    run.clipHeight ?? '',
    run.font ?? '',
    run.color ?? '',
  ].join(':')
}

function getDefaultFont(run: TextQuadRun): string {
  return run.font?.trim()
    ? run.font
    : `400 ${run.fontSize ?? workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)}px ${WORKBOOK_FONT_SANS}`
}

export function resolveNativeTextRunFontStyleV3(run: TextQuadRun): NativeTextRunFontStyleV3 {
  const font = getDefaultFont(run)
  const sizeCssPx = Math.max(1, run.fontSize ?? parseNativeTextRunFontSize(font))
  const sizeMatch = font.match(/\b\d+(?:\.\d+)?px\s+(.+)$/i)
  const fontFamily = sizeMatch?.[1]?.trim() || WORKBOOK_FONT_SANS
  const fontStyle = /\bitalic\b/i.test(font) ? 'italic' : 'normal'
  const weightMatch = font.match(/\b([1-9]00|normal|bold)\b/i)
  const rawWeight = weightMatch?.[1]?.toLowerCase()
  const fontWeight = rawWeight === 'bold' ? 'bold' : rawWeight === 'normal' || rawWeight === undefined ? 400 : Number(rawWeight)
  return {
    fontFamily,
    fontSize: sizeCssPx,
    fontStyle,
    fontWeight,
  }
}

function parseNativeTextRunFontSize(font: string): number {
  const match = font.match(/\b(\d+(?:\.\d+)?)px\b/i)
  return match ? Number(match[1]) : workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
}

function snapCssPixel(value: number, dpr: number): number {
  return workbookSnapCssPixel(value, dpr)
}

function snapCssViewportPixel(value: number, dpr: number, viewportOffset: NativeTextViewportOffsetV3, axis: 'x' | 'y'): number {
  const offset = axis === 'x' ? viewportOffset.x : viewportOffset.y
  return snapCssPixel(value + offset, dpr) - offset
}

export function snapNativeTextDisplayFontSizeV3(fontSize: number, dpr = 1): number {
  return workbookDisplayFontCssPx(fontSize, dpr)
}

function resolveNativeTextLineBoxV3(input: { readonly run: TextQuadRun; readonly dpr: number }): {
  readonly height: number
  readonly topInset: number
} {
  const fontStyle = resolveNativeTextRunFontStyleV3(input.run)
  const contentHeight = input.run.height ?? 0
  const lineHeight = workbookDisplayLineHeightCssPx(fontStyle.fontSize, input.dpr)
  return {
    height: lineHeight,
    topInset: snapCssPixel(Math.max(0, (contentHeight - lineHeight) / 2), input.dpr),
  }
}

export function resolveNativeTextRunVisibleClipV3(input: {
  readonly geometry?: GridGeometrySnapshot | null | undefined
  readonly pane: TextLayerPane
  readonly run: TextQuadRun
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly selectionOcclusionRanges?: readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] | null | undefined
  readonly dpr?: number | undefined
  readonly viewportOffset?: NativeTextViewportOffsetV3 | undefined
}): NativeTextRunVisibleClipV3 | null {
  const dpr = input.dpr ?? getWorkbookDevicePixelRatio()
  const viewportOffset = input.viewportOffset ?? EMPTY_VIEWPORT_OFFSET
  const offset = resolvePaneRenderOffset(input.pane, input.scrollSnapshot)
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clip = resolveNativeTextRunSelectionOccludedClipV3({
    clipHeight: input.run.clipHeight ?? height,
    clipWidth: input.run.clipWidth ?? width,
    clipX: input.run.clipX ?? input.run.x,
    clipY: input.run.clipY ?? input.run.y,
    geometry: input.geometry,
    pane: input.pane,
    run: input.run,
    selectionOcclusionRanges: input.selectionOcclusionRanges,
  })
  if (!clip) {
    return null
  }
  const { clipX, clipY, clipWidth, clipHeight } = clip
  const clipLeft = input.pane.frame.x + offset.x + clipX
  const clipTop = input.pane.frame.y + offset.y + clipY
  const clipRight = clipLeft + clipWidth
  const clipBottom = clipTop + clipHeight
  const paneLeft = input.pane.frame.x
  const paneTop = input.pane.frame.y
  const paneRight = paneLeft + input.pane.frame.width
  const paneBottom = paneTop + input.pane.frame.height
  const visibleLeft = Math.max(clipLeft, paneLeft)
  const visibleTop = Math.max(clipTop, paneTop)
  const visibleRight = Math.min(clipRight, paneRight)
  const visibleBottom = Math.min(clipBottom, paneBottom)
  if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) {
    return null
  }

  const contentLeft = input.pane.frame.x + offset.x + input.run.x
  const contentTop = input.pane.frame.y + offset.y + input.run.y
  const outerLeft = snapCssViewportPixel(visibleLeft, dpr, viewportOffset, 'x')
  const outerTop = snapCssViewportPixel(visibleTop, dpr, viewportOffset, 'y')
  const outerRight = snapCssViewportPixel(visibleRight, dpr, viewportOffset, 'x')
  const outerBottom = snapCssViewportPixel(visibleBottom, dpr, viewportOffset, 'y')
  const innerRight = Math.min(contentLeft + width, visibleRight)
  const innerBottom = Math.min(contentTop + height, visibleBottom)

  return {
    innerHeight: Math.max(0, innerBottom - contentTop),
    innerLeft: contentLeft - outerLeft,
    innerTop: contentTop - outerTop,
    innerWidth: Math.max(0, innerRight - contentLeft),
    outerHeight: Math.max(0, outerBottom - outerTop),
    outerLeft,
    outerTop,
    outerWidth: Math.max(0, outerRight - outerLeft),
  }
}

export function resolveNativeTextRunSelectionOccludedClipV3(input: {
  readonly clipHeight: number
  readonly clipWidth: number
  readonly clipX: number
  readonly clipY: number
  readonly geometry?: GridGeometrySnapshot | null | undefined
  readonly pane?: TextLayerPane | null | undefined
  readonly run: TextQuadRun
  readonly selectionOcclusionRanges?: readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] | null | undefined
}): { readonly clipX: number; readonly clipY: number; readonly clipWidth: number; readonly clipHeight: number } | null {
  let resolvedClip: { readonly clipX: number; readonly clipY: number; readonly clipWidth: number; readonly clipHeight: number } = {
    clipHeight: input.clipHeight,
    clipWidth: input.clipWidth,
    clipX: input.clipX,
    clipY: input.clipY,
  }
  const ranges = input.selectionOcclusionRanges ?? []
  const geometry = input.geometry
  const pane = input.pane
  const row = input.run.row
  const col = input.run.col
  if (!geometry || ranges.length === 0 || row === undefined || col === undefined || input.clipWidth <= 0 || input.clipHeight <= 0) {
    return resolvedClip
  }
  if (!pane || !('tile' in pane)) {
    return resolvedClip
  }

  const tileBaseX = geometry.columns.offsetOf(pane.tile.bounds.colStart)
  const tileBaseY = geometry.rows.offsetOf(pane.tile.bounds.rowStart)
  for (const range of ranges) {
    const nextClip = resolveNativeTextRunClipAgainstSelectionRangeV3({
      clip: resolvedClip,
      col,
      geometry,
      range,
      row,
      tileBaseX,
      tileBaseY,
    })
    if (!nextClip) {
      return null
    }
    resolvedClip = nextClip
  }

  return resolvedClip
}

function resolveNativeTextRunClipAgainstSelectionRangeV3(input: {
  readonly clip: { readonly clipX: number; readonly clipY: number; readonly clipWidth: number; readonly clipHeight: number }
  readonly col: number
  readonly geometry: GridGeometrySnapshot
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  readonly row: number
  readonly tileBaseX: number
  readonly tileBaseY: number
}): { readonly clipX: number; readonly clipY: number; readonly clipWidth: number; readonly clipHeight: number } | null {
  const rangeEndColExclusive = input.range.x + input.range.width
  const rangeEndRowExclusive = input.range.y + input.range.height
  const sourceInsideSelection =
    input.col >= input.range.x && input.col < rangeEndColExclusive && input.row >= input.range.y && input.row < rangeEndRowExclusive
  if (sourceInsideSelection) {
    return input.clip
  }

  const selectionTop = input.geometry.rows.offsetOf(input.range.y) - input.tileBaseY
  const selectionBottom = selectionTop + input.geometry.rows.span(input.range.y, rangeEndRowExclusive)
  const clipTop = input.clip.clipY
  const clipBottom = input.clip.clipY + input.clip.clipHeight
  if (clipBottom <= selectionTop || clipTop >= selectionBottom) {
    return input.clip
  }

  const selectionLeft = input.geometry.columns.offsetOf(input.range.x) - input.tileBaseX
  const selectionRight = selectionLeft + input.geometry.columns.span(input.range.x, rangeEndColExclusive)
  const clipLeft = input.clip.clipX
  const clipRight = input.clip.clipX + input.clip.clipWidth
  if (input.col < input.range.x && clipRight > selectionLeft) {
    const nextRight = Math.min(clipRight, selectionLeft)
    if (nextRight <= clipLeft) {
      return null
    }
    return {
      ...input.clip,
      clipWidth: nextRight - clipLeft,
    }
  }

  if (input.col >= rangeEndColExclusive && clipLeft < selectionRight) {
    const nextLeft = Math.max(clipLeft, selectionRight)
    if (clipRight <= nextLeft) {
      return null
    }
    return {
      ...input.clip,
      clipWidth: clipRight - nextLeft,
      clipX: nextLeft,
    }
  }

  return input.clip
}

export function resolveNativeTextRunOuterStyleV3(input: {
  readonly pane: TextLayerPane
  readonly run: TextQuadRun
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly dpr?: number | undefined
  readonly visibleClip?: NativeTextRunVisibleClipV3 | null | undefined
}): CSSProperties {
  const visibleClip = input.visibleClip ?? resolveNativeTextRunVisibleClipV3(input)
  if (!visibleClip) {
    return { display: 'none' }
  }
  return {
    height: visibleClip.outerHeight,
    left: visibleClip.outerLeft,
    overflow: 'hidden',
    position: 'absolute',
    top: visibleClip.outerTop,
    width: visibleClip.outerWidth,
  }
}

export function resolveNativeTextRunInnerStyleV3(input: {
  readonly run: TextQuadRun
  readonly dpr?: number | undefined
  readonly visibleClip?: NativeTextRunVisibleClipV3 | null | undefined
  readonly viewportOffset?: NativeTextViewportOffsetV3 | undefined
}): CSSProperties {
  const dpr = input.dpr ?? getWorkbookDevicePixelRatio()
  const viewportOffset = input.viewportOffset ?? EMPTY_VIEWPORT_OFFSET
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clipX = input.run.clipX ?? input.run.x
  const clipY = input.run.clipY ?? input.run.y
  const visibleClip = input.visibleClip ?? null
  const fontStyle = resolveNativeTextRunFontStyleV3(input.run)
  const displayFontSize = snapNativeTextDisplayFontSizeV3(fontStyle.fontSize, dpr)
  const lineBox = resolveNativeTextLineBoxV3({ dpr, run: input.run })
  const baseTop = visibleClip?.innerTop ?? input.run.y - clipY
  const baseLeft = visibleClip?.innerLeft ?? input.run.x - clipX
  const baseWidth = visibleClip?.innerWidth ?? width
  const baseHeight = visibleClip?.innerHeight ?? height
  const outerLeft = visibleClip?.outerLeft ?? 0
  const outerTop = visibleClip?.outerTop ?? 0
  const textAlign = input.run.align ?? 'left'
  const anchorOffset =
    textAlign === 'right' ? baseWidth - CELL_TEXT_PADDING_X : textAlign === 'center' ? baseWidth / 2 : CELL_TEXT_PADDING_X
  const anchorPosition = outerLeft + baseLeft + anchorOffset
  const anchorDelta = snapCssViewportPixel(anchorPosition, dpr, viewportOffset, 'x') - anchorPosition
  const paddingLeft = textAlign === 'left' || textAlign === 'center' ? CELL_TEXT_PADDING_X + anchorDelta : CELL_TEXT_PADDING_X
  const paddingRight = textAlign === 'right' || textAlign === 'center' ? CELL_TEXT_PADDING_X - anchorDelta : CELL_TEXT_PADDING_X
  const rawTextTop = input.run.wrap ? baseTop : baseTop + lineBox.topInset
  const textTopPosition = outerTop + rawTextTop
  const textTopDelta = snapCssViewportPixel(textTopPosition, dpr, viewportOffset, 'y') - textTopPosition
  const textTop = rawTextTop + textTopDelta
  return {
    ...workbookNativeTextQualityStyle,
    boxSizing: 'border-box',
    color: input.run.color ?? '#111827',
    display: 'block',
    fontFamily: fontStyle.fontFamily,
    fontFeatureSettings: 'normal',
    fontSize: displayFontSize,
    fontStyle: fontStyle.fontStyle,
    fontVariantNumeric: 'tabular-nums',
    fontWeight: fontStyle.fontWeight,
    height: input.run.wrap ? Math.max(0, baseHeight + Math.abs(textTopDelta)) : lineBox.height,
    letterSpacing: 0,
    left: baseLeft,
    lineHeight: `${lineBox.height}px`,
    overflow: 'hidden',
    paddingLeft,
    paddingRight,
    position: 'absolute',
    textAlign,
    textDecorationLine: input.run.underline ? 'underline' : input.run.strike ? 'line-through' : undefined,
    top: textTop,
    whiteSpace: input.run.wrap ? 'pre-wrap' : 'pre',
    width: baseWidth,
  }
}

function readNativeTextLayerViewportOffset(layerRef: { readonly current: HTMLDivElement | null }): NativeTextViewportOffsetV3 {
  const rect = layerRef.current?.getBoundingClientRect()
  return rect ? { x: rect.left, y: rect.top } : EMPTY_VIEWPORT_OFFSET
}

function useNativeTextLayerViewportOffset(layerRef: { readonly current: HTMLDivElement | null }): NativeTextViewportOffsetV3 {
  const [viewportOffset, setViewportOffset] = useState<NativeTextViewportOffsetV3>(EMPTY_VIEWPORT_OFFSET)
  const updateViewportOffset = useCallback((next: NativeTextViewportOffsetV3) => {
    setViewportOffset((current) => (Math.abs(current.x - next.x) < 0.001 && Math.abs(current.y - next.y) < 0.001 ? current : next))
  }, [])

  useLayoutEffect(() => {
    updateViewportOffset(readNativeTextLayerViewportOffset(layerRef))
  })

  useLayoutEffect(() => {
    if (typeof window === 'undefined') {
      return
    }
    let frame = 0
    const measure = () => updateViewportOffset(readNativeTextLayerViewportOffset(layerRef))
    const scheduleMeasure = () => {
      if (frame !== 0) {
        window.cancelAnimationFrame(frame)
      }
      frame = window.requestAnimationFrame(() => {
        frame = 0
        measure()
      })
    }

    measure()
    window.addEventListener('resize', scheduleMeasure)
    window.visualViewport?.addEventListener('resize', scheduleMeasure)
    window.visualViewport?.addEventListener('scroll', scheduleMeasure)
    return () => {
      if (frame !== 0) {
        window.cancelAnimationFrame(frame)
      }
      window.removeEventListener('resize', scheduleMeasure)
      window.visualViewport?.removeEventListener('resize', scheduleMeasure)
      window.visualViewport?.removeEventListener('scroll', scheduleMeasure)
    }
  }, [layerRef, updateViewportOffset])

  return viewportOffset
}

export const WorkbookPaneNativeTextLayerV3 = memo(function WorkbookPaneNativeTextLayerV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes,
  presentedScrollSnapshot = null,
  scrollTransformStore,
  selectionOcclusionRanges = null,
  suppressedTextCell = null,
  tilePanes,
}: WorkbookPaneNativeTextLayerV3Props) {
  const liveGeometry = useSyncExternalStore(
    cameraStore ? cameraStore.subscribe.bind(cameraStore) : subscribeNoop,
    cameraStore ? cameraStore.getSnapshot.bind(cameraStore) : getNullSnapshot,
    getNullSnapshot,
  )
  const scrollSnapshot = useSyncExternalStore(
    scrollTransformStore ? scrollTransformStore.subscribe.bind(scrollTransformStore) : subscribeNoop,
    scrollTransformStore ? scrollTransformStore.getSnapshot.bind(scrollTransformStore) : () => EMPTY_SCROLL_SNAPSHOT,
    () => EMPTY_SCROLL_SNAPSHOT,
  )
  const resolvedGeometry = resolveLatestGeometry(geometry, liveGeometry)
  const drawScrollSnapshot = useMemo(
    () =>
      resolveNativeTextLayerDrawScrollSnapshotV3({
        geometry: resolvedGeometry,
        liveScrollSnapshot: scrollSnapshot,
        panes: tilePanes,
        presentedScrollSnapshot,
      }),
    [presentedScrollSnapshot, resolvedGeometry, scrollSnapshot, tilePanes],
  )
  const panes = useMemo<readonly TextLayerPane[]>(() => [...tilePanes, ...headerPanes], [headerPanes, tilePanes])
  const dpr = useSyncExternalStore(subscribeWorkbookDevicePixelRatioChange, getWorkbookDevicePixelRatio, () => 1)
  const layerRef = useRef<HTMLDivElement | null>(null)
  const viewportOffset = useNativeTextLayerViewportOffset(layerRef)
  const renderedRuns = useMemo(
    () =>
      active
        ? panes.flatMap((pane) => {
            if (!isPaneDrawVisible(pane)) {
              return []
            }
            return getPaneTextRuns(pane).flatMap((run) => {
              if (isNativeTextRunSuppressed(pane, run, suppressedTextCell)) {
                return []
              }
              const width = run.width ?? 0
              const height = run.height ?? 0
              const clipWidth = run.clipWidth ?? width
              const clipHeight = run.clipHeight ?? height
              if (!run.text || clipWidth <= 0 || clipHeight <= 0) {
                return []
              }
              const visibleClip = resolveNativeTextRunVisibleClipV3({
                dpr,
                geometry: resolvedGeometry,
                pane,
                run,
                scrollSnapshot: drawScrollSnapshot,
                selectionOcclusionRanges,
                viewportOffset,
              })
              return visibleClip ? [{ pane, run, visibleClip }] : []
            })
          })
        : [],
    [active, dpr, drawScrollSnapshot, panes, resolvedGeometry, selectionOcclusionRanges, suppressedTextCell, viewportOffset],
  )
  const textRunCount = renderedRuns.length

  if (!active || textRunCount === 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-[15] overflow-hidden"
      data-v3-native-text-layer="mounted"
      data-v3-native-text-presented-scroll-left={drawScrollSnapshot.scrollLeft ?? ''}
      data-v3-native-text-presented-scroll-top={drawScrollSnapshot.scrollTop ?? ''}
      data-v3-native-text-render-tx={drawScrollSnapshot.renderTx ?? drawScrollSnapshot.tx}
      data-v3-native-text-render-ty={drawScrollSnapshot.renderTy ?? drawScrollSnapshot.ty}
      data-v3-native-text-run-count={textRunCount}
      data-testid="grid-native-text-layer"
      ref={layerRef}
      style={{ contain: 'strict' }}
    >
      {renderedRuns.map(({ pane, run, visibleClip }) => (
        <div
          data-native-text-run=""
          data-native-text-run-col={run.col ?? ''}
          data-native-text-run-row={run.row ?? ''}
          key={resolveNativeTextRunKey(pane, run)}
          style={resolveNativeTextRunOuterStyleV3({ dpr, pane, run, scrollSnapshot: drawScrollSnapshot, visibleClip })}
        >
          <div style={resolveNativeTextRunInnerStyleV3({ dpr, run, visibleClip, viewportOffset })}>{run.text}</div>
        </div>
      ))}
    </div>
  )
})
