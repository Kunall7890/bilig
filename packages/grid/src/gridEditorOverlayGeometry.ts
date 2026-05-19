import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { Rectangle } from './gridTypes.js'

export function applyEditorOverlayBounds(bounds: Rectangle, hostElement?: HTMLElement | null): void {
  if (typeof document === 'undefined') {
    return
  }
  const element = findEditorOverlayElement(hostElement)
  if (!(element instanceof HTMLElement)) {
    return
  }
  element.style.display = ''
  element.style.height = `${bounds.height}px`
  element.style.left = `${bounds.x}px`
  element.style.top = `${bounds.y}px`
  element.style.width = `${bounds.width}px`
}

export function clearEditorOverlayBounds(hostElement?: HTMLElement | null): void {
  if (typeof document === 'undefined') {
    return
  }
  const element = findEditorOverlayElement(hostElement)
  if (!(element instanceof HTMLElement)) {
    return
  }
  element.style.display = 'none'
  element.style.height = '0px'
  element.style.left = '-100000px'
  element.style.top = '-100000px'
  element.style.width = '0px'
}

function findEditorOverlayElement(hostElement?: HTMLElement | null): HTMLElement | null {
  const scoped = hostElement?.querySelector<HTMLElement>('[data-testid="cell-editor-overlay"]')
  if (scoped) {
    return scoped
  }
  const root = hostElement?.parentElement ?? hostElement?.ownerDocument?.body ?? document
  return root.querySelector<HTMLElement>('[data-testid="cell-editor-overlay"]')
}

export function resolveEditorOverlayScreenBounds(input: {
  readonly col: number
  readonly row: number
  readonly geometry: GridGeometrySnapshot | null
  readonly hostElement: HTMLElement | null
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
}): Rectangle | null {
  const localBounds = input.geometry?.editorScreenRect(input.col, input.row) ?? input.getCellLocalBounds(input.col, input.row)
  const hostBounds = input.hostElement?.getBoundingClientRect()
  if (!localBounds || !hostBounds) {
    return null
  }
  return {
    height: localBounds.height,
    width: localBounds.width,
    x: hostBounds.left + localBounds.x,
    y: hostBounds.top + localBounds.y,
  }
}
