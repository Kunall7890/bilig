import { describe, expect, it, vi } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { createGridGeometrySnapshot } from '../gridGeometry.js'
import { handleGridBodyDoubleClick, handleGridPointerDown, handleGridPointerMove } from '../gridInteractionController.js'

function createInteractionState() {
  return {
    ignoreNextPointerSelectionRef: { current: false },
    pendingPointerCellRef: { current: null },
    dragAnchorCellRef: { current: null },
    dragPointerCellRef: { current: null },
    dragHeaderSelectionRef: { current: null },
    dragDidMoveRef: { current: false },
    postDragSelectionExpiryRef: { current: 0 },
    columnResizeActiveRef: { current: false },
  }
}

describe('gridInteractionController', () => {
  it('asks the commit hook before applying a body-click selection change even when edit state is stale', () => {
    const order: string[] = []
    const onCommitEdit = vi.fn(() => {
      order.push('commit')
    })
    const onSelectionChange = vi.fn(() => {
      order.push('selection')
    })

    handleGridPointerDown({
      event: {
        button: 0,
        clientX: 20,
        clientY: 30,
        shiftKey: false,
        preventDefault: vi.fn(),
        stopPropagation: vi.fn(),
      },
      columnWidths: {},
      defaultColumnWidth: 120,
      focusGrid: vi.fn(),
      interactionState: createInteractionState(),
      isEditingCell: false,
      onCommitEdit,
      onSelectionChange,
      resolvePointerGeometry: vi.fn(() => null),
      resolveColumnResizeTargetAtPointer: vi.fn(() => null),
      resolveHeaderSelectionAtPointer: vi.fn(() => null),
      resolvePointerCell: vi.fn(() => [3, 4] as const),
      selectedCell: [1, 1],
      setGridSelection: vi.fn(),
      visibleRegion: {
        range: { x: 0, y: 0, width: 10, height: 20 },
        tx: 0,
        ty: 0,
      },
    })

    expect(onCommitEdit).toHaveBeenCalledTimes(1)
    expect(onSelectionChange).toHaveBeenCalledTimes(1)
    expect(order).toEqual(['commit', 'selection'])
  })

  it('publishes the live rectangular drag selection instead of keeping it only in local grid state', () => {
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()

    handleGridPointerMove({
      event: {
        clientX: 80,
        clientY: 120,
        buttons: 1,
      },
      dragAnchorCell: [1, 22],
      dragHeaderSelection: null,
      dragPointerCell: [1, 22],
      interactionState: createInteractionState(),
      resolvePointerCell: vi.fn(() => [4, 31] as const),
      resolveHeaderSelectionForPointerDrag: vi.fn(),
      selectedCell: [1, 22],
      setGridSelection,
      visibleRegion: {
        range: { x: 0, y: 0, width: 20, height: 40 },
        tx: 0,
        ty: 0,
      },
      onSelectionChange,
      isEditingCell: false,
      onCommitEdit: vi.fn(),
    })

    expect(setGridSelection).toHaveBeenCalledTimes(1)
    expect(onSelectionChange).toHaveBeenCalledWith(
      expect.objectContaining({
        current: expect.objectContaining({
          cell: [1, 22],
          range: { x: 1, y: 22, width: 4, height: 10 },
        }),
      }),
    )
  })

  it('begins double-click editing from the clicked cell instead of a stale surface editor value', () => {
    const beginEditAt = vi.fn()
    const geometry = createGridGeometrySnapshot({
      sheetName: 'Sheet1',
      scrollLeft: 0,
      scrollTop: 0,
      hostWidth: 800,
      hostHeight: 600,
      dpr: 1,
      gridMetrics: getGridMetrics(),
    })

    handleGridBodyDoubleClick({
      event: {
        clientX: 20,
        clientY: 30,
        preventDefault: vi.fn(),
        stopPropagation: vi.fn(),
      },
      applyColumnWidth: vi.fn(),
      beginEditAt,
      columnWidths: {},
      computeAutofitColumnWidth: vi.fn(() => 120),
      defaultColumnWidth: 120,
      interactionState: createInteractionState(),
      isEditingCell: false,
      lastBodyClickCell: [2, 4],
      onCommitEdit: vi.fn(),
      onSelectionChange: vi.fn(),
      resolveColumnResizeTargetAtPointer: vi.fn(() => null),
      resolvePointerCell: vi.fn(() => [2, 4] as const),
      resolvePointerGeometry: vi.fn(() => geometry),
      selectedCell: [0, 0],
      setGridSelection: vi.fn(),
      visibleRegion: {
        range: { x: 0, y: 0, width: 10, height: 20 },
        tx: 0,
        ty: 0,
      },
    })

    expect(beginEditAt.mock.calls).toEqual([['C5']])
  })
})
