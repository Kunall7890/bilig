// @vitest-environment jsdom
import { act, createElement, useEffect, useRef } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridSelectionSnapshot } from '@bilig/grid'
import type { WorkerRuntimeSelection } from '../runtime-session.js'
import { useWorkerWorkbookAgentContext } from '../use-worker-workbook-agent-context.js'

function AgentContextHarness(props: {
  selection: WorkerRuntimeSelection
  selectionSnapshot: GridSelectionSnapshot
  capture: (value: ReturnType<typeof useWorkerWorkbookAgentContext>) => void
  subscribeViewport?: (sheetName: string, viewport: { rowStart: number; rowEnd: number; colStart: number; colEnd: number }) => () => void
}) {
  const selectionRangeRef = useRef({
    sheetName: props.selectionSnapshot.sheetName,
    startAddress: props.selectionSnapshot.range.startAddress,
    endAddress: props.selectionSnapshot.range.endAddress,
  })
  const selectionSnapshotRef = useRef(props.selectionSnapshot)
  const selectionRef = useRef(props.selection)
  const workerHandleRef = useRef({
    viewportStore: createViewportStoreStub(),
  })
  const runtimeControllerRef = useRef({
    subscribeViewport: props.subscribeViewport ?? vi.fn(() => () => undefined),
  })

  useEffect(() => {
    selectionSnapshotRef.current = props.selectionSnapshot
    selectionRef.current = props.selection
    selectionRangeRef.current = {
      sheetName: props.selectionSnapshot.sheetName,
      startAddress: props.selectionSnapshot.range.startAddress,
      endAddress: props.selectionSnapshot.range.endAddress,
    }
    runtimeControllerRef.current.subscribeViewport = props.subscribeViewport ?? runtimeControllerRef.current.subscribeViewport
  }, [props.selection, props.selectionSnapshot, props.subscribeViewport])

  const state = useWorkerWorkbookAgentContext({
    selection: props.selection,
    selectionRangeRef,
    selectionSnapshotRef,
    selectionRef,
    workerHandleRef,
    runtimeControllerRef,
  })

  useEffect(() => {
    props.capture(state)
  }, [props, state])

  return createElement('div')
}

function mountHarness(): {
  root: Root
  render: (props: Parameters<typeof AgentContextHarness>[0]) => Promise<void>
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  const render = async (props: Parameters<typeof AgentContextHarness>[0]) => {
    await act(async () => {
      root.render(createElement(AgentContextHarness, props))
    })
  }
  return { root, render }
}

describe('useWorkerWorkbookAgentContext', () => {
  afterEach(() => {
    document.body.innerHTML = ''
    vi.clearAllMocks()
  })

  it('builds agent context from the current selection snapshot and viewport', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookAgentContext> | null = null
    await harness.render({
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectionSnapshot: {
        sheetName: 'Sheet1',
        address: 'A1',
        kind: 'cell',
        range: {
          startAddress: 'A1',
          endAddress: 'A1',
        },
      },
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected agent context capture')
    }

    const context = captured.getAgentContext()
    expect(context.selection.sheetName).toBe('Sheet1')
    expect(context.selection.address).toBe('A1')
    expect(context.rendered.selection?.range.startAddress).toBe('A1')
    expect(context.rendered.selection?.range.endAddress).toBe('A1')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('resets the visible viewport when the sheet changes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookAgentContext> | null = null
    await harness.render({
      selection: { sheetName: 'Sheet1', address: 'C3' },
      selectionSnapshot: {
        sheetName: 'Sheet1',
        address: 'C3',
        kind: 'cell',
        range: {
          startAddress: 'C3',
          endAddress: 'C3',
        },
      },
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected agent context capture')
    }

    await act(async () => {
      captured?.resetVisibleViewportForSheet({ sheetName: 'Sheet2', address: 'B4' })
    })

    const context = captured.getAgentContext()
    expect(context.viewport.rowStart).toBe(3)
    expect(context.viewport.colStart).toBe(1)
    expect(context.selection.sheetName).toBe('Sheet1')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('prewarms same-sheet deep selections before visible viewport feedback arrives', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeViewport = vi.fn(() => () => undefined)
    const harness = mountHarness()
    const capture = vi.fn()
    await harness.render({
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectionSnapshot: createCellSelectionSnapshot('Sheet1', 'A1'),
      subscribeViewport,
      capture,
    })

    await harness.render({
      selection: { sheetName: 'Sheet1', address: 'D250' },
      selectionSnapshot: createCellSelectionSnapshot('Sheet1', 'D250'),
      subscribeViewport,
      capture,
    })

    expect(subscribeViewport).toHaveBeenCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 153,
        rowEnd: 345,
        colStart: 0,
        colEnd: 131,
      }),
      expect.any(Function),
      { initialPatch: 'full' },
    )

    await act(async () => {
      harness.root.unmount()
    })
  })
})

function createCellSelectionSnapshot(sheetName: string, address: string): GridSelectionSnapshot {
  return {
    sheetName,
    address,
    kind: 'cell',
    range: {
      startAddress: address,
      endAddress: address,
    },
  }
}

function createViewportStoreStub() {
  return {
    peekCell(sheetName: string, address: string): CellSnapshot | undefined {
      return stringCell(sheetName, address, `${sheetName}:${address}`)
    },
    getCellStyle: vi.fn(() => null),
    getLastAuthoritativeRevision: vi.fn(() => 7),
    getLastMetrics: vi.fn(() => ({ batchId: 11 })),
  }
}

function stringCell(sheetName: string, address: string, value: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: {
      tag: ValueTag.String,
      value,
    },
    input: value,
    flags: 0,
    version: 1,
  }
}
