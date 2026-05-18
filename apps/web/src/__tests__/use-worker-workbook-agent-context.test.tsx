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
  viewportStore?: ReturnType<typeof createViewportStoreStub>
}) {
  const selectionRangeRef = useRef({
    sheetName: props.selectionSnapshot.sheetName,
    startAddress: props.selectionSnapshot.range.startAddress,
    endAddress: props.selectionSnapshot.range.endAddress,
  })
  const selectionSnapshotRef = useRef(props.selectionSnapshot)
  const selectionRef = useRef(props.selection)
  const workerHandleRef = useRef({
    viewportStore: props.viewportStore ?? createViewportStoreStub(),
  })
  const runtimeControllerRef = useRef({
    subscribeViewport: props.subscribeViewport ?? vi.fn(() => () => undefined),
  })
  selectionSnapshotRef.current = props.selectionSnapshot
  selectionRef.current = props.selection
  selectionRangeRef.current = {
    sheetName: props.selectionSnapshot.sheetName,
    startAddress: props.selectionSnapshot.range.startAddress,
    endAddress: props.selectionSnapshot.range.endAddress,
  }
  runtimeControllerRef.current.subscribeViewport = props.subscribeViewport ?? runtimeControllerRef.current.subscribeViewport
  workerHandleRef.current.viewportStore = props.viewportStore ?? workerHandleRef.current.viewportStore

  useEffect(() => {
    runtimeControllerRef.current.subscribeViewport = props.subscribeViewport ?? runtimeControllerRef.current.subscribeViewport
  }, [props.selection, props.selectionSnapshot, props.subscribeViewport, props.viewportStore])

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

  it('does not advance the agent context version for rendered metadata churn when visible content is unchanged', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let viewportListener: (() => void) | null = null
    const subscribeViewport = vi.fn(
      (_sheetName: string, _viewport: { rowStart: number; rowEnd: number; colStart: number; colEnd: number }, listener: () => void) => {
        viewportListener = listener
        return () => {
          viewportListener = null
        }
      },
    )
    const viewportStore = createViewportStoreStub()
    const harness = mountHarness()
    const capturedStates: Array<ReturnType<typeof useWorkerWorkbookAgentContext>> = []
    await harness.render({
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectionSnapshot: createCellSelectionSnapshot('Sheet1', 'A1'),
      subscribeViewport,
      viewportStore,
      capture: (value) => {
        capturedStates.push(value)
      },
    })

    const initialVersion = capturedStates.at(-1)?.agentContextVersion
    viewportStore.setMetrics(8, 12)
    await act(async () => {
      viewportListener?.()
    })

    expect(capturedStates.at(-1)?.agentContextVersion).toBe(initialVersion)

    viewportStore.setCellValue('Sheet1:A1', 'changed')
    viewportStore.setMetrics(9, 13)
    await act(async () => {
      viewportListener?.()
    })

    expect(capturedStates.at(-1)?.agentContextVersion).not.toBe(initialVersion)

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('changes the agent context version when a range selection changes without moving the active cell', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const harness = mountHarness()
    const capturedStates: Array<ReturnType<typeof useWorkerWorkbookAgentContext>> = []
    const baseSelection = { sheetName: 'Sheet1', address: 'A1' }
    await harness.render({
      selection: baseSelection,
      selectionSnapshot: createCellSelectionSnapshot('Sheet1', 'A1'),
      capture: (value) => {
        capturedStates.push(value)
      },
    })

    const initialVersion = capturedStates.at(-1)?.agentContextVersion
    await harness.render({
      selection: baseSelection,
      selectionSnapshot: {
        sheetName: 'Sheet1',
        address: 'A1',
        kind: 'range',
        range: {
          startAddress: 'A1',
          endAddress: 'B2',
        },
      },
      capture: (value) => {
        capturedStates.push(value)
      },
    })

    expect(capturedStates.at(-1)?.agentContextVersion).not.toBe(initialVersion)
    expect(capturedStates.at(-1)?.getAgentContext().selection.range).toEqual({
      startAddress: 'A1',
      endAddress: 'B2',
    })

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
  let revision = 7
  let batchId = 11
  const cellValues = new Map<string, string>()
  return {
    peekCell(sheetName: string, address: string): CellSnapshot | undefined {
      return stringCell(sheetName, address, cellValues.get(`${sheetName}:${address}`) ?? `${sheetName}:${address}`)
    },
    getCellStyle: vi.fn(() => null),
    getLastAuthoritativeRevision: vi.fn(() => revision),
    getLastMetrics: vi.fn(() => ({ batchId })),
    setCellValue(key: string, value: string): void {
      cellValues.set(key, value)
    },
    setMetrics(nextRevision: number, nextBatchId: number): void {
      revision = nextRevision
      batchId = nextBatchId
    },
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
