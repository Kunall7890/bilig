import { describe, expect, it } from 'vitest'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames, type ProtocolFrame } from '@bilig/binary-protocol'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { createAppendBatchFrame } from './sync-frame-shared.js'
import {
  acceptLocalSnapshotChunk,
  getLocalCachedSnapshot,
  invalidateLocalSnapshotCache,
  maybeCompactLocalSession,
  publishLocalSnapshot,
  type LocalSnapshotSessionState,
} from './local-session-snapshot-store.js'

function createSession(documentId: string): LocalSnapshotSessionState {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `worksheet-host:${documentId}`,
  })
  engine.createSheet('Sheet1')
  return {
    documentId,
    engine,
    batches: [],
    latestSnapshot: null,
    snapshotCache: null,
    snapshotDirty: true,
    cursor: 0,
    replicaSnapshot: null,
    compactScheduled: false,
  }
}

describe('local-session-snapshot-store', () => {
  it('reuses cached snapshots until invalidated', () => {
    const session = createSession('doc-cache')

    const first = getLocalCachedSnapshot(session)
    const second = getLocalCachedSnapshot(session)

    expect(second).toBe(first)

    session.engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      [[123]],
    )
    invalidateLocalSnapshotCache(session)

    const third = getLocalCachedSnapshot(session)
    expect(third).not.toBe(first)
    expect(third.sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'A1', value: 123 }))
  })

  it('publishes and compacts snapshots through the shared helper', () => {
    const session = createSession('doc-compact')
    const broadcastFrames: string[] = []
    let capturedSnapshot: WorkbookSnapshot | null = null
    const appendFrames: ReturnType<typeof createAppendBatchFrame>[] = []

    session.engine.subscribeBatches((batch) => {
      appendFrames.push(createAppendBatchFrame(session.documentId, 1, batch))
    })
    session.engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      [[456]],
    )
    const appendFrame = appendFrames[0]
    if (!appendFrame) {
      throw new Error('Expected engine batch')
    }
    session.cursor = 1
    session.batches = [{ cursor: appendFrame.cursor, frame: appendFrame }]

    maybeCompactLocalSession(session, {
      broadcast: (_documentId, protocolFrame) => {
        broadcastFrames.push(protocolFrame.kind)
      },
      getSession: () => session,
      snapshotAssemblies: new Map(),
      maxBatchBacklog: 0,
      schedule: (callback) => callback(),
    })

    expect(session.latestSnapshot).not.toBeNull()
    expect(session.batches).toEqual([])
    expect(broadcastFrames).toContain('snapshotChunk')
    expect(broadcastFrames.at(-1)).toBe('cursorWatermark')

    capturedSnapshot = getLocalCachedSnapshot(session)
    publishLocalSnapshot(session, capturedSnapshot, (_documentId, protocolFrame) => {
      broadcastFrames.push(protocolFrame.kind)
    })

    expect(session.latestSnapshot?.cursor).toBeGreaterThan(1)
    expect(broadcastFrames.at(-1)).toBe('cursorWatermark')
  })

  it('does not import assembled snapshots that would rewind the local session cursor', () => {
    const session = createSession('doc-local')
    session.cursor = 10
    session.engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      [[999]],
    )
    const frames = createSnapshotChunkFrames({
      documentId: 'doc-local',
      snapshotId: 'stale-snapshot',
      cursor: 4,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode(
        JSON.stringify({
          version: 1,
          workbook: { name: 'doc-local' },
          sheets: [
            {
              name: 'Sheet1',
              order: 0,
              cells: [{ address: 'A1', value: 123 }],
            },
          ],
        }),
      ),
      chunkSize: 16,
    })
    const broadcastFrames: string[] = []
    const context = {
      broadcast: (_documentId: string, frame: ProtocolFrame) => {
        broadcastFrames.push(frame.kind)
      },
      snapshotAssemblies: new Map(),
    }

    frames.forEach((frame) => {
      acceptLocalSnapshotChunk(session, frame, context)
    })

    expect(session.cursor).toBe(10)
    expect(getLocalCachedSnapshot(session).sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'A1', value: 999 }))
    expect(broadcastFrames).toEqual([])
  })
})
