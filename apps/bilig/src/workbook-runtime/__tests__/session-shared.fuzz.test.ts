import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from '@bilig/binary-protocol'
import { fuzzWorkbookSnapshotArbitrary, runProperty, type FuzzWorkbookSnapshot } from '@bilig/test-fuzz'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  acceptSnapshotChunk,
  buildBrowserUrl,
  createSnapshotPublicationFromBytes,
  decodeWorkbookBase64,
  decodeWorkbookSnapshotBytes,
  type CompletedSnapshotAssembly,
  type SnapshotAssemblyRegistry,
} from '../session-shared.js'

const snapshotEncoder = new TextEncoder()

describe('workbook runtime session fuzz', () => {
  it('roundtrips workbook upload base64 bytes and rejects generated corrupt encodings', async () => {
    await runProperty({
      suite: 'bilig/workbook-runtime/base64-roundtrip-and-rejection',
      arbitrary: fc.uint8Array({ minLength: 1, maxLength: 256 }),
      predicate: async (bytes) => {
        const encoded = Buffer.from(bytes).toString('base64')
        expect(decodeWorkbookBase64(`\n${encoded}\t`)).toEqual(bytes)

        const invalidCharacters = encoded.replace(/[A-Za-z0-9+/=]/u, '*')
        expect(() => decodeWorkbookBase64(invalidCharacters)).toThrow(/base64/)
        expect(() => decodeWorkbookBase64(encoded.slice(1))).toThrow(/base64/)
      },
    })
  })

  it('assembles generated snapshot chunks exactly once across reordered runtime session frames', async () => {
    await runProperty({
      suite: 'bilig/workbook-runtime/snapshot-assembly-reordered-chunks',
      arbitrary: snapshotAssemblyCaseArbitrary,
      predicate: async ({ documentId, snapshotId, cursor, bytes, chunkSize, order }) => {
        const publication = createSnapshotPublicationFromBytes({
          documentId,
          snapshotId,
          cursor,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
        })
        const frames = createSnapshotChunkFrames({
          documentId,
          snapshotId,
          cursor,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
          chunkSize,
        })
        const orderedFrames = reorderFrames(frames, order)
        const registry: SnapshotAssemblyRegistry = new Map()
        let completed: CompletedSnapshotAssembly | null = null

        for (let index = 0; index < orderedFrames.length; index += 1) {
          completed = acceptSnapshotChunk(registry, orderedFrames[index], {
            nowUnixMs: 1_000 + index,
            maxChunks: 512,
            maxBytes: 512,
          })
          if (index < orderedFrames.length - 1) {
            expect(completed).toBeNull()
          }
        }

        expect(completed).toEqual({
          documentId,
          snapshotId,
          cursor,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
        })
        expect(registry.has(snapshotId)).toBe(false)
        expect(publication.bytes).toEqual(bytes)
      },
      parameters: { numRuns: 400 },
    })
  })

  it('rejects contradictory duplicate snapshot chunks and clears stale assemblies', async () => {
    await runProperty({
      suite: 'bilig/workbook-runtime/snapshot-assembly-corrupt-duplicate-rejection',
      arbitrary: fc.record({
        bytes: fc.uint8Array({ minLength: 2, maxLength: 64 }),
        cursor: fc.integer({ min: 0, max: 10_000 }),
      }),
      predicate: async ({ bytes, cursor }) => {
        const snapshotId = `snapshot-${cursor}`
        const frames = createSnapshotChunkFrames({
          documentId: 'doc-corrupt',
          snapshotId,
          cursor,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
          chunkSize: 1,
        })
        const registry: SnapshotAssemblyRegistry = new Map()
        const first = frames[0]
        expect(acceptSnapshotChunk(registry, first, { maxBytes: 512 })).toBeNull()
        expect(registry.has(snapshotId)).toBe(true)

        const corruptBytes = new Uint8Array(first.bytes)
        corruptBytes[0] ^= 0xff
        expect(acceptSnapshotChunk(registry, { ...first, bytes: corruptBytes }, { maxBytes: 512 })).toBeNull()
        expect(registry.has(snapshotId)).toBe(false)
      },
    })
  })

  it('decodes generated workbook snapshots after chunk publication and rejects content-type drift', async () => {
    await runProperty({
      suite: 'bilig/workbook-runtime/snapshot-publication-decode-guards',
      arbitrary: fc.record({
        snapshot: fuzzWorkbookSnapshotArbitrary,
        chunkSize: fc.integer({ min: 1, max: 48 }),
      }),
      predicate: async ({ snapshot, chunkSize }) => {
        const typedSnapshot = toWorkbookSnapshot(snapshot)
        const bytes = snapshotEncoder.encode(JSON.stringify(typedSnapshot))
        const publication = createSnapshotPublicationFromBytes({
          documentId: 'doc-decode',
          snapshotId: 'snapshot-decode',
          cursor: 7,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
        })
        const registry: SnapshotAssemblyRegistry = new Map()
        let completed: CompletedSnapshotAssembly | null = null
        for (const frame of createSnapshotChunkFrames({
          documentId: 'doc-decode',
          snapshotId: 'snapshot-decode',
          cursor: 7,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
          chunkSize,
        })) {
          completed = acceptSnapshotChunk(registry, frame, { maxBytes: Math.max(512, bytes.byteLength) })
        }

        const directAssembly: CompletedSnapshotAssembly = {
          documentId: 'doc-decode',
          snapshotId: 'snapshot-decode',
          cursor: 7,
          contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
          bytes,
        }
        expect(publication.bytes).toEqual(bytes)
        expect(completed?.bytes).toEqual(bytes)
        expect(completed ? decodeWorkbookSnapshotBytes(completed) : null).toEqual(typedSnapshot)
        expect(decodeWorkbookSnapshotBytes(directAssembly)).toEqual(typedSnapshot)
        expect(() => decodeWorkbookSnapshotBytes({ ...directAssembly, contentType: 'application/octet-stream' })).toThrow(
          /Unsupported snapshot/,
        )
      },
      parameters: { numRuns: 250 },
    })
  })

  it('normalizes generated browser session URLs without losing document or server identity', async () => {
    await runProperty({
      suite: 'bilig/workbook-runtime/browser-url-session-params',
      arbitrary: fc.record({
        baseUrl: fc.constantFrom('http://localhost:3000', 'http://localhost:3000/', 'https://app.example.test/workbook/'),
        serverUrl: fc.constantFrom('http://127.0.0.1:4173', 'https://server.example.test/sync?token=abc'),
        documentId: fc.string({ minLength: 1, maxLength: 32 }).filter((value) => !containsControlCharacter(value)),
      }),
      predicate: async ({ baseUrl, serverUrl, documentId }) => {
        const built = buildBrowserUrl(baseUrl, serverUrl, documentId)
        expect(built).toBeDefined()
        const url = new URL(built ?? '')
        expect(url.searchParams.get('document')).toBe(documentId)
        expect(url.searchParams.get('server')).toBe(serverUrl)
        expect(buildBrowserUrl(undefined, serverUrl, documentId)).toBeUndefined()
      },
    })
  })
})

const snapshotAssemblyCaseArbitrary = fc.record({
  documentId: fc.constantFrom('doc-a', 'doc-b', 'doc-runtime'),
  snapshotId: fc.uuid(),
  cursor: fc.integer({ min: 0, max: 10_000 }),
  bytes: fc.uint8Array({ minLength: 0, maxLength: 320 }),
  chunkSize: fc.integer({ min: 1, max: 48 }),
  order: fc.constantFrom('forward', 'reverse', 'rotate'),
})

function reorderFrames<T>(frames: readonly T[], order: 'forward' | 'reverse' | 'rotate'): T[] {
  if (order === 'reverse') {
    return frames.toReversed()
  }
  if (order === 'rotate' && frames.length > 1) {
    return [...frames.slice(1), frames[0]]
  }
  return [...frames]
}

function toWorkbookSnapshot(snapshot: FuzzWorkbookSnapshot): WorkbookSnapshot {
  return {
    version: snapshot.version,
    workbook: {
      name: snapshot.workbook.name,
    },
    sheets: snapshot.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      order: sheet.order,
      cells: sheet.cells.map((cell) => ({
        address: cell.address,
        row: cell.row,
        col: cell.col,
        value: cell.value,
        formula: cell.formula,
        format: cell.format,
      })),
    })),
  }
}

function containsControlCharacter(value: string): boolean {
  for (const character of value) {
    if (character.charCodeAt(0) <= 0x1f) {
      return true
    }
  }
  return false
}
