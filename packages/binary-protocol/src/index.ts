import type { EngineOpBatch } from '@bilig/workbook'
import { BinaryProtocolError, BinaryReader, BinaryWriter } from './binary-io.js'
import { decodeBatch, encodeBatch } from './binary-engine-op-codec.js'
import { assertNever } from './binary-value-codec.js'

export { BinaryProtocolError, BinaryReader, BinaryWriter }

export const PROTOCOL_MAGIC = 0x424c4731
export const PROTOCOL_VERSION = 1
export const WORKBOOK_SNAPSHOT_CONTENT_TYPE = 'application/vnd.bilig.workbook+json'

export type FrameKind = 'hello' | 'appendBatch' | 'ack' | 'snapshotChunk' | 'cursorWatermark' | 'heartbeat' | 'error'

export interface HelloFrame {
  kind: 'hello'
  documentId: string
  replicaId: string
  sessionId: string
  protocolVersion: number
  lastServerCursor: number
  capabilities: string[]
}

export interface BatchAppendFrame {
  kind: 'appendBatch'
  documentId: string
  cursor: number
  batch: EngineOpBatch
}

export interface AckFrame {
  kind: 'ack'
  documentId: string
  batchId: string
  cursor: number
  acceptedAtUnixMs: number
}

export interface SnapshotChunkFrame {
  kind: 'snapshotChunk'
  documentId: string
  snapshotId: string
  cursor: number
  chunkIndex: number
  chunkCount: number
  contentType: string
  bytes: Uint8Array
}

export interface SnapshotChunkOptions {
  documentId: string
  snapshotId: string
  cursor: number
  contentType: string
  bytes: Uint8Array
  chunkSize?: number
}

export interface CursorWatermarkFrame {
  kind: 'cursorWatermark'
  documentId: string
  cursor: number
  compactedCursor: number
}

export interface HeartbeatFrame {
  kind: 'heartbeat'
  documentId: string
  cursor: number
  sentAtUnixMs: number
}

export interface ErrorFrame {
  kind: 'error'
  documentId: string
  code: string
  message: string
  retryable: boolean
}

export type ProtocolFrame =
  | HelloFrame
  | BatchAppendFrame
  | AckFrame
  | SnapshotChunkFrame
  | CursorWatermarkFrame
  | HeartbeatFrame
  | ErrorFrame

const FRAME_TAGS: Record<FrameKind, number> = {
  hello: 1,
  appendBatch: 2,
  ack: 3,
  snapshotChunk: 4,
  cursorWatermark: 5,
  heartbeat: 6,
  error: 7,
}

const FRAME_ENTRIES: ReadonlyArray<readonly [FrameKind, number]> = [
  ['hello', 1],
  ['appendBatch', 2],
  ['ack', 3],
  ['snapshotChunk', 4],
  ['cursorWatermark', 5],
  ['heartbeat', 6],
  ['error', 7],
]

const FRAME_BY_TAG = new Map<number, FrameKind>(FRAME_ENTRIES.map(([kind, tag]) => [tag, kind]))

function encodePayload(frame: ProtocolFrame): Uint8Array {
  const writer = new BinaryWriter()
  switch (frame.kind) {
    case 'hello':
      writer.string(frame.documentId)
      writer.string(frame.replicaId)
      writer.string(frame.sessionId)
      writer.u32(frame.protocolVersion)
      writer.u32(frame.lastServerCursor)
      writer.stringArray(frame.capabilities)
      return writer.finish()
    case 'appendBatch':
      writer.string(frame.documentId)
      writer.u32(frame.cursor)
      encodeBatch(writer, frame.batch)
      return writer.finish()
    case 'ack':
      writer.string(frame.documentId)
      writer.string(frame.batchId)
      writer.u32(frame.cursor)
      writer.f64(frame.acceptedAtUnixMs)
      return writer.finish()
    case 'snapshotChunk':
      writer.string(frame.documentId)
      writer.string(frame.snapshotId)
      writer.u32(frame.cursor)
      writer.u32(frame.chunkIndex)
      writer.u32(frame.chunkCount)
      writer.string(frame.contentType)
      writer.bytes(frame.bytes)
      return writer.finish()
    case 'cursorWatermark':
      writer.string(frame.documentId)
      writer.u32(frame.cursor)
      writer.u32(frame.compactedCursor)
      return writer.finish()
    case 'heartbeat':
      writer.string(frame.documentId)
      writer.u32(frame.cursor)
      writer.f64(frame.sentAtUnixMs)
      return writer.finish()
    case 'error':
      writer.string(frame.documentId)
      writer.string(frame.code)
      writer.string(frame.message)
      writer.bool(frame.retryable)
      return writer.finish()
    default:
      assertNever(frame)
  }
}

function decodePayload(kind: FrameKind, payload: Uint8Array): ProtocolFrame {
  const reader = new BinaryReader(payload)
  switch (kind) {
    case 'hello':
      return {
        kind,
        documentId: reader.string(),
        replicaId: reader.string(),
        sessionId: reader.string(),
        protocolVersion: reader.u32(),
        lastServerCursor: reader.u32(),
        capabilities: reader.stringArray(),
      }
    case 'appendBatch':
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        batch: decodeBatch(reader),
      }
    case 'ack':
      return {
        kind,
        documentId: reader.string(),
        batchId: reader.string(),
        cursor: reader.u32(),
        acceptedAtUnixMs: reader.f64(),
      }
    case 'snapshotChunk':
      return {
        kind,
        documentId: reader.string(),
        snapshotId: reader.string(),
        cursor: reader.u32(),
        chunkIndex: reader.u32(),
        chunkCount: reader.u32(),
        contentType: reader.string(),
        bytes: reader.bytesView(),
      }
    case 'cursorWatermark':
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        compactedCursor: reader.u32(),
      }
    case 'heartbeat':
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        sentAtUnixMs: reader.f64(),
      }
    case 'error':
      return {
        kind,
        documentId: reader.string(),
        code: reader.string(),
        message: reader.string(),
        retryable: reader.bool(),
      }
    default:
      assertNever(kind)
  }
}

export function encodeFrame(frame: ProtocolFrame): Uint8Array {
  const payload = encodePayload(frame)
  const output = new Uint8Array(11 + payload.byteLength)
  const view = new DataView(output.buffer)
  view.setUint32(0, PROTOCOL_MAGIC, true)
  view.setUint16(4, PROTOCOL_VERSION, true)
  view.setUint8(6, FRAME_TAGS[frame.kind])
  view.setUint32(7, payload.byteLength, true)
  output.set(payload, 11)
  return output
}

export function decodeFrame(bytesLike: Uint8Array | ArrayBuffer): ProtocolFrame {
  const bytes = bytesLike instanceof Uint8Array ? bytesLike : new Uint8Array(bytesLike)
  if (bytes.byteLength < 11) {
    throw new BinaryProtocolError('Binary frame too short')
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const magic = view.getUint32(0, true)
  if (magic !== PROTOCOL_MAGIC) {
    throw new BinaryProtocolError('Binary frame magic mismatch')
  }

  const version = view.getUint16(4, true)
  if (version !== PROTOCOL_VERSION) {
    throw new BinaryProtocolError(`Unsupported protocol version ${version}`)
  }

  const kind = FRAME_BY_TAG.get(view.getUint8(6))
  if (!kind) {
    throw new BinaryProtocolError('Unknown frame tag')
  }

  const payloadLength = view.getUint32(7, true)
  if (bytes.byteLength !== 11 + payloadLength) {
    throw new BinaryProtocolError('Binary frame length mismatch')
  }

  return decodePayload(kind, bytes.subarray(11))
}

export function createSnapshotChunkFrames(options: SnapshotChunkOptions): SnapshotChunkFrame[] {
  const chunkSize = Math.max(1, options.chunkSize ?? 64 * 1024)
  const chunkCount = Math.max(1, Math.ceil(options.bytes.byteLength / chunkSize))
  const frames: SnapshotChunkFrame[] = []
  for (let chunkIndex = 0; chunkIndex < chunkCount; chunkIndex += 1) {
    const start = chunkIndex * chunkSize
    const end = Math.min(options.bytes.byteLength, start + chunkSize)
    frames.push({
      kind: 'snapshotChunk',
      documentId: options.documentId,
      snapshotId: options.snapshotId,
      cursor: options.cursor,
      chunkIndex,
      chunkCount,
      contentType: options.contentType,
      bytes: options.bytes.subarray(start, end),
    })
  }
  return frames
}
