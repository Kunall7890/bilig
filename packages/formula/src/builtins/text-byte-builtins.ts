import { ErrorCode, type CellValue } from '@bilig/protocol'
import type { TextBuiltin } from './text.js'

interface TextByteBuiltinDeps {
  error: (code: ErrorCode) => CellValue
  stringResult: (value: string) => CellValue
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined
  coerceText: (value: CellValue) => string
  coerceLength: (value: CellValue | undefined, defaultValue: number) => number | CellValue
  isErrorValue: (value: number | CellValue) => value is CellValue
}

const utf8Encoder = new TextEncoder()
const utf8Decoder = new TextDecoder()

export function utf8Bytes(value: string): Uint8Array {
  return utf8Encoder.encode(value)
}

function utf8Text(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes)
}

export function findSubBytes(haystack: Uint8Array, needle: Uint8Array, start: number): number {
  if (needle.length === 0) {
    return Math.max(0, Math.min(start, haystack.length))
  }

  for (let index = start; index + needle.length <= haystack.length; index += 1) {
    let match = true
    for (let offset = 0; offset < needle.length; offset += 1) {
      if (haystack[index + offset] !== needle[offset]) {
        match = false
        break
      }
    }
    if (match) {
      return index
    }
  }
  return -1
}

function leftBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text)
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length))
  return utf8Text(bytes.slice(0, normalizedCount))
}

function rightBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text)
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length))
  return utf8Text(bytes.slice(bytes.length - normalizedCount))
}

function midBytes(text: string, start: number, byteCount: number): string {
  const bytes = utf8Bytes(text)
  if (byteCount <= 0) {
    return ''
  }

  const zeroBasedStart = Math.max(0, start - 1)
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + byteCount)
  if (zeroBasedStart >= bytes.length) {
    return ''
  }
  return utf8Text(bytes.slice(zeroBasedStart, zeroBasedEnd))
}

function replaceBytes(text: string, start: number, byteCount: number, replacement: string): string {
  const bytes = utf8Bytes(text)
  const replacementBytes = utf8Bytes(replacement)
  const zeroBasedStart = Math.max(0, start - 1)
  if (zeroBasedStart >= bytes.length) {
    return text
  }
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + Math.max(0, byteCount))
  return utf8Text(new Uint8Array([...bytes.slice(0, zeroBasedStart), ...replacementBytes, ...bytes.slice(zeroBasedEnd)]))
}

export function bytePositionToCharPosition(text: string, startByte: number): number {
  if (startByte <= 1) {
    return 1
  }
  return utf8Text(utf8Bytes(text).slice(0, startByte - 1)).length + 1
}

export function charPositionToBytePosition(text: string, charPosition: number): number {
  return utf8Bytes(text.slice(0, Math.max(0, charPosition - 1))).length + 1
}

export function createTextByteBuiltins(deps: TextByteBuiltinDeps): Record<string, TextBuiltin> {
  return {
    LEFTB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, countValue] = args
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const count = deps.coerceLength(countValue, 1)
      if (deps.isErrorValue(count)) {
        return count
      }
      return deps.stringResult(leftBytes(deps.coerceText(textValue), count))
    },
    MIDB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, startValue, countValue] = args
      if (textValue === undefined || startValue === undefined || countValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const start = deps.coerceLength(startValue, 1)
      if (deps.isErrorValue(start) || start < 1) {
        return deps.error(ErrorCode.Value)
      }
      const count = deps.coerceLength(countValue, 0)
      if (deps.isErrorValue(count)) {
        return count
      }
      return deps.stringResult(midBytes(deps.coerceText(textValue), start, count))
    },
    RIGHTB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, countValue] = args
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const count = deps.coerceLength(countValue, 1)
      if (deps.isErrorValue(count)) {
        return count
      }
      return deps.stringResult(rightBytes(deps.coerceText(textValue), count))
    },
    REPLACEB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, startValue, countValue, replacementValue] = args
      if (textValue === undefined || startValue === undefined || countValue === undefined || replacementValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const start = deps.coerceLength(startValue, 1)
      if (deps.isErrorValue(start) || start < 1) {
        return deps.error(ErrorCode.Value)
      }
      const count = deps.coerceLength(countValue, 0)
      if (deps.isErrorValue(count)) {
        return count
      }
      return deps.stringResult(replaceBytes(deps.coerceText(textValue), start, count, deps.coerceText(replacementValue)))
    },
  }
}
