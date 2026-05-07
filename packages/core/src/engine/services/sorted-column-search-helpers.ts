import { ValueTag, type CellValue } from '@bilig/protocol'
import { parseRangeAddress } from '@bilig/formula'
import type { PreparedApproximateVectorLookup } from '../runtime-state.js'

export interface ApproximateColumnIndexEntry {
  sheetName: string
  rowStart: number
  rowEnd: number
  col: number
  columnVersion: number
  structureVersion: number
  comparableKind: 'numeric' | 'text' | undefined
  uniformStart: number | undefined
  uniformStep: number | undefined
  repeatedUniformStart: number | undefined
  repeatedUniformStep: number | undefined
  repeatedUniformRunLength: number | undefined
  sortedAscending: boolean
  sortedDescending: boolean
  numericValues: Float64Array | undefined
  textValues: string[] | undefined
}

export type ApproximateComparable =
  | { kind: 'empty' }
  | { kind: 'numeric'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'invalid' }

type UniformNumericApproximateMatchResult =
  | { handled: false }
  | {
      handled: true
      position: number | undefined
    }

interface SingleColumnBoundsRequest {
  sheetName: string
  start: string
  end: string
  startRow?: number
  endRow?: number
  startCol?: number
  endCol?: number
}

export interface SingleColumnBounds {
  rowStart: number
  rowEnd: number
  col: number
}

export function getColumnCacheKey(sheetName: string, col: number, rowStart: number, rowEnd: number): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`
}

export function normalizeApproximateComparableValue(
  value: CellValue,
  lookupString: (id: number) => string,
  stringId = 0,
): ApproximateComparable {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'empty' }
    case ValueTag.Number:
      return { kind: 'numeric', value: Object.is(value.value, -0) ? 0 : value.value }
    case ValueTag.Boolean:
      return { kind: 'numeric', value: value.value ? 1 : 0 }
    case ValueTag.String:
      return {
        kind: 'text',
        value: (stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase(),
      }
    case ValueTag.Error:
      return { kind: 'invalid' }
  }
}

export function compareApproximateNumeric(left: number, right: number): number {
  if (left === right) {
    return 0
  }
  return left < right ? -1 : 1
}

export function compareApproximateText(left: string, right: string): number {
  if (left === right) {
    return 0
  }
  return left < right ? -1 : 1
}

export function findUniformNumericApproximatePosition(args: {
  lookupValue: number
  length: number
  matchMode: 1 | -1
  uniformStart: number | undefined
  uniformStep: number | undefined
}): UniformNumericApproximateMatchResult {
  if (args.uniformStart === undefined || args.uniformStep === undefined) {
    return { handled: false }
  }
  if (args.length <= 0) {
    return { handled: true, position: undefined }
  }
  const { lookupValue, length, matchMode, uniformStart, uniformStep } = args
  const lastValue = uniformStart + uniformStep * (length - 1)
  if (matchMode === 1 && uniformStep > 0) {
    if (lookupValue < uniformStart) {
      return { handled: true, position: undefined }
    }
    if (lookupValue >= lastValue) {
      return { handled: true, position: length }
    }
    const position = Math.floor((lookupValue - uniformStart) / uniformStep) + 1
    return { handled: true, position: Math.min(length, Math.max(1, position)) }
  }
  if (matchMode === -1 && uniformStep < 0) {
    if (lookupValue > uniformStart) {
      return { handled: true, position: undefined }
    }
    if (lookupValue <= lastValue) {
      return { handled: true, position: length }
    }
    const position = Math.floor((uniformStart - lookupValue) / -uniformStep) + 1
    return { handled: true, position: Math.min(length, Math.max(1, position)) }
  }
  return { handled: false }
}

export function findRepeatedUniformNumericApproximatePosition(args: {
  lookupValue: number
  length: number
  matchMode: 1 | -1
  repeatedUniformStart: number | undefined
  repeatedUniformStep: number | undefined
  repeatedUniformRunLength: number | undefined
}): UniformNumericApproximateMatchResult {
  if (args.repeatedUniformStart === undefined || args.repeatedUniformStep === undefined || args.repeatedUniformRunLength === undefined) {
    return { handled: false }
  }
  if (args.length <= 0) {
    return { handled: true, position: undefined }
  }
  const { lookupValue, length, matchMode, repeatedUniformStart, repeatedUniformStep, repeatedUniformRunLength } = args
  const groupCount = Math.ceil(length / repeatedUniformRunLength)
  const lastValue = repeatedUniformStart + repeatedUniformStep * (groupCount - 1)
  if (matchMode === 1 && repeatedUniformStep > 0) {
    if (lookupValue < repeatedUniformStart) {
      return { handled: true, position: undefined }
    }
    if (lookupValue >= lastValue) {
      return { handled: true, position: length }
    }
    const group = Math.floor((lookupValue - repeatedUniformStart) / repeatedUniformStep)
    return { handled: true, position: Math.min(length, (group + 1) * repeatedUniformRunLength) }
  }
  if (matchMode === -1 && repeatedUniformStep < 0) {
    if (lookupValue > repeatedUniformStart) {
      return { handled: true, position: undefined }
    }
    if (lookupValue <= lastValue) {
      return { handled: true, position: length }
    }
    const group = Math.floor((repeatedUniformStart - lookupValue) / -repeatedUniformStep)
    return { handled: true, position: Math.min(length, (group + 1) * repeatedUniformRunLength) }
  }
  return { handled: false }
}

export function supportsPreparedApproximateKind(
  prepared: PreparedApproximateVectorLookup,
  kind: 'numeric' | 'text',
  matchMode: 1 | -1,
): boolean {
  return prepared.comparableKind === kind && (matchMode === 1 ? prepared.sortedAscending : prepared.sortedDescending)
}

export function detectUniformNumericStep(values: Float64Array): { start: number; step: number } | undefined {
  if (values.length < 2) {
    return undefined
  }
  const start = values[0]!
  const step = values[1]! - start
  if (!Number.isFinite(step) || step === 0) {
    return undefined
  }
  for (let index = 2; index < values.length; index += 1) {
    if (values[index]! - values[index - 1]! !== step) {
      return undefined
    }
  }
  return { start, step }
}

export function detectRepeatedUniformNumericStep(values: Float64Array): { start: number; step: number; runLength: number } | undefined {
  if (values.length < 4) {
    return undefined
  }
  const start = values[0]!
  let runLength = 1
  while (runLength < values.length && values[runLength] === start) {
    runLength += 1
  }
  if (runLength <= 1 || runLength >= values.length) {
    return undefined
  }
  const step = values[runLength]! - start
  if (!Number.isFinite(step) || step === 0) {
    return undefined
  }
  for (let index = 0; index < values.length; index += 1) {
    const group = Math.floor(index / runLength)
    if (values[index]! !== start + step * group) {
      return undefined
    }
  }
  return { start, step, runLength }
}

export function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

export function resolveSingleColumnBounds(request: SingleColumnBoundsRequest): SingleColumnBounds | undefined {
  if (request.startRow !== undefined && request.endRow !== undefined && request.startCol !== undefined && request.endCol !== undefined) {
    if (request.startCol !== request.endCol) {
      return undefined
    }
    return {
      rowStart: request.startRow,
      rowEnd: request.endRow,
      col: request.startCol,
    }
  }

  const parsedRange = parseRangeAddress(`${request.start}:${request.end}`, request.sheetName)
  if (parsedRange.kind !== 'cells' || parsedRange.start.col !== parsedRange.end.col) {
    return undefined
  }
  return {
    rowStart: parsedRange.start.row,
    rowEnd: parsedRange.end.row,
    col: parsedRange.start.col,
  }
}

export function columnRegistryKey(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`
}
