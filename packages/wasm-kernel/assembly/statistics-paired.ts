import { ErrorCode, ValueTag } from './protocol'
import { inputCellNumeric, inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'

export let pairedSampleCount: i32 = 0
export let pairedSumX: f64 = 0
export let pairedSumY: f64 = 0
export let pairedSumXX: f64 = 0
export let pairedSumYY: f64 = 0
export let pairedSumXY: f64 = 0

function normalizeNearZero(value: f64): f64 {
  return Math.abs(value) < 1e-12 ? 0.0 : value
}

export function collectPairedNumericStats(
  ySlot: i32,
  xSlot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  const yRows = inputRowsFromSlot(ySlot, kindStack, rangeIndexStack, rangeRowCounts)
  const yCols = inputColsFromSlot(ySlot, kindStack, rangeIndexStack, rangeColCounts)
  const xRows = inputRowsFromSlot(xSlot, kindStack, rangeIndexStack, rangeRowCounts)
  const xCols = inputColsFromSlot(xSlot, kindStack, rangeIndexStack, rangeColCounts)
  if (yRows < 1 || yCols < 1 || xRows < 1 || xCols < 1) {
    return ErrorCode.Value
  }

  const yCount = yRows * yCols
  const xCount = xRows * xCols
  if (yCount != xCount || yCount <= 0) {
    return ErrorCode.Value
  }

  pairedSampleCount = yCount
  pairedSumX = 0
  pairedSumY = 0
  pairedSumXX = 0
  pairedSumYY = 0
  pairedSumXY = 0

  for (let offset = 0; offset < yCount; offset += 1) {
    const yRow = yCols == 0 ? 0 : <i32>Math.floor(<f64>offset / <f64>yCols)
    const yCol = yCols == 0 ? 0 : offset - yRow * yCols
    const xRow = xCols == 0 ? 0 : <i32>Math.floor(<f64>offset / <f64>xCols)
    const xCol = xCols == 0 ? 0 : offset - xRow * xCols

    const yTag = inputCellTag(
      ySlot,
      yRow,
      yCol,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
    )
    if (yTag == ValueTag.Error) {
      return <i32>(
        inputCellScalarValue(
          ySlot,
          yRow,
          yCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
      )
    }

    const xTag = inputCellTag(
      xSlot,
      xRow,
      xCol,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
    )
    if (xTag == ValueTag.Error) {
      return <i32>(
        inputCellScalarValue(
          xSlot,
          xRow,
          xCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
      )
    }

    const yNumeric = inputCellNumeric(
      ySlot,
      yRow,
      yCol,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
    )
    const xNumeric = inputCellNumeric(
      xSlot,
      xRow,
      xCol,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
    )
    if (isNaN(yNumeric) || isNaN(xNumeric)) {
      return ErrorCode.Value
    }

    pairedSumX += xNumeric
    pairedSumY += yNumeric
    pairedSumXX += xNumeric * xNumeric
    pairedSumYY += yNumeric * yNumeric
    pairedSumXY += xNumeric * yNumeric
  }

  return 0
}

export function pairedCenteredSumSquaresX(): f64 {
  return normalizeNearZero(pairedSumXX - (pairedSumX * pairedSumX) / <f64>pairedSampleCount)
}

export function pairedCenteredSumSquaresY(): f64 {
  return normalizeNearZero(pairedSumYY - (pairedSumY * pairedSumY) / <f64>pairedSampleCount)
}

export function pairedCenteredCrossProducts(): f64 {
  return normalizeNearZero(pairedSumXY - (pairedSumX * pairedSumY) / <f64>pairedSampleCount)
}
