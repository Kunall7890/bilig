import { ErrorCode, ValueTag } from './protocol'
import { errors, numbers, stringIds, tags } from './vm'

const ROW_PAIR_LEFT_PLUS_RIGHT: u8 = 1
const ROW_PAIR_LEFT_MINUS_RIGHT: u8 = 2
const ROW_PAIR_RIGHT_MINUS_LEFT: u8 = 3
const ROW_PAIR_LEFT_TIMES_RIGHT: u8 = 4

function evalRowPairCode(code: u8, leftValue: f64, rightValue: f64): f64 {
  if (code == ROW_PAIR_LEFT_PLUS_RIGHT) {
    return leftValue + rightValue
  }
  if (code == ROW_PAIR_LEFT_MINUS_RIGHT) {
    return leftValue - rightValue
  }
  if (code == ROW_PAIR_RIGHT_MINUS_LEFT) {
    return rightValue - leftValue
  }
  if (code == ROW_PAIR_LEFT_TIMES_RIGHT) {
    return leftValue * rightValue
  }
  return NaN
}

function writeNumber(targetCellIndex: i32, value: f64): void {
  tags[targetCellIndex] = <u8>ValueTag.Number
  numbers[targetCellIndex] = value
  stringIds[targetCellIndex] = 0
  errors[targetCellIndex] = ErrorCode.None
}

export function evalDenseDirectScalarRowChainStoreTargetBatch(
  leftValues: Float64Array,
  rightValues: Float64Array,
  firstTargets: Uint32Array,
  secondTargets: Uint32Array,
  rowCount: i32,
  firstFormulaCode: u8,
  secondFormulaScale: f64,
  secondFormulaOffset: f64,
): void {
  if (rowCount <= 0) {
    return
  }
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const first = evalRowPairCode(firstFormulaCode, leftValues[rowOffset], rightValues[rowOffset])
    writeNumber(<i32>firstTargets[rowOffset], first)
    writeNumber(<i32>secondTargets[rowOffset], first * secondFormulaScale + secondFormulaOffset)
  }
}
