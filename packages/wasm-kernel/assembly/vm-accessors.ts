import {
  constantArena,
  constantLengths,
  constantOffsets,
  errors,
  numbers,
  programLengths,
  programOffsets,
  rangeLengths,
  rangeMembers,
  rangeOffsets,
  stringIds,
  tags,
} from './vm'

export function getTagsPtr(): usize {
  return changetype<usize>(tags.dataStart)
}

export function getNumbersPtr(): usize {
  return changetype<usize>(numbers.dataStart)
}

export function getStringIdsPtr(): usize {
  return changetype<usize>(stringIds.dataStart)
}

export function getErrorsPtr(): usize {
  return changetype<usize>(errors.dataStart)
}

export function getProgramOffsetsPtr(): usize {
  return changetype<usize>(programOffsets.dataStart)
}

export function getProgramLengthsPtr(): usize {
  return changetype<usize>(programLengths.dataStart)
}

export function getConstantOffsetsPtr(): usize {
  return changetype<usize>(constantOffsets.dataStart)
}

export function getConstantLengthsPtr(): usize {
  return changetype<usize>(constantLengths.dataStart)
}

export function getConstantArenaPtr(): usize {
  return changetype<usize>(constantArena.dataStart)
}

export function getRangeOffsetsPtr(): usize {
  return changetype<usize>(rangeOffsets.dataStart)
}

export function getRangeLengthsPtr(): usize {
  return changetype<usize>(rangeLengths.dataStart)
}

export function getRangeMembersPtr(): usize {
  return changetype<usize>(rangeMembers.dataStart)
}

export function getCellCapacity(): i32 {
  return tags.length
}

export function getFormulaCapacity(): i32 {
  return programOffsets.length
}

export function getConstantCapacity(): i32 {
  return constantArena.length
}

export function getRangeCapacity(): i32 {
  return rangeOffsets.length
}

export function getMemberCapacity(): i32 {
  return rangeMembers.length
}
