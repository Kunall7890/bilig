import type {
  FormulaFamilyKey,
  FormulaFamilyMember,
  FormulaFamilyMembership,
  FormulaFamilyRunAxis,
  FormulaFamilyRunId,
} from './formula-family-store.js'

export interface MutableFormulaFamilyMemberRun {
  id: FormulaFamilyRunId
  axis: FormulaFamilyRunAxis
  fixedIndex: number
  start: number
  end: number
  step: number
  cellIndices: number[]
}

export interface MutableFormulaFamily {
  id: number
  sheetId: number
  templateId: number
  shapeKey: string
  key: string
  runs: MutableFormulaFamilyMemberRun[]
  rowRunsByFixedIndex: Map<number, MutableFormulaFamilyMemberRun[]>
  columnRunsByFixedIndex: Map<number, MutableFormulaFamilyMemberRun[]>
  singletonRunsByRow: Map<number, MutableFormulaFamilyMemberRun[]>
  rowAppendRunByFixedIndex: Array<MutableFormulaFamilyMemberRun | undefined>
  recentAppendRun: MutableFormulaFamilyMemberRun | undefined
}

export interface FormulaFamilyCellRecord extends FormulaFamilyKey, FormulaFamilyMember {}

export interface FormulaFamilyRunDescriptor {
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly end: number
  readonly step: number
  readonly members: readonly FormulaFamilyMember[]
}

export type FormulaRunFastPath =
  | {
      readonly kind: 'append'
      readonly run: MutableFormulaFamilyMemberRun
    }
  | {
      readonly kind: 'create'
    }

export function indexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  appendMapArray(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.axis === 'row') {
    family.rowAppendRunByFixedIndex[run.fixedIndex] = run
  }
  if (run.cellIndices.length === 1) {
    appendMapArray(family.singletonRunsByRow, runRowStart(run), run)
  }
}

export function unindexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  removeMapArrayValue(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.axis === 'row' && family.rowAppendRunByFixedIndex[run.fixedIndex] === run) {
    family.rowAppendRunByFixedIndex[run.fixedIndex] = undefined
  }
  if (run.cellIndices.length === 1) {
    removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
  }
  if (family.recentAppendRun === run) {
    family.recentAppendRun = undefined
  }
}

export function hasDuplicateRunMemberCellIndex(members: readonly FormulaFamilyMember[]): boolean {
  return new Set(members.map((member) => member.cellIndex)).size !== members.length
}

export function describeFreshOrderedUniformRun(
  members: readonly FormulaFamilyMember[],
  membershipFamilyIds: readonly number[],
): Omit<FormulaFamilyRunDescriptor, 'members'> | undefined {
  const first = members[0]
  if (!first || (membershipFamilyIds[first.cellIndex] ?? 0) !== 0) {
    return undefined
  }
  if (members.length === 1) {
    return {
      axis: 'row',
      fixedIndex: first.col,
      start: first.row,
      end: first.row,
      step: 1,
    }
  }
  const second = members[1]!
  if ((membershipFamilyIds[second.cellIndex] ?? 0) !== 0 || second.cellIndex <= first.cellIndex) {
    return undefined
  }
  let axis: FormulaFamilyRunAxis
  let fixedIndex: number
  let previousIndex: number
  let step: number
  if (second.col === first.col && second.row > first.row) {
    axis = 'row'
    fixedIndex = first.col
    previousIndex = first.row
    step = second.row - first.row
  } else if (second.row === first.row && second.col > first.col) {
    axis = 'column'
    fixedIndex = first.row
    previousIndex = first.col
    step = second.col - first.col
  } else {
    return undefined
  }
  for (let index = 1; index < members.length; index += 1) {
    const member = members[index]!
    if ((membershipFamilyIds[member.cellIndex] ?? 0) !== 0 || member.cellIndex <= members[index - 1]!.cellIndex) {
      return undefined
    }
    const memberIndex = axis === 'row' ? member.row : member.col
    if (memberIndex !== previousIndex + step || (axis === 'row' ? member.col !== fixedIndex : member.row !== fixedIndex)) {
      return undefined
    }
    previousIndex = memberIndex
  }
  return {
    axis,
    fixedIndex,
    start: axis === 'row' ? first.row : first.col,
    end: previousIndex,
    step,
  }
}

export function describeFreshFormulaRunMemberSegmentsInOrder(
  members: readonly FormulaFamilyMember[],
  membershipFamilyIds: readonly number[],
): FormulaFamilyRunDescriptor[] | undefined {
  const first = members[0]
  if (!first || (membershipFamilyIds[first.cellIndex] ?? 0) !== 0) {
    return undefined
  }
  const second = members[1]!
  if (!second) {
    return undefined
  }
  if ((membershipFamilyIds[second.cellIndex] ?? 0) !== 0 || second.cellIndex <= first.cellIndex) {
    return undefined
  }
  let axis: FormulaFamilyRunAxis
  let fixedIndex: number
  let previousIndex: number
  let step: number
  if (second.col === first.col && second.row > first.row) {
    axis = 'row'
    fixedIndex = first.col
    previousIndex = first.row
    step = second.row - first.row
  } else if (second.row === first.row && second.col > first.col) {
    axis = 'column'
    fixedIndex = first.row
    previousIndex = first.col
    step = second.col - first.col
  } else {
    return undefined
  }
  const indices = [previousIndex]
  for (let index = 1; index < members.length; index += 1) {
    const member = members[index]!
    if ((membershipFamilyIds[member.cellIndex] ?? 0) !== 0) {
      return undefined
    }
    if (index > 0 && member.cellIndex <= members[index - 1]!.cellIndex) {
      return undefined
    }
    if (axis === 'row') {
      if (member.col !== fixedIndex || member.row <= previousIndex) {
        return undefined
      }
      previousIndex = member.row
    } else {
      if (member.row !== fixedIndex || member.col <= previousIndex) {
        return undefined
      }
      previousIndex = member.col
    }
    indices.push(previousIndex)
  }
  const lastIndex = indices[indices.length - 1]!
  if (isUniformIndexRun(indices, step)) {
    return [
      {
        axis,
        fixedIndex,
        start: axis === 'row' ? first.row : first.col,
        end: lastIndex,
        step,
        members,
      },
    ]
  }
  for (let cycleLength = 2; cycleLength <= Math.min(8, Math.floor(members.length / 2)); cycleLength += 1) {
    const descriptors = describeDeinterleavedFormulaRunSegments(axis, fixedIndex, members, cycleLength)
    if (descriptors) {
      return descriptors
    }
  }
  return undefined
}

function describeDeinterleavedFormulaRunSegments(
  axis: FormulaFamilyRunAxis,
  fixedIndex: number,
  members: readonly FormulaFamilyMember[],
  cycleLength: number,
): FormulaFamilyRunDescriptor[] | undefined {
  const descriptors: FormulaFamilyRunDescriptor[] = []
  for (let offset = 0; offset < cycleLength; offset += 1) {
    const segment: FormulaFamilyMember[] = []
    for (let index = offset; index < members.length; index += cycleLength) {
      segment.push(members[index]!)
    }
    const indices = segment.map((member) => (axis === 'row' ? member.row : member.col))
    const step = indices[1]! - indices[0]!
    if (step <= 0 || !isUniformIndexRun(indices, step)) {
      return undefined
    }
    descriptors.push({
      axis,
      fixedIndex,
      start: indices[0]!,
      end: indices[indices.length - 1]!,
      step,
      members: segment,
    })
  }
  return descriptors
}

function isUniformIndexRun(indices: readonly number[], step: number): boolean {
  for (let index = 1; index < indices.length; index += 1) {
    if (indices[index]! !== indices[index - 1]! + step) {
      return false
    }
  }
  return true
}

export function describeFormulaRunMembers(members: readonly FormulaFamilyMember[]): FormulaFamilyRunDescriptor | undefined {
  const first = members[0]
  if (!first) {
    return undefined
  }
  if (members.length === 1) {
    return {
      axis: 'row',
      fixedIndex: first.col,
      start: first.row,
      end: first.row,
      step: 1,
      members: [first],
    }
  }
  const isRowRun = members.every((member) => member.col === first.col)
  const isColumnRun = members.every((member) => member.row === first.row)
  if (isRowRun === isColumnRun) {
    return undefined
  }
  const axis: FormulaFamilyRunAxis = isRowRun ? 'row' : 'column'
  const sorted = sortRunMembers(axis, members)
  const firstIndex = axis === 'row' ? sorted[0]!.row : sorted[0]!.col
  const secondIndex = axis === 'row' ? sorted[1]!.row : sorted[1]!.col
  const step = secondIndex - firstIndex
  if (step <= 0) {
    return undefined
  }
  for (let index = 2; index < sorted.length; index += 1) {
    const previous = sorted[index - 1]!
    const current = sorted[index]!
    const previousIndex = axis === 'row' ? previous.row : previous.col
    const currentIndex = axis === 'row' ? current.row : current.col
    if (currentIndex - previousIndex !== step) {
      return undefined
    }
  }
  return {
    axis,
    fixedIndex: axis === 'row' ? first.col : first.row,
    start: firstIndex,
    end: axis === 'row' ? sorted[sorted.length - 1]!.row : sorted[sorted.length - 1]!.col,
    step,
    members: sorted,
  }
}

export function getFormulaRunFastPath(
  family: MutableFormulaFamily,
  descriptor: FormulaFamilyRunDescriptor,
): FormulaRunFastPath | undefined {
  const targetRuns =
    descriptor.axis === 'row'
      ? family.rowRunsByFixedIndex.get(descriptor.fixedIndex)
      : family.columnRunsByFixedIndex.get(descriptor.fixedIndex)
  let appendRun: MutableFormulaFamilyMemberRun | undefined
  let appendRunCount = 0
  targetRuns?.forEach((run) => {
    if (!canAppendFormulaRunDescriptor(run, descriptor)) {
      return
    }
    appendRun = run
    appendRunCount += 1
  })
  if (appendRunCount > 1) {
    return undefined
  }
  for (const member of descriptor.members) {
    for (const candidate of candidateRunsForMember(family, member)) {
      if (candidate !== appendRun) {
        return undefined
      }
    }
  }
  return appendRun ? { kind: 'append', run: appendRun } : { kind: 'create' }
}

function canAppendFormulaRunDescriptor(run: MutableFormulaFamilyMemberRun, descriptor: FormulaFamilyRunDescriptor): boolean {
  return (
    run.axis === descriptor.axis &&
    run.fixedIndex === descriptor.fixedIndex &&
    run.step === descriptor.step &&
    (descriptor.start === run.end + run.step || descriptor.end === run.start - run.step)
  )
}

export function appendFormulaRunMembers(
  family: MutableFormulaFamily,
  run: MutableFormulaFamilyMemberRun,
  descriptor: FormulaFamilyRunDescriptor,
): void {
  if (run.cellIndices.length === 1) {
    removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
  }
  const cellIndices = descriptor.members.map((member) => member.cellIndex)
  if (descriptor.end < run.start) {
    run.start = descriptor.start
    run.cellIndices.unshift(...cellIndices)
  } else {
    run.end = descriptor.end
    run.cellIndices.push(...cellIndices)
  }
  family.recentAppendRun = run
}

export function canRegisterFreshOrderedRunWithoutMerging(
  family: MutableFormulaFamily,
  descriptor: Omit<FormulaFamilyRunDescriptor, 'members'>,
  members: readonly FormulaFamilyMember[],
): boolean {
  const targetRuns =
    descriptor.axis === 'row'
      ? family.rowRunsByFixedIndex.get(descriptor.fixedIndex)
      : family.columnRunsByFixedIndex.get(descriptor.fixedIndex)
  if (targetRuns && targetRuns.length > 0) {
    return false
  }
  return members.every((member) => candidateRunsForMember(family, member).length === 0)
}

export function runRowStart(run: MutableFormulaFamilyMemberRun): number {
  return run.axis === 'row' ? run.start : run.fixedIndex
}

function appendMapArray<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const existing = map.get(key)
  if (existing) {
    existing.push(value)
    return
  }
  map.set(key, [value])
}

export function removeMapArrayValue<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const existing = map.get(key)
  if (!existing) {
    return
  }
  const index = existing.indexOf(value)
  if (index >= 0) {
    existing.splice(index, 1)
  }
  if (existing.length === 0) {
    map.delete(key)
  }
}

export function candidateRunsForMember(family: MutableFormulaFamily, member: FormulaFamilyMember): MutableFormulaFamilyMemberRun[] {
  const candidates: MutableFormulaFamilyMemberRun[] = []
  const seen = new Set<FormulaFamilyRunId>()
  const appendCandidates = (runs: readonly MutableFormulaFamilyMemberRun[] | undefined): void => {
    runs?.forEach((run) => {
      if (seen.has(run.id)) {
        return
      }
      seen.add(run.id)
      candidates.push(run)
    })
  }
  appendCandidates(family.rowRunsByFixedIndex.get(member.col))
  appendCandidates(family.columnRunsByFixedIndex.get(member.row))
  appendCandidates(family.singletonRunsByRow.get(member.row))
  return candidates
}

export function tryMergeRun(
  family: MutableFormulaFamily,
  runIndex: number,
  run: MutableFormulaFamilyMemberRun,
  member: FormulaFamilyMember,
  appendMemberToRun: (
    family: MutableFormulaFamily,
    run: MutableFormulaFamilyMemberRun,
    member: FormulaFamilyMember,
  ) => FormulaFamilyMembership,
  replaceRunWithMembers: (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ) => FormulaFamilyMembership,
  getCellRecord: (cellIndex: number) => FormulaFamilyCellRecord | undefined,
): FormulaFamilyMembership | undefined {
  if (run.axis === 'row' && run.fixedIndex === member.col) {
    if (canAppendStridedRunMember(run, member.row)) {
      return appendMemberToRun(family, run, member)
    }
    const membership = tryReshapeStridedRun(family, runIndex, run, member, replaceRunWithMembers, getCellRecord)
    if (membership) {
      return membership
    }
  }
  if (run.axis === 'column' && run.fixedIndex === member.row) {
    if (canAppendStridedRunMember(run, member.col)) {
      return appendMemberToRun(family, run, member)
    }
    const membership = tryReshapeStridedRun(family, runIndex, run, member, replaceRunWithMembers, getCellRecord)
    if (membership) {
      return membership
    }
  }
  if (run.cellIndices.length === 1) {
    const existingRecord = getCellRecord(run.cellIndices[0]!)
    if (!existingRecord) {
      return undefined
    }
    const existing = {
      cellIndex: existingRecord.cellIndex,
      row: existingRecord.row,
      col: existingRecord.col,
    }
    if (existing.col === member.col && existing.row !== member.row) {
      return replaceRunWithMembers(family, runIndex, 'row', member.col, [existing, member], member.cellIndex)
    }
    if (existing.row === member.row && existing.col !== member.col) {
      return replaceRunWithMembers(family, runIndex, 'column', member.row, [existing, member], member.cellIndex)
    }
  }
  return undefined
}

export function inferRunStep(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): number {
  if (members.length < 2) {
    return 1
  }
  const first = members[0]!
  const second = members[1]!
  return Math.max(1, axis === 'row' ? second.row - first.row : second.col - first.col)
}

export function canAppendStridedRunMember(run: MutableFormulaFamilyMemberRun, memberIndex: number): boolean {
  return memberIndex === run.start - run.step || memberIndex === run.end + run.step
}

function tryReshapeStridedRun(
  family: MutableFormulaFamily,
  runIndex: number,
  run: MutableFormulaFamilyMemberRun,
  member: FormulaFamilyMember,
  replaceRunWithMembers: (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ) => FormulaFamilyMembership,
  getCellRecord: (cellIndex: number) => FormulaFamilyCellRecord | undefined,
): FormulaFamilyMembership | undefined {
  const memberIndex = run.axis === 'row' ? member.row : member.col
  if (run.cellIndices.length !== 2 || memberIndex <= run.start || memberIndex >= run.end) {
    return undefined
  }
  const existingMembers = run.cellIndices.flatMap((cellIndex): FormulaFamilyMember[] => {
    const record = getCellRecord(cellIndex)
    return record ? [{ cellIndex, row: record.row, col: record.col }] : []
  })
  if (existingMembers.length !== run.cellIndices.length) {
    return undefined
  }
  const members = [...existingMembers, member]
  if (!isUniformRun(run.axis, members)) {
    return undefined
  }
  return replaceRunWithMembers(family, runIndex, run.axis, run.fixedIndex, members, member.cellIndex)
}

function isUniformRun(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): boolean {
  const sorted = sortRunMembers(axis, members)
  if (sorted.length < 2) {
    return true
  }
  const first = sorted[0]!
  const second = sorted[1]!
  const step = axis === 'row' ? second.row - first.row : second.col - first.col
  if (step <= 0) {
    return false
  }
  for (let index = 2; index < sorted.length; index += 1) {
    const previous = sorted[index - 1]!
    const current = sorted[index]!
    const delta = axis === 'row' ? current.row - previous.row : current.col - previous.col
    if (delta !== step) {
      return false
    }
  }
  return true
}

export function sortRunMembers(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): FormulaFamilyMember[] {
  return [...members].toSorted((left, right) =>
    axis === 'row' ? left.row - right.row || left.cellIndex - right.cellIndex : left.col - right.col || left.cellIndex - right.cellIndex,
  )
}

export function groupRunMembersByStep(
  axis: FormulaFamilyRunAxis,
  members: readonly FormulaFamilyMember[],
  step: number,
): FormulaFamilyMember[][] {
  const sorted = sortRunMembers(axis, members)
  const groups: FormulaFamilyMember[][] = []
  for (const member of sorted) {
    const current = groups[groups.length - 1]
    const previous = current?.[current.length - 1]
    const memberIndex = axis === 'row' ? member.row : member.col
    const previousIndex = previous ? (axis === 'row' ? previous.row : previous.col) : undefined
    if (!current || previousIndex === undefined || memberIndex !== previousIndex + step) {
      groups.push([member])
      continue
    }
    current.push(member)
  }
  return groups
}
