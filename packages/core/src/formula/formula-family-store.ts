import type { StructuralAxisTransform } from '@bilig/formula'
import {
  appendFormulaRunMembers,
  canAppendStridedRunMember,
  canRegisterFreshOrderedRunWithoutMerging,
  candidateRunsForMember,
  describeFormulaRunMembers,
  describeFreshFormulaRunMemberSegmentsInOrder,
  describeFreshOrderedUniformRun,
  getFormulaRunFastPath,
  groupRunMembersByStep,
  hasDuplicateRunMemberCellIndex,
  indexRun,
  inferRunStep,
  removeMapArrayValue,
  runRowStart,
  sortRunMembers,
  tryMergeRun,
  unindexRun,
  type FormulaFamilyCellRecord,
  type FormulaFamilyRunDescriptor,
  type FormulaRunFastPath,
  type MutableFormulaFamily,
  type MutableFormulaFamilyMemberRun,
} from './formula-family-store-runs.js'

export type FormulaFamilyId = number
export type FormulaFamilyRunId = number

export type FormulaFamilyRunAxis = 'row' | 'column'

export interface FormulaFamilyKey {
  readonly sheetId: number
  readonly templateId: number
  readonly shapeKey: string
}

export interface FormulaFamilyMember {
  readonly cellIndex: number
  readonly row: number
  readonly col: number
}

export interface FormulaFamilyRunUpsertArgs extends FormulaFamilyKey {
  readonly members: readonly FormulaFamilyMember[]
}

export interface FormulaFamilyFreshUniformRunRegistrationArgs extends FormulaFamilyKey {
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly step: number
  readonly cellIndices: readonly number[] | Uint32Array
}

export interface FormulaFamilyMemberRun {
  readonly id: FormulaFamilyRunId
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly end: number
  readonly step: number
  readonly cellIndices: readonly number[]
}

export interface FormulaFamily {
  readonly id: FormulaFamilyId
  readonly sheetId: number
  readonly templateId: number
  readonly shapeKey: string
  readonly runs: readonly FormulaFamilyMemberRun[]
}

export interface FormulaFamilyMembership {
  readonly familyId: FormulaFamilyId
  readonly runId: FormulaFamilyRunId
}

export interface FormulaFamilyStats {
  readonly familyCount: number
  readonly runCount: number
  readonly memberCount: number
}

export interface FormulaFamilyStructuralSourceTransform {
  readonly ownerSheetName: string
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
  readonly preservesValue: boolean
}

export interface FormulaFamilyStructuralSourceTransformEntry {
  readonly cellIndices: readonly number[]
  readonly transform: FormulaFamilyStructuralSourceTransform
}

export function composeFormulaFamilyStructuralSourceTransform(
  existing: FormulaFamilyStructuralSourceTransform,
  next: FormulaFamilyStructuralSourceTransform,
): FormulaFamilyStructuralSourceTransform | undefined {
  if (
    existing.ownerSheetName !== next.ownerSheetName ||
    existing.targetSheetName !== next.targetSheetName ||
    !existing.preservesValue ||
    !next.preservesValue ||
    existing.transform.kind !== 'insert' ||
    next.transform.kind !== 'insert' ||
    existing.transform.axis !== 'column' ||
    next.transform.axis !== 'column' ||
    next.transform.start < existing.transform.start ||
    next.transform.start > existing.transform.start + existing.transform.count
  ) {
    return undefined
  }
  return {
    ownerSheetName: existing.ownerSheetName,
    targetSheetName: existing.targetSheetName,
    preservesValue: true,
    transform: {
      kind: 'insert',
      axis: 'column',
      start: existing.transform.start,
      count: existing.transform.count + next.transform.count,
    },
  }
}

export interface FormulaFamilyStore {
  readonly upsertFormula: (args: FormulaFamilyKey & FormulaFamilyMember) => FormulaFamilyMembership
  readonly upsertFormulaRun: (args: FormulaFamilyRunUpsertArgs) => FormulaFamilyMembership[]
  readonly registerFormulaRun: (args: FormulaFamilyRunUpsertArgs) => void
  readonly registerFreshUniformRun: (args: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly unregisterFormula: (cellIndex: number) => boolean
  readonly getMembership: (cellIndex: number) => FormulaFamilyMembership | undefined
  readonly countSheetMembers: (sheetId: number) => number
  readonly forEachFamily: (fn: (family: FormulaFamily) => void) => void
  readonly setStructuralSourceTransform: (familyId: FormulaFamilyId, transform: FormulaFamilyStructuralSourceTransform) => void
  readonly getStructuralSourceTransform: (cellIndex: number) => FormulaFamilyStructuralSourceTransform | undefined
  readonly hasStructuralSourceTransforms: () => boolean
  readonly peekStructuralSourceTransforms: () => FormulaFamilyStructuralSourceTransformEntry[]
  readonly consumeStructuralSourceTransforms: () => FormulaFamilyStructuralSourceTransformEntry[]
  readonly getStats: () => FormulaFamilyStats
  readonly listFamilies: () => FormulaFamily[]
  readonly invalidateSheet: (sheetId: number) => void
  readonly applyStructuralInvalidation: (args: {
    readonly sheetId: number
    readonly axis: 'row' | 'column'
    readonly start: number
    readonly end: number
  }) => void
  readonly clear: () => void
}

function keyForFormulaFamily(args: FormulaFamilyKey): string {
  return `${args.sheetId}\t${args.templateId}\t${args.shapeKey}`
}

export function createFormulaFamilyStore(): FormulaFamilyStore {
  const familiesById = new Map<FormulaFamilyId, MutableFormulaFamily>()
  const familyIdByKey = new Map<string, FormulaFamilyId>()
  const recentFamilyByTemplateId = new Map<number, MutableFormulaFamily>()
  const cellRecordSheetIds: Array<number | undefined> = []
  const cellRecordTemplateIds: Array<number | undefined> = []
  const cellRecordShapeKeys: Array<string | undefined> = []
  const cellRecordRows: Array<number | undefined> = []
  const cellRecordCols: Array<number | undefined> = []
  const membershipFamilyIds: number[] = []
  const membershipRunIds: number[] = []
  const sheetMemberCounts = new Map<number, number>()
  const structuralSourceTransforms = new Map<FormulaFamilyId, FormulaFamilyStructuralSourceTransform>()
  const noMemberships: FormulaFamilyMembership[] = []
  let memberCount = 0
  let nextFamilyId = 1
  let nextRunId = 1

  const setMembership = (cellIndex: number, familyId: FormulaFamilyId, runId: FormulaFamilyRunId): FormulaFamilyMembership => {
    membershipFamilyIds[cellIndex] = familyId
    membershipRunIds[cellIndex] = runId
    return { familyId, runId }
  }

  const getMembershipRecord = (cellIndex: number): FormulaFamilyMembership | undefined => {
    const familyId = membershipFamilyIds[cellIndex] ?? 0
    if (familyId === 0) {
      return undefined
    }
    return { familyId, runId: membershipRunIds[cellIndex]! }
  }

  const getCellRecord = (cellIndex: number): FormulaFamilyCellRecord | undefined => {
    const sheetId = cellRecordSheetIds[cellIndex]
    const templateId = cellRecordTemplateIds[cellIndex]
    const shapeKey = cellRecordShapeKeys[cellIndex]
    const row = cellRecordRows[cellIndex]
    const col = cellRecordCols[cellIndex]
    if (sheetId === undefined || templateId === undefined || shapeKey === undefined || row === undefined || col === undefined) {
      return undefined
    }
    return { sheetId, templateId, shapeKey, cellIndex, row, col }
  }

  const clearCellRecord = (cellIndex: number): void => {
    cellRecordSheetIds[cellIndex] = undefined
    cellRecordTemplateIds[cellIndex] = undefined
    cellRecordShapeKeys[cellIndex] = undefined
    cellRecordRows[cellIndex] = undefined
    cellRecordCols[cellIndex] = undefined
  }

  const getExistingFamily = (args: FormulaFamilyKey): MutableFormulaFamily | undefined => {
    const recent = recentFamilyByTemplateId.get(args.templateId)
    if (recent && recent.sheetId === args.sheetId && recent.shapeKey === args.shapeKey && familiesById.get(recent.id) === recent) {
      return recent
    }
    const key = keyForFormulaFamily(args)
    const existingId = familyIdByKey.get(key)
    if (existingId === undefined) {
      return undefined
    }
    const existing = familiesById.get(existingId)
    if (existing) {
      recentFamilyByTemplateId.set(args.templateId, existing)
    }
    return existing
  }

  const getOrCreateFamily = (args: FormulaFamilyKey): MutableFormulaFamily => {
    const existing = getExistingFamily(args)
    if (existing) {
      return existing
    }
    const key = keyForFormulaFamily(args)
    const family: MutableFormulaFamily = {
      id: nextFamilyId,
      sheetId: args.sheetId,
      templateId: args.templateId,
      shapeKey: args.shapeKey,
      key,
      runs: [],
      rowRunsByFixedIndex: new Map(),
      columnRunsByFixedIndex: new Map(),
      singletonRunsByRow: new Map(),
      rowAppendRunByFixedIndex: [],
      recentAppendRun: undefined,
    }
    nextFamilyId += 1
    familiesById.set(family.id, family)
    familyIdByKey.set(key, family.id)
    recentFamilyByTemplateId.set(args.templateId, family)
    return family
  }

  const recordFormulaMemberAt = (key: FormulaFamilyKey, cellIndex: number, row: number, col: number): void => {
    cellRecordSheetIds[cellIndex] = key.sheetId
    cellRecordTemplateIds[cellIndex] = key.templateId
    cellRecordShapeKeys[cellIndex] = key.shapeKey
    cellRecordRows[cellIndex] = row
    cellRecordCols[cellIndex] = col
    memberCount += 1
    sheetMemberCounts.set(key.sheetId, (sheetMemberCounts.get(key.sheetId) ?? 0) + 1)
  }

  const recordFormulaMember = (key: FormulaFamilyKey, member: FormulaFamilyMember): void => {
    recordFormulaMemberAt(key, member.cellIndex, member.row, member.col)
  }

  const makeRun = (
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
  ): MutableFormulaFamilyMemberRun => {
    const sorted = sortRunMembers(axis, members)
    const first = sorted[0]!
    const last = sorted[sorted.length - 1]!
    return {
      id: nextRunId++,
      axis,
      fixedIndex,
      start: axis === 'row' ? first.row : first.col,
      end: axis === 'row' ? last.row : last.col,
      step: inferRunStep(axis, sorted),
      cellIndices: sorted.map((member) => member.cellIndex),
    }
  }

  const replaceRunWithMembers = (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ): FormulaFamilyMembership => {
    const previousRun = family.runs[runIndex]
    if (previousRun) {
      unindexRun(family, previousRun)
    }
    const run = makeRun(axis, fixedIndex, members)
    family.runs.splice(runIndex, 1, run)
    indexRun(family, run)
    family.recentAppendRun = run
    run.cellIndices.forEach((cellIndex) => {
      setMembership(cellIndex, family.id, run.id)
    })
    return getMembershipRecord(memberCellIndex)!
  }

  const appendMemberToRun = (
    family: MutableFormulaFamily,
    run: MutableFormulaFamilyMemberRun,
    member: FormulaFamilyMember,
  ): FormulaFamilyMembership => {
    const wasSingleton = run.cellIndices.length === 1
    if (wasSingleton) {
      removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
    }
    const memberIndex = run.axis === 'row' ? member.row : member.col
    if (memberIndex < run.start) {
      run.start = memberIndex
      run.cellIndices.unshift(member.cellIndex)
    } else {
      run.end = Math.max(run.end, memberIndex)
      run.cellIndices.push(member.cellIndex)
    }
    family.recentAppendRun = run
    return setMembership(member.cellIndex, family.id, run.id)
  }

  const splitRunAfterRemoval = (family: MutableFormulaFamily, runIndex: number, removedCellIndex: number): void => {
    const run = family.runs[runIndex]
    if (!run) {
      return
    }
    const remainingMembers = run.cellIndices
      .filter((cellIndex) => cellIndex !== removedCellIndex)
      .flatMap((cellIndex): FormulaFamilyMember[] => {
        const record = getCellRecord(cellIndex)
        return record ? [{ cellIndex, row: record.row, col: record.col }] : []
      })
    unindexRun(family, run)
    family.runs.splice(runIndex, 1)
    if (remainingMembers.length === 0) {
      return
    }
    const groups = groupRunMembersByStep(run.axis, remainingMembers, run.step)
    groups.forEach((group, offset) => {
      const nextRun = makeRun(run.axis, run.fixedIndex, group)
      family.runs.splice(runIndex + offset, 0, nextRun)
      indexRun(family, nextRun)
      nextRun.cellIndices.forEach((cellIndex) => {
        setMembership(cellIndex, family.id, nextRun.id)
      })
    })
  }

  const unregisterFormula = (cellIndex: number): boolean => {
    const membership = getMembershipRecord(cellIndex)
    const record = getCellRecord(cellIndex)
    if (!membership || !record) {
      return false
    }
    membershipFamilyIds[cellIndex] = 0
    membershipRunIds[cellIndex] = 0
    clearCellRecord(cellIndex)
    memberCount -= 1
    const sheetMemberCount = sheetMemberCounts.get(record.sheetId) ?? 0
    if (sheetMemberCount <= 1) {
      sheetMemberCounts.delete(record.sheetId)
    } else {
      sheetMemberCounts.set(record.sheetId, sheetMemberCount - 1)
    }
    const family = familiesById.get(membership.familyId)
    if (!family) {
      return true
    }
    const runIndex = family.runs.findIndex((run) => run.id === membership.runId)
    splitRunAfterRemoval(family, runIndex, cellIndex)
    if (family.runs.length === 0) {
      familiesById.delete(family.id)
      familyIdByKey.delete(family.key)
      if (recentFamilyByTemplateId.get(record.templateId) === family) {
        recentFamilyByTemplateId.delete(record.templateId)
      }
      structuralSourceTransforms.delete(family.id)
    }
    return true
  }

  const upsertFormula = (args: FormulaFamilyKey & FormulaFamilyMember): FormulaFamilyMembership => {
    if ((membershipFamilyIds[args.cellIndex] ?? 0) !== 0) {
      unregisterFormula(args.cellIndex)
    }
    const family = getOrCreateFamily(args)
    const member: FormulaFamilyMember = args
    recordFormulaMember(args, member)

    const rowAppendRun = family.rowAppendRunByFixedIndex[member.col]
    if (rowAppendRun && canAppendStridedRunMember(rowAppendRun, member.row)) {
      return appendMemberToRun(family, rowAppendRun, member)
    }

    const recentRun = family.recentAppendRun
    if (recentRun?.axis === 'row' && recentRun.fixedIndex === member.col && canAppendStridedRunMember(recentRun, member.row)) {
      return appendMemberToRun(family, recentRun, member)
    }

    const rowRuns = family.rowRunsByFixedIndex.get(member.col)
    if (rowRuns?.length === 1) {
      const run = rowRuns[0]!
      if (canAppendStridedRunMember(run, member.row)) {
        return appendMemberToRun(family, run, member)
      }
    }

    for (const run of candidateRunsForMember(family, member)) {
      const runIndex = family.runs.indexOf(run)
      if (runIndex < 0) {
        continue
      }
      const maybeMembership = tryMergeRun(family, runIndex, run, member, appendMemberToRun, replaceRunWithMembers, getCellRecord)
      if (maybeMembership) {
        return maybeMembership
      }
    }
    for (let runIndex = 0; runIndex < family.runs.length; runIndex += 1) {
      const run = family.runs[runIndex]!
      const maybeMembership = tryMergeRun(family, runIndex, run, member, appendMemberToRun, replaceRunWithMembers, getCellRecord)
      if (maybeMembership) {
        return maybeMembership
      }
    }

    const run = makeRun('row', args.col, [member])
    family.runs.push(run)
    indexRun(family, run)
    family.recentAppendRun = run
    return setMembership(args.cellIndex, family.id, run.id)
  }

  const fallbackUpsertFormulaRun = (args: FormulaFamilyRunUpsertArgs): FormulaFamilyMembership[] => {
    args.members.forEach((member) => {
      upsertFormula({ sheetId: args.sheetId, templateId: args.templateId, shapeKey: args.shapeKey, ...member })
    })
    return args.members.map((member) => getMembershipRecord(member.cellIndex)!)
  }

  const fallbackRegisterFormulaRun = (args: FormulaFamilyRunUpsertArgs): void => {
    args.members.forEach((member) => {
      upsertFormula({ sheetId: args.sheetId, templateId: args.templateId, shapeKey: args.shapeKey, ...member })
    })
  }

  const tryRegisterFreshOrderedUniformRun = (args: FormulaFamilyRunUpsertArgs): boolean => {
    const descriptor = describeFreshOrderedUniformRun(args.members, membershipFamilyIds)
    if (!descriptor) {
      return false
    }
    const family = getOrCreateFamily(args)
    if (!canRegisterFreshOrderedRunWithoutMerging(family, descriptor, args.members)) {
      return false
    }
    const cellIndices: number[] = []
    cellIndices.length = args.members.length
    const run: MutableFormulaFamilyMemberRun = {
      id: nextRunId++,
      axis: descriptor.axis,
      fixedIndex: descriptor.fixedIndex,
      start: descriptor.start,
      end: descriptor.end,
      step: descriptor.step,
      cellIndices,
    }
    for (let index = 0; index < args.members.length; index += 1) {
      const member = args.members[index]!
      recordFormulaMember(args, member)
      run.cellIndices[index] = member.cellIndex
      membershipFamilyIds[member.cellIndex] = family.id
      membershipRunIds[member.cellIndex] = run.id
    }
    family.runs.push(run)
    indexRun(family, run)
    family.recentAppendRun = run
    return true
  }

  const registerFreshUniformRun = (args: FormulaFamilyFreshUniformRunRegistrationArgs): boolean => {
    const runLength = args.cellIndices.length
    if (runLength === 0 || (runLength > 1 && args.step <= 0)) {
      return false
    }
    const existingFamily = getExistingFamily(args)
    if (existingFamily && existingFamily.runs.length > 0) {
      return false
    }
    const cellIndices: number[] = []
    cellIndices.length = runLength
    for (let index = 0; index < runLength; index += 1) {
      const cellIndex = args.cellIndices[index]!
      if ((membershipFamilyIds[cellIndex] ?? 0) !== 0) {
        return false
      }
      cellIndices[index] = cellIndex
    }
    const family = existingFamily ?? getOrCreateFamily(args)
    const step = runLength === 1 ? 1 : args.step
    const run: MutableFormulaFamilyMemberRun = {
      id: nextRunId++,
      axis: args.axis,
      fixedIndex: args.fixedIndex,
      start: args.start,
      end: args.start + step * (runLength - 1),
      step,
      cellIndices,
    }
    for (let index = 0; index < runLength; index += 1) {
      const cellIndex = cellIndices[index]!
      const variableIndex = args.start + step * index
      const row = args.axis === 'row' ? variableIndex : args.fixedIndex
      const col = args.axis === 'row' ? args.fixedIndex : variableIndex
      recordFormulaMemberAt(args, cellIndex, row, col)
      membershipFamilyIds[cellIndex] = family.id
      membershipRunIds[cellIndex] = run.id
    }
    family.runs.push(run)
    indexRun(family, run)
    family.recentAppendRun = run
    return true
  }

  const tryUpsertFormulaRunDescriptors = (
    args: FormulaFamilyRunUpsertArgs,
    descriptors: readonly FormulaFamilyRunDescriptor[],
    materializeMemberships = true,
  ): FormulaFamilyMembership[] | undefined => {
    const existingFamily = getExistingFamily(args)
    const fastPaths: FormulaRunFastPath[] = []
    for (const descriptor of descriptors) {
      const fastPath = existingFamily ? getFormulaRunFastPath(existingFamily, descriptor) : { kind: 'create' as const }
      if (!fastPath) {
        return undefined
      }
      fastPaths.push(fastPath)
    }
    const family = existingFamily ?? getOrCreateFamily(args)
    args.members.forEach((member) => {
      recordFormulaMember(args, member)
    })
    descriptors.forEach((descriptor, index) => {
      const fastPath = fastPaths[index]!
      const run = fastPath.kind === 'append' ? fastPath.run : makeRun(descriptor.axis, descriptor.fixedIndex, descriptor.members)
      if (fastPath.kind === 'append') {
        appendFormulaRunMembers(family, run, descriptor)
      } else {
        family.runs.push(run)
        indexRun(family, run)
        family.recentAppendRun = run
      }
      descriptor.members.forEach((member) => {
        setMembership(member.cellIndex, family.id, run.id)
      })
    })
    if (!materializeMemberships) {
      return noMemberships
    }
    return args.members.map((member) => getMembershipRecord(member.cellIndex)!)
  }

  const upsertFormulaRun = (args: FormulaFamilyRunUpsertArgs): FormulaFamilyMembership[] => {
    const inOrderDescriptors = describeFreshFormulaRunMemberSegmentsInOrder(args.members, membershipFamilyIds)
    if (inOrderDescriptors) {
      const memberships = tryUpsertFormulaRunDescriptors(args, inOrderDescriptors)
      if (memberships) {
        return memberships
      }
    }
    if (hasDuplicateRunMemberCellIndex(args.members)) {
      return fallbackUpsertFormulaRun(args)
    }
    if (args.members.some((member) => (membershipFamilyIds[member.cellIndex] ?? 0) !== 0)) {
      return fallbackUpsertFormulaRun(args)
    }
    const descriptor = describeFormulaRunMembers(args.members)
    if (!descriptor) {
      return fallbackUpsertFormulaRun(args)
    }
    return tryUpsertFormulaRunDescriptors(args, [descriptor]) ?? fallbackUpsertFormulaRun(args)
  }

  const registerFormulaRun = (args: FormulaFamilyRunUpsertArgs): void => {
    if (tryRegisterFreshOrderedUniformRun(args)) {
      return
    }
    const inOrderDescriptors = describeFreshFormulaRunMemberSegmentsInOrder(args.members, membershipFamilyIds)
    if (inOrderDescriptors && tryUpsertFormulaRunDescriptors(args, inOrderDescriptors, false)) {
      return
    }
    if (hasDuplicateRunMemberCellIndex(args.members)) {
      fallbackRegisterFormulaRun(args)
      return
    }
    if (args.members.some((member) => (membershipFamilyIds[member.cellIndex] ?? 0) !== 0)) {
      fallbackRegisterFormulaRun(args)
      return
    }
    const descriptor = describeFormulaRunMembers(args.members)
    if (!descriptor || !tryUpsertFormulaRunDescriptors(args, [descriptor], false)) {
      fallbackRegisterFormulaRun(args)
    }
  }

  const peekStructuralSourceTransforms = (): FormulaFamilyStructuralSourceTransformEntry[] => {
    const entries: FormulaFamilyStructuralSourceTransformEntry[] = []
    structuralSourceTransforms.forEach((transform, familyId) => {
      const family = familiesById.get(familyId)
      if (!family) {
        return
      }
      entries.push({
        cellIndices: family.runs.flatMap((run) => run.cellIndices),
        transform,
      })
    })
    return entries
  }

  return {
    upsertFormula,
    upsertFormulaRun,
    registerFormulaRun,
    registerFreshUniformRun,
    unregisterFormula,
    getMembership(cellIndex) {
      return getMembershipRecord(cellIndex)
    },
    countSheetMembers(sheetId) {
      return sheetMemberCounts.get(sheetId) ?? 0
    },
    forEachFamily(fn) {
      familiesById.forEach((family) => {
        fn({
          id: family.id,
          sheetId: family.sheetId,
          templateId: family.templateId,
          shapeKey: family.shapeKey,
          runs: family.runs,
        })
      })
    },
    setStructuralSourceTransform(familyId, transform) {
      if (familiesById.has(familyId)) {
        const existing = structuralSourceTransforms.get(familyId)
        structuralSourceTransforms.set(
          familyId,
          existing ? (composeFormulaFamilyStructuralSourceTransform(existing, transform) ?? transform) : transform,
        )
      }
    },
    getStructuralSourceTransform(cellIndex) {
      const membership = getMembershipRecord(cellIndex)
      return membership ? structuralSourceTransforms.get(membership.familyId) : undefined
    },
    hasStructuralSourceTransforms() {
      return structuralSourceTransforms.size > 0
    },
    peekStructuralSourceTransforms,
    consumeStructuralSourceTransforms() {
      const entries = peekStructuralSourceTransforms()
      structuralSourceTransforms.clear()
      return entries
    },
    getStats() {
      let runCount = 0
      familiesById.forEach((family) => {
        runCount += family.runs.length
      })
      return {
        familyCount: familiesById.size,
        runCount,
        memberCount,
      }
    },
    listFamilies() {
      return [...familiesById.values()]
        .toSorted((left, right) => left.sheetId - right.sheetId || left.templateId - right.templateId || left.id - right.id)
        .map((family) => ({
          id: family.id,
          sheetId: family.sheetId,
          templateId: family.templateId,
          shapeKey: family.shapeKey,
          runs: family.runs.map((run) => ({ ...run, cellIndices: [...run.cellIndices] })),
        }))
    },
    invalidateSheet(sheetId) {
      const removedCellIndices: number[] = []
      for (let cellIndex = 0; cellIndex < cellRecordSheetIds.length; cellIndex += 1) {
        if (cellRecordSheetIds[cellIndex] === sheetId) {
          removedCellIndices.push(cellIndex)
        }
      }
      removedCellIndices.forEach((cellIndex) => {
        unregisterFormula(cellIndex)
      })
    },
    applyStructuralInvalidation(args) {
      const removedCellIndices: number[] = []
      for (let cellIndex = 0; cellIndex < cellRecordSheetIds.length; cellIndex += 1) {
        if (cellRecordSheetIds[cellIndex] !== args.sheetId) {
          continue
        }
        const axisIndex = args.axis === 'row' ? cellRecordRows[cellIndex] : cellRecordCols[cellIndex]
        if (axisIndex === undefined) {
          continue
        }
        if (axisIndex >= args.start && axisIndex < args.end) {
          removedCellIndices.push(cellIndex)
        }
      }
      removedCellIndices.forEach((cellIndex) => {
        unregisterFormula(cellIndex)
      })
    },
    clear() {
      familiesById.clear()
      familyIdByKey.clear()
      recentFamilyByTemplateId.clear()
      cellRecordSheetIds.length = 0
      cellRecordTemplateIds.length = 0
      cellRecordShapeKeys.length = 0
      cellRecordRows.length = 0
      cellRecordCols.length = 0
      membershipFamilyIds.length = 0
      membershipRunIds.length = 0
      sheetMemberCounts.clear()
      structuralSourceTransforms.clear()
      memberCount = 0
    },
  }
}
