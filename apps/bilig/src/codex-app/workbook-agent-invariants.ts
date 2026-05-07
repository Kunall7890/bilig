import { SpreadsheetEngine } from '@bilig/core'
import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, WorkbookSnapshot } from '@bilig/protocol'
import { isDeepStrictEqual } from 'node:util'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

interface InvariantProblem {
  code: string
  message: string
  sheetName?: string
  address?: string
}

export interface WorkbookInvariantVerificationReport {
  summary: {
    ok: boolean
    problemCount: number
    roundTripChecked: boolean
    roundTripStable: boolean
  }
  problems: Array<{
    code: string
    message: string
    sheetName?: string
    address?: string
  }>
}

function addProblem(
  problems: InvariantProblem[],
  code: string,
  message: string,
  options: Pick<InvariantProblem, 'sheetName' | 'address'> = {},
): void {
  const problem: InvariantProblem = {
    code,
    message,
  }
  if (options.sheetName !== undefined) {
    problem.sheetName = options.sheetName
  }
  if (options.address !== undefined) {
    problem.address = options.address
  }
  problems.push(problem)
}

function ensureSheetExists(sheetNames: ReadonlySet<string>, problems: InvariantProblem[], sheetName: string, context: string): boolean {
  if (sheetNames.has(sheetName)) {
    return true
  }
  addProblem(problems, 'missingSheet', `${context} references missing sheet ${sheetName}`)
  return false
}

function validateRangeRef(problems: InvariantProblem[], expectedSheetName: string | undefined, range: CellRangeRef, context: string): void {
  if (expectedSheetName !== undefined && range.sheetName !== expectedSheetName) {
    addProblem(problems, 'rangeSheetMismatch', `${context} range points at ${range.sheetName} instead of ${expectedSheetName}`, {
      sheetName: expectedSheetName,
    })
  }
  try {
    parseCellAddress(range.startAddress, range.sheetName)
    parseCellAddress(range.endAddress, range.sheetName)
  } catch (error) {
    addProblem(problems, 'invalidRangeAddress', `${context} range is invalid: ${error instanceof Error ? error.message : String(error)}`, {
      sheetName: expectedSheetName ?? range.sheetName,
    })
  }
}

function validateAddressRef(problems: InvariantProblem[], sheetName: string, address: string, context: string): void {
  try {
    parseCellAddress(address, sheetName)
  } catch (error) {
    addProblem(problems, 'invalidCellAddress', `${context} address is invalid: ${error instanceof Error ? error.message : String(error)}`, {
      sheetName,
      address,
    })
  }
}

function collectInvariantProblems(snapshot: WorkbookSnapshot): InvariantProblem[] {
  const problems: InvariantProblem[] = []
  const sheetNames = new Set<string>()
  const sheetOrders = new Set<number>()

  for (const sheet of snapshot.sheets) {
    if (sheetNames.has(sheet.name)) {
      addProblem(problems, 'duplicateSheetName', `Duplicate sheet name ${sheet.name}`, {
        sheetName: sheet.name,
      })
    } else {
      sheetNames.add(sheet.name)
    }
    if (sheetOrders.has(sheet.order)) {
      addProblem(problems, 'duplicateSheetOrder', `Duplicate sheet order ${String(sheet.order)}`, {
        sheetName: sheet.name,
      })
    } else {
      sheetOrders.add(sheet.order)
    }
    const addresses = new Set<string>()
    for (const cell of sheet.cells) {
      validateAddressRef(problems, sheet.name, cell.address, 'Cell')
      if (addresses.has(cell.address)) {
        addProblem(problems, 'duplicateCellAddress', `Duplicate cell ${sheet.name}!${cell.address}`, {
          sheetName: sheet.name,
          address: cell.address,
        })
      } else {
        addresses.add(cell.address)
      }
    }

    for (const styleRange of sheet.metadata?.styleRanges ?? []) {
      validateRangeRef(problems, sheet.name, styleRange.range, 'Style range')
    }
    for (const formatRange of sheet.metadata?.formatRanges ?? []) {
      validateRangeRef(problems, sheet.name, formatRange.range, 'Number format range')
    }
    for (const filter of sheet.metadata?.filters ?? []) {
      validateRangeRef(problems, sheet.name, filter, 'Filter')
    }
    for (const sort of sheet.metadata?.sorts ?? []) {
      validateRangeRef(problems, sheet.name, sort.range, 'Sort')
      for (const key of sort.keys) {
        validateAddressRef(problems, sheet.name, key.keyAddress, 'Sort key')
      }
    }
    for (const validation of sheet.metadata?.validations ?? []) {
      validateRangeRef(problems, sheet.name, validation.range, 'Validation')
    }
    for (const conditionalFormat of sheet.metadata?.conditionalFormats ?? []) {
      validateRangeRef(problems, sheet.name, conditionalFormat.range, 'Conditional format')
    }
    for (const protectedRange of sheet.metadata?.protectedRanges ?? []) {
      validateRangeRef(problems, sheet.name, protectedRange.range, 'Protected range')
    }
    for (const commentThread of sheet.metadata?.commentThreads ?? []) {
      if (commentThread.sheetName !== sheet.name) {
        addProblem(
          problems,
          'commentThreadSheetMismatch',
          `Comment thread ${commentThread.threadId} points at ${commentThread.sheetName} instead of ${sheet.name}`,
          { sheetName: sheet.name, address: commentThread.address },
        )
      }
      validateAddressRef(problems, sheet.name, commentThread.address, 'Comment thread')
    }
    for (const note of sheet.metadata?.notes ?? []) {
      validateAddressRef(problems, sheet.name, note.address, 'Note')
    }
    for (const record of sheet.metadata?.rowMetadata ?? []) {
      if (record.count <= 0 || record.start < 0) {
        addProblem(
          problems,
          'invalidRowMetadata',
          `Invalid row metadata region start=${String(record.start)} count=${String(record.count)}`,
          { sheetName: sheet.name },
        )
      }
    }
    for (const record of sheet.metadata?.columnMetadata ?? []) {
      if (record.count <= 0 || record.start < 0) {
        addProblem(
          problems,
          'invalidColumnMetadata',
          `Invalid column metadata region start=${String(record.start)} count=${String(record.count)}`,
          { sheetName: sheet.name },
        )
      }
    }
  }

  for (const table of snapshot.workbook.metadata?.tables ?? []) {
    if (!ensureSheetExists(sheetNames, problems, table.sheetName, `Table ${table.name}`)) {
      continue
    }
    validateRangeRef(
      problems,
      table.sheetName,
      {
        sheetName: table.sheetName,
        startAddress: table.startAddress,
        endAddress: table.endAddress,
      },
      `Table ${table.name}`,
    )
  }
  for (const pivot of snapshot.workbook.metadata?.pivots ?? []) {
    const hasPivotSheet = ensureSheetExists(sheetNames, problems, pivot.sheetName, `Pivot ${pivot.name}`)
    const hasSourceSheet = ensureSheetExists(sheetNames, problems, pivot.source.sheetName, `Pivot ${pivot.name} source`)
    if (hasPivotSheet) {
      validateAddressRef(problems, pivot.sheetName, pivot.address, `Pivot ${pivot.name}`)
    }
    if (hasSourceSheet) {
      validateRangeRef(problems, pivot.source.sheetName, pivot.source, `Pivot ${pivot.name} source`)
    }
    if (pivot.rows <= 0 || pivot.cols <= 0) {
      addProblem(
        problems,
        'invalidPivotExtent',
        `Pivot ${pivot.name} has invalid output extent ${String(pivot.rows)}x${String(pivot.cols)}`,
      )
    }
  }
  for (const chart of snapshot.workbook.metadata?.charts ?? []) {
    const hasChartSheet = ensureSheetExists(sheetNames, problems, chart.sheetName, `Chart ${chart.id}`)
    const hasSourceSheet = ensureSheetExists(sheetNames, problems, chart.source.sheetName, `Chart ${chart.id} source`)
    if (hasChartSheet) {
      validateAddressRef(problems, chart.sheetName, chart.address, `Chart ${chart.id}`)
    }
    if (hasSourceSheet) {
      validateRangeRef(problems, chart.source.sheetName, chart.source, `Chart ${chart.id} source`)
    }
    if (chart.rows <= 0 || chart.cols <= 0) {
      addProblem(problems, 'invalidChartExtent', `Chart ${chart.id} has invalid footprint ${String(chart.rows)}x${String(chart.cols)}`)
    }
  }
  for (const image of snapshot.workbook.metadata?.images ?? []) {
    if (ensureSheetExists(sheetNames, problems, image.sheetName, `Image ${image.id}`)) {
      validateAddressRef(problems, image.sheetName, image.address, `Image ${image.id}`)
    }
    if (image.rows <= 0 || image.cols <= 0) {
      addProblem(problems, 'invalidImageExtent', `Image ${image.id} has invalid footprint ${String(image.rows)}x${String(image.cols)}`, {
        sheetName: image.sheetName,
        address: image.address,
      })
    }
  }
  for (const shape of snapshot.workbook.metadata?.shapes ?? []) {
    if (ensureSheetExists(sheetNames, problems, shape.sheetName, `Shape ${shape.id}`)) {
      validateAddressRef(problems, shape.sheetName, shape.address, `Shape ${shape.id}`)
    }
    if (shape.rows <= 0 || shape.cols <= 0) {
      addProblem(problems, 'invalidShapeExtent', `Shape ${shape.id} has invalid footprint ${String(shape.rows)}x${String(shape.cols)}`, {
        sheetName: shape.sheetName,
        address: shape.address,
      })
    }
  }
  for (const spill of snapshot.workbook.metadata?.spills ?? []) {
    if (ensureSheetExists(sheetNames, problems, spill.sheetName, 'Spill range')) {
      validateAddressRef(problems, spill.sheetName, spill.address, 'Spill range')
    }
    if (spill.rows <= 0 || spill.cols <= 0) {
      addProblem(
        problems,
        'invalidSpillExtent',
        `Spill ${spill.sheetName}!${spill.address} has invalid extent ${String(spill.rows)}x${String(spill.cols)}`,
        { sheetName: spill.sheetName, address: spill.address },
      )
    }
  }
  for (const definedName of snapshot.workbook.metadata?.definedNames ?? []) {
    const value = definedName.value
    if (typeof value !== 'object' || value === null) {
      continue
    }
    if (value.kind === 'cell-ref') {
      if (ensureSheetExists(sheetNames, problems, value.sheetName, `Defined name ${definedName.name}`)) {
        validateAddressRef(problems, value.sheetName, value.address, `Defined name ${definedName.name}`)
      }
    }
    if (value.kind === 'range-ref') {
      if (ensureSheetExists(sheetNames, problems, value.sheetName, `Defined name ${definedName.name}`)) {
        validateRangeRef(
          problems,
          value.sheetName,
          {
            sheetName: value.sheetName,
            startAddress: value.startAddress,
            endAddress: value.endAddress,
          },
          `Defined name ${definedName.name}`,
        )
      }
    }
  }

  return problems
}

export async function verifyWorkbookInvariants(
  runtime: WorkbookRuntime,
  input: {
    roundTrip?: boolean | undefined
  } = {},
): Promise<WorkbookInvariantVerificationReport> {
  const snapshot = runtime.engine.exportSnapshot()
  const problems = collectInvariantProblems(snapshot)
  const roundTripChecked = input.roundTrip !== false
  let roundTripStable = true
  if (roundTripChecked) {
    const restored = new SpreadsheetEngine({
      workbookName: `${runtime.documentId}:audit-roundtrip`,
      replicaId: `audit:${runtime.documentId}`,
    })
    await restored.ready()
    restored.importSnapshot(snapshot)
    const exported = restored.exportSnapshot()
    if (!isDeepStrictEqual(exported, snapshot)) {
      roundTripStable = false
      addProblem(problems, 'roundTripMismatch', 'Snapshot export/import round-trip changed workbook state')
    }
  }
  return {
    summary: {
      ok: problems.length === 0,
      problemCount: problems.length,
      roundTripChecked,
      roundTripStable,
    },
    problems,
  }
}
