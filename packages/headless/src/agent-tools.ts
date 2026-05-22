import { Buffer } from 'node:buffer'

import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import type { WorkPaper } from './work-paper.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperCellRange, WorkPaperChange } from './work-paper-types.js'

export interface WorkPaperAgentToolOptions {
  readonly defaultReadRange?: string
  readonly trackedRanges?: readonly string[]
  readonly allowedInputSheets?: readonly string[]
  readonly writable?: boolean
}

export interface WorkPaperAgentReadRangeArgs {
  readonly range?: string
}

export interface WorkPaperAgentSetCellArgs {
  readonly target?: string
  readonly sheetName?: string
  readonly address?: string
  readonly value: RawCellContent
}

export interface WorkPaperAgentRangeReadback {
  readonly range: string
  readonly values: ReturnType<WorkPaper['getRangeValues']>
  readonly serialized: ReturnType<WorkPaper['getRangeSerialized']>
}

export interface WorkPaperAgentSetCellReadback {
  readonly editedCell: string
  readonly changes: readonly WorkPaperChange[]
  readonly trackedRanges: Readonly<
    Record<
      string,
      {
        readonly before: WorkPaperAgentRangeReadback
        readonly after: WorkPaperAgentRangeReadback
        readonly restored: WorkPaperAgentRangeReadback
      }
    >
  >
  readonly checks: {
    readonly previousValue: ReturnType<WorkPaper['getCellSerialized']>
    readonly newValue: ReturnType<WorkPaper['getCellSerialized']>
    readonly cellValueChanged: boolean
    readonly trackedRangesChanged: boolean
    readonly formulasPersisted: boolean
    readonly restoredMatchesAfter: boolean
    readonly serializedBytes: number
  }
}

export interface WorkPaperAgentToolSet {
  readWorkPaperRange(args?: WorkPaperAgentReadRangeArgs): WorkPaperAgentRangeReadback
  setWorkPaperCell(args: WorkPaperAgentSetCellArgs): WorkPaperAgentSetCellReadback
}

export const workPaperAgentToolSchemas = {
  readWorkPaperRange: {
    name: 'read_workpaper_range',
    description: 'Read a WorkPaper range as computed values plus serialized cell contents.',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          default: 'Summary!A1:B5',
          description: 'Sheet-qualified A1 range such as Summary!A1:B5.',
        },
      },
      additionalProperties: false,
    },
  },
  setWorkPaperCell: {
    name: 'set_workpaper_cell',
    description: 'Set one allowed WorkPaper input cell and return recalculation, persistence, and restore proof.',
    inputSchema: {
      type: 'object',
      required: ['value'],
      properties: {
        target: {
          type: 'string',
          description: 'Optional sheet-qualified A1 cell such as Inputs!B3.',
        },
        sheetName: {
          type: 'string',
          description: 'Sheet name when target is split into sheetName plus address.',
        },
        address: {
          type: 'string',
          description: 'A1 cell address such as B3 when sheetName is provided.',
        },
        value: {
          type: ['string', 'number', 'boolean', 'null'],
          description: 'Literal cell value to write.',
        },
      },
      additionalProperties: false,
    },
  },
} as const

export function createWorkPaperAgentTools(workpaper: WorkPaper, options: WorkPaperAgentToolOptions = {}): WorkPaperAgentToolSet {
  const defaultReadRange = options.defaultReadRange ?? 'Summary!A1:B5'
  const trackedRanges = options.trackedRanges && options.trackedRanges.length > 0 ? options.trackedRanges : [defaultReadRange]

  return {
    readWorkPaperRange(args: WorkPaperAgentReadRangeArgs = {}) {
      return readRange(workpaper, args.range ?? defaultReadRange)
    },

    setWorkPaperCell(args: WorkPaperAgentSetCellArgs) {
      if (options.writable !== true) {
        throw new Error('WorkPaper agent tools are read-only by default. Pass writable: true before exposing setWorkPaperCell.')
      }

      const target = parseSetCellTarget(workpaper, args)
      requireAllowedSheet(workpaper, target, options.allowedInputSheets)

      const before = readTrackedRanges(workpaper, trackedRanges)
      const previousValue = workpaper.getCellSerialized(target)
      const changes = workpaper.setCellContents(target, args.value)
      const after = readTrackedRanges(workpaper, trackedRanges)
      const serialized = serializeWorkPaperDocument(
        exportWorkPaperDocument(workpaper, {
          includeConfig: true,
        }),
      )
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))

      try {
        const restoredRanges = readTrackedRanges(restored, trackedRanges)
        const trackedRangeReadback = mergeTrackedRanges(before, after, restoredRanges)

        return {
          editedCell: workpaper.simpleCellAddressToString(target, { includeSheetName: true }),
          changes,
          trackedRanges: trackedRangeReadback,
          checks: {
            previousValue,
            newValue: workpaper.getCellSerialized(target),
            cellValueChanged: !sameJson(previousValue, workpaper.getCellSerialized(target)),
            trackedRangesChanged: !sameJson(rangeValues(before), rangeValues(after)),
            formulasPersisted: sameJson(rangeSerialized(after), rangeSerialized(restoredRanges)),
            restoredMatchesAfter: sameJson(rangeValues(after), rangeValues(restoredRanges)),
            serializedBytes: Buffer.byteLength(serialized, 'utf8'),
          },
        }
      } finally {
        restored.dispose()
      }
    },
  }
}

function readRange(workpaper: WorkPaper, range: string): WorkPaperAgentRangeReadback {
  const parsed = workpaper.simpleCellRangeFromString(range)
  if (parsed === undefined) {
    throw new Error(`Invalid WorkPaper range: ${range}`)
  }

  return readParsedRange(workpaper, parsed)
}

function readParsedRange(workpaper: WorkPaper, range: WorkPaperCellRange): WorkPaperAgentRangeReadback {
  return {
    range: workpaper.simpleCellRangeToString(range, { includeSheetName: true }),
    values: workpaper.getRangeValues(range),
    serialized: workpaper.getRangeSerialized(range),
  }
}

function readTrackedRanges(workpaper: WorkPaper, ranges: readonly string[]): Readonly<Record<string, WorkPaperAgentRangeReadback>> {
  const readback: Record<string, WorkPaperAgentRangeReadback> = {}
  for (const range of ranges) {
    const item = readRange(workpaper, range)
    readback[item.range] = item
  }
  return readback
}

function mergeTrackedRanges(
  before: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
  after: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
  restored: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
): WorkPaperAgentSetCellReadback['trackedRanges'] {
  const merged: Record<
    string,
    {
      before: WorkPaperAgentRangeReadback
      after: WorkPaperAgentRangeReadback
      restored: WorkPaperAgentRangeReadback
    }
  > = {}

  for (const range of Object.keys(after)) {
    const beforeRange = before[range]
    const afterRange = after[range]
    const restoredRange = restored[range]
    if (beforeRange === undefined || afterRange === undefined || restoredRange === undefined) {
      throw new Error(`Tracked range disappeared during WorkPaper agent proof: ${range}`)
    }
    merged[range] = {
      before: beforeRange,
      after: afterRange,
      restored: restoredRange,
    }
  }

  return merged
}

function parseSetCellTarget(workpaper: WorkPaper, args: WorkPaperAgentSetCellArgs): WorkPaperCellAddress {
  if (args.target !== undefined) {
    const parsed = workpaper.simpleCellAddressFromString(args.target)
    if (parsed === undefined) {
      throw new Error(`Invalid WorkPaper cell target: ${args.target}`)
    }
    return parsed
  }

  if (args.sheetName === undefined || args.address === undefined) {
    throw new Error('Expected setWorkPaperCell target or sheetName plus address.')
  }

  const sheetId = workpaper.getSheetId(args.sheetName)
  if (sheetId === undefined) {
    throw new Error(`Unknown WorkPaper sheet: ${args.sheetName}`)
  }
  const parsed = workpaper.simpleCellAddressFromString(args.address, sheetId)
  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid WorkPaper cell target: ${args.sheetName}!${args.address}`)
  }
  return parsed
}

function requireAllowedSheet(workpaper: WorkPaper, target: WorkPaperCellAddress, allowedInputSheets: readonly string[] | undefined): void {
  if (allowedInputSheets === undefined || allowedInputSheets.length === 0) {
    return
  }

  const allowedSheetIds = new Set(
    allowedInputSheets.map((sheetName) => {
      const sheetId = workpaper.getSheetId(sheetName)
      if (sheetId === undefined) {
        throw new Error(`Unknown allowed input sheet: ${sheetName}`)
      }
      return sheetId
    }),
  )

  if (!allowedSheetIds.has(target.sheet)) {
    throw new Error(
      `WorkPaper agent write is outside allowed input sheets: ${workpaper.simpleCellAddressToString(target, { includeSheetName: true })}`,
    )
  }
}

function rangeValues(
  ranges: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
): Readonly<Record<string, WorkPaperAgentRangeReadback['values']>> {
  return mapRangeReadback(ranges, (range) => range.values)
}

function rangeSerialized(
  ranges: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
): Readonly<Record<string, WorkPaperAgentRangeReadback['serialized']>> {
  return mapRangeReadback(ranges, (range) => range.serialized)
}

function mapRangeReadback<T>(
  ranges: Readonly<Record<string, WorkPaperAgentRangeReadback>>,
  mapper: (range: WorkPaperAgentRangeReadback) => T,
): Readonly<Record<string, T>> {
  const mapped: Record<string, T> = {}
  for (const [range, readback] of Object.entries(ranges)) {
    mapped[range] = mapper(readback)
  }
  return mapped
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
