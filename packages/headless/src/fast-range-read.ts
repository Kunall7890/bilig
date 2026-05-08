import { BLOCK_COLS, BLOCK_ROWS, type SpreadsheetEngine } from '@bilig/core'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkPaperCellRange } from './work-paper-types.js'

const EMPTY_CELL_VALUE: CellValue = Object.freeze({ tag: ValueTag.Empty })
const FAST_PHYSICAL_RANGE_AREA_LIMIT = 262_144

export function readFastPhysicalRangeValues(engine: SpreadsheetEngine, range: WorkPaperCellRange): CellValue[][] | undefined {
  const sheet = engine.workbook.getSheetById(range.start.sheet)
  if (!sheet || sheet.structureVersion !== 1) {
    return undefined
  }
  const height = range.end.row - range.start.row + 1
  const width = range.end.col - range.start.col + 1
  if (height <= 0 || width <= 0 || height * width > FAST_PHYSICAL_RANGE_AREA_LIMIT) {
    return undefined
  }

  const rows: CellValue[][] = []
  rows.length = height
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const row: CellValue[] = []
    row.length = width
    rows[rowOffset] = row
  }
  const cellStore = engine.workbook.cellStore
  const cellTags = cellStore.tags
  const cellNumbers = cellStore.numbers
  const cellStringIds = cellStore.stringIds
  const cellErrors = cellStore.errors
  const strings = engine.strings
  let filledCells = 0
  const blockRowStart = Math.floor(range.start.row / BLOCK_ROWS)
  const blockRowEnd = Math.floor(range.end.row / BLOCK_ROWS)
  const blockColStart = Math.floor(range.start.col / BLOCK_COLS)
  const blockColEnd = Math.floor(range.end.col / BLOCK_COLS)
  for (let blockRow = blockRowStart; blockRow <= blockRowEnd; blockRow += 1) {
    const absoluteBlockRow = blockRow * BLOCK_ROWS
    const localRowStart = Math.max(range.start.row - absoluteBlockRow, 0)
    const localRowEnd = Math.min(range.end.row - absoluteBlockRow, BLOCK_ROWS - 1)
    for (let blockCol = blockColStart; blockCol <= blockColEnd; blockCol += 1) {
      const block = sheet.grid.blocks.get(blockRow * 1_000_000 + blockCol)
      if (!block) {
        continue
      }
      const absoluteBlockCol = blockCol * BLOCK_COLS
      const localColStart = Math.max(range.start.col - absoluteBlockCol, 0)
      const localColEnd = Math.min(range.end.col - absoluteBlockCol, BLOCK_COLS - 1)
      for (let localRow = localRowStart; localRow <= localRowEnd; localRow += 1) {
        const row = rows[absoluteBlockRow + localRow - range.start.row]!
        const blockRowOffset = localRow * BLOCK_COLS
        for (let localCol = localColStart; localCol <= localColEnd; localCol += 1) {
          const encodedCellIndex = block[blockRowOffset + localCol]!
          if (encodedCellIndex === 0) {
            continue
          }
          const cellIndex = encodedCellIndex - 1
          const outputCol = absoluteBlockCol + localCol - range.start.col
          switch ((cellTags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
            case ValueTag.Number:
              row[outputCol] = { tag: ValueTag.Number, value: cellNumbers[cellIndex] ?? 0 }
              break
            case ValueTag.Boolean:
              row[outputCol] = {
                tag: ValueTag.Boolean,
                value: (cellNumbers[cellIndex] ?? 0) !== 0,
              }
              break
            case ValueTag.String: {
              const stringId = cellStringIds[cellIndex] ?? 0
              row[outputCol] = {
                tag: ValueTag.String,
                value: stringId === 0 ? '' : strings.get(stringId),
                stringId,
              }
              break
            }
            case ValueTag.Error:
              row[outputCol] = {
                tag: ValueTag.Error,
                code: (cellErrors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None,
              }
              break
            case ValueTag.Empty:
            default:
              row[outputCol] = EMPTY_CELL_VALUE
              break
          }
          filledCells += 1
        }
      }
    }
  }
  if (filledCells < height * width) {
    for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
      const row = rows[rowOffset]!
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        row[colOffset] ??= EMPTY_CELL_VALUE
      }
    }
  }
  return rows
}
