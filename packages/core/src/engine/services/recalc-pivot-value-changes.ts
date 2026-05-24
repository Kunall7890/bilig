import type { CellValue } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { areCellValuesEqual, emptyValue } from '../../engine-value-utils.js'
import type { RecalcValueChangeCollector } from './recalc-value-change-collector.js'

export interface RecalcPivotValueChangeWorkbook {
  readonly cellStore: {
    readonly getValue: (cellIndex: number, resolveString: (stringId: number) => string) => CellValue
  }
  readonly getCellIndex: (sheetName: string, address: string) => number | undefined
}

export interface RecalcPivotValueChangeStringPool {
  readonly get: (stringId: number) => string
}

export interface RecalcPivotOutputShape {
  readonly sheetName: string
  readonly address: string
  readonly rows: number
  readonly cols: number
}

export function capturePivotOutputValues(input: {
  readonly pivot: RecalcPivotOutputShape
  readonly workbook: RecalcPivotValueChangeWorkbook
  readonly strings: RecalcPivotValueChangeStringPool
}): Map<number, CellValue> {
  const values = new Map<number, CellValue>()
  const owner = parseCellAddress(input.pivot.address, input.pivot.sheetName)
  for (let rowOffset = 0; rowOffset < input.pivot.rows; rowOffset += 1) {
    for (let colOffset = 0; colOffset < input.pivot.cols; colOffset += 1) {
      const cellIndex = input.workbook.getCellIndex(input.pivot.sheetName, formatAddress(owner.row + rowOffset, owner.col + colOffset))
      if (cellIndex !== undefined) {
        values.set(
          cellIndex,
          input.workbook.cellStore.getValue(cellIndex, (id) => input.strings.get(id)),
        )
      }
    }
  }
  return values
}

export function notePivotValueChanges(input: {
  readonly changed: readonly number[] | Uint32Array
  readonly before: ReadonlyMap<number, CellValue>
  readonly workbook: RecalcPivotValueChangeWorkbook
  readonly strings: RecalcPivotValueChangeStringPool
  readonly collector: RecalcValueChangeCollector
}): void {
  for (let index = 0; index < input.changed.length; index += 1) {
    const cellIndex = input.changed[index]!
    const beforeValue = input.before.get(cellIndex) ?? emptyValue()
    const afterValue = input.workbook.cellStore.getValue(cellIndex, (id) => input.strings.get(id))
    if (!areCellValuesEqual(beforeValue, afterValue)) {
      input.collector.note(cellIndex)
    }
  }
}
