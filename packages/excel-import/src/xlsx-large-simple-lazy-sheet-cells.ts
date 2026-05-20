import type { WorkbookSnapshot } from '@bilig/protocol'

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']
type WorkbookSheetCell = WorkbookSheetCells[number]

const lazyCellsBrand = Symbol('bilig.lazyImportedXlsxCells')

export interface ImportedWorkbookLazySheetCells extends Array<WorkbookSheetCell> {
  readonly [lazyCellsBrand]: true
}

export function createLazyWorkbookSheetCells(
  cellCount: number,
  materialize: (index: number) => WorkbookSheetCell | undefined,
): WorkbookSheetCells {
  const iterate = function* (): IterableIterator<WorkbookSheetCell> {
    for (let index = 0; index < cellCount; index += 1) {
      const cell = materialize(index)
      if (cell) {
        yield cell
      }
    }
  }
  const target: ImportedWorkbookLazySheetCells = Object.assign([], { [lazyCellsBrand]: true as const })
  let proxy: WorkbookSheetCells
  proxy = new Proxy<ImportedWorkbookLazySheetCells>(target, {
    get: (_target, property) => {
      if (property === lazyCellsBrand) {
        return true
      }
      if (property === 'length') {
        return cellCount
      }
      if (property === Symbol.iterator || property === 'values') {
        return iterate
      }
      if (property === 'entries') {
        return function* entries(): IterableIterator<[number, WorkbookSheetCell]> {
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell) {
              yield [index, cell]
            }
          }
        }
      }
      if (property === 'keys') {
        return function* keys(): IterableIterator<number> {
          for (let index = 0; index < cellCount; index += 1) {
            yield index
          }
        }
      }
      if (property === 'at') {
        return (index: number) => materialize(index < 0 ? cellCount + index : index)
      }
      if (property === 'forEach') {
        return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => void, thisArg?: unknown) => {
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell) {
              callback.call(thisArg, cell, index, proxy)
            }
          }
        }
      }
      if (property === 'map') {
        return <T>(callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => T, thisArg?: unknown): T[] => {
          const output: T[] = []
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell) {
              output.push(callback.call(thisArg, cell, index, proxy))
            }
          }
          return output
        }
      }
      if (property === 'filter') {
        return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
          const output: WorkbookSheetCell[] = []
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell && callback.call(thisArg, cell, index, proxy)) {
              output.push(cell)
            }
          }
          return output
        }
      }
      if (property === 'reduce') {
        return function reduce(
          callback: (previous: unknown, cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => unknown,
          initialValue?: unknown,
        ): unknown {
          let index = 0
          let accumulator = initialValue
          if (arguments.length < 2) {
            const first = materialize(0)
            if (!first) {
              throw new TypeError('Reduce of empty array with no initial value')
            }
            accumulator = first
            index = 1
          }
          for (; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell) {
              accumulator = callback(accumulator, cell, index, proxy)
            }
          }
          return accumulator
        }
      }
      if (property === 'some') {
        return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell && callback.call(thisArg, cell, index, proxy)) {
              return true
            }
          }
          return false
        }
      }
      if (property === 'every') {
        return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell && !callback.call(thisArg, cell, index, proxy)) {
              return false
            }
          }
          return true
        }
      }
      if (property === 'find') {
        return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
          for (let index = 0; index < cellCount; index += 1) {
            const cell = materialize(index)
            if (cell && callback.call(thisArg, cell, index, proxy)) {
              return cell
            }
          }
          return undefined
        }
      }
      if (property === 'slice') {
        return (start?: number, end?: number) => {
          const from = normalizeSliceIndex(start ?? 0, cellCount)
          const to = normalizeSliceIndex(end ?? cellCount, cellCount)
          const output: WorkbookSheetCell[] = []
          for (let index = from; index < to; index += 1) {
            const cell = materialize(index)
            if (cell) {
              output.push(cell)
            }
          }
          return output
        }
      }
      if (property === 'toJSON' || property === 'toArray') {
        return () => Array.from(iterate())
      }
      if (typeof property === 'string' && isArrayIndexProperty(property)) {
        return materialize(Number(property))
      }
      return Reflect.get(Array.prototype, property)
    },
    has: (_target, property) => property === 'length' || (typeof property === 'string' && isArrayIndexProperty(property)),
    getOwnPropertyDescriptor: (_target, property) => {
      if (property === 'length') {
        return { configurable: true, enumerable: false, value: cellCount }
      }
      if (typeof property === 'string' && isArrayIndexProperty(property)) {
        const value = materialize(Number(property))
        return value === undefined ? undefined : { configurable: true, enumerable: true, value }
      }
      return undefined
    },
  })
  return proxy
}

function isArrayIndexProperty(property: string): boolean {
  if (property.length === 0 || !/^(?:0|[1-9][0-9]*)$/u.test(property)) {
    return false
  }
  const index = Number(property)
  return Number.isSafeInteger(index) && index >= 0 && index < 2 ** 32 - 1
}

function normalizeSliceIndex(index: number, length: number): number {
  const integer = Math.trunc(index)
  if (integer < 0) {
    return Math.max(length + integer, 0)
  }
  return Math.min(integer, length)
}
