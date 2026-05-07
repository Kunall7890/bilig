import {
  isWorkPaperAxisIntervalEditPossible,
  isWorkPaperAxisOrderPossible,
  isWorkPaperAxisSwapPossible,
  isWorkPaperMoveAxisPossible,
  isWorkPaperMoveCellsPossible,
  isWorkPaperSetCellContentsPossible,
  isWorkPaperSheetContentReplaceable,
  isWorkPaperSheetNameAvailable,
} from './work-paper-capability-checks.js'
import type {
  RawCellContent,
  WorkPaperAddressLike,
  WorkPaperAxisInterval,
  WorkPaperAxisSwapMapping,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperSheet,
} from './work-paper-types.js'

type WorkPaperCapabilityContext = Parameters<typeof isWorkPaperSetCellContentsPossible>[0]

export interface WorkPaperCapabilityOperationsRuntime {
  readonly assertNotDisposed: () => void
  readonly getCapabilityContext: () => WorkPaperCapabilityContext
  readonly doesSheetIdExist: (sheetId: number) => boolean
  readonly validateNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => void
  readonly hasNamedExpression: (expressionName: string, scope?: number) => boolean
}

export interface WorkPaperCapabilityOperations {
  readonly isItPossibleToSetCellContents: (addressOrRange: WorkPaperAddressLike, content?: RawCellContent | WorkPaperSheet) => boolean
  readonly isItPossibleToSwapRowIndexes: (
    sheetId: number,
    rowAOrMappings: number | readonly WorkPaperAxisSwapMapping[],
    rowB?: number,
  ) => boolean
  readonly isItPossibleToSetRowOrder: (sheetId: number, rowOrder: readonly number[]) => boolean
  readonly isItPossibleToSwapColumnIndexes: (
    sheetId: number,
    columnAOrMappings: number | readonly WorkPaperAxisSwapMapping[],
    columnB?: number,
  ) => boolean
  readonly isItPossibleToSetColumnOrder: (sheetId: number, columnOrder: readonly number[]) => boolean
  readonly isItPossibleToAddRows: (
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    restIntervals?: readonly WorkPaperAxisInterval[],
  ) => boolean
  readonly isItPossibleToRemoveRows: (
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    restIntervals?: readonly WorkPaperAxisInterval[],
  ) => boolean
  readonly isItPossibleToAddColumns: (
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    restIntervals?: readonly WorkPaperAxisInterval[],
  ) => boolean
  readonly isItPossibleToRemoveColumns: (
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    restIntervals?: readonly WorkPaperAxisInterval[],
  ) => boolean
  readonly isItPossibleToMoveCells: (source: WorkPaperCellRange, target: WorkPaperCellAddress) => boolean
  readonly isItPossibleToMoveRows: (sheetId: number, start: number, count: number, target: number) => boolean
  readonly isItPossibleToMoveColumns: (sheetId: number, start: number, count: number, target: number) => boolean
  readonly isItPossibleToAddSheet: (sheetName: string) => boolean
  readonly isItPossibleToRemoveSheet: (sheetId: number) => boolean
  readonly isItPossibleToClearSheet: (sheetId: number) => boolean
  readonly isItPossibleToReplaceSheetContent: (sheetId: number, content: WorkPaperSheet) => boolean
  readonly isItPossibleToRenameSheet: (sheetId: number, nextName: string) => boolean
  readonly isItPossibleToAddNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => boolean
  readonly isItPossibleToChangeNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => boolean
  readonly isItPossibleToRemoveNamedExpression: (expressionName: string, scope?: number) => boolean
}

export function createWorkPaperCapabilityOperations(runtime: WorkPaperCapabilityOperationsRuntime): WorkPaperCapabilityOperations {
  const axisIntervalPossible = (
    axis: 'row' | 'column',
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval: number | WorkPaperAxisInterval | undefined,
    restIntervals: readonly WorkPaperAxisInterval[] = [],
  ): boolean =>
    isWorkPaperAxisIntervalEditPossible(runtime.getCapabilityContext(), axis, sheetId, startOrInterval, countOrInterval, restIntervals)

  return {
    isItPossibleToSetCellContents(addressOrRange, content) {
      runtime.assertNotDisposed()
      return isWorkPaperSetCellContentsPossible(runtime.getCapabilityContext(), addressOrRange, content)
    },

    isItPossibleToSwapRowIndexes(sheetId, rowAOrMappings, rowB) {
      return isWorkPaperAxisSwapPossible(runtime.getCapabilityContext(), 'row', sheetId, rowAOrMappings, rowB)
    },

    isItPossibleToSetRowOrder(sheetId, rowOrder) {
      return isWorkPaperAxisOrderPossible(runtime.getCapabilityContext(), 'row', sheetId, rowOrder)
    },

    isItPossibleToSwapColumnIndexes(sheetId, columnAOrMappings, columnB) {
      return isWorkPaperAxisSwapPossible(runtime.getCapabilityContext(), 'column', sheetId, columnAOrMappings, columnB)
    },

    isItPossibleToSetColumnOrder(sheetId, columnOrder) {
      return isWorkPaperAxisOrderPossible(runtime.getCapabilityContext(), 'column', sheetId, columnOrder)
    },

    isItPossibleToAddRows(sheetId, startOrInterval, countOrInterval, restIntervals) {
      return axisIntervalPossible('row', sheetId, startOrInterval, countOrInterval, restIntervals)
    },

    isItPossibleToRemoveRows(sheetId, startOrInterval, countOrInterval, restIntervals) {
      return axisIntervalPossible('row', sheetId, startOrInterval, countOrInterval, restIntervals)
    },

    isItPossibleToAddColumns(sheetId, startOrInterval, countOrInterval, restIntervals) {
      return axisIntervalPossible('column', sheetId, startOrInterval, countOrInterval, restIntervals)
    },

    isItPossibleToRemoveColumns(sheetId, startOrInterval, countOrInterval, restIntervals) {
      return axisIntervalPossible('column', sheetId, startOrInterval, countOrInterval, restIntervals)
    },

    isItPossibleToMoveCells(source, target) {
      return isWorkPaperMoveCellsPossible(runtime.getCapabilityContext(), source, target)
    },

    isItPossibleToMoveRows(sheetId, start, count, target) {
      return isWorkPaperMoveAxisPossible(runtime.getCapabilityContext(), 'row', sheetId, start, count, target)
    },

    isItPossibleToMoveColumns(sheetId, start, count, target) {
      return isWorkPaperMoveAxisPossible(runtime.getCapabilityContext(), 'column', sheetId, start, count, target)
    },

    isItPossibleToAddSheet(sheetName) {
      return isWorkPaperSheetNameAvailable(runtime.getCapabilityContext(), sheetName)
    },

    isItPossibleToRemoveSheet(sheetId) {
      return runtime.doesSheetIdExist(sheetId)
    },

    isItPossibleToClearSheet(sheetId) {
      return runtime.doesSheetIdExist(sheetId)
    },

    isItPossibleToReplaceSheetContent(sheetId, content) {
      return isWorkPaperSheetContentReplaceable(runtime.getCapabilityContext(), sheetId, content)
    },

    isItPossibleToRenameSheet(sheetId, nextName) {
      return isWorkPaperSheetNameAvailable(runtime.getCapabilityContext(), nextName, sheetId)
    },

    isItPossibleToAddNamedExpression(expressionName, expression, scope) {
      runtime.validateNamedExpression(expressionName, expression, scope)
      return !runtime.hasNamedExpression(expressionName, scope)
    },

    isItPossibleToChangeNamedExpression(expressionName, expression, scope) {
      runtime.validateNamedExpression(expressionName, expression, scope)
      return runtime.hasNamedExpression(expressionName, scope)
    },

    isItPossibleToRemoveNamedExpression(expressionName, scope) {
      return runtime.hasNamedExpression(expressionName, scope)
    },
  }
}
