import { toAppleScriptString, toAppleScriptValue } from './macos-excel-oracle-applescript-helpers.js'
import type {
  MacosExcelSortHeader,
  MacosExcelSortOrder,
  MacosExcelSortOrientation,
  MacosExcelStructuralOperation,
} from './macos-excel-oracle-types.js'

export function structuralOperationAppleScript(operation: MacosExcelStructuralOperation): string {
  switch (operation.kind) {
    case 'insertRows':
      return `insert into range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift down`
    case 'insertColumns':
      return `insert into range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift to right`
    case 'deleteRows':
      return `delete range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift up`
    case 'deleteColumns':
      return `delete range (range ${toAppleScriptString(operation.range)} of targetWorksheet) shift shift to left`
    case 'setCellValue':
      return `set value of range ${toAppleScriptString(operation.address)} of targetWorksheet to ${toAppleScriptValue(operation.value)}`
    case 'clearCell':
      return `clear contents range ${toAppleScriptString(operation.address)} of targetWorksheet`
    case 'createSheet':
      return [
        'set createdWorksheet to make new worksheet at after worksheet (count of worksheets of targetWorkbook) of targetWorkbook',
        `set name of createdWorksheet to ${toAppleScriptString(operation.name)}`,
      ].join('\n      ')
    case 'renameSheet':
      return `set name of targetWorksheet to ${toAppleScriptString(operation.newName)}`
    case 'deleteSheet':
      return [
        `if exists worksheet ${toAppleScriptString(operation.name)} of targetWorkbook then`,
        `  delete worksheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        `else if exists chart sheet ${toAppleScriptString(operation.name)} of targetWorkbook then`,
        `  delete chart sheet ${toAppleScriptString(operation.name)} of targetWorkbook`,
        'else',
        `  error ${toAppleScriptString(`Sheet not found: ${operation.name}`)} number -1728`,
        'end if',
      ].join('\n      ')
    case 'moveSheet':
      return moveSheetAppleScript(operation)
    case 'moveRows':
      return [
        `cut range (range ${toAppleScriptString(operation.sourceRange)} of targetWorksheet)`,
        `insert into range (range ${toAppleScriptString(operation.destinationRange)} of targetWorksheet) shift shift down`,
      ].join('\n      ')
    case 'moveColumns':
      return [
        `cut range (range ${toAppleScriptString(operation.sourceRange)} of targetWorksheet)`,
        `insert into range (range ${toAppleScriptString(operation.destinationRange)} of targetWorksheet) shift shift to right`,
      ].join('\n      ')
    case 'createDataTable':
      if (!operation.rowInput && !operation.columnInput) {
        throw new Error('macOS Excel data table operation requires a row input, column input, or both')
      }
      return [
        `data table (range ${toAppleScriptString(operation.range)} of targetWorksheet)`,
        operation.rowInput ? `row input (range ${toAppleScriptString(operation.rowInput)} of targetWorksheet)` : '',
        operation.columnInput ? `column input (range ${toAppleScriptString(operation.columnInput)} of targetWorksheet)` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
    case 'deleteTable':
      return `delete list object ${toAppleScriptString(operation.tableName)} of targetWorksheet`
    case 'applySort':
      if (operation.keys.length === 0 || operation.keys.length > 3) {
        throw new Error('macOS Excel sort operation requires one to three sort keys')
      }
      return [
        `sort (range ${toAppleScriptString(operation.range)} of targetWorksheet)`,
        ...operation.keys.flatMap((key, index) => {
          const position = String(index + 1)
          return [
            `key${position} (range ${toAppleScriptString(key.key)} of targetWorksheet)`,
            `order${position} ${sortOrderAppleScript(key.order ?? 'ascending')}`,
          ]
        }),
        operation.header ? `header ${sortHeaderAppleScript(operation.header)}` : '',
        operation.orientation ? `orientation ${sortOrientationAppleScript(operation.orientation)}` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
    case 'applyTableSort':
      if (operation.keys.length === 0 || operation.keys.length > 3) {
        throw new Error('macOS Excel table sort operation requires one to three sort keys')
      }
      return [
        `set tableSort to sort object of list object ${toAppleScriptString(operation.tableName)} of targetWorksheet`,
        `clear sortfieldset (sortfieldset of tableSort)`,
        ...operation.keys.map((key) =>
          [
            `add sortfield (sortfieldset of tableSort)`,
            `key (range ${toAppleScriptString(key.key)} of targetWorksheet)`,
            `order ${sortOrderAppleScript(key.order ?? 'ascending')}`,
          ].join(' '),
        ),
        operation.header ? `set sort header of tableSort to ${sortHeaderAppleScript(operation.header)}` : '',
        operation.orientation ? `set sort orientation of tableSort to ${sortOrientationAppleScript(operation.orientation)}` : '',
        `apply sort tableSort`,
      ]
        .filter((part) => part.length > 0)
        .join('\n      ')
    case 'applyTableAutoFilter':
      if (!Number.isSafeInteger(operation.field) || operation.field <= 0) {
        throw new Error('macOS Excel table AutoFilter operation requires a positive one-based field')
      }
      return [
        `autofilter range (range object of autofilter object of list object ${toAppleScriptString(operation.tableName)} of targetWorksheet)`,
        `field ${String(operation.field)}`,
        operation.criteria1 !== undefined ? `criteria1 ${toAppleScriptValue(operation.criteria1)}` : '',
        operation.operator ? `operator ${operation.operator}` : '',
        operation.criteria2 !== undefined ? `criteria2 ${toAppleScriptValue(operation.criteria2)}` : '',
        operation.visibleDropDown !== undefined ? `visible drop down ${toAppleScriptValue(operation.visibleDropDown)}` : '',
      ]
        .filter((part) => part.length > 0)
        .join(' ')
  }
}

function sortHeaderAppleScript(header: MacosExcelSortHeader): string {
  switch (header) {
    case 'guess':
      return 'header guess'
    case 'no':
      return 'header no'
    case 'yes':
      return 'header yes'
  }
}

function moveSheetAppleScript(operation: Extract<MacosExcelStructuralOperation, { readonly kind: 'moveSheet' }>): string {
  if ((operation.before === undefined) === (operation.after === undefined)) {
    throw new Error('macOS Excel moveSheet operation requires exactly one before or after anchor')
  }
  const anchor = operation.before
    ? `to before worksheet ${toAppleScriptString(operation.before)} of targetWorkbook`
    : `to after worksheet ${toAppleScriptString(operation.after!)} of targetWorkbook`
  return `move worksheet ${toAppleScriptString(operation.name)} of targetWorkbook ${anchor}`
}

function sortOrderAppleScript(order: MacosExcelSortOrder): string {
  switch (order) {
    case 'ascending':
      return 'sort ascending'
    case 'descending':
      return 'sort descending'
  }
}

function sortOrientationAppleScript(orientation: MacosExcelSortOrientation): string {
  switch (orientation) {
    case 'columns':
      return 'sort rows'
    case 'rows':
      return 'sort columns'
  }
}
