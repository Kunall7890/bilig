import { describe, expect, it } from 'vitest'

import { parseLargeSimpleWorksheetCellsFromChunks } from '../xlsx-large-simple-worksheet-stream-scanner.js'

const encoder = new TextEncoder()

describe('large simple worksheet WASM scan storage', () => {
  it('stores numeric streamed cells, styles, and formula refs in WASM before arena projection', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(wasmNumericWorksheetXml()), 0, {
      hasSharedStrings: false,
      useWasmScanStorage: true,
    })
    const coordinates: Array<{ row: number; column: number; styleIndex: number }> = []

    scan?.cellScan.styleIndexes.forEach((row, column, styleIndex) => coordinates.push({ row, column, styleIndex }))

    expect(scan?.cellScan.scanStorageKind).toBe('wasm')
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2 },
      { address: 'C1', value: 3, formula: 'A1+B1' },
      { address: 'D1', formula: 'A1' },
    ])
    expect(coordinates).toEqual([
      { row: 0, column: 0, styleIndex: 3 },
      { row: 0, column: 2, styleIndex: 4 },
    ])
  })

  it('flushes WASM records once a streamed sheet needs JS-only value storage', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(wasmFallbackWorksheetXml()), 0, {
      hasSharedStrings: false,
      useWasmScanStorage: true,
    })

    expect(scan?.cellScan.scanStorageKind).toBe('wasm-fallback')
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 'Inline label' },
      { address: 'C1', value: 2 },
    ])
  })

  it('hydrates WASM formula coordinates when a later cell forces JS fallback', () => {
    const scan = parseLargeSimpleWorksheetCellsFromChunks(splitAfterTagOpen(wasmFormulaBeforeFallbackWorksheetXml()), 0, {
      hasSharedStrings: false,
      useWasmScanStorage: true,
    })

    expect(scan?.cellScan.scanStorageKind).toBe('wasm-fallback')
    expect(scan?.cellScan.arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2, formula: 'A1+1' },
      { address: 'C1', value: 'Inline label' },
      { address: 'D1', value: 4, formula: 'A1+3' },
    ])
  })
})

function splitAfterTagOpen(xml: string): (onChunk: (chunk: Uint8Array) => void) => boolean {
  const bytes = encoder.encode(xml)
  return (onChunk) => {
    let start = 0
    for (let index = 0; index < bytes.byteLength; index += 1) {
      if (bytes[index] !== 60) {
        continue
      }
      onChunk(bytes.subarray(start, index + 1))
      start = index + 1
    }
    if (start < bytes.byteLength) {
      onChunk(bytes.subarray(start))
    }
    return true
  }
}

function wasmNumericWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:D1"/>',
    '<sheetData><row r="1">',
    '<c r="A1" s="3"><v>1</v></c>',
    '<c r="B1"><v>2</v></c>',
    '<c r="C1" s="4"><f>A1+B1</f><v>3</v></c>',
    '<c r="D1"><f>A1</f></c>',
    '</row></sheetData>',
    '</worksheet>',
  ].join('')
}

function wasmFallbackWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:C1"/>',
    '<sheetData><row r="1">',
    '<c r="A1"><v>1</v></c>',
    '<c r="B1" t="inlineStr"><is><t>Inline label</t></is></c>',
    '<c r="C1"><v>2</v></c>',
    '</row></sheetData>',
    '</worksheet>',
  ].join('')
}

function wasmFormulaBeforeFallbackWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:D1"/>',
    '<sheetData><row r="1">',
    '<c r="A1"><v>1</v></c>',
    '<c r="B1"><f>A1+1</f><v>2</v></c>',
    '<c r="C1" t="inlineStr"><is><t>Inline label</t></is></c>',
    '<c r="D1"><f>A1+3</f><v>4</v></c>',
    '</row></sheetData>',
    '</worksheet>',
  ].join('')
}
