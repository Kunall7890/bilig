const RUNTIME_IMAGE_COORD_STRIDE = 1_048_576

export function toFormulaInstanceKey(row: number, col: number): number {
  return row * RUNTIME_IMAGE_COORD_STRIDE + col
}

export function getOrCreateSheetFormulaMap<T>(maps: Map<string, Map<number, T>>, sheetName: string): Map<number, T> {
  let sheetMap = maps.get(sheetName)
  if (!sheetMap) {
    sheetMap = new Map()
    maps.set(sheetName, sheetMap)
  }
  return sheetMap
}
