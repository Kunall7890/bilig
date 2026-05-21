export function excelCompatibleTableColumnName(requestedName: string, columnNames: readonly string[], columnIndex: number): string {
  const usedColumnNames = collectUsedTableColumnNames(columnNames, columnIndex)
  const baseName = requestedName.trim()
  if (baseName.length === 0) {
    return nextGeneratedTableColumnName(usedColumnNames)
  }
  if (!usedColumnNames.has(normalizeTableColumnName(baseName))) {
    return baseName
  }
  let suffix = 2
  while (usedColumnNames.has(normalizeTableColumnName(`${baseName}${String(suffix)}`))) {
    suffix += 1
  }
  return `${baseName}${String(suffix)}`
}

export function nextGeneratedTableColumnName(usedColumnNames: ReadonlySet<string>): string {
  let suffix = 1
  while (usedColumnNames.has(normalizeTableColumnName(`Column${String(suffix)}`))) {
    suffix += 1
  }
  return `Column${String(suffix)}`
}

export function normalizeTableColumnName(name: string): string {
  return name.trim().toUpperCase()
}

function collectUsedTableColumnNames(columnNames: readonly string[], excludedColumnIndex: number): Set<string> {
  const usedColumnNames = new Set<string>()
  columnNames.forEach((name, index) => {
    if (index === excludedColumnIndex || name.trim().length === 0) {
      return
    }
    usedColumnNames.add(normalizeTableColumnName(name))
  })
  return usedColumnNames
}
