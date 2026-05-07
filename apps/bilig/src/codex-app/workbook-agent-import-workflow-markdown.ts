import type { CellNumberFormatPreset } from '@bilig/protocol'

export interface NumberFormatRecommendation {
  readonly headerLabel: string
  readonly columnLabel: string
  readonly startAddress: string
  readonly endAddress: string
  readonly preset: CellNumberFormatPreset
  readonly numericCount: number
}

export interface FormulaFillRecommendation {
  readonly columnLabel: string
  readonly sourceAddress: string
  readonly targetStartAddress: string
  readonly targetEndAddress: string
  readonly filledRowCount: number
}

function describeNumberFormatPreset(preset: CellNumberFormatPreset): string {
  switch (preset.kind) {
    case 'general':
      return 'general'
    case 'text':
      return 'text'
    case 'number':
      return `number (${String(preset.decimals ?? 0)} decimals)`
    case 'currency':
      return `currency ${preset.currency ?? 'USD'} (${String(preset.decimals ?? 2)} decimals)`
    case 'accounting':
      return `accounting ${preset.currency ?? 'USD'} (${String(preset.decimals ?? 2)} decimals)`
    case 'date':
      return `date (${preset.dateStyle ?? 'short'})`
    case 'time':
      return `time (${preset.dateStyle ?? 'short'})`
    case 'datetime':
      return `datetime (${preset.dateStyle ?? 'short'})`
    case 'percent':
      return `percent (${String(preset.decimals ?? 2)} decimals)`
  }
}

export function summarizeNumberFormatNormalizationMarkdown(input: {
  readonly sheetName: string
  readonly dataStartAddress: string
  readonly dataEndAddress: string
  readonly recommendations: readonly NumberFormatRecommendation[]
}): string {
  const lines = [
    '## Number Format Normalization Preview',
    '',
    `Sheet: ${input.sheetName}`,
    `Data range: ${input.dataStartAddress}:${input.dataEndAddress}`,
    `Columns staged: ${String(input.recommendations.length)}`,
    '',
  ]
  if (input.recommendations.length === 0) {
    lines.push('No number-format changes were needed on the current sheet.')
    return lines.join('\n')
  }
  lines.push('### Staged number formats')
  for (const recommendation of input.recommendations) {
    lines.push(
      `- ${recommendation.columnLabel} (${recommendation.headerLabel || 'Untitled'}): ${describeNumberFormatPreset(recommendation.preset)} across ${recommendation.startAddress}:${recommendation.endAddress} (${String(recommendation.numericCount)} numeric cell${recommendation.numericCount === 1 ? '' : 's'})`,
    )
  }
  lines.push('', 'The staged workbook change set applies semantic number-format commands through the normal workbook mutation path.')
  return lines.join('\n')
}

export function summarizeHeaderNormalizationMarkdown(input: {
  readonly sheetName: string
  readonly headerStartAddress: string
  readonly headerEndAddress: string
  readonly totalColumns: number
  readonly changes: readonly {
    readonly address: string
    readonly before: string
    readonly after: string
  }[]
}): string {
  const lines = [
    '## Header Normalization Preview',
    '',
    `Sheet: ${input.sheetName}`,
    `Header row: ${input.headerStartAddress}:${input.headerEndAddress}`,
    `Columns inspected: ${String(input.totalColumns)}`,
    `Headers changed: ${String(input.changes.length)}`,
    '',
  ]
  if (input.changes.length === 0) {
    lines.push('No header changes were needed. The current header row is already normalized.')
    return lines.join('\n')
  }
  lines.push('### Changed headers')
  for (const change of input.changes) {
    lines.push(`- ${change.address}: ${change.before} -> ${change.after}`)
  }
  lines.push('', 'The staged workbook change set writes the normalized header row through the normal workbook mutation path.')
  return lines.join('\n')
}

export function summarizeWhitespaceNormalizationMarkdown(input: {
  readonly sheetName: string
  readonly rangeStartAddress: string
  readonly rangeEndAddress: string
  readonly changes: readonly {
    readonly address: string
    readonly before: string
    readonly after: string
  }[]
}): string {
  const lines = [
    '## Whitespace Normalization Preview',
    '',
    `Sheet: ${input.sheetName}`,
    `Inspected range: ${input.rangeStartAddress}:${input.rangeEndAddress}`,
    `Text cells changed: ${String(input.changes.length)}`,
    '',
  ]
  if (input.changes.length === 0) {
    lines.push('No text whitespace changes were needed on the current sheet.')
    return lines.join('\n')
  }
  lines.push('### Normalized cells')
  for (const change of input.changes) {
    lines.push(`- ${change.address}: ${JSON.stringify(change.before)} -> ${JSON.stringify(change.after)}`)
  }
  lines.push('', 'The staged workbook change set writes the normalized text cells through the normal workbook mutation path.')
  return lines.join('\n')
}

export function summarizeFormulaFillMarkdown(input: {
  readonly sheetName: string
  readonly recommendations: readonly FormulaFillRecommendation[]
}): string {
  const lines = [
    '## Formula Fill-Down Preview',
    '',
    `Sheet: ${input.sheetName}`,
    `Formula regions staged: ${String(input.recommendations.length)}`,
    '',
  ]
  if (input.recommendations.length === 0) {
    lines.push('No fill-down changes were needed on the current sheet.')
    return lines.join('\n')
  }
  lines.push('### Filled ranges')
  for (const recommendation of input.recommendations) {
    lines.push(
      `- ${recommendation.columnLabel}: fill ${recommendation.sourceAddress} down through ${recommendation.targetStartAddress}:${recommendation.targetEndAddress} (${String(recommendation.filledRowCount)} row${recommendation.filledRowCount === 1 ? '' : 's'})`,
    )
  }
  lines.push('', 'The staged workbook change set applies semantic fill commands through the normal workbook mutation path.')
  return lines.join('\n')
}
