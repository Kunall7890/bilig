import { MAX_COLS, MAX_ROWS, type CellRangeRef, type SheetFormatRangeSnapshot, type SheetStyleRangeSnapshot } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { mapStructuralBoundary } from '../../engine-structural-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

type StructureMetadataRewriteArgs = Pick<CreateEngineStructureServiceArgs, 'state' | 'clearOwnedPivot'>

type MetadataRangeLike = {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

export function rewriteDefinedNamesForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): Set<string> {
  const workbook = args.state.workbook
  const changedNames = new Set<string>()
  workbook.listDefinedNames().forEach((record) => {
    if (typeof record.value === 'string' && record.value.startsWith('=')) {
      const nextFormula = rewriteDefinedNameFormulaOrNull(record.value.slice(1), sheetName, transform)
      if (nextFormula === null) {
        return
      }
      if (`=${nextFormula}` !== record.value) {
        workbook.setDefinedName(record.name, `=${nextFormula}`, record.scopeSheetName)
      }
      return
    }
    if (typeof record.value !== 'object' || !record.value) {
      return
    }
    switch (record.value.kind) {
      case 'formula': {
        const nextFormula = rewriteDefinedNameFormulaOrNull(
          record.value.formula.startsWith('=') ? record.value.formula.slice(1) : record.value.formula,
          sheetName,
          transform,
        )
        if (nextFormula === null) {
          return
        }
        const nextValue = {
          ...record.value,
          formula: record.value.formula.startsWith('=') ? `=${nextFormula}` : nextFormula,
        }
        if (nextValue.formula !== record.value.formula) {
          workbook.setDefinedName(record.name, nextValue, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'cell-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextAddress = rewriteAddressForStructuralTransform(record.value.address, transform)
        if (!nextAddress) {
          workbook.deleteDefinedName(record.name, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextAddress !== record.value.address) {
          workbook.setDefinedName(
            record.name,
            {
              ...record.value,
              address: nextAddress,
            },
            record.scopeSheetName,
          )
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'range-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextRange = rewriteMetadataRangeForStructuralTransform(record.value, transform)
        if (!nextRange) {
          workbook.deleteDefinedName(record.name, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
          workbook.setDefinedName(record.name, nextRange, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'scalar':
      case 'structured-ref':
        return
    }
  })
  return changedNames
}

function rewriteDefinedNameFormulaOrNull(formula: string, sheetName: string, transform: StructuralAxisTransform): string | null {
  try {
    return rewriteFormulaForStructuralTransform(formula, sheetName, sheetName, transform)
  } catch {
    return null
  }
}

export function rewriteWorkbookMetadataForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): { changedTableNames: Set<string> } {
  const workbook = args.state.workbook
  const changedTableNames = new Set<string>()
  workbook.listTables().forEach((table) => {
    if (table.sheetName !== sheetName) {
      return
    }
    const range = rewriteMetadataRangeForStructuralTransform(table, transform)
    if (!range) {
      changedTableNames.add(table.name)
      workbook.deleteTable(table.name)
      return
    }
    changedTableNames.add(table.name)
    workbook.setTable(range)
  })
  const rewrittenMergeRanges: CellRangeRef[] = []
  workbook.listMergeRanges(sheetName).forEach((merge) => {
    const range = rewriteMetadataRangeForStructuralTransform(merge, transform)
    if (!range) {
      return
    }
    rewrittenMergeRanges.push(range)
  })
  workbook.setMergeRanges(sheetName, rewrittenMergeRanges)
  workbook.listFilters(sheetName).forEach((filter) => {
    const range = rewriteMetadataRangeForStructuralTransform(filter.range, transform)
    workbook.deleteFilter(sheetName, filter.range)
    if (range) {
      workbook.setFilter(sheetName, range)
    }
  })
  workbook.listSorts(sheetName).forEach((sort) => {
    const range = rewriteMetadataRangeForStructuralTransform(sort.range, transform)
    workbook.deleteSort(sheetName, sort.range)
    if (!range) {
      return
    }
    workbook.setSort(
      sheetName,
      range,
      sort.keys.map((key) => ({
        ...key,
        keyAddress: rewriteMetadataAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
      })),
    )
  })
  workbook.listDataValidations(sheetName).forEach((validation) => {
    const range = rewriteMetadataRangeForStructuralTransform(validation.range, transform)
    workbook.deleteDataValidation(sheetName, validation.range)
    if (!range) {
      return
    }
    const nextValidation = structuredClone(validation)
    nextValidation.range = range
    if (nextValidation.rule.kind === 'list' && nextValidation.rule.source) {
      switch (nextValidation.rule.source.kind) {
        case 'cell-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextAddress = rewriteMetadataAddressForStructuralTransform(nextValidation.rule.source.address, transform)
          if (!nextAddress) {
            return
          }
          nextValidation.rule.source.address = nextAddress
          break
        }
        case 'range-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextSourceRange = rewriteMetadataRangeForStructuralTransform(nextValidation.rule.source, transform)
          if (!nextSourceRange) {
            return
          }
          nextValidation.rule.source = nextSourceRange
          break
        }
        case 'named-range':
        case 'structured-ref':
          break
      }
    }
    workbook.setDataValidation(nextValidation)
  })
  workbook.listConditionalFormats(sheetName).forEach((format) => {
    const nextFormat = rewriteMetadataRangeRecord(format, transform)
    workbook.deleteConditionalFormat(format.id)
    if (!nextFormat) {
      return
    }
    workbook.setConditionalFormat(nextFormat)
  })
  workbook.listRangeProtections(sheetName).forEach((protection) => {
    const nextProtection = rewriteMetadataRangeRecord(protection, transform)
    workbook.deleteRangeProtection(protection.id)
    if (!nextProtection) {
      return
    }
    workbook.setRangeProtection(nextProtection)
  })
  workbook.listCommentThreads(sheetName).forEach((thread) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(thread.address, transform)
    workbook.deleteCommentThread(sheetName, thread.address)
    if (!nextAddress) {
      return
    }
    workbook.setCommentThread({
      ...thread,
      address: nextAddress,
    })
  })
  workbook.listNotes(sheetName).forEach((note) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(note.address, transform)
    workbook.deleteNote(sheetName, note.address)
    if (!nextAddress) {
      return
    }
    workbook.setNote({
      ...note,
      address: nextAddress,
    })
  })
  const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
  const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
  workbook.listStyleRanges(sheetName).forEach((record) => {
    const nextRecord = rewriteMetadataRangeRecord(record, transform)
    if (nextRecord) {
      rewrittenStyleRanges.push(nextRecord)
    }
  })
  workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
  workbook.listFormatRanges(sheetName).forEach((record) => {
    const nextRecord = rewriteMetadataRangeRecord(record, transform)
    if (nextRecord) {
      rewrittenFormatRanges.push(nextRecord)
    }
  })
  workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
  const freezePane = workbook.getFreezePane(sheetName)
  if (freezePane) {
    const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
    const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
    if (nextRows <= 0 && nextCols <= 0) {
      workbook.clearFreezePane(sheetName)
    } else {
      workbook.setFreezePane(sheetName, nextRows, nextCols)
    }
  }
  workbook.listPivots().forEach((pivot) => {
    const nextAddress =
      pivot.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(pivot.address, transform) : pivot.address
    const nextSource =
      pivot.source.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(pivot.source, transform) : pivot.source
    if (!nextAddress || !nextSource) {
      args.clearOwnedPivot(pivot)
      workbook.deletePivot(pivot.sheetName, pivot.address)
      return
    }
    if (nextAddress !== pivot.address) {
      args.clearOwnedPivot(pivot)
      workbook.deletePivot(pivot.sheetName, pivot.address)
    }
    workbook.setPivot({
      ...pivot,
      address: nextAddress,
      source: nextSource,
    })
  })
  workbook.listCharts().forEach((chart) => {
    const nextAddress =
      chart.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(chart.address, transform) : chart.address
    const nextSource =
      chart.source.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(chart.source, transform) : chart.source
    if (!nextAddress || !nextSource) {
      workbook.deleteChart(chart.id)
      return
    }
    workbook.setChart({
      ...chart,
      address: nextAddress,
      source: nextSource,
    })
  })
  workbook.listImages().forEach((image) => {
    if (image.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(image.address, transform)
    if (!nextAddress) {
      workbook.deleteImage(image.id)
      return
    }
    workbook.setImage({
      ...image,
      address: nextAddress,
    })
  })
  workbook.listShapes().forEach((shape) => {
    if (shape.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(shape.address, transform)
    if (!nextAddress) {
      workbook.deleteShape(shape.id)
      return
    }
    workbook.setShape({
      ...shape,
      address: nextAddress,
    })
  })
  return { changedTableNames }
}

function rewriteMetadataRangeForStructuralTransform<T extends MetadataRangeLike>(
  range: T,
  transform: StructuralAxisTransform,
): T | undefined {
  const rewritten = rewriteRangeForStructuralTransform(range.startAddress, range.endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  const clipped = clipMetadataRangeToSheetGrid(range.sheetName, rewritten.startAddress, rewritten.endAddress)
  return clipped ? withRewrittenMetadataRange(range, clipped) : undefined
}

function withRewrittenMetadataRange<T extends MetadataRangeLike>(range: T, rewritten: CellRangeRef): T {
  return {
    ...range,
    startAddress: rewritten.startAddress,
    endAddress: rewritten.endAddress,
  }
}

function rewriteMetadataRangeRecord<T extends { readonly range: MetadataRangeLike }>(
  record: T,
  transform: StructuralAxisTransform,
): T | undefined {
  const range = rewriteMetadataRangeForStructuralTransform(record.range, transform)
  return range ? { ...record, range } : undefined
}

function rewriteMetadataAddressForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseUnboundedMetadataCellAddress(rewritten)
  if (!parsed) {
    throw new Error('Invalid metadata reference')
  }
  if (parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function clipMetadataRangeToSheetGrid(sheetName: string, startAddress: string, endAddress: string): CellRangeRef | undefined {
  const start = parseUnboundedMetadataCellAddress(startAddress)
  const end = parseUnboundedMetadataCellAddress(endAddress)
  if (!start || !end) {
    throw new Error('Invalid metadata reference')
  }
  const startRow = Math.min(start[0], end[0])
  const endRow = Math.min(MAX_ROWS - 1, Math.max(start[0], end[0]))
  const startCol = Math.min(start[1], end[1])
  const endCol = Math.min(MAX_COLS - 1, Math.max(start[1], end[1]))
  if (startRow > endRow || startCol > endCol) {
    return undefined
  }
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function parseUnboundedMetadataCellAddress(address: string): [number, number] | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
