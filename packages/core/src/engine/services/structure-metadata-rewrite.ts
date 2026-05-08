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

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9][0-9]*)$/i

export function rewriteDefinedNamesForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): Set<string> {
  const changedNames = new Set<string>()
  args.state.workbook.listDefinedNames().forEach((record) => {
    if (typeof record.value === 'string' && record.value.startsWith('=')) {
      const nextFormula = rewriteDefinedNameFormulaOrNull(record.value.slice(1), sheetName, transform)
      if (nextFormula === null) {
        return
      }
      if (`=${nextFormula}` !== record.value) {
        args.state.workbook.setDefinedName(record.name, `=${nextFormula}`, record.scopeSheetName)
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
          args.state.workbook.setDefinedName(record.name, nextValue, record.scopeSheetName)
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
          args.state.workbook.deleteDefinedName(record.name, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextAddress !== record.value.address) {
          args.state.workbook.setDefinedName(
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
          args.state.workbook.deleteDefinedName(record.name, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
          args.state.workbook.setDefinedName(
            record.name,
            {
              ...record.value,
              startAddress: nextRange.startAddress,
              endAddress: nextRange.endAddress,
            },
            record.scopeSheetName,
          )
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
  const changedTableNames = new Set<string>()
  args.state.workbook
    .listTables()
    .filter((table) => table.sheetName === sheetName)
    .forEach((table) => {
      const range = rewriteMetadataRangeForStructuralTransform(table, transform)
      if (!range) {
        changedTableNames.add(table.name)
        args.state.workbook.deleteTable(table.name)
        return
      }
      changedTableNames.add(table.name)
      args.state.workbook.setTable({
        ...table,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      })
    })
  const mergeRanges = args.state.workbook.listMergeRanges(sheetName)
  const rewrittenMergeRanges: CellRangeRef[] = []
  mergeRanges.forEach((merge) => {
    const range = rewriteMetadataRangeForStructuralTransform(merge, transform)
    if (!range) {
      return
    }
    rewrittenMergeRanges.push({
      ...merge,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    })
  })
  args.state.workbook.setMergeRanges(sheetName, rewrittenMergeRanges)
  args.state.workbook.listFilters(sheetName).forEach((filter) => {
    const range = rewriteMetadataRangeForStructuralTransform(filter.range, transform)
    args.state.workbook.deleteFilter(sheetName, filter.range)
    if (range) {
      args.state.workbook.setFilter(sheetName, {
        ...filter.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      })
    }
  })
  args.state.workbook.listSorts(sheetName).forEach((sort) => {
    const range = rewriteMetadataRangeForStructuralTransform(sort.range, transform)
    args.state.workbook.deleteSort(sheetName, sort.range)
    if (!range) {
      return
    }
    args.state.workbook.setSort(
      sheetName,
      { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
      sort.keys.map((key) => ({
        ...key,
        keyAddress: rewriteMetadataAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
      })),
    )
  })
  args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
    const range = rewriteMetadataRangeForStructuralTransform(validation.range, transform)
    args.state.workbook.deleteDataValidation(sheetName, validation.range)
    if (!range) {
      return
    }
    const nextValidation = structuredClone(validation)
    nextValidation.range = {
      ...validation.range,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    }
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
          nextValidation.rule.source.startAddress = nextSourceRange.startAddress
          nextValidation.rule.source.endAddress = nextSourceRange.endAddress
          break
        }
        case 'named-range':
        case 'structured-ref':
          break
      }
    }
    args.state.workbook.setDataValidation(nextValidation)
  })
  args.state.workbook.listConditionalFormats(sheetName).forEach((format) => {
    const range = rewriteMetadataRangeForStructuralTransform(format.range, transform)
    args.state.workbook.deleteConditionalFormat(format.id)
    if (!range) {
      return
    }
    args.state.workbook.setConditionalFormat({
      ...format,
      range: {
        ...format.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
    })
  })
  args.state.workbook.listRangeProtections(sheetName).forEach((protection) => {
    const range = rewriteMetadataRangeForStructuralTransform(protection.range, transform)
    args.state.workbook.deleteRangeProtection(protection.id)
    if (!range) {
      return
    }
    args.state.workbook.setRangeProtection({
      ...protection,
      range: {
        ...protection.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
    })
  })
  args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(thread.address, transform)
    args.state.workbook.deleteCommentThread(sheetName, thread.address)
    if (!nextAddress) {
      return
    }
    args.state.workbook.setCommentThread({
      ...thread,
      address: nextAddress,
    })
  })
  args.state.workbook.listNotes(sheetName).forEach((note) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(note.address, transform)
    args.state.workbook.deleteNote(sheetName, note.address)
    if (!nextAddress) {
      return
    }
    args.state.workbook.setNote({
      ...note,
      address: nextAddress,
    })
  })
  const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
  const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
  args.state.workbook.listStyleRanges(sheetName).forEach((record) => {
    const range = rewriteMetadataRangeForStructuralTransform(record.range, transform)
    if (!range) {
      return
    }
    rewrittenStyleRanges.push({
      range: {
        ...record.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
      styleId: record.styleId,
    })
  })
  args.state.workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
  args.state.workbook.listFormatRanges(sheetName).forEach((record) => {
    const range = rewriteMetadataRangeForStructuralTransform(record.range, transform)
    if (!range) {
      return
    }
    rewrittenFormatRanges.push({
      range: {
        ...record.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
      formatId: record.formatId,
    })
  })
  args.state.workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
  const freezePane = args.state.workbook.getFreezePane(sheetName)
  if (freezePane) {
    const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
    const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
    if (nextRows <= 0 && nextCols <= 0) {
      args.state.workbook.clearFreezePane(sheetName)
    } else {
      args.state.workbook.setFreezePane(sheetName, nextRows, nextCols)
    }
  }
  args.state.workbook.listPivots().forEach((pivot) => {
    const nextAddress =
      pivot.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(pivot.address, transform) : pivot.address
    const nextSource =
      pivot.source.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(pivot.source, transform) : pivot.source
    if (!nextAddress || !nextSource) {
      args.clearOwnedPivot(pivot)
      args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
      return
    }
    if (nextAddress !== pivot.address) {
      args.clearOwnedPivot(pivot)
      args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
    }
    args.state.workbook.setPivot({
      ...pivot,
      address: nextAddress,
      source: {
        ...pivot.source,
        startAddress: nextSource.startAddress,
        endAddress: nextSource.endAddress,
      },
    })
  })
  args.state.workbook.listCharts().forEach((chart) => {
    const nextAddress =
      chart.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(chart.address, transform) : chart.address
    const nextSource =
      chart.source.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(chart.source, transform) : chart.source
    if (!nextAddress || !nextSource) {
      args.state.workbook.deleteChart(chart.id)
      return
    }
    args.state.workbook.setChart({
      ...chart,
      address: nextAddress,
      source: {
        ...chart.source,
        startAddress: nextSource.startAddress,
        endAddress: nextSource.endAddress,
      },
    })
  })
  args.state.workbook.listImages().forEach((image) => {
    if (image.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(image.address, transform)
    if (!nextAddress) {
      args.state.workbook.deleteImage(image.id)
      return
    }
    args.state.workbook.setImage({
      ...image,
      address: nextAddress,
    })
  })
  args.state.workbook.listShapes().forEach((shape) => {
    if (shape.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(shape.address, transform)
    if (!nextAddress) {
      args.state.workbook.deleteShape(shape.id)
      return
    }
    args.state.workbook.setShape({
      ...shape,
      address: nextAddress,
    })
  })
  return { changedTableNames }
}

function rewriteMetadataRangeForStructuralTransform(
  range: MetadataRangeLike,
  transform: StructuralAxisTransform,
): CellRangeRef | undefined {
  const rewritten = rewriteRangeForStructuralTransform(range.startAddress, range.endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  return clipMetadataRangeToSheetGrid({
    sheetName: range.sheetName,
    startAddress: rewritten.startAddress,
    endAddress: rewritten.endAddress,
  })
}

function rewriteMetadataAddressForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseUnboundedMetadataCellAddress(rewritten)
  if (!parsed) {
    throw new Error(`Invalid metadata cell reference '${rewritten}'`)
  }
  if (parsed.row < 0 || parsed.row >= MAX_ROWS || parsed.col < 0 || parsed.col >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed.row, parsed.col)
}

function clipMetadataRangeToSheetGrid(range: MetadataRangeLike): CellRangeRef | undefined {
  const start = parseUnboundedMetadataCellAddress(range.startAddress)
  const end = parseUnboundedMetadataCellAddress(range.endAddress)
  if (!start || !end) {
    throw new Error(`Invalid metadata range reference '${range.startAddress}:${range.endAddress}'`)
  }
  const startRow = Math.max(0, Math.min(start.row, end.row))
  const endRow = Math.min(MAX_ROWS - 1, Math.max(start.row, end.row))
  const startCol = Math.max(0, Math.min(start.col, end.col))
  const endCol = Math.min(MAX_COLS - 1, Math.max(start.col, end.col))
  if (startRow > endRow || startCol > endCol) {
    return undefined
  }
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function parseUnboundedMetadataCellAddress(address: string): { row: number; col: number } | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  if (!match) {
    return undefined
  }
  return {
    col: columnToIndex(match[1]!.toUpperCase()),
    row: Number.parseInt(match[2]!, 10) - 1,
  }
}
