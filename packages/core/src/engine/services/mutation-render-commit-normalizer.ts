import { parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import type { CommitOp, PreparedCellAddress } from '../runtime-state.js'

export interface NormalizedRenderCommitOps {
  readonly engineOps: EngineOp[]
  readonly potentialNewCells: number
  readonly preparedCellAddressesByOpIndex: Array<PreparedCellAddress | null>
}

export function normalizeRenderCommitOps(ops: readonly CommitOp[]): NormalizedRenderCommitOps {
  const maxEngineOpCount = ops.length * 2
  const engineOps: EngineOp[] = []
  engineOps.length = maxEngineOpCount
  const preparedCellAddressesByOpIndex: Array<PreparedCellAddress | null> = []
  preparedCellAddressesByOpIndex.length = maxEngineOpCount
  let engineOpCount = 0
  let potentialNewCells = 0
  const pushEngineOp = (engineOp: EngineOp, preparedCellAddress: PreparedCellAddress | null = null): void => {
    engineOps[engineOpCount] = engineOp
    preparedCellAddressesByOpIndex[engineOpCount] = preparedCellAddress
    engineOpCount += 1
  }

  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index]
    if (!op) {
      continue
    }
    switch (op.kind) {
      case 'upsertWorkbook':
        if (op.name) {
          pushEngineOp({ kind: 'upsertWorkbook', name: op.name })
        }
        break
      case 'upsertSheet':
        if (op.name) {
          pushEngineOp({ kind: 'upsertSheet', name: op.name, order: op.order ?? 0 })
        }
        break
      case 'renameSheet':
        if (op.oldName && op.newName) {
          pushEngineOp({
            kind: 'renameSheet',
            oldName: op.oldName,
            newName: op.newName,
          })
        }
        break
      case 'deleteSheet':
        if (op.name) {
          pushEngineOp({ kind: 'deleteSheet', name: op.name })
        }
        break
      case 'upsertCell': {
        if (!op.sheetName || !op.addr) {
          break
        }
        const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
        if (op.formula !== undefined) {
          pushEngineOp(
            {
              kind: 'setCellFormula',
              sheetName: op.sheetName,
              address: op.addr,
              formula: op.formula,
            },
            { row: preparedCellAddress.row, col: preparedCellAddress.col },
          )
        } else {
          pushEngineOp(
            {
              kind: 'setCellValue',
              sheetName: op.sheetName,
              address: op.addr,
              value: op.value ?? null,
            },
            { row: preparedCellAddress.row, col: preparedCellAddress.col },
          )
        }
        potentialNewCells += 1
        if (op.format !== undefined) {
          pushEngineOp({
            kind: 'setCellFormat',
            sheetName: op.sheetName,
            address: op.addr,
            format: op.format,
          })
        }
        break
      }
      case 'deleteCell': {
        if (op.sheetName && op.addr) {
          const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
          pushEngineOp(
            {
              kind: 'clearCell',
              sheetName: op.sheetName,
              address: op.addr,
            },
            { row: preparedCellAddress.row, col: preparedCellAddress.col },
          )
          pushEngineOp({
            kind: 'setCellFormat',
            sheetName: op.sheetName,
            address: op.addr,
            format: null,
          })
        }
        break
      }
    }
  }

  engineOps.length = engineOpCount
  preparedCellAddressesByOpIndex.length = engineOpCount
  return { engineOps, potentialNewCells, preparedCellAddressesByOpIndex }
}
