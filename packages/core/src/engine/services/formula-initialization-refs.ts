import type { CompiledFormula } from '@bilig/formula'
import type { U32 } from '../runtime-state.js'

export type InitialFormulaCellIndexList = readonly number[] | U32
export type InitialFormulaEntryRefSource<Entry> = readonly Entry[] | { readonly length: number; readonly at: (index: number) => Entry }

export interface InitialResolvedFormulaEntry {
  cellIndex: number
  sheetId: number
  row: number
  col: number
  ownerSheetName: string
  source: string
  compiled: CompiledFormula
  templateId?: number
}

export function initialFormulaEntryRefAt<Entry>(refs: InitialFormulaEntryRefSource<Entry>, index: number): Entry {
  return Array.isArray(refs) ? refs[index]! : refs.at(index)!
}
