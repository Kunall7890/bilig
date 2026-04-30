import type { Rectangle } from './gridTypes.js'

export interface VisibleRegionState {
  range: Rectangle
  tx: number
  ty: number
  freezeRows?: number
  freezeCols?: number
}

export type HeaderSelection = { kind: 'column'; index: number } | { kind: 'row'; index: number }
