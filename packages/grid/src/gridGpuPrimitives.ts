import { parseGridCssColor, type GridColorChannels } from './gridColorParser.js'

export interface GridGpuColor {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

export interface GridGpuRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: GridGpuColor
}

export interface GridGpuScene {
  readonly fillRects: readonly GridGpuRect[]
  readonly borderRects: readonly GridGpuRect[]
}

export function parseGpuColor(input: string | undefined): GridGpuColor {
  return parseGridCssColor(input) satisfies GridColorChannels
}
