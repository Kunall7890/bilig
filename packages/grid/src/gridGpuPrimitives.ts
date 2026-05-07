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

const FALLBACK_COLOR: GridGpuColor = Object.freeze({ r: 0, g: 0, b: 0, a: 1 })

export function parseGpuColor(input: string | undefined): GridGpuColor {
  if (!input) {
    return FALLBACK_COLOR
  }

  const color = input.trim()
  if (color === 'transparent') {
    return { r: 0, g: 0, b: 0, a: 0 }
  }

  if (color.startsWith('#')) {
    return parseHexGpuColor(color)
  }

  const rgbaMatch = color.match(/^rgba?\(([^)]+)\)$/i)
  if (rgbaMatch) {
    const parts = (rgbaMatch[1] ?? '')
      .split(',')
      .map((part) => part.trim())
      .filter(Boolean)
    const [r = '0', g = '0', b = '0', a = '1'] = parts
    return {
      r: clampColorChannel(Number.parseFloat(r) / 255),
      g: clampColorChannel(Number.parseFloat(g) / 255),
      b: clampColorChannel(Number.parseFloat(b) / 255),
      a: clampColorChannel(Number.parseFloat(a)),
    }
  }

  return FALLBACK_COLOR
}

function parseHexGpuColor(input: string): GridGpuColor {
  const hex = input.slice(1)
  switch (hex.length) {
    case 3:
      return {
        r: hexPairToChannel((hex.slice(0, 1) || '0').repeat(2)),
        g: hexPairToChannel((hex.slice(1, 2) || '0').repeat(2)),
        b: hexPairToChannel((hex.slice(2, 3) || '0').repeat(2)),
        a: 1,
      }
    case 4:
      return {
        r: hexPairToChannel((hex.slice(0, 1) || '0').repeat(2)),
        g: hexPairToChannel((hex.slice(1, 2) || '0').repeat(2)),
        b: hexPairToChannel((hex.slice(2, 3) || '0').repeat(2)),
        a: hexPairToChannel((hex.slice(3, 4) || 'f').repeat(2)),
      }
    case 6:
      return {
        r: hexPairToChannel(hex.slice(0, 2)),
        g: hexPairToChannel(hex.slice(2, 4)),
        b: hexPairToChannel(hex.slice(4, 6)),
        a: 1,
      }
    case 8:
      return {
        r: hexPairToChannel(hex.slice(0, 2)),
        g: hexPairToChannel(hex.slice(2, 4)),
        b: hexPairToChannel(hex.slice(4, 6)),
        a: hexPairToChannel(hex.slice(6, 8)),
      }
    default:
      return FALLBACK_COLOR
  }
}

function hexPairToChannel(value: string): number {
  return clampColorChannel(Number.parseInt(value, 16) / 255)
}

function clampColorChannel(value: number): number {
  if (!Number.isFinite(value)) {
    return 0
  }
  return Math.min(1, Math.max(0, value))
}
