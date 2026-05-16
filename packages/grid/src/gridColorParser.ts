export interface GridColorChannels {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

const FALLBACK_COLOR: GridColorChannels = Object.freeze({ r: 0, g: 0, b: 0, a: 1 })

export function parseGridCssColor(input: string | undefined): GridColorChannels {
  if (!input) {
    return FALLBACK_COLOR
  }

  const color = input.trim()
  if (color.length === 0) {
    return FALLBACK_COLOR
  }
  if (color.toLowerCase() === 'transparent') {
    return { r: 0, g: 0, b: 0, a: 0 }
  }
  if (color.startsWith('#')) {
    return parseHexColor(color.slice(1)) ?? FALLBACK_COLOR
  }
  return parseRgbColor(color) ?? FALLBACK_COLOR
}

export function gridColorToTuple(color: GridColorChannels): readonly [number, number, number, number] {
  return [color.r, color.g, color.b, color.a]
}

function parseHexColor(hex: string): GridColorChannels | null {
  if (!/^[0-9a-fA-F]+$/.test(hex)) {
    return null
  }

  switch (hex.length) {
    case 3:
      return {
        r: hexChannel(`${hex[0]!}${hex[0]!}`),
        g: hexChannel(`${hex[1]!}${hex[1]!}`),
        b: hexChannel(`${hex[2]!}${hex[2]!}`),
        a: 1,
      }
    case 4:
      return {
        r: hexChannel(`${hex[0]!}${hex[0]!}`),
        g: hexChannel(`${hex[1]!}${hex[1]!}`),
        b: hexChannel(`${hex[2]!}${hex[2]!}`),
        a: hexChannel(`${hex[3]!}${hex[3]!}`),
      }
    case 6:
      return {
        r: hexChannel(hex.slice(0, 2)),
        g: hexChannel(hex.slice(2, 4)),
        b: hexChannel(hex.slice(4, 6)),
        a: 1,
      }
    case 8:
      return {
        r: hexChannel(hex.slice(0, 2)),
        g: hexChannel(hex.slice(2, 4)),
        b: hexChannel(hex.slice(4, 6)),
        a: hexChannel(hex.slice(6, 8)),
      }
    default:
      return null
  }
}

function parseRgbColor(color: string): GridColorChannels | null {
  const match = color.match(/^rgba?\(([^)]+)\)$/i)
  if (!match) {
    return null
  }

  const parts = (match[1] ?? '').split(',').map((part) => part.trim())
  if (parts.length < 3 || parts.length > 4 || parts.some((part) => part.length === 0)) {
    return null
  }

  const [r, g, b, a = '1'] = parts
  return {
    r: parseRgbChannel(r!),
    g: parseRgbChannel(g!),
    b: parseRgbChannel(b!),
    a: parseAlphaChannel(a),
  }
}

function hexChannel(value: string): number {
  return clampColorChannel(Number.parseInt(value, 16) / 255)
}

function parseRgbChannel(value: string): number {
  return clampColorChannel(Number.parseFloat(value) / 255)
}

function parseAlphaChannel(value: string): number {
  return clampColorChannel(Number.parseFloat(value))
}

function clampColorChannel(value: number): number {
  if (!Number.isFinite(value)) {
    return 0
  }
  return Math.min(1, Math.max(0, value))
}
