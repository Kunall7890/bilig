export interface GridColorChannels {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

const FALLBACK_COLOR: GridColorChannels = Object.freeze({ r: 0, g: 0, b: 0, a: 1 })
const NAMED_COLORS: Readonly<Record<string, GridColorChannels>> = Object.freeze({
  black: { r: 0, g: 0, b: 0, a: 1 },
  blue: { r: 0, g: 0, b: 1, a: 1 },
  cyan: { r: 0, g: 1, b: 1, a: 1 },
  gray: { r: 128 / 255, g: 128 / 255, b: 128 / 255, a: 1 },
  green: { r: 0, g: 128 / 255, b: 0, a: 1 },
  grey: { r: 128 / 255, g: 128 / 255, b: 128 / 255, a: 1 },
  lime: { r: 0, g: 1, b: 0, a: 1 },
  magenta: { r: 1, g: 0, b: 1, a: 1 },
  orange: { r: 1, g: 165 / 255, b: 0, a: 1 },
  purple: { r: 128 / 255, g: 0, b: 128 / 255, a: 1 },
  red: { r: 1, g: 0, b: 0, a: 1 },
  white: { r: 1, g: 1, b: 1, a: 1 },
  yellow: { r: 1, g: 1, b: 0, a: 1 },
})

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
  const named = NAMED_COLORS[color.toLowerCase()]
  if (named) {
    return named
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

  const parts = splitRgbFunctionParts(match[1] ?? '')
  if (!parts || parts.length < 3 || parts.length > 4 || parts.some((part) => part.length === 0)) {
    return null
  }

  const [r, g, b, a = '1'] = parts
  const channels = [parseRgbChannel(r!), parseRgbChannel(g!), parseRgbChannel(b!)]
  const alpha = parseAlphaChannel(a)
  if (channels.some((channel) => channel === null) || alpha === null) {
    return null
  }
  return {
    r: channels[0]!,
    g: channels[1]!,
    b: channels[2]!,
    a: alpha,
  }
}

function splitRgbFunctionParts(value: string): readonly string[] | null {
  if (value.includes(',')) {
    const parts = value.split(',').map((part) => part.trim())
    if (parts.length !== 3 && parts.length !== 4) {
      return null
    }
    return parts
  }
  const [channels, alpha] = value.split('/').map((part) => part.trim())
  const channelParts = (channels ?? '').split(/\s+/).filter(Boolean)
  if (channelParts.length !== 3) {
    return null
  }
  return alpha ? [...channelParts, alpha] : channelParts
}

function hexChannel(value: string): number {
  return clampColorChannel(Number.parseInt(value, 16) / 255)
}

function parseRgbChannel(value: string): number | null {
  const trimmed = value.trim()
  if (trimmed.endsWith('%')) {
    return parsePercentageChannel(trimmed)
  }
  if (!isStrictNumericToken(trimmed)) {
    return null
  }
  return clampColorChannel(Number.parseFloat(trimmed) / 255)
}

function parseAlphaChannel(value: string): number | null {
  const trimmed = value.trim()
  if (trimmed.endsWith('%')) {
    return parsePercentageChannel(trimmed)
  }
  if (!isStrictNumericToken(trimmed)) {
    return null
  }
  return clampColorChannel(Number.parseFloat(trimmed))
}

function parsePercentageChannel(value: string): number | null {
  const numeric = value.slice(0, -1).trim()
  if (!isStrictNumericToken(numeric)) {
    return null
  }
  return clampColorChannel(Number.parseFloat(numeric) / 100)
}

function isStrictNumericToken(value: string): boolean {
  return /^[-+]?(?:\d+|\d*\.\d+)(?:e[-+]?\d+)?$/i.test(value)
}

function clampColorChannel(value: number): number {
  if (!Number.isFinite(value)) {
    return 0
  }
  return Math.min(1, Math.max(0, value))
}
