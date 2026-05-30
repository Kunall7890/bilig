const ooxmlIndexedColors = [
  '#000000',
  '#ffffff',
  '#ff0000',
  '#00ff00',
  '#0000ff',
  '#ffff00',
  '#ff00ff',
  '#00ffff',
  '#000000',
  '#ffffff',
  '#ff0000',
  '#00ff00',
  '#0000ff',
  '#ffff00',
  '#ff00ff',
  '#00ffff',
  '#800000',
  '#008000',
  '#000080',
  '#808000',
  '#800080',
  '#008080',
  '#c0c0c0',
  '#808080',
  '#9999ff',
  '#993366',
  '#ffffcc',
  '#ccffff',
  '#660066',
  '#ff8080',
  '#0066cc',
  '#ccccff',
  '#000080',
  '#ff00ff',
  '#ffff00',
  '#00ffff',
  '#800080',
  '#800000',
  '#008080',
  '#0000ff',
  '#00ccff',
  '#ccffff',
  '#ccffcc',
  '#ffff99',
  '#99ccff',
  '#ff99cc',
  '#cc99ff',
  '#ffcc99',
  '#3366ff',
  '#33cccc',
  '#99cc00',
  '#ffcc00',
  '#ff9900',
  '#ff6600',
  '#666699',
  '#969696',
  '#003366',
  '#339966',
  '#003300',
  '#333300',
  '#993300',
  '#993366',
  '#333399',
  '#333333',
] as const

export function readOoxmlIndexedColor(indexedValue: string | null, tintValue: string | null): string | null {
  if (indexedValue === null) {
    return null
  }
  const index = Number(indexedValue)
  if (!Number.isInteger(index) || index < 0 || index >= ooxmlIndexedColors.length) {
    return null
  }
  return applyOoxmlTint(ooxmlIndexedColors[index], tintValue)
}

function applyOoxmlTint(hexColor: string, tintValue: string | null): string {
  const tint = tintValue === null ? 0 : Number(tintValue)
  if (!Number.isFinite(tint) || tint === 0) {
    return hexColor
  }
  const hex = hexColor.replace(/^#/u, '')
  const channels = [hex.slice(0, 2), hex.slice(2, 4), hex.slice(4, 6)].map((channel) => Number.parseInt(channel, 16))
  const tintedChannels = channels.map((channel) => {
    const transformed = tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint
    return Math.min(255, Math.max(0, Math.round(transformed)))
      .toString(16)
      .padStart(2, '0')
  })
  return `#${tintedChannels.join('')}`
}
