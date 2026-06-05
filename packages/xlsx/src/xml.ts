export const worksheetCellElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?c)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu

export const worksheetCellOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u

export function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/gu, '&quot;')
}

export function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

export function readXmlAttribute(tag: string, name: string): string | null {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  const doubleQuoted = new RegExp(`(?:^|\\s)${escapedName}="([^"]*)"`, 'u').exec(tag)
  if (doubleQuoted) {
    return decodeXmlAttribute(doubleQuoted[1] ?? '')
  }
  const singleQuoted = new RegExp(`(?:^|\\s)${escapedName}='([^']*)'`, 'u').exec(tag)
  return singleQuoted ? decodeXmlAttribute(singleQuoted[1] ?? '') : null
}

export function setXmlAttribute(tag: string, name: string, value: string): string {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  const attribute = `${name}="${escapeXmlAttribute(value)}"`
  const existingAttribute = new RegExp(`\\s${escapedName}=(["'])[\\s\\S]*?\\1`, 'u')
  if (existingAttribute.test(tag)) {
    return tag.replace(existingAttribute, ` ${attribute}`)
  }
  return tag.replace(/\/?>$/u, (ending) => ` ${attribute}${ending}`)
}
