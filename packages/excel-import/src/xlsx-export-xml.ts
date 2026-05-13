import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

export const customNumberFormatStartId = 164

export function escapeXmlAttribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function xmlPrefixForTagName(tagName: string): string {
  const separator = tagName.indexOf(':')
  return separator >= 0 ? tagName.slice(0, separator + 1) : ''
}

function xmlTagName(prefix: string, localName: string): string {
  return `${prefix}${localName}`
}

export function addCustomNumberFormatsToStylesXml(stylesXml: string, formatIdsByCode: ReadonlyMap<string, number>): string {
  if (formatIdsByCode.size === 0) {
    return stylesXml
  }
  const buildNumFmtEntries = (prefix: string) =>
    [...formatIdsByCode.entries()]
      .map(
        ([formatCode, id]) => `<${xmlTagName(prefix, 'numFmt')} numFmtId="${String(id)}" formatCode="${escapeXmlAttribute(formatCode)}"/>`,
      )
      .join('')
  const selfClosingNumFmts = /<((?:[A-Za-z_][\w.-]*:)?numFmts)\b[^>]*\/>/u
  const selfClosingMatch = selfClosingNumFmts.exec(stylesXml)
  if (selfClosingMatch) {
    const tagName = selfClosingMatch[1]!
    const prefix = xmlPrefixForTagName(tagName)
    return stylesXml.replace(
      selfClosingNumFmts,
      () => `<${tagName} count="${String(formatIdsByCode.size)}">${buildNumFmtEntries(prefix)}</${tagName}>`,
    )
  }
  const existingNumFmts = /<((?:[A-Za-z_][\w.-]*:)?numFmts)\b[^>]*>/u.exec(stylesXml)
  if (existingNumFmts) {
    const openingTag = existingNumFmts[0]
    const tagName = existingNumFmts[1]!
    const prefix = xmlPrefixForTagName(tagName)
    const count = readXmlNumberAttribute(openingTag, 'count') ?? 0
    const nextCount = Number.isFinite(count) ? count + formatIdsByCode.size : formatIdsByCode.size
    const closingTag = `</${tagName}>`
    return stylesXml
      .replace(openingTag, () => setXmlAttribute(openingTag, 'count', String(nextCount)))
      .replace(closingTag, () => `${buildNumFmtEntries(prefix)}${closingTag}`)
  }
  const fontsMatch = /<((?:[A-Za-z_][\w.-]*:)?fonts)\b/u.exec(stylesXml)
  const prefix = fontsMatch ? xmlPrefixForTagName(fontsMatch[1]!) : ''
  const numFmtsTagName = xmlTagName(prefix, 'numFmts')
  const numFmtsXml = `<${numFmtsTagName} count="${String(formatIdsByCode.size)}">${buildNumFmtEntries(prefix)}</${numFmtsTagName}>`
  return fontsMatch ? stylesXml.replace(fontsMatch[0], (match) => `${numFmtsXml}${match}`) : stylesXml
}

export function repairLeadingZeroNumberFormatIds(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const styles = zip['xl/styles.xml']
  if (!styles) {
    return bytes
  }
  let stylesXml = strFromU8(styles)
  const leadingZeroFormatCodes = [...new Set([...stylesXml.matchAll(/\bnumFmtId="(0[0-9]+)"/gu)].map((match) => match[1]!))]
  if (leadingZeroFormatCodes.length === 0) {
    return bytes
  }
  const usedIds = new Set([...stylesXml.matchAll(/\bnumFmtId="([0-9]+)"/gu)].map((match) => Number(match[1])))
  const formatIdsByCode = new Map<string, number>()
  let nextId = customNumberFormatStartId
  for (const formatCode of leadingZeroFormatCodes) {
    while (usedIds.has(nextId)) {
      nextId += 1
    }
    formatIdsByCode.set(formatCode, nextId)
    usedIds.add(nextId)
  }
  for (const [formatCode, id] of formatIdsByCode.entries()) {
    stylesXml = stylesXml.replaceAll(`numFmtId="${formatCode}"`, `numFmtId="${String(id)}"`)
  }
  const customIds = [...formatIdsByCode.values()].map(String).join('|')
  const xfWithCustomNumberFormatPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?xf\\b[^>]*\\bnumFmtId="(${customIds})"[^>]*(?:/?>)`, 'gu')
  stylesXml = stylesXml.replace(xfWithCustomNumberFormatPattern, (tag: string) =>
    tag.includes('applyNumberFormat=') ? tag : setXmlAttribute(tag, 'applyNumberFormat', '1'),
  )
  stylesXml = addCustomNumberFormatsToStylesXml(stylesXml, formatIdsByCode)
  zip['xl/styles.xml'] = strToU8(stylesXml)
  return zipSync(zip)
}

export function getZipText(zip: Record<string, Uint8Array>, path: string): string | null {
  const file = zip[path]
  return file ? strFromU8(file) : null
}

export function setZipText(zip: Record<string, Uint8Array>, path: string, text: string): void {
  zip[path] = strToU8(text)
}

export function setXmlAttribute(tag: string, name: string, value: string): string {
  const attribute = `${name}="${escapeXmlAttribute(value)}"`
  const existingAttribute = new RegExp(`\\s${name}="[^"]*"`, 'u')
  if (existingAttribute.test(tag)) {
    return tag.replace(existingAttribute, ` ${attribute}`)
  }
  return tag.replace(/\/?>$/u, (ending) => ` ${attribute}${ending}`)
}

export function readXmlNumberAttribute(tag: string, name: string): number | null {
  const match = new RegExp(`\\s${name}="([0-9]+)"`, 'u').exec(tag)
  if (!match) {
    return null
  }
  const value = Number(match[1])
  return Number.isSafeInteger(value) && value >= 0 ? value : null
}
