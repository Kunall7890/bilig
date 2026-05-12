import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'
import { themeContentType, themeRelationshipType } from '../xlsx-theme-artifacts.js'

describe('xlsx workbook theme roundtrip', () => {
  it('preserves imported workbook theme colors and font scheme', () => {
    const source = buildCustomThemeWorkbook()
    const sourceTheme = readThemeSignature(source)

    const imported = importXlsx(source, 'custom-theme.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(sourceTheme.colors.slice(0, customThemeColors.length)).toEqual(customThemeColors)
    expect(sourceTheme.majorFonts).toContain(customThemeFonts.major)
    expect(sourceTheme.minorFonts).toContain(customThemeFonts.minor)
    expect(imported.snapshot.workbook.metadata?.styleArtifacts?.theme?.path).toBe('xl/theme/theme1.xml')
    expect(readThemeSignature(exported)).toEqual(sourceTheme)
    expect(readThemeXml(exported)).toBe(readThemeXml(source))
    expect(readWorkbookThemeRelationship(exported)).toMatchObject({
      target: 'theme/theme1.xml',
      type: themeRelationshipType,
    })
    expect(readContentTypeOverride(exported, '/xl/theme/theme1.xml')).toBe(themeContentType)
  })
})

const customThemeColors = ['123456', 'F0E4D7', '225577', 'AA5500', '667788', '44AA88', 'CC6677', '889944', '3366CC', '7733AA']
const customThemeFonts = {
  major: 'Aptos Display',
  minor: 'Aptos',
}

function buildCustomThemeWorkbook(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  const themeXml = readZipText(zip, 'xl/theme/theme1.xml')
  zip['xl/theme/theme1.xml'] = strToU8(customizeThemeXml(themeXml))
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Custom theme' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'B1', value: 'Value' },
        ],
      },
    ],
  }
}

function customizeThemeXml(themeXml: string): string {
  let colorIndex = 0
  return themeXml
    .replace(/\bname="[^"]*Theme"/u, 'name="Bilig Custom Theme"')
    .replace(/(<a:srgbClr\b[^>]*\bval=")[0-9A-Fa-f]{6}("[^>]*\/>)/gu, (match: string, before: string, after: string) => {
      const color = customThemeColors[colorIndex]
      colorIndex += 1
      return color ? `${before}${color}${after}` : match
    })
    .replace(/(<a:majorFont>[\s\S]*?<a:latin\b[^>]*\btypeface=")[^"]*("[\s\S]*?<\/a:majorFont>)/u, `$1${customThemeFonts.major}$2`)
    .replace(/(<a:minorFont>[\s\S]*?<a:latin\b[^>]*\btypeface=")[^"]*("[\s\S]*?<\/a:minorFont>)/u, `$1${customThemeFonts.minor}$2`)
}

function readThemeSignature(bytes: Uint8Array): { colors: string[]; majorFonts: string[]; minorFonts: string[] } {
  const themeXml = readThemeXml(bytes)
  return {
    colors: [...themeXml.matchAll(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/gu)].map((match) => match[1] ?? ''),
    majorFonts: readFontScheme(themeXml, 'majorFont'),
    minorFonts: readFontScheme(themeXml, 'minorFont'),
  }
}

function readFontScheme(themeXml: string, tagName: 'majorFont' | 'minorFont'): string[] {
  const section = new RegExp(`<a:${tagName}>[\\s\\S]*?<\\/a:${tagName}>`, 'u').exec(themeXml)?.[0] ?? ''
  return [...section.matchAll(/<a:latin\b[^>]*\btypeface="([^"]*)"/gu)].map((match) => match[1] ?? '')
}

function readThemeXml(bytes: Uint8Array): string {
  return readZipText(unzipSync(bytes), 'xl/theme/theme1.xml')
}

function readWorkbookThemeRelationship(bytes: Uint8Array): { target: string; type: string; targetMode?: string } | undefined {
  const relsXml = readZipText(unzipSync(bytes), 'xl/_rels/workbook.xml.rels')
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Type') === themeRelationshipType) {
      return {
        target: readXmlAttribute(attributes, 'Target') ?? '',
        type: readXmlAttribute(attributes, 'Type') ?? '',
        ...(readXmlAttribute(attributes, 'TargetMode') ? { targetMode: readXmlAttribute(attributes, 'TargetMode') ?? undefined } : {}),
      }
    }
  }
  return undefined
}

function readContentTypeOverride(bytes: Uint8Array, partName: string): string | undefined {
  const contentTypesXml = readZipText(unzipSync(bytes), '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function readZipText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}
