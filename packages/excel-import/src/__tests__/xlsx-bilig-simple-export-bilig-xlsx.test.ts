import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const repoRoot = fileURLToPath(new URL('../../../..', import.meta.url))

describe('@bilig/xlsx simple export boundary', () => {
  it('exports bordered generated snapshots without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const snapshot = {
      version: 1,
      workbook: {
        name: 'bilig-xlsx-bordered-export',
        metadata: {
          styles: [
            {
              id: 'total-border',
              font: { bold: true },
              borders: {
                bottom: { style: 'solid', weight: 'thin', color: '#000000' },
              },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Report',
          order: 0,
          cells: [
            { address: 'A1', value: 'Total' },
            { address: 'B1', formula: 'SUM(B2:B3)', value: 30 },
          ],
          metadata: {
            styleRanges: [
              {
                range: { sheetName: 'Report', startAddress: 'A1', endAddress: 'B1' },
                styleId: 'total-border',
              },
            ],
          },
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'])
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({ loaded, bytes: exported.byteLength, stylesXml, sheetXml }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isBiligXlsxExportOutput(output)).toBe(true)
    if (!isBiligXlsxExportOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.stylesXml).toContain('<bottom style="thin"><color rgb="FF000000"/></bottom>')
    expect(output.stylesXml).toContain('applyBorder="1"')
    expect(output.sheetXml).toContain('<c r="A1" s="1"')
    expect(output.sheetXml).toContain('<c r="B1" s="1"')
  }, 15_000)

  it('exports sparse style artifacts with @bilig/xlsx without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const minimalRawStylesXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>',
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
      '</styleSheet>',
    ].join('')
    const snapshot = {
      version: 1,
      workbook: {
        name: 'bilig-xlsx-sparse-style-artifacts',
        metadata: { styleArtifacts: { stylesXml: minimalRawStylesXml } },
      },
      sheets: [
        {
          id: 1,
          name: 'Sparse',
          order: 0,
          cells: [{ address: 'A1', value: 'Header' }],
          metadata: {
            styleArtifacts: {
              cellStyleIndexes: Array.from({ length: 65_000 }, (_entry, index) => ({
                address: 'CF' + String(index + 1),
                styleIndex: 1,
              })),
            },
          },
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'])
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      bytes: exported.byteLength,
      hasRawStyles: stylesXml === minimalRawStylesXml,
      hasTailStyle: sheetXml.includes('<c r="CF65000" s="1"/>'),
      hasDimension: sheetXml.includes('<dimension ref="A1:CF65000"/>'),
      hasSharedStrings: zip['xl/sharedStrings.xml'] !== undefined,
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isSparseStyleArtifactOutput(output)).toBe(true)
    if (!isSparseStyleArtifactOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.hasRawStyles).toBe(true)
    expect(output.hasTailStyle).toBe(true)
    expect(output.hasDimension).toBe(true)
    expect(output.hasSharedStrings).toBe(false)
  }, 15_000)
})

function isBiligXlsxExportOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly stylesXml: string
  readonly sheetXml: string
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['stylesXml'] === 'string' &&
    typeof value['sheetXml'] === 'string'
  )
}

function isSparseStyleArtifactOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly hasRawStyles: boolean
  readonly hasTailStyle: boolean
  readonly hasDimension: boolean
  readonly hasSharedStrings: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['hasRawStyles'] === 'boolean' &&
    typeof value['hasTailStyle'] === 'boolean' &&
    typeof value['hasDimension'] === 'boolean' &&
    typeof value['hasSharedStrings'] === 'boolean'
  )
}

function isRecord(value: unknown): value is Record<PropertyKey, unknown> {
  return typeof value === 'object' && value !== null
}
