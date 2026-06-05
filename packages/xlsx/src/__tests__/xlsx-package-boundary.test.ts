import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  decodeCellAddress,
  decodeCellRange,
  decodeColumnAddress,
  encodeCellAddress,
  encodeCellRange,
  normalizeCellAddress,
  readXmlAttribute,
  worksheetCellElementPattern,
} from '../index.js'

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readPackageJson(): Record<string, unknown> {
  const packageJson: unknown = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'))
  if (!isObjectRecord(packageJson)) {
    throw new TypeError('Expected package.json to parse to an object')
  }
  return packageJson
}

function sourceFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) {
      return entry.name === '__tests__' ? [] : sourceFiles(path)
    }
    return entry.isFile() && entry.name.endsWith('.ts') ? [path] : []
  })
}

describe('@bilig/xlsx package boundary', () => {
  it('does not depend on SheetJS or the xlsx CDN tarball', () => {
    const packageJson = readPackageJson()
    expect(packageJson.name).toBe('@bilig/xlsx')
    expect(packageJson.dependencies).toBeUndefined()

    const manifestSource = readFileSync(join(packageDir, 'package.json'), 'utf8')
    expect(manifestSource).not.toContain('cdn.sheetjs.com')
    expect(manifestSource).not.toMatch(/"xlsx"\s*:/u)
  })

  it('keeps native source free of SheetJS imports', () => {
    for (const sourceFile of sourceFiles(join(packageDir, 'src'))) {
      const source = readFileSync(sourceFile, 'utf8')
      expect(source, sourceFile).not.toMatch(/from\s+["']xlsx["']/u)
      expect(source, sourceFile).not.toMatch(/require\(["']xlsx["']\)/u)
      expect(source, sourceFile).not.toContain('cdn.sheetjs.com')
    }
  })

  it('normalizes XLSX addresses and ranges without SheetJS', () => {
    expect(decodeCellAddress('$AA$42')).toEqual({ r: 41, c: 26 })
    expect(decodeColumnAddress('$XFD')).toBe(16383)
    expect(encodeCellAddress({ r: 0, c: 701 })).toBe('ZZ1')
    expect(normalizeCellAddress('$b$7')).toBe('B7')
    expect(decodeCellRange('C3:A1')).toEqual({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } })
    expect(encodeCellRange({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } })).toBe('A1:C3')
  })

  it('exports native XML helpers for package readers without SheetJS', () => {
    expect(readXmlAttribute('<sheet name="A &amp; B" r:id="rId1"/>', 'name')).toBe('A & B')
    const cells = [
      ...'<sheetData><row><c r="A1"><v>1</v></c><c r="B2" t="str"><v>x</v></c></row></sheetData>'.matchAll(worksheetCellElementPattern),
    ].map((match) => match[0])

    expect(cells).toHaveLength(2)
  })
})
