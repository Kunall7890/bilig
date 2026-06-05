import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const srcDir = join(packageDir, 'src')

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readPackageJson(): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'))
  if (!isObjectRecord(parsed)) {
    throw new TypeError('Expected package.json to parse to an object')
  }
  return parsed
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

function relativeSourcePath(path: string): string {
  return relative(packageDir, path).split('/').join('/')
}

describe('@bilig/excel-import package boundary', () => {
  it('does not publish SheetJS as a production dependency', () => {
    const manifest = readPackageJson()
    expect(manifest.name).toBe('@bilig/excel-import')
    expect(manifest.dependencies).toEqual(
      expect.objectContaining({
        '@bilig/xlsx': 'workspace:*',
      }),
    )
    expect(manifest.dependencies).not.toHaveProperty('xlsx')
    expect(manifest.dependencies).not.toHaveProperty('xlsx-js-style')
  })

  it('keeps SheetJS behind a single optional fallback loader', () => {
    const sourceFilesWithOptionalRequire: string[] = []
    for (const sourceFile of sourceFiles(srcDir)) {
      const source = readFileSync(sourceFile, 'utf8')
      expect(source, sourceFile).not.toMatch(/^\s*import\s+type\b[^;\n]*\sfrom\s+["']xlsx["']/mu)
      expect(source, sourceFile).not.toMatch(/^\s*import\s+(?!type\b)[^;\n]*\sfrom\s+["']xlsx["']/mu)
      expect(source, sourceFile).not.toMatch(/^\s*import\s+["']xlsx["']/mu)
      expect(source, sourceFile).not.toMatch(/^\s*import\s+(?!type\b)[^;\n]*\sfrom\s+["']xlsx-js-style["']/mu)
      expect(source, sourceFile).not.toMatch(/\brequire\(["']xlsx["']\)/u)
      if (source.includes("requireModule('xlsx')") || source.includes('requireModule("xlsx")')) {
        sourceFilesWithOptionalRequire.push(relativeSourcePath(sourceFile))
      }
    }
    expect(sourceFilesWithOptionalRequire).toEqual(['src/xlsx-optional-sheetjs.ts'])
  })
})
