import { readFileSync, readdirSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)
const packageManifestDirs = ['.', 'packages', 'apps', 'examples'] as const
const forbiddenDependencySources = Object.freeze(['cdn.sheetjs.com'])
const excelImportRuntimeXlsxImportAllowlist = new Set([
  'packages/excel-import/src/xlsx-export.ts',
  'packages/excel-import/src/xlsx-sheetjs-import.ts',
])
const liveScorecardFixtureScripts = [
  'scripts/gen-google-sheets-live-calculation-scorecard.ts',
  'scripts/gen-google-sheets-live-recalculation-scorecard.ts',
  'scripts/gen-google-sheets-live-structural-scorecard.ts',
  'scripts/gen-microsoft-excel-live-calculation-scorecard.ts',
  'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts',
  'scripts/gen-microsoft-excel-live-structural-scorecard.ts',
] as const
const nativeXlsxFixtureScripts = [
  ...liveScorecardFixtureScripts,
  'e2e/tests/web-shell-import.pw.ts',
  'scripts/gen-import-export-fidelity-scorecard.ts',
  'scripts/gen-security-posture-scorecard.ts',
  'scripts/gen-workpaper-xlsx-corpus-fixtures.ts',
] as const
const nativeXlsxCorpusProofScripts = ['scripts/check-workpaper-xlsx-corpus.ts', 'scripts/workpaper-xlsx-volatile-dependencies.ts'] as const
const nativeXlsxExampleScripts = [
  'examples/recalc-bridge-workflows/smoke.mjs',
  'examples/recalc-bridge-workflows/stackoverflow-sheetjs-63085785.mjs',
] as const
const nativeXlsxExcelImportFixtureTests = [
  'packages/excel-import/src/__tests__/xlsx-byte-source-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-calculation-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-date1904-system.test.ts',
  'packages/excel-import/src/__tests__/xlsx-defined-names-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-error-cell-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-external-defined-names.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-text-normalization.test.ts',
  'packages/excel-import/src/__tests__/xlsx-hyperlink-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-axis-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-name-whitespace.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-view-state-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-workbook-sheet-paths.test.ts',
  'packages/excel-import/src/__tests__/xlsx-worksheet-dimensions-roundtrip.test.ts',
] as const

function packageManifestPaths(): string[] {
  return packageManifestDirs.flatMap((dir) => {
    if (dir === '.') {
      return ['package.json']
    }

    return readdirSync(join(repoRoot, dir), { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => join(dir, entry.name, 'package.json'))
      .filter((path) => {
        try {
          readFileSync(join(repoRoot, path), 'utf8')
          return true
        } catch {
          return false
        }
      })
  })
}

function packageDependencySourceViolations(path: string): string[] {
  const source = readFileSync(join(repoRoot, path), 'utf8')
  return forbiddenDependencySources.filter((forbidden) => source.includes(forbidden)).map((forbidden) => `${path}: ${forbidden}`)
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

function relativePath(path: string): string {
  return path.slice(repoRoot.length + 1)
}

function runtimeXlsxImportViolations(path: string): string[] {
  const source = readFileSync(path, 'utf8')
  if (!/(?:^|\n)\s*import\s+\*\s+as\s+\w+\s+from\s+['"]xlsx['"]|require\(['"]xlsx['"]\)|import\(['"]xlsx['"]\)/u.test(source)) {
    return []
  }
  const relative = relativePath(path)
  return excelImportRuntimeXlsxImportAllowlist.has(relative) ? [] : [relative]
}

function hasRuntimeXlsxImport(source: string): boolean {
  return /(?:^|\n)\s*import\s+\*\s+as\s+\w+\s+from\s+['"]xlsx['"]|require\(['"]xlsx['"]\)|import\(['"]xlsx['"]\)/u.test(source)
}

describe('repository dependency policy', () => {
  it('does not pin package dependencies to the SheetJS CDN tarball', () => {
    const violations = packageManifestPaths().flatMap(packageDependencySourceViolations)

    expect(violations).toEqual([])
  })

  it('does not resolve lockfile entries from the SheetJS CDN tarball', () => {
    const lockfile = readFileSync(join(repoRoot, 'pnpm-lock.yaml'), 'utf8')
    const violations = forbiddenDependencySources.filter((forbidden) => lockfile.includes(forbidden))

    expect(violations).toEqual([])
  })

  it('keeps @bilig/excel-import runtime SheetJS imports isolated to the parser and writer boundary', () => {
    const violations = sourceFiles(join(repoRoot, 'packages/excel-import/src')).flatMap(runtimeXlsxImportViolations)

    expect(violations).toEqual([])
  })

  it('keeps native XLSX fixture generation on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxFixtureScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps WorkPaper XLSX corpus proof readback on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxCorpusProofScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps recalculation bridge examples on @bilig/xlsx instead of SheetJS fixture edits', () => {
    const violations = nativeXlsxExampleScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps simple excel-import fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxExcelImportFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })
})
