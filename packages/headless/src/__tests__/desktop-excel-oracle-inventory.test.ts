import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const rootDir = fileURLToPath(new URL('../../../..', import.meta.url))
const packageOracleGlob = 'src/__tests__/macos-desktop-excel-*.test.ts'
const rootOracleGlob = 'packages/headless/src/__tests__/macos-desktop-excel-*.test.ts'
const inventoryTestPath = 'packages/headless/src/__tests__/desktop-excel-oracle-inventory.test.ts'
const oracleRunGuard = "BILIG_EXCEL_ORACLE_RUN === '1'"

interface PackageJson {
  readonly scripts?: Record<string, string>
}

describe('Desktop Excel oracle inventory', () => {
  it('keeps the macOS oracle surface package-owned and visible from the correctness corpus', () => {
    const headlessPackage = readPackageJson('packages/headless/package.json')
    expect(headlessPackage.scripts?.['test:excel-oracle']).toBe(`vitest run --root ../.. ${packageOracleGlob} --maxWorkers=1`)
    expect(headlessPackage.scripts?.['test:excel-oracle:live']).toBe(
      `BILIG_EXCEL_ORACLE_RUN=1 vitest run --root ../.. ${packageOracleGlob} --maxWorkers=1`,
    )

    const oracleFiles = listMacosDesktopExcelOracleFiles()
    expect(oracleFiles.length).toBeGreaterThanOrEqual(16)
    expect(listLiveGuardedOracleFiles()).toEqual(oracleFiles)

    const rootPackage = readPackageJson('package.json')
    const corpusScript = rootPackage.scripts?.['test:correctness:corpus'] ?? ''
    expect(corpusScript).toContain(inventoryTestPath)
    expect(corpusScript).toContain('pnpm --filter @bilig/headless test:excel-oracle')
    expect(corpusScript).not.toContain(`${rootOracleGlob.slice(0, rootOracleGlob.indexOf('*'))}xlsx-oracle.test.ts`)
  })
})

function listMacosDesktopExcelOracleFiles(): readonly string[] {
  return readdirSync(join(rootDir, 'packages/headless/src/__tests__'))
    .filter((file) => /^macos-desktop-excel-.*\.test\.ts$/u.test(file))
    .toSorted()
}

function listLiveGuardedOracleFiles(): readonly string[] {
  return listMacosDesktopExcelOracleFiles()
    .filter((file) => readFileSync(join(rootDir, 'packages/headless/src/__tests__', file), 'utf8').includes(oracleRunGuard))
    .toSorted()
}

function readPackageJson(path: string): PackageJson {
  const parsed: unknown = JSON.parse(readFileSync(join(rootDir, path), 'utf8'))
  if (!isPackageJson(parsed)) {
    throw new Error(`${path} is not a package.json object with string scripts`)
  }
  return parsed
}

function isPackageJson(value: unknown): value is PackageJson {
  if (!isRecord(value)) {
    return false
  }
  const scripts = value['scripts']
  return scripts === undefined || (isRecord(scripts) && Object.values(scripts).every((script) => typeof script === 'string'))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
