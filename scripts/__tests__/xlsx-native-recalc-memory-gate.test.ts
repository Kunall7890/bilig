import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import {
  buildNativeRecalcCliArgs,
  nativeRecalcMemoryGateBudgets,
  writeSyntheticNativeRecalcWorkbook,
  type NativeRecalcGateTarget,
} from '../xlsx-native-recalc-memory-gate.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'

const mib = 1024 * 1024

describe('xlsx native recalc memory gate', () => {
  it('keeps native recalc budgets explicit for synthetic and issue 442 targets', () => {
    expect(nativeRecalcMemoryGateBudgets.syntheticRowChainMaxRssBytes).toBe(320 * mib)
    expect(nativeRecalcMemoryGateBudgets.issue442MaxRssBytes).toBe(350 * mib)
  })

  it('builds streaming-native CLI args with fallback disabled', () => {
    const target: NativeRecalcGateTarget = {
      id: 'issue-442-ocha-native-recalc',
      label: 'ocha.xlsx',
      inputPath: '/tmp/ocha.xlsx',
      outputPath: '/tmp/ocha.native.xlsx',
      maxRssBytes: 350 * mib,
      edits: [{ target: 'Data!R57152', value: 16 }],
      reads: ['Data!U57152', 'Data!V57152'],
      expectedReads: { 'Data!U57152': 168.75, 'Data!V57152': 28.125 },
    }

    expect(buildNativeRecalcCliArgs(target)).toEqual([
      '/tmp/ocha.xlsx',
      '--out',
      '/tmp/ocha.native.xlsx',
      '--engine',
      'streaming-native',
      '--fallback-policy',
      'error',
      '--max-rss-bytes',
      String(350 * mib),
      '--json',
      '--set',
      'Data!R57152=16',
      '--read',
      'Data!U57152',
      '--read',
      'Data!V57152',
    ])
  })

  it('writes a synthetic row-chain workbook that exercises dependent native cache patches', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-recalc-gate-'))
    try {
      const artifact = writeSyntheticNativeRecalcWorkbook(tempDir, { rowCount: 12 })

      expect(existsSync(artifact.filePath)).toBe(true)
      expect(artifact.editTarget).toBe('Data!A12')
      expect(artifact.reads).toEqual(['Data!C12', 'Data!B12'])
      expect(artifact.expectedReads).toEqual({ 'Data!C12': 168.75, 'Data!B12': 28.125 })

      const files = unzipSync(readFileSync(artifact.filePath))
      const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<dimension ref="A1:C12"/>')
      expect(sheetXml).toContain('<c r="B12"><f>A12*1.7578125</f><v>0</v></c>')
      expect(sheetXml).toContain('<c r="C12"><f>B12*6</f><v>0</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('exposes the focused memory gate as a package script', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(join(process.cwd(), 'package.json'), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])
    const issue442Script = String(scripts['xlsx-native-recalc:issue-442-gate'])

    expect(scripts['xlsx-native-recalc:memory-gate']).toBe('bun scripts/xlsx-native-recalc-memory-gate.ts')
    expect(issue442Script).toContain('bun scripts/xlsx-native-recalc-memory-gate.ts')
    expect(issue442Script).toContain('--issue-442-path .cache/issue-442/ocha-operational-partners-presence-jan-sep-2024.xlsx')
    expect(issue442Script).toContain('--require-issue-442')
    expect(issue442Script).toContain('--issue-442-only')
  })
})
