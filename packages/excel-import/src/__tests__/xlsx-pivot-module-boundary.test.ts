import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const testDir = dirname(fileURLToPath(import.meta.url))
const excelImportSrc = join(testDir, '..')

function sourceLineCount(relativePath: string): number {
  return readFileSync(join(excelImportSrc, relativePath), 'utf8').split('\n').length
}

describe('xlsx pivot module boundary', () => {
  it('keeps pivot semantic import separate from pivot package export wiring', () => {
    expect(sourceLineCount('xlsx-pivots.ts')).toBeLessThan(900)
    expect(sourceLineCount('xlsx-pivot-export.ts')).toBeLessThan(550)
  })
})
