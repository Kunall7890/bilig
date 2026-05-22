import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const testDir = dirname(fileURLToPath(import.meta.url))

function source(relativePath: string): string {
  return readFileSync(join(testDir, relativePath), 'utf8')
}

function sourceLineCount(relativePath: string): number {
  return source(relativePath).split('\n').length
}

describe('workbook agent mutation receipt module boundary', () => {
  it('keeps mutation proof authority outside the tool-result serializer', () => {
    const receiptSource = source('workbook-agent-mutation-receipt.ts')
    const proofSource = source('workbook-agent-mutation-proof.ts')

    expect(sourceLineCount('workbook-agent-mutation-receipt.ts')).toBeLessThan(360)
    expect(sourceLineCount('workbook-agent-mutation-proof.ts')).toBeLessThan(620)
    expect(receiptSource).not.toContain('listWorkbookChanges')
    expect(receiptSource).not.toContain('inspectWorkbookRange')
    expect(receiptSource).not.toContain('buildCellNumberFormatCode')
    expect(proofSource).toContain('listWorkbookChanges')
    expect(proofSource).toContain('inspectWorkbookRange')
    expect(proofSource).toContain('buildCellNumberFormatCode')
  })
})
