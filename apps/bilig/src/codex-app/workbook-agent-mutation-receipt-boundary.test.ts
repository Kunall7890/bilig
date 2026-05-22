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
    const authoritativeProofSource = source('workbook-agent-mutation-proof.ts')
    const semanticProofSource = source('workbook-agent-mutation-authoritative-semantic.ts')
    const runtimeProofSource = source('workbook-agent-mutation-runtime-proof.ts')

    expect(sourceLineCount('workbook-agent-mutation-receipt.ts')).toBeLessThan(340)
    expect(sourceLineCount('workbook-agent-mutation-proof.ts')).toBeLessThan(620)
    expect(sourceLineCount('workbook-agent-mutation-authoritative-semantic.ts')).toBeLessThan(320)
    expect(sourceLineCount('workbook-agent-mutation-runtime-proof.ts')).toBeLessThan(260)
    expect(receiptSource).not.toContain('listWorkbookChanges')
    expect(receiptSource).not.toContain('inspectWorkbookRange')
    expect(receiptSource).not.toContain('buildCellNumberFormatCode')
    expect(authoritativeProofSource).toContain('inspectWorkbookRange')
    expect(authoritativeProofSource).toContain('buildCellNumberFormatCode')
    expect(authoritativeProofSource).not.toContain('listWorkbookChanges')
    expect(semanticProofSource).toContain('diffWorkbookSemanticSnapshots')
    expect(runtimeProofSource).toContain('listWorkbookChanges')
    expect(runtimeProofSource).toContain('verifyWorkbookInvariants')
  })
})
