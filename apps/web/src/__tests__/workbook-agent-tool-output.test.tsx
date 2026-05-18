// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { StructuredToolOutput, summarizeToolEntry } from '../workbook-agent-tool-output.js'

beforeEach(() => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(() => {
  document.body.innerHTML = ''
})

function toolEntry(output: unknown): WorkbookAgentTimelineEntry {
  return {
    id: 'tool-write-1',
    kind: 'tool',
    turnId: 'turn-1',
    text: null,
    phase: null,
    toolName: 'write_range',
    toolStatus: 'completed',
    argumentsText: null,
    outputText: JSON.stringify(output),
    success: true,
    citations: [],
  }
}

describe('workbook agent tool output trust summaries', () => {
  it('summarizes verification-incomplete mutation receipts before applied wording', () => {
    const summary = summarizeToolEntry(
      toolEntry({
        applied: true,
        status: 'verification_incomplete',
        summary: 'Applied workbook change set at revision r2: Write cells in Sheet1!B2',
        mutationReceipt: {
          status: 'verification_incomplete',
          warnings: ['No browser-rendered context was attached to this tool call.'],
        },
      }),
    )

    expect(summary).toBe('Verification incomplete: No browser-rendered context was attached to this tool call.')
  })

  it('renders mutation receipt proof state and warnings instead of a generic object card', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <StructuredToolOutput
          outputText={JSON.stringify({
            applied: true,
            status: 'verification_incomplete',
            summary: 'Applied workbook change set at revision r2: Write cells in Sheet1!B2',
            mutationReceipt: {
              status: 'verification_incomplete',
              revision: {
                before: 1,
                after: 2,
              },
              authoritativeReadback: {
                requested: true,
                matched: true,
              },
              renderedReadback: {
                requested: true,
                matched: null,
                incompleteReason: 'No browser-rendered context was attached to this tool call.',
              },
              undo: {
                available: false,
                lookupFailed: true,
                reasonUnavailable: 'Undo metadata lookup failed for applied revision r2: history store unavailable',
              },
              warnings: [
                'No browser-rendered context was attached to this tool call.',
                'Undo metadata lookup failed for applied revision r2: history store unavailable',
              ],
            },
          })}
          toolName="write_range"
        />,
      )
    })

    expect(host.textContent).toContain('Verification incomplete')
    expect(host.textContent).toContain('Authoritative proof')
    expect(host.textContent).toContain('Rendered proof')
    expect(host.textContent).toContain('Undo proof')
    expect(host.textContent).toContain('No browser-rendered context was attached to this tool call.')
    expect(host.textContent).toContain('Undo metadata lookup failed')
    expect(host.textContent).not.toContain('Applied workbook change set at revision r2')

    await act(async () => {
      root.unmount()
    })
  })
})
