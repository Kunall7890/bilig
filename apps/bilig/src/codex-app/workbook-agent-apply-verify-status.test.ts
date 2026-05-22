import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { createWorkbookAgentCommandBundle, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { z } from 'zod'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-apply-verify-status',
    replicaId: 'server:apply-verify-status-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'B2', 'visible value')
  return engine
}

function createZeroSyncHarness(
  engine: SpreadsheetEngine,
  input: {
    readonly headRevision: number
    readonly calculatedRevision?: number
  },
): ZeroSyncService {
  const calculatedRevision = input.calculatedRevision ?? input.headRevision
  return {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error('not used')
    },
    async handleMutate() {
      throw new Error('not used')
    },
    async inspectWorkbook(_documentId, task) {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: input.headRevision,
          calculatedRevision,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-05-17T12:00:00.000Z',
        }),
        headRevision: input.headRevision,
        calculatedRevision,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {
      throw new Error('not used')
    },
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error('not used')
    },
    async getWorkbookHeadRevision() {
      return input.headRevision
    },
    async loadAuthoritativeEvents() {
      throw new Error('not used')
    },
  }
}

function createBundle(): WorkbookAgentCommandBundle {
  return createWorkbookAgentCommandBundle({
    bundleId: 'bundle-apply-verify-status',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Verify apply and verify status',
    baseRevision: 1,
    context: null,
    commands: [{ kind: 'createSheet', name: 'unused' }],
    now: 1,
  })
}

function renderedContext(capturedRevision: number, source: 'selection' | 'visibleRange' = 'selection'): WorkbookAgentUiContext {
  const renderedRange = {
    range: {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'B2',
    },
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    truncated: false,
    rows: [
      [
        {
          address: 'B2',
          input: 'visible value',
          value: { tag: ValueTag.String, value: 'visible value' },
          formula: null,
          displayFormat: 'visible value',
          styleId: null,
          numberFormatId: null,
          style: null,
        },
      ],
    ],
  }
  return {
    selection: {
      sheetName: 'Sheet1',
      address: source === 'selection' ? 'B2' : 'A1',
      range: {
        startAddress: source === 'selection' ? 'B2' : 'A1',
        endAddress: source === 'selection' ? 'B2' : 'A1',
      },
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
    rendered: {
      capturedAtUnixMs: 10,
      capturedRevision,
      batchId: capturedRevision,
      surfaceProof: {
        mode: 'typegpu-v3',
        backendStatus: 'ready',
        frameProofStatus: 'presented',
        hasPresentedFrame: true,
        hasPresentedVisibleFrame: true,
        frameProofSignature: `frame:${String(capturedRevision)}`,
        presentedFrameProofSignature: `frame:${String(capturedRevision)}`,
        currentSceneOwnershipSignature: `scene:${String(capturedRevision)}`,
        presentedSceneOwnershipSignature: `scene:${String(capturedRevision)}`,
        authoritativeRevision: capturedRevision,
        localRevision: null,
        projectedRevision: capturedRevision,
        visibleRenderRevision: capturedRevision,
        tileSceneRevision: capturedRevision,
        tileSceneCameraSeq: 1,
        currentTilePaneCount: 1,
        currentHeaderPaneCount: 1,
        presentedTilePaneCount: 1,
        presentedHeaderPaneCount: 1,
        surfaceWidth: 800,
        surfaceHeight: 600,
        surfacePixelWidth: 1600,
        surfacePixelHeight: 1200,
        devicePixelRatio: 2,
        capturedAtUnixMs: 10,
      },
      selection: source === 'selection' ? renderedRange : null,
      visibleRange: source === 'visibleRange' ? renderedRange : null,
    },
  }
}

function readToolJson(response: { readonly success: boolean; readonly contentItems: readonly unknown[] }): unknown {
  expect(response.success).toBe(true)
  const item = response.contentItems[0]
  if (!item || typeof item !== 'object' || !('text' in item) || typeof item.text !== 'string') {
    throw new Error('Expected text tool result')
  }
  return JSON.parse(item.text)
}

describe('apply_and_verify proof status', () => {
  it('keeps verification incomplete when formula and invariant audits are explicitly skipped', async () => {
    const engine = await createEngine()
    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: renderedContext(5),
        zeroSyncService: createZeroSyncHarness(engine, {
          headRevision: 5,
          calculatedRevision: 5,
        }),
        awaitRenderedRevision: vi.fn(async () => undefined),
        stageCommand: vi.fn(async () => createBundle()),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-apply-and-verify-skipped-audits',
        tool: 'apply_and_verify',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          includeFormulaIssues: false,
          includeInvariants: false,
        },
      },
    )

    const payload = z
      .object({
        status: z.literal('verification_incomplete'),
        verificationComplete: z.literal(false),
        verificationMissingChecks: z.array(z.string()),
        renderedReadback: z.array(
          z.object({
            matched: z.literal(true),
            stale: z.literal(false),
          }),
        ),
        formulaIssues: z.null(),
        invariants: z.null(),
      })
      .parse(readToolJson(response))
    expect(payload.verificationMissingChecks).toEqual(['formulaIssues', 'invariants'])
  })

  it('keeps verification incomplete when rendered readback only comes from the viewport', async () => {
    const engine = await createEngine()
    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: renderedContext(5, 'visibleRange'),
        zeroSyncService: createZeroSyncHarness(engine, {
          headRevision: 5,
          calculatedRevision: 5,
        }),
        awaitRenderedRevision: vi.fn(async () => undefined),
        stageCommand: vi.fn(async () => createBundle()),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-apply-and-verify-viewport-only',
        tool: 'apply_and_verify',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
        },
      },
    )

    const payload = z
      .object({
        status: z.literal('verification_incomplete'),
        verificationComplete: z.literal(false),
        verificationMissingChecks: z.array(z.string()),
        renderedReadback: z.array(
          z.object({
            matched: z.literal(true),
            sourceKind: z.literal('visibleRange'),
          }),
        ),
      })
      .parse(readToolJson(response))
    expect(payload.verificationMissingChecks).toEqual(['renderedSelection'])
  })
})
