import { SpreadsheetEngine } from '@bilig/core'
import { checkWorkbookRunResultDescription, defineModel, formula, planWorkbookAction, toPlanData } from '@bilig/workbook'
import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import {
  runStrictWorkbookPlanData,
  workbookPlanRunAppliedOps,
  workbookPlanRunResultProof,
  workbookPlanRunUndoBundle,
} from '../workbook-plan-data-apply.js'

function plannedData(model: unknown, actionName: string) {
  const planned = planWorkbookAction(model, actionName)
  if (planned.status !== 'planned') {
    throw new Error(planned.errors.map((error) => error.message).join('\n'))
  }
  return toPlanData(planned.plan)
}

describe('workbook plan data apply', () => {
  it('runs transported generic workbook plan data through strict core proof', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'plan-data-apply' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    const model = defineModel({
      name: 'generic-plan-data',
      find(workbook) {
        return {
          left: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          right: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'C1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.valueEquals(refs.output, 5)]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.add(refs.left, refs.right))
        },
      },
    })

    const result = await runStrictWorkbookPlanData(engine, plannedData(model, 'calculate'))

    expect(result.status).toBe('done')
    expect(workbookPlanRunAppliedOps(result)).toEqual([
      {
        kind: 'setCellFormula',
        sheetName: 'Sheet1',
        address: 'C1',
        formula: 'Sheet1!A1+Sheet1!B1',
      },
    ])
    expect(result.apply?.commandReceipts?.[0]).toMatchObject({
      resolvedRefs: {
        target: expect.objectContaining({ kind: 'range', label: 'Sheet1!C1' }),
        inputs: [
          expect.objectContaining({ kind: 'range', label: 'Sheet1!A1' }),
          expect.objectContaining({ kind: 'range', label: 'Sheet1!B1' }),
        ],
      },
    })
    expect(workbookPlanRunUndoBundle(result)).toMatchObject({
      kind: 'engineOps',
      ops: expect.any(Array),
    })
    const proof = workbookPlanRunResultProof(result)
    const checkedProof = checkWorkbookRunResultDescription(proof)
    expect(checkedProof.status).toBe('valid')
    if (checkedProof.status !== 'valid') {
      throw new Error(checkedProof.issues.map((issue) => issue.message).join('\n'))
    }
    expect(checkedProof.description.apply).toMatchObject({
      baseRevision: 0,
      revision: 1,
      commandReceipts: [
        expect.objectContaining({
          commandIndex: 0,
          commandKind: 'writeFormula',
          resolvedRefs: expect.objectContaining({
            target: expect.objectContaining({ label: 'Sheet1!C1' }),
          }),
        }),
      ],
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('rolls back applied ops when transported plan data fails post-apply checks', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'plan-data-rollback' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 0)
    const model = defineModel({
      name: 'generic-plan-data-failure',
      find(workbook) {
        return {
          divisor: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.noFormulaErrors(refs.output)]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.divide(1, refs.divisor))
        },
      },
    })

    const result = await runStrictWorkbookPlanData(engine, plannedData(model, 'calculate'))

    expect(result.status).toBe('failed')
    expect(result.errors.map((error) => error.code)).toContain('check_failed')
    expect(result.undo?.ops?.length).toBeGreaterThan(0)
    const proof = workbookPlanRunResultProof(result)
    const checkedProof = checkWorkbookRunResultDescription(proof)
    expect(checkedProof.status).toBe('valid')
    if (checkedProof.status !== 'valid') {
      throw new Error(checkedProof.issues.map((issue) => issue.message).join('\n'))
    }
    expect(checkedProof.description).toMatchObject({
      status: 'failed',
      apply: expect.objectContaining({
        baseRevision: 0,
        revision: 1,
      }),
      undo: expect.objectContaining({
        id: expect.stringContaining('generic-plan-data-failure.calculate.undo.'),
      }),
    })
    expect(engine.getCell('Sheet1', 'B1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })
  })
})
