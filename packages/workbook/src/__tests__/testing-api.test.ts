import { describe, expect, it } from 'vitest'
import {
  defineModel,
  formula,
  prepareWorkbookAction,
  workbookActionCommandDigest,
  workbookPlanId,
  type EngineOp,
  type WorkbookActionPlan,
  type WorkbookRunAdapter,
  type WorkbookRunApplyCommandReceipt,
} from '../index.js'
import { assertWorkbookRunAdapter, checkWorkbookRunAdapter } from '../testing.js'

function rangeRef(label: string, startAddress: string, endAddress = startAddress) {
  return {
    kind: 'range' as const,
    id: `range_${label}`,
    label,
    range: {
      sheetName: 'Resolved',
      startAddress,
      endAddress,
    },
  }
}

function labelSource(labelName: string): string {
  if (labelName === 'input') {
    return 'Resolved!A1'
  }
  if (labelName === 'factor') {
    return 'Resolved!B1'
  }
  return labelName
}

function model() {
  return defineModel({
    name: 'testing-adapter-model',
    find(workbook) {
      return {
        input: workbook.findName('input'),
        factor: workbook.findName('factor'),
        result: workbook.findName('result'),
      }
    },
    checks({ refs, workbook }) {
      const expected = formula.multiply(refs.input, refs.factor)
      return [workbook.check.formulaEquals(refs.result, expected)]
    },
    actions: {
      calculate({ refs, workbook }) {
        workbook.writeFormula(refs.result, formula.multiply(refs.input, refs.factor))
      },
    },
  })
}

function receipt<Refs>(plan: WorkbookActionPlan<Refs>, commandIndex: number, op: EngineOp): WorkbookRunApplyCommandReceipt {
  const command = plan.commands[commandIndex]
  if (command?.kind !== 'writeFormula') {
    throw new Error('expected formula command')
  }
  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: [op],
    appliedOps: [op],
    resolvedRefs: {
      target: rangeRef('Resolved!C1', 'C1'),
      inputs: [rangeRef('Resolved!A1', 'A1'), rangeRef('Resolved!B1', 'B1')],
    },
    formulaLabels: command.labels.map((label) => ({ name: label.name, source: labelSource(label.name) })),
  }
}

function passingAdapter(): WorkbookRunAdapter<{ readonly refsUsed: ReturnType<typeof prepare>['plan']['refsUsed'] }> {
  return {
    apply(plan) {
      const op: EngineOp = {
        kind: 'setCellFormula',
        sheetName: 'Resolved',
        address: 'C1',
        formula: 'Resolved!A1*Resolved!B1',
      }
      return {
        status: 'applied',
        planId: workbookPlanId(plan),
        baseRevision: 4,
        revision: 5,
        previewOps: [op],
        appliedOps: [op],
        commandReceipts: [receipt(plan, 0, op)],
      }
    },
    read(targets) {
      return targets.map((target) => ({
        target,
        formula: 'Resolved!A1*Resolved!B1',
        formulaLabels: [
          { name: 'input', source: 'Resolved!A1' },
          { name: 'factor', source: 'Resolved!B1' },
        ],
      }))
    },
  }
}

function prepare() {
  const prepared = prepareWorkbookAction(model(), 'calculate')
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared fixture')
  }
  return prepared
}

function prepareMultiCellFormula() {
  const prepared = prepareWorkbookAction(
    defineModel({
      name: 'testing-adapter-multi-cell-model',
      find(workbook) {
        return {
          input: workbook.findName('input'),
          factor: workbook.findName('factor'),
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.formulaEquals(refs.result, formula.multiply(refs.input, refs.factor))]
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.result, formula.multiply(refs.input, refs.factor))
        },
      },
    }),
    'calculate',
  )
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared multi-cell fixture')
  }
  return prepared
}

function prepareFormatCommand() {
  const prepared = prepareWorkbookAction(
    defineModel({
      name: 'testing-adapter-format-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },
      actions: {
        format({ refs, workbook }) {
          workbook.format(refs.result, { numberFormat: '0.00' })
        },
      },
    }),
    'format',
  )
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared format fixture')
  }
  return prepared
}

function prepareStyleFormatCommand() {
  const prepared = prepareWorkbookAction(
    defineModel({
      name: 'testing-adapter-style-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },
      actions: {
        format({ refs, workbook }) {
          workbook.format(refs.result, { style: { font: { bold: true } } })
        },
      },
    }),
    'format',
  )
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared style fixture')
  }
  return prepared
}

function prepareFullFormatCommand() {
  const prepared = prepareWorkbookAction(
    defineModel({
      name: 'testing-adapter-full-format-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },
      actions: {
        format({ refs, workbook }) {
          workbook.format(refs.result, {
            numberFormat: '0.00',
            style: { font: { bold: true } },
          })
        },
      },
    }),
    'format',
  )
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared full format fixture')
  }
  return prepared
}

describe('@bilig/workbook testing api', () => {
  it('checks a runtime adapter against a strict transported plan', async () => {
    const prepared = prepare()

    const check = await checkWorkbookRunAdapter(prepared.planData, passingAdapter())

    expect(check.status).toBe('passed')
    if (check.status !== 'passed') {
      throw new Error('adapter did not pass')
    }
    expect(Object.isFrozen(check)).toBe(true)
    expect(check.result.status).toBe('done')
    expect(check.description).toMatchObject({
      status: 'done',
      apply: {
        matched: true,
        baseRevision: 4,
        revision: 5,
        commandReceipts: [expect.objectContaining({ commandKind: 'writeFormula' })],
      },
    })
  })

  it('returns adapter capability issues without mutating', async () => {
    const prepared = prepare()
    const adapter = passingAdapter()

    const check = await checkWorkbookRunAdapter(prepared.planData, { apply: (plan) => adapter.apply?.(plan) })

    expect(check).toEqual({
      status: 'failed',
      errors: [],
      issues: [
        {
          code: 'adapter_missing_capability',
          path: 'adapter.read',
          message: 'Adapter is missing read for read',
        },
      ],
    })
  })

  it('returns transported plan issues without throwing', async () => {
    const invalidPlan = JSON.parse(
      JSON.stringify({
        modelName: 'bad-plan-model',
        actionName: 'calculate',
        refsUsed: 'not-an-array',
        commands: [],
        ops: [],
        changed: [],
        checks: [],
      }),
    )

    const check = await checkWorkbookRunAdapter(invalidPlan, passingAdapter())

    expect(check).toEqual({
      status: 'failed',
      errors: [],
      issues: [
        {
          code: 'invalid_plan_data',
          path: 'plan',
          message: 'Workbook plan data is invalid: Workbook plan data refsUsed must be an array',
        },
      ],
    })
  })

  it('returns strict proof failures as boring issues', async () => {
    const prepared = prepare()

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          previewOps: [],
          appliedOps: [],
          commandReceipts: [],
        }
      },
      read() {
        return []
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message: 'Workbook action testing-adapter-model.calculate returned invalid command receipts: expected 1 command receipts, got 0',
      },
    ])
    expect(check.result?.status).toBe('failed')
    expect(check.description?.status).toBe('failed')
  })

  it('rejects symbolic ref receipts that are not bound to concrete ranges', async () => {
    const prepared = prepare()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const op: EngineOp = {
      kind: 'setCellFormula',
      sheetName: 'Resolved',
      address: 'C1',
      formula: command.formula,
    }

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: [op],
          appliedOps: [op],
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: [op],
              appliedOps: [op],
              resolvedRefs: {
                target: {
                  kind: 'name',
                  id: 'name_result',
                  label: 'result',
                  name: 'result',
                },
              },
            },
          ],
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'input*factor',
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-model.calculate returned invalid command receipts: commandReceipts[0].resolvedRefs.target must be concrete range ref data',
      },
    ])
  })

  it('rejects materialized ops that do not match the resolved concrete target', async () => {
    const prepared = prepare()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const op: EngineOp = {
      kind: 'setCellFormula',
      sheetName: 'Resolved',
      address: 'D1',
      formula: command.formula,
    }

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: [op],
          appliedOps: [op],
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: [op],
              appliedOps: [op],
              resolvedRefs: {
                target: rangeRef('Resolved!C1', 'C1'),
                inputs: [rangeRef('Resolved!A1', 'A1'), rangeRef('Resolved!B1', 'B1')],
              },
            },
          ],
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'input*factor',
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-model.calculate returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('rejects multi-cell formula receipts whose formulas do not match resolved inputs', async () => {
    const prepared = prepareMultiCellFormula()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'writeFormula') {
      throw new Error('expected formula command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'setCellFormula',
        sheetName: 'Resolved',
        address: 'C1',
        formula: 'Resolved!A1*Resolved!B1',
      },
      {
        kind: 'setCellFormula',
        sheetName: 'Resolved',
        address: 'C2',
        formula: 'Resolved!A1*Resolved!B1',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
                inputs: [rangeRef('Resolved!A1:A2', 'A1', 'A2'), rangeRef('Resolved!B1:B2', 'B1', 'B2')],
              },
            },
          ],
        }
      },
      read(targets) {
        return targets.map((target) => ({
          target,
          formula: 'input*factor',
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-multi-cell-model.calculate returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('rejects symbolic format receipts whose ops do not touch the resolved target', async () => {
    const prepared = prepareFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format_0_00', code: '0.00', kind: 'number' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'D1',
          endAddress: 'D2',
        },
        formatId: 'format_0_00',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-format-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('rejects symbolic format receipts that only cover part of the resolved target', async () => {
    const prepared = prepareFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format_0_00', code: '0.00', kind: 'number' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C1',
        },
        formatId: 'format_0_00',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-format-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('rejects symbolic format receipts missing one requested component', async () => {
    const prepared = prepareFullFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format_0_00', code: '0.00', kind: 'number' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        formatId: 'format_0_00',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-full-format-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('accepts symbolic format receipts that cover every requested component', async () => {
    const prepared = prepareFullFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format_0_00', code: '0.00', kind: 'number' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        formatId: 'format_0_00',
      },
      {
        kind: 'upsertCellStyle',
        style: { id: 'style_bold', font: { bold: true } },
      },
      {
        kind: 'setStyleRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        styleId: 'style_bold',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('passed')
  })

  it('rejects symbolic format receipts whose range format id has the wrong payload', async () => {
    const prepared = prepareFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format_text', code: 'text', kind: 'text' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        formatId: 'format_text',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-format-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('rejects symbolic style receipts whose range style id has the wrong payload', async () => {
    const prepared = prepareStyleFormatCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellStyle',
        style: { id: 'style_italic', font: { italic: true } },
      },
      {
        kind: 'setStyleRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        styleId: 'style_italic',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-style-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('throws from the assertion helper with the first issue message', async () => {
    const prepared = prepare()
    const adapter = passingAdapter()
    await expect(assertWorkbookRunAdapter(prepared.planData, { apply: (plan) => adapter.apply?.(plan) })).rejects.toThrow(
      'Adapter is missing read for read',
    )

    await expect(assertWorkbookRunAdapter(prepared.planData, passingAdapter())).resolves.toMatchObject({
      status: 'done',
    })
  })
})
