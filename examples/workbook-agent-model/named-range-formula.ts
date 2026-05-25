import {
  defineModel,
  describeModel,
  describePlan,
  describeRunResult,
  formula,
  normalizeWorkbookActionInput,
  prepareWorkbookAction,
  runWorkbookPlan,
  toWorkbookRefData,
  verifyPlanData,
  workbookActionCommandDigest,
  workbookPlanId,
  type EngineOp,
  type WorkbookActionPlan,
  type WorkbookRunAdapter,
  type WorkbookRunApplyCommandReceipt,
} from '@bilig/workbook'

export const model = defineModel({
  name: 'named-range-formula',
  description: 'Bind named workbook refs and write a generic formula with strict runtime proof.',

  find(workbook) {
    return {
      input: workbook.findName('input'),
      factor: workbook.findName('factor'),
      result: workbook.findName('result'),
    }
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.result), workbook.check.noFormulaErrors(refs.result)]
  },

  actions: {
    calculate: {
      description: 'Write input times factor into result.',
      run({ refs, workbook }) {
        const expected = formula.multiply(refs.input, refs.factor)
        workbook.writeFormula(refs.result, expected)
        workbook.check.formulaEquals(refs.result, expected)
      },
    },
  },
})

function requiredFormula<Refs>(plan: WorkbookActionPlan<Refs>): string {
  const command = plan.commands.find((entry) => entry.kind === 'writeFormula')
  if (command?.kind !== 'writeFormula') {
    throw new Error('example plan did not produce a formula write')
  }
  return command.formula
}

function materializedOps(formulaText: string): readonly EngineOp[] {
  return [
    {
      kind: 'setCellFormula',
      sheetName: 'Resolved',
      address: 'C1',
      formula: formulaText,
    },
  ]
}

function commandReceipt<Refs>(
  plan: WorkbookActionPlan<Refs>,
  commandIndex: number,
  ops: readonly EngineOp[],
): WorkbookRunApplyCommandReceipt {
  const command = plan.commands[commandIndex]
  if (command === undefined) {
    throw new Error(`example plan is missing command ${String(commandIndex)}`)
  }

  const rawResolvedRefs: Record<string, unknown> = {}
  if (command.target !== undefined) {
    rawResolvedRefs['target'] = toWorkbookRefData(command.target)
  }
  if (command.kind === 'writeFormula' && command.inputs.length > 0) {
    rawResolvedRefs['inputs'] = command.inputs.map((input) => toWorkbookRefData(input))
  }
  const resolvedRefs = Object.keys(rawResolvedRefs).length > 0 ? normalizeWorkbookActionInput(rawResolvedRefs) : undefined

  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: ops,
    appliedOps: ops,
    ...(resolvedRefs !== undefined ? { resolvedRefs } : {}),
  }
}

const prepared = prepareWorkbookAction(model, 'calculate')
if (prepared.status === 'failed') {
  throw new Error(JSON.stringify(prepared.errors, null, 2))
}

const plannedFormula = requiredFormula(prepared.plan)
const describedPlan = describePlan(prepared.plan)
const transportedPlan = JSON.parse(JSON.stringify(prepared.planData))
const ops = materializedOps(plannedFormula)
const adapter: WorkbookRunAdapter<{ readonly refsUsed: typeof prepared.plan.refsUsed }> = {
  apply(plan) {
    return {
      status: 'applied',
      planId: workbookPlanId(plan),
      baseRevision: 0,
      revision: 1,
      previewOps: ops,
      appliedOps: ops,
      commandReceipts: plan.commands.map((_, commandIndex) => commandReceipt(plan, commandIndex, ops)),
      proof: {
        source: 'example-adapter',
        mode: 'strict-proof-fixture',
      },
    }
  },
  read(targets) {
    return targets.map((target) => ({
      target,
      formula: plannedFormula,
    }))
  },
  verifyChecks(checks) {
    return checks.map((check) =>
      check.status === 'planned'
        ? {
            ...check,
            status: 'passed',
            proof: { source: 'example-adapter' },
          }
        : check,
    )
  },
}

const result = await runWorkbookPlan(transportedPlan, adapter, { strict: true })

console.log(
  JSON.stringify(
    {
      model: describeModel(model),
      plan: describedPlan,
      verification: prepared.verification,
      transportVerification: verifyPlanData(transportedPlan),
      requirements: prepared.requirements,
      result: describeRunResult(result),
    },
    null,
    2,
  ),
)
