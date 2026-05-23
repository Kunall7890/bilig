import {
  defineModel,
  describeModel,
  describePlan,
  describeRunResult,
  describeRuntimeRequirements,
  formula,
  planWorkbookAction,
  runWorkbookPlan,
  verifyPlan,
  verifyPlanData,
  type WorkbookActionPlan,
  type WorkbookRunAdapter,
} from '@bilig/workbook'

export const model = defineModel({
  name: 'generic-row-calculator',
  description: 'Find rows by headers and status, then calculate a generic result column.',

  find(workbook) {
    const table = workbook.findTable({
      headers: ['Item', 'Quantity', 'Rate', 'Status', 'Total'],
    })
    const readyRows = workbook.findRows({
      table,
      where: {
        column: 'Status',
        op: 'eq',
        value: 'ready',
      },
    })

    return {
      table,
      readyRows,
      quantity: readyRows.column('Quantity'),
      rate: readyRows.column('Rate'),
      total: readyRows.column('Total'),
    }
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.table), workbook.check.noFormulaErrors(refs.total)]
  },

  actions: {
    recompute: {
      description: 'Write Quantity times Rate into Total for the matching rows.',
      run({ refs, workbook }) {
        const expected = formula.multiply(refs.quantity, refs.rate)
        workbook.writeFormula(refs.total, expected)
        workbook.check.formulaEquals(refs.total, expected)
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

const planned = planWorkbookAction(model, 'recompute')
if (planned.status === 'failed') {
  throw new Error(JSON.stringify(planned.errors, null, 2))
}

const plannedFormula = requiredFormula(planned.plan)
const describedPlan = describePlan(planned.plan)
const transportedPlan = JSON.parse(JSON.stringify(describedPlan))
const adapter: WorkbookRunAdapter<typeof planned.plan.refs> = {
  apply() {
    return { status: 'applied' }
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

const result = await runWorkbookPlan(planned.plan, adapter)

console.log(
  JSON.stringify(
    {
      model: describeModel(model),
      plan: describedPlan,
      verification: verifyPlan(planned.plan),
      transportVerification: verifyPlanData(transportedPlan),
      requirements: describeRuntimeRequirements(planned.plan),
      result: describeRunResult(result),
    },
    null,
    2,
  ),
)
