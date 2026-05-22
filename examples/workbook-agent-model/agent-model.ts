import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createWorkbookRunAdapter, SpreadsheetEngine } from '@bilig/core'
import {
  defineModel,
  describeModel,
  describePlan,
  describeRunResult,
  describeRuntimeRequirements,
  formula,
  planWorkbookAction,
  runWorkbookAction,
  verifyPlan,
  type WorkbookModelDescription,
  type WorkbookPlanVerification,
  type WorkbookRuntimeRequirements,
  type WorkbookActionPlanDescription,
  type WorkbookRunResultDescription,
} from '@bilig/workbook'

type CellValue = ReturnType<SpreadsheetEngine['getCellValue']>

export interface WorkbookAgentModelExampleOutput {
  readonly model: WorkbookModelDescription
  readonly planning: {
    readonly status: 'planned'
    readonly plan: WorkbookActionPlanDescription
    readonly verification: WorkbookPlanVerification
    readonly requirements: WorkbookRuntimeRequirements
  }
  readonly run: WorkbookRunResultDescription
  readonly workbook: {
    readonly formulas: Record<string, string | null>
    readonly values: Record<string, number | null>
  }
}

export const model = defineModel({
  name: 'consumer-table-calculation',
  description: 'Consumer-defined generic table model. Replace the selectors, not the library.',

  find(workbook) {
    const table = workbook.findTable({ name: 'Inputs' })
    const rows = workbook.findRows({
      table,
      where: { column: 'Kind', op: 'eq', value: 'actual' },
    })

    return {
      table,
      rows,
      base: rows.column('Base'),
      rate: rows.column('Rate'),
      result: rows.column('Result'),
    }
  },

  checks({ refs, workbook }) {
    return [
      workbook.check.exists(refs.table),
      workbook.check.exists(refs.rows),
      workbook.check.noFormulaErrors(refs.result),
      workbook.check.valuesEqual(refs.result, [[6], [20]]),
    ]
  },

  actions: {
    calculate: {
      description: 'Fill Result for the consumer-selected rows.',
      run({ refs, workbook }) {
        workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
      },
    },
  },
})

function readNumber(value: CellValue): number | null {
  return typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number' ? value.value : null
}

async function seedWorkbook(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'Agent workbook model example' })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 'Kind')
  engine.setCellValue('Sheet1', 'B1', 'Base')
  engine.setCellValue('Sheet1', 'C1', 'Rate')
  engine.setCellValue('Sheet1', 'D1', 'Result')
  engine.setCellValue('Sheet1', 'A2', 'actual')
  engine.setCellValue('Sheet1', 'B2', 2)
  engine.setCellValue('Sheet1', 'C2', 3)
  engine.setCellValue('Sheet1', 'A3', 'draft')
  engine.setCellValue('Sheet1', 'B3', 10)
  engine.setCellValue('Sheet1', 'C3', 10)
  engine.setCellValue('Sheet1', 'A4', 'actual')
  engine.setCellValue('Sheet1', 'B4', 4)
  engine.setCellValue('Sheet1', 'C4', 5)
  engine.setTable({
    name: 'Inputs',
    sheetName: 'Sheet1',
    startAddress: 'A1',
    endAddress: 'D4',
    columnNames: ['Kind', 'Base', 'Rate', 'Result'],
    headerRow: true,
    totalsRow: false,
  })
  return engine
}

export async function runWorkbookAgentModelExample(): Promise<WorkbookAgentModelExampleOutput> {
  const engine = await seedWorkbook()

  const planned = planWorkbookAction(model, 'calculate')
  if (planned.status !== 'planned') {
    throw new Error(planned.errors.map((error) => error.message).join('\n'))
  }

  const verification = verifyPlan(planned.plan)
  if (verification.status !== 'valid') {
    throw new Error(verification.issues.map((issue) => issue.message).join('\n'))
  }

  const run = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))
  if (run.status !== 'done') {
    throw new Error(run.errors.map((error) => error.message).join('\n'))
  }

  return {
    model: describeModel(model),
    planning: {
      status: 'planned',
      plan: describePlan(planned.plan),
      verification,
      requirements: describeRuntimeRequirements(planned.plan),
    },
    run: describeRunResult(run),
    workbook: {
      formulas: {
        D2: engine.getCell('Sheet1', 'D2').formula ?? null,
        D3: engine.getCell('Sheet1', 'D3').formula ?? null,
        D4: engine.getCell('Sheet1', 'D4').formula ?? null,
      },
      values: {
        D2: readNumber(engine.getCellValue('Sheet1', 'D2')),
        D3: readNumber(engine.getCellValue('Sheet1', 'D3')),
        D4: readNumber(engine.getCellValue('Sheet1', 'D4')),
      },
    },
  }
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  console.log(JSON.stringify(await runWorkbookAgentModelExample(), null, 2))
}
