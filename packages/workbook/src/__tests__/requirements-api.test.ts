import { describe, expect, it } from 'vitest'
import { buildWorkbookActionPlan, defineModel, describeRuntimeRequirements, findRange, formula, type WorkbookActionPlan } from '../index.js'

function summary(requirement: ReturnType<typeof describeRuntimeRequirements>['requirements'][number]) {
  return {
    kind: requirement.kind,
    capability: requirement.capability,
    commandIndex: requirement.commandIndex,
    checkIndex: requirement.checkIndex,
    opIndex: requirement.opIndex,
    opKind: requirement.opKind,
    checkKind: requirement.checkKind,
    target: requirement.target?.label,
    refs: requirement.refs?.map((ref) => ref.label),
    message: requirement.message,
  }
}

describe('@bilig/workbook runtime requirements api', () => {
  it('describes apply, read, and check proof requirements as JSON-safe handoff data', () => {
    const model = defineModel({
      name: 'requirements-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs', headers: ['Base', 'Rate', 'Result'] })
        const activeRows = workbook.findRows({
          table,
          where: {
            column: 'Status',
            op: 'eq',
            value: 'Active',
          },
        })

        return {
          table,
          base: activeRows.column('Base'),
          rate: activeRows.column('Rate'),
          result: activeRows.column('Result'),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },

      actions: {
        calculate({ refs, workbook }) {
          const expected = formula.multiply(refs.base, refs.rate)
          workbook.writeFormula(refs.result, expected)
          workbook.check.formulaEquals(refs.result, expected)
          workbook.check.noFormulaErrors(refs.result)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'calculate')
    const requirements = describeRuntimeRequirements(plan)

    expect(requirements.modelName).toBe('requirements-model')
    expect(requirements.actionName).toBe('calculate')
    expect(requirements.requirements.map(summary)).toEqual([
      {
        kind: 'apply',
        capability: 'writeFormula',
        commandIndex: 0,
        checkIndex: undefined,
        opIndex: undefined,
        opKind: undefined,
        checkKind: undefined,
        target: 'table_Inputs_Base_Rate_Result rows where Status eq "Active".Result',
        refs: [
          'table_Inputs_Base_Rate_Result rows where Status eq "Active".Base',
          'table_Inputs_Base_Rate_Result rows where Status eq "Active".Rate',
        ],
        message: 'Apply formula write to table_Inputs_Base_Rate_Result rows where Status eq "Active".Result',
      },
      {
        kind: 'read',
        capability: 'read',
        commandIndex: undefined,
        checkIndex: 1,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'formulaEquals',
        target: 'table_Inputs_Base_Rate_Result rows where Status eq "Active".Result',
        refs: [
          'table_Inputs_Base_Rate_Result rows where Status eq "Active".Base',
          'table_Inputs_Base_Rate_Result rows where Status eq "Active".Rate',
        ],
        message: 'Read table_Inputs_Base_Rate_Result rows where Status eq "Active".Result for formulaEquals',
      },
      {
        kind: 'verify',
        capability: 'verifyCheck',
        commandIndex: undefined,
        checkIndex: 0,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'exists',
        target: 'Inputs',
        refs: undefined,
        message: 'Verify exists for Inputs',
      },
      {
        kind: 'verify',
        capability: 'verifyCheck',
        commandIndex: undefined,
        checkIndex: 2,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'noFormulaErrors',
        target: 'table_Inputs_Base_Rate_Result rows where Status eq "Active".Result',
        refs: undefined,
        message: 'Verify noFormulaErrors for table_Inputs_Base_Rate_Result rows where Status eq "Active".Result',
      },
    ])
    expect(JSON.parse(JSON.stringify(requirements))).toEqual(requirements)
  })

  it('describes explicit ops on manually assembled plans', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const plan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'manual-plan',
      actionName: 'seed',
      refs: { target },
      refsUsed: [target],
      commands: [],
      ops: [
        {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 1,
        },
      ],
      changed: [],
      checks: [],
    }

    expect(describeRuntimeRequirements(plan).requirements.map(summary)).toEqual([
      {
        kind: 'apply',
        capability: 'applyOp',
        commandIndex: undefined,
        checkIndex: undefined,
        opIndex: 0,
        opKind: 'setCellValue',
        checkKind: undefined,
        target: undefined,
        refs: undefined,
        message: 'Apply workbook op setCellValue',
      },
    ])
  })
})
