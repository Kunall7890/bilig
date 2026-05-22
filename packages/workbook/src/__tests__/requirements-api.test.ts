import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describeRuntimeRequirements,
  findRange,
  formula,
  verifyModel,
  type WorkbookActionPlan,
} from '../index.js'

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
        target: 'Inputs rows where Status eq "Active".Result',
        refs: ['Inputs rows where Status eq "Active".Base', 'Inputs rows where Status eq "Active".Rate'],
        message: 'Apply formula write to Inputs rows where Status eq "Active".Result',
      },
      {
        kind: 'read',
        capability: 'read',
        commandIndex: undefined,
        checkIndex: 1,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'formulaEquals',
        target: 'Inputs rows where Status eq "Active".Result',
        refs: ['Inputs rows where Status eq "Active".Base', 'Inputs rows where Status eq "Active".Rate'],
        message: 'Read Inputs rows where Status eq "Active".Result for formulaEquals',
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
        target: 'Inputs rows where Status eq "Active".Result',
        refs: undefined,
        message: 'Verify noFormulaErrors for Inputs rows where Status eq "Active".Result',
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

  it('includes runtime requirements in whole-model verification for each planned action', () => {
    const model = defineModel({
      name: 'verification-requirements-model',

      find(workbook) {
        return {
          input: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.add(refs.input, 1))
          workbook.check.formulaEquals(refs.output, formula.add(refs.input, 1))
        },
      },
    })

    const verification = verifyModel(model)

    expect(verification.actions[0]?.requirements?.requirements.map(summary)).toEqual([
      {
        kind: 'apply',
        capability: 'writeFormula',
        commandIndex: 0,
        checkIndex: undefined,
        opIndex: undefined,
        opKind: undefined,
        checkKind: undefined,
        target: 'Sheet1!B2',
        refs: ['Sheet1!A2'],
        message: 'Apply formula write to Sheet1!B2',
      },
      {
        kind: 'read',
        capability: 'read',
        commandIndex: undefined,
        checkIndex: 0,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'formulaEquals',
        target: 'Sheet1!B2',
        refs: ['Sheet1!A2'],
        message: 'Read Sheet1!B2 for formulaEquals',
      },
    ])
    expect(JSON.parse(JSON.stringify(verification))).toEqual(verification)
  })
})
