import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  defineModel,
  describeRuntimeRequirements,
  findRange,
  formula,
  isWorkbookRuntimeCapability,
  verifyRuntimeRequirements,
  workbookRuntimeCapabilities,
  type WorkbookActionPlan,
} from '../index.js'

function summary(requirement: ReturnType<typeof describeRuntimeRequirements>['requirements'][number]) {
  return {
    kind: requirement.kind,
    capability: requirement.capability,
    path: requirement.path,
    materialization: requirement.materialization,
    commandIndex: requirement.commandIndex,
    checkIndex: requirement.checkIndex,
    opIndex: requirement.opIndex,
    opIndexes: requirement.opIndexes,
    opKind: requirement.opKind,
    checkKind: requirement.checkKind,
    target: requirement.target?.label,
    refs: requirement.refs?.map((ref) => ref.label),
    expectation: requirement.expectation?.kind,
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
        path: 'commands[0]',
        materialization: 'adapterMaterialization',
        commandIndex: 0,
        checkIndex: undefined,
        opIndex: undefined,
        opIndexes: undefined,
        opKind: undefined,
        checkKind: undefined,
        target: 'Inputs rows where Status eq "Active".Result',
        refs: ['Inputs rows where Status eq "Active".Base', 'Inputs rows where Status eq "Active".Rate'],
        expectation: undefined,
        message: 'Apply formula write to Inputs rows where Status eq "Active".Result',
      },
      {
        kind: 'read',
        capability: 'read',
        path: 'checks[1]',
        materialization: undefined,
        commandIndex: undefined,
        checkIndex: 1,
        opIndex: undefined,
        opIndexes: undefined,
        opKind: undefined,
        checkKind: 'formulaEquals',
        target: 'Inputs rows where Status eq "Active".Result',
        refs: ['Inputs rows where Status eq "Active".Base', 'Inputs rows where Status eq "Active".Rate'],
        expectation: 'formulaEquals',
        message: 'Read Inputs rows where Status eq "Active".Result for formulaEquals',
      },
      {
        kind: 'verify',
        capability: 'verifyCheck',
        path: 'checks[0]',
        materialization: undefined,
        commandIndex: undefined,
        checkIndex: 0,
        opIndex: undefined,
        opIndexes: undefined,
        opKind: undefined,
        checkKind: 'exists',
        target: 'Inputs',
        refs: undefined,
        expectation: undefined,
        message: 'Verify exists for Inputs',
      },
      {
        kind: 'verify',
        capability: 'verifyCheck',
        path: 'checks[2]',
        materialization: undefined,
        commandIndex: undefined,
        checkIndex: 2,
        opIndex: undefined,
        opIndexes: undefined,
        opKind: undefined,
        checkKind: 'noFormulaErrors',
        target: 'Inputs rows where Status eq "Active".Result',
        refs: undefined,
        expectation: undefined,
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
        path: 'ops[0]',
        materialization: 'providedOp',
        commandIndex: undefined,
        checkIndex: undefined,
        opIndex: 0,
        opIndexes: undefined,
        opKind: 'setCellValue',
        checkKind: undefined,
        target: undefined,
        refs: undefined,
        expectation: undefined,
        message: 'Apply workbook op setCellValue',
      },
    ])
  })

  it('marks known single-cell commands as concrete op backed handoff requirements', () => {
    const model = defineModel({
      name: 'single-cell-requirements',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeValue(refs.output, 12)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'write')

    expect(describeRuntimeRequirements(plan).requirements.map(summary)).toEqual([
      {
        kind: 'apply',
        capability: 'writeValue',
        path: 'commands[0]',
        materialization: 'concreteOp',
        commandIndex: 0,
        checkIndex: undefined,
        opIndex: undefined,
        opIndexes: [0],
        opKind: undefined,
        checkKind: undefined,
        target: 'Sheet1!B2',
        refs: undefined,
        expectation: undefined,
        message: 'Apply value write to Sheet1!B2',
      },
    ])
  })

  it('verifies runtime capability coverage before preview or apply', () => {
    const model = defineModel({
      name: 'runtime-capability-preflight',

      find(workbook) {
        return {
          input: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B1' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.valueEquals(refs.output, 10), workbook.check.noFormulaErrors(refs.output)]
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.multiply(refs.input, 2))
        },
      },
    })

    const requirements = describeRuntimeRequirements(buildWorkbookActionPlan(model, 'calculate'))

    expect(Object.isFrozen(workbookRuntimeCapabilities)).toBe(true)
    expect(isWorkbookRuntimeCapability('writeFormula')).toBe(true)
    expect(isWorkbookRuntimeCapability('pivotMagic')).toBe(false)
    expect(verifyRuntimeRequirements(requirements, ['writeFormula', 'read', 'verifyCheck'])).toEqual({
      status: 'supported',
      missing: [],
    })
    expect(verifyRuntimeRequirements(requirements, ['writeValue', 'read'])).toEqual({
      status: 'unsupported',
      missing: [
        {
          capability: 'writeFormula',
          path: 'commands[0]',
          message: 'Runtime is missing writeFormula for commands[0]: Apply formula write to Sheet1!B1',
          requirement: requirements.requirements[0],
        },
        {
          capability: 'verifyCheck',
          path: 'checks[1]',
          message: 'Runtime is missing verifyCheck for checks[1]: Verify noFormulaErrors for Sheet1!B1',
          requirement: requirements.requirements[2],
        },
      ],
    })
  })
})
