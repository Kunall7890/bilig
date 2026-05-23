import { describe, expect, it } from 'vitest'
import {
  buildWorkbookActionPlan,
  checkRuntimeRequirements,
  checkRuntimeAdapter,
  defineModel,
  describeRuntimeRequirements,
  findRange,
  formula,
  isWorkbookRuntimeCapability,
  isWorkbookRuntimeRequirementKind,
  toPlanData,
  verifyModel,
  workbookRuntimeCapabilities,
  workbookRuntimeRequirementKinds,
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

function accessorArray(get: () => unknown): unknown[] {
  const value = Array.from<unknown>({ length: 1 })
  Object.defineProperty(value, '0', {
    enumerable: true,
    get,
  })
  return value
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

  it('validates transported runtime requirements with stable path issues', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const plan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'runtime-validation-model',
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
    const requirements = JSON.parse(JSON.stringify(describeRuntimeRequirements(plan)))

    expect(checkRuntimeRequirements(requirements)).toEqual({
      status: 'valid',
      requirements,
      issues: [],
    })
    expect(workbookRuntimeRequirementKinds).toEqual(['apply', 'read', 'verify'])
    expect(workbookRuntimeCapabilities).toEqual(['writeFormula', 'writeValue', 'format', 'clear', 'applyOp', 'read', 'verifyCheck'])
    expect(isWorkbookRuntimeRequirementKind('apply')).toBe(true)
    expect(isWorkbookRuntimeRequirementKind('other')).toBe(false)
    expect(isWorkbookRuntimeCapability('read')).toBe(true)
    expect(isWorkbookRuntimeCapability('other')).toBe(false)

    const result = checkRuntimeRequirements({
      modelName: 'runtime-validation-model',
      actionName: 12,
      requirements: [
        {
          kind: 'apply',
          capability: 'read',
          commandIndex: -1,
          target: { kind: 'range' },
          message: 42,
        },
        'not-an-object',
        {
          kind: 'verify',
          capability: 'verifyCheck',
          refs: [{ kind: 'table', id: 'table-1' }],
          message: 'Verify custom check',
        },
      ],
    })

    expect(result).toEqual({
      status: 'invalid',
      issues: expect.arrayContaining([
        {
          code: 'invalid_runtime_requirements',
          path: 'actionName',
          message: 'Workbook runtime requirements actionName must be a string',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0].capability',
          message: 'Workbook runtime requirement capability read does not match kind apply',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0].message',
          message: 'Workbook runtime requirement message must be a string',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0].commandIndex',
          message: 'Workbook runtime requirement commandIndex must be a non-negative integer',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0].target',
          message: 'Workbook runtime requirement target must be workbook ref data',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[1]',
          message: 'Workbook runtime requirement at requirements[1] must be an object',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[2].refs[0]',
          message: 'Workbook runtime requirement ref must be workbook ref data',
        },
      ]),
    })
  })

  it('treats transported runtime requirement fields as own data, not inherited prototype data', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'A1' })
    const plan: WorkbookActionPlan<{ readonly target: typeof target }> = {
      modelName: 'runtime-own-data-model',
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
    const requirements = structuredClone(describeRuntimeRequirements(plan))
    const inheritedRequirements = Object.create(requirements) as unknown

    expect(checkRuntimeRequirements(inheritedRequirements)).toEqual({
      status: 'invalid',
      issues: expect.arrayContaining([
        {
          code: 'invalid_runtime_requirements',
          path: 'modelName',
          message: 'Workbook runtime requirements modelName must be a string',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'actionName',
          message: 'Workbook runtime requirements actionName must be a string',
        },
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements',
          message: 'Workbook runtime requirements requirements must be an array',
        },
      ]),
    })
  })

  it('rejects accessor-backed runtime requirement arrays without invoking getters', () => {
    let requirementGetterInvoked = false
    const requirements = accessorArray(() => {
      requirementGetterInvoked = true
      throw new Error('getter must not run')
    })

    expect(
      checkRuntimeRequirements({
        modelName: 'runtime-accessor-array-model',
        actionName: 'seed',
        requirements,
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0]',
          message: 'Workbook runtime requirement at requirements[0] must be an object',
        },
      ],
    })
    expect(requirementGetterInvoked).toBe(false)

    let refGetterInvoked = false
    const refs = accessorArray(() => {
      refGetterInvoked = true
      throw new Error('getter must not run')
    })

    expect(
      checkRuntimeRequirements({
        modelName: 'runtime-accessor-ref-array-model',
        actionName: 'seed',
        requirements: [
          {
            kind: 'verify',
            capability: 'verifyCheck',
            checkIndex: 0,
            checkKind: 'custom',
            refs,
            message: 'Verify custom check',
          },
        ],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_runtime_requirements',
          path: 'requirements[0].refs[0]',
          message: 'Workbook runtime requirement ref must be workbook ref data',
        },
      ],
    })
    expect(refGetterInvoked).toBe(false)
  })

  it('checks runtime adapter capabilities before mutation handoff', () => {
    const model = defineModel({
      name: 'adapter-capability-model',

      find(workbook) {
        return {
          input: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.output)]
      },

      actions: {
        write({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.add(refs.input, 1))
          workbook.check.formulaEquals(refs.output, formula.add(refs.input, 1))
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'write')
    const transportedPlan = JSON.parse(JSON.stringify(toPlanData(plan)))

    expect(checkRuntimeAdapter(transportedPlan, {})).toEqual({
      status: 'invalid',
      modelName: 'adapter-capability-model',
      actionName: 'write',
      requiredCapabilities: ['writeFormula', 'read', 'verifyCheck'],
      issues: [
        {
          code: 'missing_apply',
          capability: 'writeFormula',
          method: 'apply',
          requirementIndexes: [0],
          message: 'Adapter is missing apply for writeFormula',
        },
        {
          code: 'missing_read',
          capability: 'read',
          method: 'read',
          requirementIndexes: [1],
          message: 'Adapter is missing read for read',
        },
        {
          code: 'missing_check_verifier',
          capability: 'verifyCheck',
          method: 'verifyChecks',
          requirementIndexes: [2],
          message: 'Adapter is missing verifyChecks for verifyCheck',
        },
      ],
    })

    expect(
      checkRuntimeAdapter(describeRuntimeRequirements(plan), {
        apply() {
          return { status: 'applied' as const }
        },
        read() {
          return []
        },
        verifyChecks(checks: unknown) {
          return checks
        },
      }),
    ).toEqual({
      status: 'valid',
      modelName: 'adapter-capability-model',
      actionName: 'write',
      requiredCapabilities: ['writeFormula', 'read', 'verifyCheck'],
      issues: [],
    })
  })

  it('does not require apply for check-only runtime handoff', () => {
    const model = defineModel({
      name: 'check-only-adapter-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        inspect({ refs, workbook }) {
          workbook.check.valueEquals(refs.output, 12)
        },
      },
    })

    const plan = buildWorkbookActionPlan(model, 'inspect')

    expect(describeRuntimeRequirements(plan).requirements.map(summary)).toEqual([
      {
        kind: 'read',
        capability: 'read',
        commandIndex: undefined,
        checkIndex: 0,
        opIndex: undefined,
        opKind: undefined,
        checkKind: 'valueEquals',
        target: 'Sheet1!B2',
        refs: undefined,
        message: 'Read Sheet1!B2 for valueEquals',
      },
    ])
    expect(checkRuntimeAdapter(plan, { read: () => [] })).toEqual({
      status: 'valid',
      modelName: 'check-only-adapter-model',
      actionName: 'inspect',
      requiredCapabilities: ['read'],
      issues: [],
    })
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
