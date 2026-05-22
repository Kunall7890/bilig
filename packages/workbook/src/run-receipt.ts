import { isWorkbookRevision, type WorkbookCommandBundle, type WorkbookRevision } from './command.js'
import { isWorkbookRef } from './find.js'
import { normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { WorkbookActionPlan } from './model.js'
import { isWorkbookReceiptProofKind, type WorkbookReceiptProof, type WorkbookRunReceipt, type WorkbookRuntimeReceipt } from './receipt.js'
import type { WorkbookRuntimePreview } from './requirements.js'
import type { WorkbookChangeSummary, WorkbookCheckResult, WorkbookRunError, WorkbookUndoRef } from './result.js'

interface ReceiptApplyResult {
  readonly status: 'applied' | 'failed'
  readonly receipt?: WorkbookRuntimeReceipt
  readonly undo?: WorkbookUndoRef
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function runtimeResultError(message: string, path = 'runtime'): WorkbookRunError {
  return {
    code: 'invalid_runtime_result',
    message,
    path,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isReceiptProofStatus(value: unknown): value is WorkbookReceiptProof['status'] {
  return value === 'passed' || value === 'failed' || value === 'skipped'
}

function validateStringArray(value: unknown, path: string): readonly string[] | WorkbookRunError {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    return runtimeResultError(`${path} must be a string array`, path)
  }
  return [...value]
}

function validateChangeSummary(value: unknown, path: string): WorkbookChangeSummary | WorkbookRunError {
  if (!isRecord(value)) {
    return runtimeResultError('Rendered receipt diff must be an object', path)
  }
  if (typeof value['kind'] !== 'string' || value['kind'].trim() === '') {
    return runtimeResultError('Rendered receipt diff kind must be a non-empty string', `${path}.kind`)
  }
  if (typeof value['message'] !== 'string' || value['message'].trim() === '') {
    return runtimeResultError('Rendered receipt diff message must be a non-empty string', `${path}.message`)
  }
  if (value['target'] !== undefined && !isWorkbookRef(value['target'])) {
    return runtimeResultError('Rendered receipt diff target must be a WorkbookRef', `${path}.target`)
  }
  return {
    kind: value['kind'],
    ...(value['target'] !== undefined ? { target: value['target'] } : {}),
    message: value['message'],
  }
}

function validateReceiptProof(value: unknown, path: string): WorkbookReceiptProof | WorkbookRunError {
  if (!isRecord(value)) {
    return runtimeResultError('Runtime receipt proof must be an object', path)
  }
  if (!isWorkbookReceiptProofKind(value['kind'])) {
    return runtimeResultError('Runtime receipt proof kind must be a known proof kind', `${path}.kind`)
  }
  if (!isReceiptProofStatus(value['status'])) {
    return runtimeResultError('Runtime receipt proof status must be passed, failed, or skipped', `${path}.status`)
  }
  if (typeof value['message'] !== 'string' || value['message'].trim() === '') {
    return runtimeResultError('Runtime receipt proof message must be a non-empty string', `${path}.message`)
  }
  if (value['revision'] !== undefined && !isWorkbookRevision(value['revision'])) {
    return runtimeResultError('Runtime receipt proof revision must be a non-empty string or finite number', `${path}.revision`)
  }
  if (value['target'] !== undefined && !isWorkbookRef(value['target'])) {
    return runtimeResultError('Runtime receipt proof target must be a WorkbookRef', `${path}.target`)
  }
  let data: WorkbookActionInput | undefined
  if (value['data'] !== undefined) {
    try {
      data = normalizeWorkbookActionInput(value['data'])
    } catch (error) {
      return runtimeResultError(errorMessage(error), `${path}.data`)
    }
  }
  return {
    kind: value['kind'],
    status: value['status'],
    message: value['message'].trim(),
    ...(value['revision'] !== undefined
      ? { revision: typeof value['revision'] === 'string' ? value['revision'].trim() : value['revision'] }
      : {}),
    ...(value['target'] !== undefined ? { target: value['target'] } : {}),
    ...(data !== undefined ? { data } : {}),
  }
}

export function validateRuntimeReceipt(value: unknown, path: string): WorkbookRuntimeReceipt | WorkbookRunError {
  if (!isRecord(value)) {
    return runtimeResultError('Runtime receipt must be an object', path)
  }
  let appliedRevision: WorkbookRuntimeReceipt['appliedRevision']
  let calculatedRevision: WorkbookRuntimeReceipt['calculatedRevision']
  let renderedRevision: WorkbookRuntimeReceipt['renderedRevision']
  let renderedOutput: WorkbookRuntimeReceipt['rendered']
  let proofOutput: WorkbookRuntimeReceipt['proof']
  let warningsOutput: WorkbookRuntimeReceipt['warnings']

  for (const key of ['appliedRevision', 'calculatedRevision', 'renderedRevision'] as const) {
    const revision = value[key]
    if (revision !== undefined) {
      if (!isWorkbookRevision(revision)) {
        return runtimeResultError(`Runtime receipt ${key} must be a non-empty string or finite number`, `${path}.${key}`)
      }
      const normalizedRevision = typeof revision === 'string' ? revision.trim() : revision
      if (key === 'appliedRevision') {
        appliedRevision = normalizedRevision
      } else if (key === 'calculatedRevision') {
        calculatedRevision = normalizedRevision
      } else {
        renderedRevision = normalizedRevision
      }
    }
  }

  const rendered = value['rendered']
  if (rendered !== undefined) {
    if (!isRecord(rendered)) {
      return runtimeResultError('Runtime receipt rendered proof must be an object', `${path}.rendered`)
    }
    let renderedProofRevision: WorkbookRevision | undefined
    let renderedProofMessage: string | undefined
    let renderedProofDiffs: WorkbookChangeSummary[] | undefined
    if (rendered['revision'] !== undefined) {
      if (!isWorkbookRevision(rendered['revision'])) {
        return runtimeResultError(
          'Runtime receipt rendered revision must be a non-empty string or finite number',
          `${path}.rendered.revision`,
        )
      }
      renderedProofRevision = typeof rendered['revision'] === 'string' ? rendered['revision'].trim() : rendered['revision']
    }
    if (rendered['message'] !== undefined) {
      if (typeof rendered['message'] !== 'string' || rendered['message'].trim() === '') {
        return runtimeResultError('Runtime receipt rendered message must be a non-empty string', `${path}.rendered.message`)
      }
      renderedProofMessage = rendered['message'].trim()
    }
    if (rendered['diffs'] !== undefined) {
      if (!Array.isArray(rendered['diffs'])) {
        return runtimeResultError('Runtime receipt rendered diffs must be an array', `${path}.rendered.diffs`)
      }
      const diffs = []
      for (let index = 0; index < rendered['diffs'].length; index += 1) {
        const diff = validateChangeSummary(rendered['diffs'][index], `${path}.rendered.diffs[${index.toString()}]`)
        if ('code' in diff) {
          return diff
        }
        diffs.push(diff)
      }
      renderedProofDiffs = diffs
    }
    renderedOutput = {
      ...(renderedProofRevision !== undefined ? { revision: renderedProofRevision } : {}),
      ...(renderedProofDiffs !== undefined ? { diffs: renderedProofDiffs } : {}),
      ...(renderedProofMessage !== undefined ? { message: renderedProofMessage } : {}),
    }
  }

  const proof = value['proof']
  if (proof !== undefined) {
    if (!Array.isArray(proof)) {
      return runtimeResultError('Runtime receipt proof must be an array', `${path}.proof`)
    }
    const normalizedProof: WorkbookReceiptProof[] = []
    for (let index = 0; index < proof.length; index += 1) {
      const entry = validateReceiptProof(proof[index], `${path}.proof[${index.toString()}]`)
      if ('code' in entry) {
        return entry
      }
      normalizedProof.push(entry)
    }
    proofOutput = normalizedProof
  }

  const warnings = value['warnings']
  if (warnings !== undefined) {
    const normalizedWarnings = validateStringArray(warnings, `${path}.warnings`)
    if ('code' in normalizedWarnings) {
      return normalizedWarnings
    }
    warningsOutput = normalizedWarnings
  }

  return {
    ...(appliedRevision !== undefined ? { appliedRevision } : {}),
    ...(calculatedRevision !== undefined ? { calculatedRevision } : {}),
    ...(renderedRevision !== undefined ? { renderedRevision } : {}),
    ...(renderedOutput !== undefined ? { rendered: renderedOutput } : {}),
    ...(proofOutput !== undefined ? { proof: proofOutput } : {}),
    ...(warningsOutput !== undefined ? { warnings: warningsOutput } : {}),
  }
}

function checkCounts(checks: readonly WorkbookCheckResult[]): {
  readonly passed: number
  readonly failed: number
  readonly unverified: number
} {
  return {
    passed: checks.filter((check) => check.status === 'passed').length,
    failed: checks.filter((check) => check.status === 'failed').length,
    unverified: checks.filter((check) => check.status === 'planned').length,
  }
}

function buildRunReceipt<Refs>(input: {
  readonly plan: WorkbookActionPlan<Refs>
  readonly command: WorkbookCommandBundle<Refs> | undefined
  readonly preview: WorkbookRuntimePreview | undefined
  readonly applyResult: ReceiptApplyResult
  readonly checks: readonly WorkbookCheckResult[]
  readonly errors: readonly WorkbookRunError[]
  readonly readbackTargetCount: number
}): WorkbookRunReceipt {
  const counts = checkCounts(input.checks)
  const runtimeReceipt = input.applyResult.receipt
  const proof: WorkbookReceiptProof[] = [
    {
      kind: 'preview',
      status: input.preview === undefined ? 'skipped' : 'passed',
      message:
        input.preview === undefined
          ? 'No preview was requested.'
          : `Preview materialized ${input.preview.materializedOps.length.toString()} ops.`,
    },
    {
      kind: 'apply',
      status: input.applyResult.status === 'applied' ? 'passed' : 'failed',
      message: input.applyResult.status === 'applied' ? 'Runtime applied the workbook action.' : 'Runtime rejected the workbook action.',
      ...(runtimeReceipt?.appliedRevision !== undefined ? { revision: runtimeReceipt.appliedRevision } : {}),
    },
    {
      kind: 'authoritativeReadback',
      status:
        input.readbackTargetCount === 0
          ? 'skipped'
          : input.errors.some(
                (error) => error.code.endsWith('_mismatch') || error.code === 'readback_missing' || error.code === 'duplicate_readback',
              )
            ? 'failed'
            : 'passed',
      message:
        input.readbackTargetCount === 0
          ? 'No readback-backed checks were planned.'
          : `Runtime read back ${input.readbackTargetCount.toString()} proof targets.`,
    },
    {
      kind: 'check',
      status: counts.failed > 0 || counts.unverified > 0 ? 'failed' : 'passed',
      message: `${counts.passed.toString()} checks passed, ${counts.failed.toString()} failed, ${counts.unverified.toString()} unverified.`,
    },
    ...(runtimeReceipt?.proof ?? []),
  ]

  return {
    ...(input.command !== undefined ? { commandId: input.command.commandId } : {}),
    ...(input.command?.idempotencyKey !== undefined ? { idempotencyKey: input.command.idempotencyKey } : {}),
    modelName: input.plan.modelName,
    actionName: input.plan.actionName,
    ...(input.command?.baseRevision !== undefined ? { baseRevision: input.command.baseRevision } : {}),
    ...(runtimeReceipt?.appliedRevision !== undefined ? { appliedRevision: runtimeReceipt.appliedRevision } : {}),
    ...(runtimeReceipt?.calculatedRevision !== undefined ? { calculatedRevision: runtimeReceipt.calculatedRevision } : {}),
    ...(runtimeReceipt?.renderedRevision !== undefined ? { renderedRevision: runtimeReceipt.renderedRevision } : {}),
    ...(runtimeReceipt?.rendered !== undefined ? { rendered: runtimeReceipt.rendered } : {}),
    previewed: input.preview !== undefined,
    applied: input.applyResult.status === 'applied',
    verified: input.applyResult.status === 'applied' && input.errors.length === 0 && counts.failed === 0 && counts.unverified === 0,
    checkCount: input.checks.length,
    passedCheckCount: counts.passed,
    failedCheckCount: counts.failed,
    unverifiedCheckCount: counts.unverified,
    proof,
    ...(runtimeReceipt?.warnings !== undefined ? { warnings: runtimeReceipt.warnings } : {}),
    ...(input.applyResult.undo !== undefined ? { undo: input.applyResult.undo } : {}),
  }
}

export function maybeBuildRunReceipt<Refs>(input: {
  readonly plan: WorkbookActionPlan<Refs>
  readonly command: WorkbookCommandBundle<Refs> | undefined
  readonly preview: WorkbookRuntimePreview | undefined
  readonly applyResult: ReceiptApplyResult
  readonly checks: readonly WorkbookCheckResult[]
  readonly errors: readonly WorkbookRunError[]
  readonly readbackTargetCount: number
}): WorkbookRunReceipt | undefined {
  if (input.command === undefined && input.applyResult.receipt === undefined) {
    return undefined
  }
  return buildRunReceipt(input)
}
