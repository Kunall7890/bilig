import {
  normalizeWorkbookActionInput,
  normalizeWorkbookCommandBundle,
  workbookCommandResultForReceipts,
  type WorkbookActionInput,
  type WorkbookCommandBundle,
  type WorkbookCommandBundleCommand,
  type WorkbookCommandReceipt,
  type WorkbookCommandResult,
  type WorkbookUndoRef,
} from '@bilig/workbook'
import type { CellRangeRef } from '@bilig/protocol'
import {
  deriveWorkbookAgentCommandPreviewRanges,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewRange,
} from './workbook-agent-bundles.js'

const WORKBOOK_AGENT_FEATURE_ID = 'workbook-agent'

function workbookAgentCommandId(command: WorkbookAgentCommand): string {
  return `workbookAgent.${command.kind}`
}

function workbookAgentBundleCommandId(bundle: WorkbookAgentCommandBundle, index: number, command: WorkbookAgentCommand): string {
  return `${bundle.id}:${String(index)}:${command.kind}`
}

function cellRangeRefForPreviewRange(range: WorkbookAgentPreviewRange): CellRangeRef {
  return {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }
}

function isPlainDataObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function childPath(path: string, key: string): string {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? `${path}.${key}` : `${path}[${JSON.stringify(key)}]`
}

function cloneCommandInput(value: unknown, path: string, seen: WeakSet<object>): unknown {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return value
  }
  if (value === undefined) {
    return undefined
  }
  if (typeof value === 'bigint' || typeof value === 'function' || typeof value === 'symbol') {
    return value
  }
  if (typeof value !== 'object') {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)

  try {
    if (Array.isArray(value)) {
      const output: unknown[] = []
      for (let index = 0; index < value.length; index += 1) {
        const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
        if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
          output.push(undefined)
          continue
        }
        output.push(cloneCommandInput(descriptor.value, `${path}[${String(index)}]`, seen))
      }
      return output
    }

    if (!isPlainDataObject(value)) {
      return value
    }

    if (Object.getOwnPropertySymbols(value).length > 0) {
      return value
    }

    const output: Record<string, unknown> = {}
    Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
      if (!descriptor.enumerable) {
        return
      }
      if (!('value' in descriptor)) {
        Object.defineProperty(output, key, {
          enumerable: true,
          configurable: true,
          get() {
            return undefined
          },
        })
        return
      }
      if (descriptor.value === undefined) {
        return
      }
      output[key] = cloneCommandInput(descriptor.value, childPath(path, key), seen)
    })
    return output
  } finally {
    seen.delete(value)
  }
}

function workbookAgentCommandInput(command: WorkbookAgentCommand): WorkbookActionInput {
  return normalizeWorkbookActionInput(cloneCommandInput(command, 'input', new WeakSet()))
}

function toWorkbookCommandBundleCommand(
  bundle: WorkbookAgentCommandBundle,
  command: WorkbookAgentCommand,
  index: number,
): WorkbookCommandBundleCommand {
  const touchedRanges = deriveWorkbookAgentCommandPreviewRanges(command).map(cellRangeRefForPreviewRange)
  return {
    id: workbookAgentBundleCommandId(bundle, index, command),
    kind: 'request',
    request: {
      featureId: WORKBOOK_AGENT_FEATURE_ID,
      commandId: workbookAgentCommandId(command),
      category: 'mutation',
      mode: 'applyAndVerify',
      input: workbookAgentCommandInput(command),
    },
    ...(touchedRanges.length > 0 ? { touchedRanges } : {}),
    destructive: true,
  }
}

export function toWorkbookCommandBundle(bundle: WorkbookAgentCommandBundle): WorkbookCommandBundle {
  return normalizeWorkbookCommandBundle({
    id: bundle.id,
    targetRevision: bundle.baseRevision,
    idempotencyKey: bundle.id,
    commands: bundle.commands.map((command, index) => toWorkbookCommandBundleCommand(bundle, command, index)),
  })
}

export interface AppliedWorkbookCommandResultInput {
  readonly bundle: WorkbookAgentCommandBundle
  readonly revision: number
  readonly undo?: WorkbookUndoRef
}

export function toAppliedWorkbookCommandResult(input: AppliedWorkbookCommandResultInput): WorkbookCommandResult {
  const commandBundle = toWorkbookCommandBundle(input.bundle)
  return workbookCommandResultForReceipts(
    commandBundle,
    commandBundle.commands.map((command) => {
      if (command.kind !== 'request') {
        throw new Error('Workbook agent command handoff produced a non-request command')
      }
      const receipt: WorkbookCommandReceipt = {
        status: 'applied' as const,
        featureId: command.request.featureId,
        commandId: command.request.commandId,
        category: command.request.category ?? 'mutation',
      }
      if (command.touchedRanges !== undefined) {
        Object.assign(receipt, { changedRanges: command.touchedRanges })
      }
      if (command.id !== undefined) {
        Object.assign(receipt, {
          proof: {
            bundleCommandId: command.id,
          },
        })
      }
      return receipt
    }),
    {
      revision: input.revision,
      ...(input.undo !== undefined ? { undo: input.undo } : {}),
    },
  )
}
