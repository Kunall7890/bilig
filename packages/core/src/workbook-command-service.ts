import { normalizeWorkbookCommandReceipt, type WorkbookCommandReceipt, type WorkbookCommandRequest } from '@bilig/workbook'
import { normalizeWorkbookCommandDescriptor, type WorkbookCommandDescriptor } from '@bilig/workbook/features'
import type { SpreadsheetEngine } from './engine.js'
import type { WorkbookProjectionInterceptorService } from './workbook-projection-interceptors.js'

export interface WorkbookCommandHandlerContext {
  readonly engine: SpreadsheetEngine
  readonly projection: WorkbookProjectionInterceptorService
}

export type WorkbookCommandHandler = (
  request: WorkbookCommandRequest,
  context: WorkbookCommandHandlerContext,
) => WorkbookCommandReceipt | Promise<WorkbookCommandReceipt>

interface RegisteredWorkbookCommand {
  readonly descriptor: WorkbookCommandDescriptor
  readonly handler: WorkbookCommandHandler
}

export class WorkbookCommandService {
  private readonly commands = new Map<string, RegisteredWorkbookCommand>()

  constructor(
    private readonly engine: SpreadsheetEngine,
    private readonly projection: WorkbookProjectionInterceptorService,
  ) {}

  registerCommand(descriptor: WorkbookCommandDescriptor, handler: WorkbookCommandHandler): () => void {
    const normalized = normalizeWorkbookCommandDescriptor(descriptor)
    const key = commandKey(normalized.featureId, normalized.id)
    if (this.commands.has(key)) {
      throw new Error(`Workbook command ${normalized.featureId}.${normalized.id} is already registered`)
    }
    const registered: RegisteredWorkbookCommand = Object.freeze({ descriptor: normalized, handler })
    this.commands.set(key, registered)
    return () => {
      if (this.commands.get(key) === registered) {
        this.commands.delete(key)
      }
    }
  }

  listCommandDescriptors(): readonly WorkbookCommandDescriptor[] {
    return [...this.commands.values()].map((command) => command.descriptor)
  }

  async execute(request: WorkbookCommandRequest): Promise<WorkbookCommandReceipt> {
    const command = this.commands.get(commandKey(request.featureId, request.commandId))
    if (!command) {
      throw new Error(`Workbook command ${request.featureId}.${request.commandId} is not registered`)
    }
    if (request.category !== undefined && request.category !== command.descriptor.category) {
      throw new Error(`Workbook command ${request.featureId}.${request.commandId} category ${request.category} does not match descriptor`)
    }
    const receipt = await command.handler(
      {
        ...request,
        category: request.category ?? command.descriptor.category,
      },
      {
        engine: this.engine,
        projection: this.projection,
      },
    )
    return normalizeWorkbookCommandReceipt(receipt)
  }

  preview(request: Omit<WorkbookCommandRequest, 'mode'>): Promise<WorkbookCommandReceipt> {
    return this.execute({ ...request, mode: 'preview' })
  }

  apply(request: Omit<WorkbookCommandRequest, 'mode'>): Promise<WorkbookCommandReceipt> {
    return this.execute({ ...request, mode: 'apply' })
  }

  applyAndVerify(request: Omit<WorkbookCommandRequest, 'mode'>): Promise<WorkbookCommandReceipt> {
    return this.execute({ ...request, mode: 'applyAndVerify' })
  }
}

function commandKey(featureId: string, commandId: string): string {
  return `${featureId}\u0000${commandId}`
}
