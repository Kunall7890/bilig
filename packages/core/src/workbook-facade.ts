import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookCommandReceipt, WorkbookCommandRequest } from '@bilig/workbook'
import type { SpreadsheetEngine } from './engine.js'
import { WorkbookCommandService } from './workbook-command-service.js'
import { WorkbookFeatureRegistry } from './workbook-feature-registry.js'
import { WorkbookProjectionInterceptorService } from './workbook-projection-interceptors.js'
import {
  createWorkbookTablesFeaturePlugin,
  registerWorkbookTablesFeature,
  WORKBOOK_TABLE_COMMAND_IDS,
  WORKBOOK_TABLES_FEATURE_ID,
} from './workbook-tables-feature.js'

export class WorkbookFacade {
  readonly features = new WorkbookFeatureRegistry()
  readonly projectionService: WorkbookProjectionInterceptorService
  readonly commands: WorkbookCommandService
  private readonly disposers: Array<() => void> = []

  constructor(readonly engine: SpreadsheetEngine) {
    this.projectionService = new WorkbookProjectionInterceptorService(engine)
    this.commands = new WorkbookCommandService(engine, this.projectionService)
    this.features.register(createWorkbookTablesFeaturePlugin())
    this.features.activateAll()
    this.disposers.push(registerWorkbookTablesFeature({ commandService: this.commands, projectionService: this.projectionService }))
  }

  table(name: string): WorkbookTableFacade {
    return new WorkbookTableFacade(this.engine, name)
  }

  selection(range: CellRangeRef): WorkbookSelectionFacade {
    return new WorkbookSelectionFacade(this.commands, range)
  }

  command(request: WorkbookCommandRequest): WorkbookCommandFacade {
    return new WorkbookCommandFacade(this.commands, request)
  }

  projection(): WorkbookProjectionFacade {
    return new WorkbookProjectionFacade(this.projectionService)
  }

  dispose(): void {
    this.disposers.toReversed().forEach((dispose) => {
      dispose()
    })
    this.features.disposeAll()
  }
}

export class WorkbookTableFacade {
  constructor(
    private readonly engine: SpreadsheetEngine,
    private readonly name: string,
  ) {}

  snapshot() {
    return this.engine.getTable(this.name)
  }

  column(columnName: string) {
    const table = this.snapshot()
    if (!table) {
      throw new Error(`Table ${this.name} does not exist`)
    }
    const columnIndex = table.columnNames.findIndex((name) => name.trim().toUpperCase() === columnName.trim().toUpperCase())
    if (columnIndex < 0) {
      throw new Error(`Table ${this.name} column ${columnName} does not exist`)
    }
    return {
      tableName: table.name,
      columnName: table.columnNames[columnIndex]!,
      columnIndex,
    }
  }
}

export class WorkbookSelectionFacade {
  constructor(
    private readonly commands: WorkbookCommandService,
    private readonly range: CellRangeRef,
  ) {}

  createTable(options: { readonly name?: string; readonly hasHeaders?: boolean } = {}): WorkbookCommandFacade {
    return new WorkbookCommandFacade(this.commands, {
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      commandId: WORKBOOK_TABLE_COMMAND_IDS.createFromSelection,
      category: 'command',
      input: {
        range: {
          sheetName: this.range.sheetName,
          startAddress: this.range.startAddress,
          endAddress: this.range.endAddress,
        },
        ...(options.name !== undefined ? { name: options.name } : {}),
        ...(options.hasHeaders !== undefined ? { hasHeaders: options.hasHeaders } : {}),
      },
    })
  }
}

export class WorkbookCommandFacade {
  constructor(
    private readonly commands: WorkbookCommandService,
    private readonly request: WorkbookCommandRequest,
  ) {}

  preview(): Promise<WorkbookCommandReceipt> {
    return this.commands.preview(this.request)
  }

  apply(): Promise<WorkbookCommandReceipt> {
    return this.commands.apply(this.request)
  }

  applyAndVerify(): Promise<WorkbookCommandReceipt> {
    return this.commands.applyAndVerify(this.request)
  }
}

export class WorkbookProjectionFacade {
  constructor(private readonly projection: WorkbookProjectionInterceptorService) {}

  rangeChrome(range: CellRangeRef) {
    return this.projection.rangeChrome(range)
  }
}

export function createWorkbookFacade(engine: SpreadsheetEngine): WorkbookFacade {
  return new WorkbookFacade(engine)
}
