import type { CellRangeRef } from '@bilig/protocol'
import type {
  WorkbookCommandMetadataProjection,
  WorkbookCommandRequest,
  WorkbookFeatureId,
  WorkbookProjectionInterceptorPoint,
  WorkbookProjectionInterceptorRegistration,
  WorkbookRangeChromeProjection,
} from '@bilig/workbook/features'
import type { SpreadsheetEngine } from './engine.js'

export interface WorkbookProjectionHandlerContext {
  readonly engine: SpreadsheetEngine
  readonly featureId: WorkbookFeatureId
}

export interface WorkbookRangeChromeProjectionInput {
  readonly range: CellRangeRef
}

export interface WorkbookCommandMetadataProjectionInput {
  readonly request: WorkbookCommandRequest
}

export interface WorkbookProjectionInterceptor extends WorkbookProjectionInterceptorRegistration {
  readonly projectRangeChrome?: (
    input: WorkbookRangeChromeProjectionInput,
    context: WorkbookProjectionHandlerContext,
  ) => readonly WorkbookRangeChromeProjection[]
  readonly projectCommandMetadata?: (
    input: WorkbookCommandMetadataProjectionInput,
    context: WorkbookProjectionHandlerContext,
  ) => WorkbookCommandMetadataProjection | undefined
}

interface RegisteredProjectionInterceptor extends WorkbookProjectionInterceptor {
  readonly registrationOrder: number
}

export class WorkbookProjectionInterceptorService {
  private readonly interceptors: RegisteredProjectionInterceptor[] = []
  private registrationCounter = 0

  constructor(private readonly engine: SpreadsheetEngine) {}

  register(interceptor: WorkbookProjectionInterceptor): () => void {
    const registered: RegisteredProjectionInterceptor = Object.freeze({
      ...interceptor,
      priority: interceptor.priority ?? 0,
      registrationOrder: this.registrationCounter,
    })
    this.registrationCounter += 1
    this.interceptors.push(registered)
    this.sortInterceptors()
    return () => {
      const index = this.interceptors.indexOf(registered)
      if (index >= 0) {
        this.interceptors.splice(index, 1)
      }
    }
  }

  list(point?: WorkbookProjectionInterceptorPoint): readonly WorkbookProjectionInterceptorRegistration[] {
    return this.interceptors
      .filter((interceptor) => point === undefined || interceptor.point === point)
      .map(
        ({
          registrationOrder: _registrationOrder,
          projectRangeChrome: _projectRangeChrome,
          projectCommandMetadata: _projectCommandMetadata,
          ...registration
        }) => Object.freeze({ ...registration }),
      )
  }

  rangeChrome(range: CellRangeRef): readonly WorkbookRangeChromeProjection[] {
    return this.interceptors.flatMap((interceptor) => {
      if (interceptor.point !== 'rangeChrome' || !interceptor.projectRangeChrome) {
        return []
      }
      return interceptor.projectRangeChrome({ range }, this.context(interceptor.featureId))
    })
  }

  commandMetadata(request: WorkbookCommandRequest): readonly WorkbookCommandMetadataProjection[] {
    return this.interceptors.flatMap((interceptor) => {
      if (interceptor.point !== 'commandMetadata' || !interceptor.projectCommandMetadata) {
        return []
      }
      const projection = interceptor.projectCommandMetadata({ request }, this.context(interceptor.featureId))
      return projection ? [projection] : []
    })
  }

  private context(featureId: WorkbookFeatureId): WorkbookProjectionHandlerContext {
    return {
      engine: this.engine,
      featureId,
    }
  }

  private sortInterceptors(): void {
    this.interceptors.sort((left, right) => {
      const priorityDelta = (right.priority ?? 0) - (left.priority ?? 0)
      return priorityDelta === 0 ? left.registrationOrder - right.registrationOrder : priorityDelta
    })
  }
}
