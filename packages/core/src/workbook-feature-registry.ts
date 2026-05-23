import {
  defineWorkbookFeaturePlugin,
  type WorkbookCommandDescriptor,
  type WorkbookFeatureId,
  type WorkbookFeatureLifecycleContext,
  type WorkbookFeaturePlugin,
  type WorkbookProjectionInterceptorRegistration,
  type WorkbookUiContribution,
} from '@bilig/workbook'

export class WorkbookFeatureRegistry {
  private readonly plugins = new Map<WorkbookFeatureId, WorkbookFeaturePlugin>()
  private readonly activeFeatureIds: WorkbookFeatureId[] = []
  private disposed = false

  register(plugin: WorkbookFeaturePlugin): WorkbookFeaturePlugin {
    this.assertNotDisposed()
    const normalized = defineWorkbookFeaturePlugin(plugin)
    if (this.plugins.has(normalized.id)) {
      throw new Error(`Workbook feature ${normalized.id} is already registered`)
    }
    this.plugins.set(normalized.id, normalized)
    normalized.register?.(this.lifecycleContext(normalized.id))
    return normalized
  }

  activateAll(): readonly WorkbookFeatureId[] {
    this.assertNotDisposed()
    const visiting = new Set<WorkbookFeatureId>()
    const visited = new Set(this.activeFeatureIds)
    for (const featureId of this.plugins.keys()) {
      this.activateFeature(featureId, visiting, visited)
    }
    return this.listActiveFeatureIds()
  }

  disposeAll(): void {
    if (this.disposed) {
      return
    }
    for (const featureId of this.activeFeatureIds.toReversed()) {
      this.plugins.get(featureId)?.dispose?.(this.lifecycleContext(featureId))
    }
    this.activeFeatureIds.length = 0
    this.plugins.clear()
    this.disposed = true
  }

  get(featureId: WorkbookFeatureId): WorkbookFeaturePlugin | undefined {
    return this.plugins.get(featureId)
  }

  listPlugins(): readonly WorkbookFeaturePlugin[] {
    return [...this.plugins.values()]
  }

  listActiveFeatureIds(): readonly WorkbookFeatureId[] {
    return Object.freeze([...this.activeFeatureIds])
  }

  listCommandDescriptors(): readonly WorkbookCommandDescriptor[] {
    const descriptors: WorkbookCommandDescriptor[] = []
    this.listPlugins().forEach((plugin) => {
      descriptors.push(...plugin.commands)
    })
    return descriptors
  }

  listProjectionInterceptors(): readonly WorkbookProjectionInterceptorRegistration[] {
    const interceptors: WorkbookProjectionInterceptorRegistration[] = []
    this.listPlugins().forEach((plugin) => {
      interceptors.push(...plugin.projectionInterceptors)
    })
    return interceptors
  }

  listUiContributions(): readonly WorkbookUiContribution[] {
    const contributions: WorkbookUiContribution[] = []
    this.listPlugins().forEach((plugin) => {
      contributions.push(...plugin.uiContributions)
    })
    return contributions.toSorted((left, right) => {
      const orderDelta = (left.order ?? 0) - (right.order ?? 0)
      return orderDelta === 0 ? left.id.localeCompare(right.id) : orderDelta
    })
  }

  private activateFeature(featureId: WorkbookFeatureId, visiting: Set<WorkbookFeatureId>, visited: Set<WorkbookFeatureId>): void {
    if (visited.has(featureId)) {
      return
    }
    const plugin = this.plugins.get(featureId)
    if (!plugin) {
      throw new Error(`Workbook feature ${featureId} is not registered`)
    }
    if (visiting.has(featureId)) {
      throw new Error(`Workbook feature dependency cycle includes ${featureId}`)
    }
    visiting.add(featureId)
    for (const dependency of plugin.dependsOn ?? []) {
      if (!this.plugins.has(dependency)) {
        throw new Error(`Workbook feature ${plugin.id} depends on missing feature ${dependency}`)
      }
      this.activateFeature(dependency, visiting, visited)
    }
    visiting.delete(featureId)
    plugin.activate?.(this.lifecycleContext(featureId))
    this.activeFeatureIds.push(featureId)
    visited.add(featureId)
  }

  private lifecycleContext(featureId: WorkbookFeatureId): WorkbookFeatureLifecycleContext {
    return {
      featureId,
      activeFeatures: this.listActiveFeatureIds(),
    }
  }

  private assertNotDisposed(): void {
    if (this.disposed) {
      throw new Error('Workbook feature registry has been disposed')
    }
  }
}
