import type { ViewportPatch } from '@bilig/worker-transport'

function readNonNegativeInteger(value: unknown): number | null {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : null
}

export class ProjectedViewportPatchRevisionGate {
  private lastAuthoritativeRevision: number | null = null
  private lastBatchId = 0

  getLastAuthoritativeRevision(): number | null {
    return this.lastAuthoritativeRevision
  }

  getLastBatchId(): number {
    return this.lastBatchId
  }

  shouldApplyViewportPatch(patch: ViewportPatch): boolean {
    const authoritativeRevision = this.readAuthoritativeRevision(patch)
    if (
      authoritativeRevision !== null &&
      this.lastAuthoritativeRevision !== null &&
      authoritativeRevision < this.lastAuthoritativeRevision
    ) {
      return false
    }

    const batchId = this.readBatchId(patch)
    if (batchId === null || batchId === 0 || this.lastBatchId === 0) {
      return true
    }
    if (
      authoritativeRevision !== null &&
      this.lastAuthoritativeRevision !== null &&
      authoritativeRevision > this.lastAuthoritativeRevision
    ) {
      return true
    }
    return batchId >= this.lastBatchId
  }

  noteAppliedViewportPatch(patch: ViewportPatch): void {
    const authoritativeRevision = this.readAuthoritativeRevision(patch)
    if (authoritativeRevision !== null) {
      this.lastAuthoritativeRevision =
        this.lastAuthoritativeRevision === null ? authoritativeRevision : Math.max(this.lastAuthoritativeRevision, authoritativeRevision)
    }
    const batchId = this.readBatchId(patch)
    if (batchId !== null) {
      this.noteObservedBatchId(batchId)
    }
  }

  noteObservedBatchId(batchId: number): void {
    const normalizedBatchId = readNonNegativeInteger(batchId)
    if (normalizedBatchId !== null) {
      this.lastBatchId = Math.max(this.lastBatchId, normalizedBatchId)
    }
  }

  private readAuthoritativeRevision(patch: ViewportPatch): number | null {
    return readNonNegativeInteger(patch.authoritativeRevision)
  }

  private readBatchId(patch: ViewportPatch): number | null {
    return readNonNegativeInteger(patch.metrics?.batchId)
  }
}
